// NewFolderWithSelection.cs — 选中文件/文件夹 → 右键 → 新建文件夹并移入
//
// 设计要点
//   * 编译成 winexe（无控制台、无 wscript、无 powershell），冷启动 ~150ms
//   * 界面全部自绘：Win11 风格圆角卡片 + 圆角输入框 + 强调色主按钮，跟随系统浅/深色
//   * Explorer 对每个选中项各调用一次本程序（最多 15 次）；首个进程抢 owner 锁后立刻弹窗，
//     窗口开着的同时每 40ms 轮询队列把后续项并进同一批 → 只建一个文件夹
//   * 队列攒到 15 项基本说明被 Explorer 截断了 → 后台 STA 线程读 Shell COM 真实选中项补齐
//   * 目录不可写 / 文件不可删 → runas 提权重试（-JobFile 子进程模式）
//
// 调试环境变量
//   NFS_NAME=名字    跳过输入框
//   NFS_NOMOVE=1     只输出建议名
//   NFS_QUIET=1      不弹任何窗口，结果写 stdout
//   NFS_NOELEVATE=1  不弹 UAC

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace NewFolderWithSelection
{
    // =================================================================
    // 小工具
    // =================================================================
    static class Util
    {
        public static bool EnvFlag(string name)
        {
            string v = Environment.GetEnvironmentVariable(name);
            return !string.IsNullOrEmpty(v) && v != "0";
        }

        public static string Env(string name)
        {
            string v = Environment.GetEnvironmentVariable(name);
            if (v == null) return null;
            v = v.Trim();
            return v.Length == 0 ? null : v;
        }

        public static bool Contains(List<string> list, string value)
        {
            for (int i = 0; i < list.Count; i++)
                if (string.Equals(list[i], value, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        // 选中项里的每一项都在 sel 里
        public static bool Covers(List<string> sel, List<string> items)
        {
            for (int i = 0; i < items.Count; i++)
                if (!Contains(sel, items[i])) return false;
            return true;
        }

        public static string Norm(string p)
        {
            if (p == null) return null;
            p = p.TrimEnd('\\', '/');
            if (p.Length == 2 && p[1] == ':') p += "\\";   // D: → D:\
            return p;
        }

        public static bool SamePath(string a, string b)
        {
            return string.Equals(Norm(a), Norm(b), StringComparison.OrdinalIgnoreCase);
        }

        // 选中项的父目录。注意：目录要取 Parent，不能取自身
        public static string ParentOf(string path)
        {
            try
            {
                string p = Norm(path);
                if (string.IsNullOrEmpty(p)) return null;
                string parent = Path.GetDirectoryName(p);
                if (string.IsNullOrEmpty(parent))
                {
                    if (p.Length >= 2 && p[1] == ':') return p.Substring(0, 2) + "\\";
                    return null;
                }
                return Norm(parent);
            }
            catch { return null; }
        }

        public static string Leaf(string path)
        {
            try
            {
                string p = Norm(path);
                if (string.IsNullOrEmpty(p)) return path;
                string leaf = Path.GetFileName(p);
                return leaf.Length == 0 ? p : leaf;
            }
            catch { return path; }
        }

        public static bool Exists(string p)
        {
            return File.Exists(p) || Directory.Exists(p);
        }

        public static string Join(string dir, string name)
        {
            if (dir.EndsWith("\\") || dir.EndsWith("/")) return dir + name;
            return dir + Path.DirectorySeparatorChar + name;
        }
    }

    // =================================================================
    // 日志 / 目录
    // =================================================================
    static class Loc
    {
        public static string QDir
        {
            get { return Path.Combine(Path.GetTempPath(), "NewFolderWithSelection"); }
        }

        public static string LogFile { get { return Path.Combine(QDir, "last-run.log"); } }

        public static void Log(string msg)
        {
            try
            {
                Directory.CreateDirectory(QDir);
                string line = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "] " + msg + "\r\n";
                byte[] bytes = new UTF8Encoding(false).GetBytes(line);
                for (int i = 0; i < 20; i++)
                {
                    try
                    {
                        using (FileStream fs = new FileStream(LogFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete))
                        { fs.Write(bytes, 0, bytes.Length); }
                        return;
                    }
                    catch { Thread.Sleep(10); }
                }
            }
            catch { }
        }
    }

    // =================================================================
    // 队列 / 锁 / done 标记（同一批 Explorer 调用共享）
    // =================================================================
    static class Store
    {
        public static string QueueFile { get { return Path.Combine(Loc.QDir, "queue.txt"); } }
        public static string LockFile { get { return Path.Combine(Loc.QDir, "owner.lock"); } }
        public static string DoneFile { get { return Path.Combine(Loc.QDir, "done.txt"); } }

        public static string Esc(string s)
        {
            if (s == null) return "";
            return s.Replace("%", "%25").Replace("\t", "%09").Replace("\r", "%0D").Replace("\n", "%0A");
        }

        public static string Unesc(string s)
        {
            if (s == null) return "";
            return s.Replace("%0A", "\n").Replace("%0D", "\r").Replace("%09", "\t").Replace("%25", "%");
        }

        public static void Append(string path)
        {
            byte[] bytes = new UTF8Encoding(false).GetBytes(Esc(path) + "\r\n");
            for (int i = 0; i < 60; i++)
            {
                try
                {
                    using (FileStream fs = new FileStream(QueueFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete))
                    { fs.Write(bytes, 0, bytes.Length); }
                    return;
                }
                catch { Thread.Sleep(15); }
            }
        }

        public static List<string> ReadQueue()
        {
            List<string> list = new List<string>();
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    using (FileStream fs = new FileStream(QueueFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    using (StreamReader sr = new StreamReader(fs, Encoding.UTF8))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            line = line.Trim();
                            if (line.Length == 0) continue;
                            string p = Unesc(line);
                            if (!Util.Contains(list, p)) list.Add(p);
                        }
                    }
                    return list;
                }
                catch (FileNotFoundException) { return list; }
                catch { Thread.Sleep(15); }
            }
            return list;
        }

        static long Length(string f)
        {
            try
            {
                using (FileStream fs = new FileStream(f, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    return fs.Length;
            }
            catch { return 0; }
        }

        // 等到队列连续 quietMs 毫秒没有增长，最多 capMs
        public static List<string> WaitForQueue(int capMs, int quietMs)
        {
            Stopwatch sw = Stopwatch.StartNew();
            long last = -1;
            long lastChange = 0;
            while (sw.ElapsedMilliseconds < capMs)
            {
                long len = Length(QueueFile);
                if (len != last) { last = len; lastChange = sw.ElapsedMilliseconds; }
                else if (len > 0 && sw.ElapsedMilliseconds - lastChange >= quietMs) break;
                Thread.Sleep(15);
            }
            return ReadQueue();
        }

        // 等 owner 释放锁（最多 capMs）
        public static bool WaitRelease(int capMs)
        {
            Stopwatch sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < capMs)
            {
                if (!File.Exists(LockFile)) return true;
                Thread.Sleep(20);
            }
            return !File.Exists(LockFile);
        }

        public static void CleanStaleLock()
        {
            try
            {
                if (!File.Exists(LockFile)) return;
                if ((DateTime.Now - File.GetLastWriteTime(LockFile)).TotalSeconds > 30)
                    File.Delete(LockFile);
            }
            catch { }
        }

        public static bool TryLock()
        {
            try
            {
                using (FileStream fs = new FileStream(LockFile, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { }
                return true;
            }
            catch { return false; }
        }

        public static void Release()
        {
            try { File.Delete(QueueFile); } catch { }
            try { File.Delete(LockFile); } catch { }
        }

        public static void WriteDone(string target, List<string> items)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
                sb.AppendLine(Esc(target));
                for (int i = 0; i < items.Count; i++) sb.AppendLine(Esc(items[i]));
                File.WriteAllText(DoneFile, sb.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }

        public static bool DoneCovers(List<string> incoming)
        {
            try
            {
                if (!File.Exists(DoneFile)) return false;
                string[] lines = File.ReadAllLines(DoneFile, Encoding.UTF8);
                if (lines.Length < 3) return false;
                DateTime when;
                if (!DateTime.TryParse(lines[0].Trim(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out when))
                    return false;
                double age = (DateTime.UtcNow - when.ToUniversalTime()).TotalSeconds;
                if (age < 0 || age > 10) return false;
                List<string> items = new List<string>();
                for (int i = 2; i < lines.Length; i++)
                {
                    string t = lines[i].Trim();
                    if (t.Length > 0) items.Add(Unesc(t));
                }
                return Util.Covers(items, incoming);
            }
            catch { return false; }
        }
    }

    // =================================================================
    // 一批待移动项（窗口显示期间持续从队列补入）
    // =================================================================
    class Batch
    {
        readonly object _gate = new object();
        readonly string _folder;
        readonly List<string> _items = new List<string>();

        public Batch(string folder) { _folder = folder; }

        public string Folder { get { return _folder; } }

        public int Count { get { lock (_gate) return _items.Count; } }

        // 只收同一父目录、仍然存在、且没重复的项；返回数量是否变化
        public bool Merge(List<string> paths)
        {
            if (paths == null) return false;
            bool changed = false;
            lock (_gate)
            {
                for (int i = 0; i < paths.Count; i++)
                {
                    string p = paths[i];
                    if (string.IsNullOrEmpty(p)) continue;
                    if (!Util.SamePath(Util.ParentOf(p), _folder)) continue;
                    if (Util.Contains(_items, p)) continue;
                    if (!Util.Exists(p)) continue;
                    _items.Add(p);
                    changed = true;
                }
            }
            return changed;
        }

        public List<string> Snapshot()
        {
            lock (_gate) return new List<string>(_items);
        }

        // 轮询队列，把新到的项并进来
        public bool PollQueue() { return Merge(Store.ReadQueue()); }
    }

    // =================================================================
    // 移动一项
    // =================================================================
    class MoveResult
    {
        public string Path;
        public bool Ok;
        public string Reason;
    }

    static class Mover
    {
        public static MoveResult MoveOne(string src, string dstDir)
        {
            try
            {
                string dst = UniqueDest(dstDir, src);
                DoMove(src, dst);
                return new MoveResult { Path = src, Ok = true, Reason = "" };
            }
            catch (Exception ex)
            {
                return new MoveResult { Path = src, Ok = false, Reason = Describe(ex) };
            }
        }

        static void DoMove(string src, string dst)
        {
            bool isDir = Directory.Exists(src);
            Exception first = null;
            try
            {
                if (isDir) Directory.Move(src, dst); else File.Move(src, dst);
                return;
            }
            catch (Exception ex) { first = ex; }

            // 超长路径重试
            string s = LongPath(src), d = LongPath(dst);
            if (s == src && d == dst) throw first;
            try
            {
                if (isDir) Directory.Move(s, d); else File.Move(s, d);
            }
            catch { throw first; }
        }

        static string LongPath(string p)
        {
            if (p.Length < 240) return p;
            if (p.StartsWith(@"\\?\")) return p;
            if (p.StartsWith(@"\\")) return @"\\?\UNC\" + p.Substring(2);
            return @"\\?\" + p;
        }

        static string UniqueDest(string dstDir, string src)
        {
            string leaf = Util.Leaf(src);
            if (leaf.Length == 0) throw new IOException("无法识别的项目名称");
            string dst = Util.Join(dstDir, leaf);
            if (!Util.Exists(dst)) return dst;

            string baseName = Path.GetFileNameWithoutExtension(leaf);
            string ext = Path.GetExtension(leaf);
            if (baseName.Length == 0) { baseName = leaf; ext = ""; }
            int i = 2;
            while (Util.Exists(Util.Join(dstDir, baseName + " (" + i + ")" + ext))) i++;
            return Util.Join(dstDir, baseName + " (" + i + ")" + ext);
        }

        public static bool IsPermission(Exception ex)
        {
            if (ex is UnauthorizedAccessException) return true;
            int code = ex.HResult & 0xFFFF;
            if (code == 5 || code == 1314) return true;
            string m = ex.Message ?? "";
            return m.IndexOf("拒绝访问", StringComparison.Ordinal) >= 0
                || m.IndexOf("Access is denied", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("UnauthorizedAccess", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("需要提升", StringComparison.Ordinal) >= 0
                || m.IndexOf("requires elevation", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsLocked(Exception ex)
        {
            int code = ex.HResult & 0xFFFF;
            if (code == 32 || code == 33) return true;
            string m = ex.Message ?? "";
            return m.IndexOf("正在使用", StringComparison.Ordinal) >= 0
                || m.IndexOf("另一进程", StringComparison.Ordinal) >= 0
                || m.IndexOf("being used", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("another process", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("sharing violation", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        static string Describe(Exception ex)
        {
            string m = ex.Message;
            if (string.IsNullOrEmpty(m)) m = ex.GetType().Name;
            if (IsPermission(ex)) return "权限不足：" + m;
            if (IsLocked(ex)) return "文件被占用：" + m;
            return m;
        }
    }

    // =================================================================
    // 提权子进程：NewFolderWithSelection.exe -JobFile <job.txt>
    //   job.txt   第 1 行 = 目标目录，第 2 行 = 结果文件，第 3 行起 = 待移动项
    //   res.txt   每行 = ok\t<路径>\t<原因>
    // =================================================================
    static class JobRunner
    {
        public static int Run(string jobFile)
        {
            string target = null;
            string resultFile = null;
            List<string> items = new List<string>();
            try
            {
                string[] lines = File.ReadAllLines(jobFile, Encoding.UTF8);
                if (lines.Length < 2) return 2;
                target = Store.Unesc(lines[0].Trim());
                resultFile = Store.Unesc(lines[1].Trim());
                for (int i = 2; i < lines.Length; i++)
                {
                    string t = lines[i].Trim();
                    if (t.Length > 0) items.Add(Store.Unesc(t));
                }
            }
            catch { return 2; }

            StringBuilder sb = new StringBuilder();
            bool dirOk = true;
            try { Directory.CreateDirectory(target); }
            catch (Exception ex)
            {
                dirOk = false;
                sb.AppendLine("0\t" + Store.Esc(target) + "\t" + Store.Esc("无法创建目标文件夹：" + ex.Message));
                Loc.Log("提权子进程建目录失败：" + ex.Message);
            }

            if (dirOk)
            {
                for (int i = 0; i < items.Count; i++)
                {
                    MoveResult r = Mover.MoveOne(items[i], target);
                    sb.AppendLine((r.Ok ? "1" : "0") + "\t" + Store.Esc(r.Path) + "\t" + Store.Esc(r.Reason));
                }
                Loc.Log("提权子进程完成 " + items.Count + " 项 → " + target);
            }

            try { File.WriteAllText(resultFile, sb.ToString(), new UTF8Encoding(false)); }
            catch (Exception ex) { Loc.Log("提权子进程写结果失败：" + ex.Message); return 3; }
            return 0;
        }
    }

    class ElevResult
    {
        public bool Cancelled;
        public List<MoveResult> Results = new List<MoveResult>();
    }

    static class Elevator
    {
        public static ElevResult Run(List<string> items, string target)
        {
            ElevResult er = new ElevResult();
            string id = Guid.NewGuid().ToString("N");
            string job = Path.Combine(Loc.QDir, "job-" + id + ".txt");
            string res = Path.Combine(Loc.QDir, "res-" + id + ".txt");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(Store.Esc(target));
            sb.AppendLine(Store.Esc(res));
            for (int i = 0; i < items.Count; i++) sb.AppendLine(Store.Esc(items[i]));
            try
            {
                Directory.CreateDirectory(Loc.QDir);
                File.WriteAllText(job, sb.ToString(), new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                Loc.Log("写提权任务失败：" + ex.Message);
                er.Cancelled = true;
                return er;
            }

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = Application.ExecutablePath;
                psi.Arguments = "-JobFile \"" + job + "\"";
                psi.Verb = "runas";
                psi.UseShellExecute = true;
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                psi.WorkingDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                using (Process p = Process.Start(psi))
                {
                    if (p != null) p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                Loc.Log("提权被拒绝或失败：" + ex.Message);
                try { File.Delete(job); } catch { }
                er.Cancelled = true;
                return er;
            }

            try
            {
                if (File.Exists(res))
                {
                    string[] lines = File.ReadAllLines(res, Encoding.UTF8);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        string[] parts = lines[i].Split('\t');
                        if (parts.Length < 3) continue;
                        er.Results.Add(new MoveResult
                        {
                            Ok = parts[0].Trim() == "1",
                            Path = Store.Unesc(parts[1]),
                            Reason = Store.Unesc(parts[2])
                        });
                    }
                }
            }
            catch (Exception ex) { Loc.Log("读提权结果失败：" + ex.Message); }

            try { File.Delete(job); } catch { }
            try { File.Delete(res); } catch { }
            return er;
        }

        public static bool IsAdmin()
        {
            try
            {
                System.Security.Principal.WindowsIdentity id = System.Security.Principal.WindowsIdentity.GetCurrent();
                return new System.Security.Principal.WindowsPrincipal(id)
                    .IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch { return false; }
        }
    }

    // =================================================================
    // 快路径：直接读 Explorer 当前选中项（不依赖队列，无需等待）
    // =================================================================
    static class ExplorerSelection
    {
        public static List<string> Get(string folderPath)
        {
            List<string> best = new List<string>();
            object shell = null;
            try
            {
                Type t = Type.GetTypeFromProgID("Shell.Application");
                if (t == null) return null;
                shell = Activator.CreateInstance(t);
                dynamic sh = shell;
                dynamic wins = sh.Windows();
                int count = (int)wins.Count;
                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        dynamic w = wins.Item(i);
                        dynamic doc = w.Document;
                        string wp = (string)doc.Folder.Self.Path;
                        if (!Util.SamePath(wp, folderPath)) continue;
                        dynamic sel = doc.SelectedItems();
                        int n = (int)sel.Count;
                        if (n == 0 || n <= best.Count) continue;
                        List<string> list = new List<string>();
                        for (int j = 0; j < n; j++)
                        {
                            dynamic it = sel.Item(j);
                            list.Add((string)it.Path);
                        }
                        best = list;
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Loc.Log("读 Explorer 选中项失败：" + ex.Message);
                return null;
            }
            finally
            {
                if (shell != null)
                {
                    try { Marshal.ReleaseComObject(shell); } catch { }
                }
            }
            return best;
        }
    }

    // =================================================================
    // 配色 / 字体 / 绘制
    // =================================================================
    class Palette
    {
        public bool Dark;
        public Color Bg, Border, Field, FieldBorder, Text, SubText;
        public Color Accent, AccentHover, AccentPress, OnAccent;
        public Color BtnBg, BtnBorder, BtnHover, BtnPress;

        static int ReadTheme()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    if (k != null)
                    {
                        object v = k.GetValue("AppsUseLightTheme");
                        if (v != null) return Convert.ToInt32(v);
                    }
                }
            }
            catch { }
            return 1;
        }

        static Color ReadAccent()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\DWM"))
                {
                    if (k != null)
                    {
                        object v = k.GetValue("ColorizationColor");
                        if (v != null)
                        {
                            int d = (int)(Convert.ToInt64(v) & 0xFFFFFFFFL);
                            int r = (d >> 16) & 0xFF, g = (d >> 8) & 0xFF, b = d & 0xFF;
                            double lum = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0;
                            if (lum > 0.06 && lum < 0.72) return Color.FromArgb(r, g, b);
                        }
                    }
                }
            }
            catch { }
            return Color.FromArgb(0, 103, 192);
        }

        public static Palette Create()
        {
            return Create(null);
        }

        public static Palette Create(bool? forceDark)
        {
            Palette p = new Palette();
            p.Dark = forceDark.HasValue ? forceDark.Value : (ReadTheme() == 0);
            Color accent = ReadAccent();

            if (!p.Dark)
            {
                p.Bg = Color.FromArgb(251, 251, 251);
                p.Border = Color.FromArgb(226, 226, 226);
                p.Field = Color.White;
                p.FieldBorder = Color.FromArgb(214, 214, 214);
                p.Text = Color.FromArgb(27, 27, 27);
                p.SubText = Color.FromArgb(110, 110, 110);
                p.Accent = accent;
                p.OnAccent = Color.White;
                p.BtnBg = Color.White;
                p.BtnBorder = Color.FromArgb(220, 220, 220);
                p.BtnHover = Color.FromArgb(244, 244, 244);
                p.BtnPress = Color.FromArgb(234, 234, 234);
            }
            else
            {
                p.Bg = Color.FromArgb(32, 32, 32);
                p.Border = Color.FromArgb(58, 58, 58);
                p.Field = Color.FromArgb(44, 44, 44);
                p.FieldBorder = Color.FromArgb(70, 70, 70);
                p.Text = Color.FromArgb(242, 242, 242);
                p.SubText = Color.FromArgb(166, 166, 166);
                p.Accent = Lighten(accent, 0.45);
                p.OnAccent = Color.FromArgb(10, 10, 10);
                p.BtnBg = Color.FromArgb(46, 46, 46);
                p.BtnBorder = Color.FromArgb(74, 74, 74);
                p.BtnHover = Color.FromArgb(56, 56, 56);
                p.BtnPress = Color.FromArgb(64, 64, 64);
            }
            p.AccentHover = Shade(p.Accent, p.Dark ? 1.12 : 0.92);
            p.AccentPress = Shade(p.Accent, p.Dark ? 0.85 : 0.80);
            return p;
        }

        static Color Shade(Color c, double f)
        {
            return Color.FromArgb(
                c.A,
                Math.Min(255, Math.Max(0, (int)Math.Round(c.R * f))),
                Math.Min(255, Math.Max(0, (int)Math.Round(c.G * f))),
                Math.Min(255, Math.Max(0, (int)Math.Round(c.B * f))));
        }

        static Color Lighten(Color c, double t)
        {
            return Color.FromArgb(
                c.A,
                (int)Math.Round(c.R + (255 - c.R) * t),
                (int)Math.Round(c.G + (255 - c.G) * t),
                (int)Math.Round(c.B + (255 - c.B) * t));
        }
    }

    static class Ui
    {
        static string Pick(string[] candidates, float pt, FontStyle style)
        {
            for (int i = 0; i < candidates.Length; i++)
            {
                try
                {
                    Font f = new Font(candidates[i], pt, style, GraphicsUnit.Point);
                    if (string.Equals(f.FontFamily.Name, candidates[i], StringComparison.OrdinalIgnoreCase))
                        return candidates[i];
                    f.Dispose();
                }
                catch { }
            }
            return null;
        }

        public static Font Font(float pt, bool semibold)
        {
            string name;
            FontStyle style = FontStyle.Regular;
            if (semibold)
            {
                name = Pick(new string[] { "Segoe UI Variable Text Semibold", "Segoe UI Semibold" }, pt, FontStyle.Regular);
                if (name == null) { name = Pick(new string[] { "Segoe UI" }, pt, FontStyle.Bold); style = FontStyle.Bold; }
            }
            else
            {
                name = Pick(new string[] { "Segoe UI Variable Text", "Segoe UI" }, pt, FontStyle.Regular);
            }
            if (name == null) return new Font(FontFamily.GenericSansSerif, pt, style, GraphicsUnit.Point);
            return new Font(name, pt, style, GraphicsUnit.Point);
        }

        public static GraphicsPath Round(RectangleF r, float radius)
        {
            GraphicsPath path = new GraphicsPath();
            float d = radius * 2f;
            if (d <= 0.5f || r.Width <= d || r.Height <= d)
            {
                path.AddRectangle(r);
                return path;
            }
            path.AddArc(r.X, r.Y, d, d, 180f, 90f);
            path.AddArc(r.Right - d, r.Y, d, d, 270f, 90f);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0f, 90f);
            path.AddArc(r.X, r.Bottom - d, d, d, 90f, 90f);
            path.CloseFigure();
            return path;
        }

        public const TextFormatFlags BaseFlags =
            TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine |
            TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis;

        public const TextFormatFlags WrapFlags =
            TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.WordBreak |
            TextFormatFlags.Left | TextFormatFlags.Top;

        // 长路径从中间省略：保留盘符 + 末尾（末尾才是用户要确认的目标目录）
        public static string FitPath(string prefix, string path, Font f, int width)
        {
            string full = prefix + path;
            if (Measure(full, f) <= width) return full;
            string head = "";
            if (path.Length > 3 && path[1] == ':') head = path.Substring(0, 3);
            string tail = path.Substring(head.Length);
            for (int cut = 1; cut < tail.Length; cut++)
            {
                string cand = prefix + head + "…" + tail.Substring(cut);
                if (Measure(cand, f) <= width) return cand;
            }
            return full;
        }

        static int Measure(string s, Font f)
        {
            return TextRenderer.MeasureText(s, f, new Size(int.MaxValue, 1000),
                TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine).Width;
        }
    }

    // =================================================================
    // 自绘按钮
    // =================================================================
    class FlatButton : Control
    {
        public Palette P;
        public float S;
        public bool Primary;
        bool _hover, _press;

        public FlatButton(Palette p, float s, string text, bool primary)
        {
            P = p; S = s; Primary = primary;
            Text = text;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Font = Ui.Font(10f, false);
            Cursor = Cursors.Hand;
            TabStop = true;
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _press = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { if (e.Button == MouseButtons.Left) { _press = true; Invalidate(); } base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _press = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        protected override bool IsInputKey(Keys keyData) { return true; }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) { OnClick(EventArgs.Empty); e.Handled = true; }
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(P.Bg);

            Color fill, border, fore;
            if (Primary)
            {
                fill = _press ? P.AccentPress : (_hover ? P.AccentHover : P.Accent);
                border = fill;
                fore = P.OnAccent;
            }
            else
            {
                fill = _press ? P.BtnPress : (_hover ? P.BtnHover : P.BtnBg);
                border = P.BtnBorder;
                fore = P.Text;
            }

            RectangleF r = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (GraphicsPath path = Ui.Round(r, 4f * S))
            {
                using (SolidBrush b = new SolidBrush(fill)) g.FillPath(b, path);
                using (Pen pen = new Pen(border, Math.Max(1f, S))) g.DrawPath(pen, path);

                if (Focused)
                {
                    RectangleF r2 = new RectangleF(r.X + 2f * S, r.Y + 2f * S, r.Width - 4f * S, r.Height - 4f * S);
                    Color ring = Primary ? Color.FromArgb(150, 255, 255, 255) : P.Accent;
                    using (GraphicsPath p2 = Ui.Round(r2, 3f * S))
                    using (Pen pen2 = new Pen(ring, Math.Max(1f, S))) g.DrawPath(pen2, p2);
                }
            }

            TextRenderer.DrawText(g, Text, Font, ClientRectangle, fore,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
        }
    }

    // =================================================================
    // 圆角输入框（内嵌真实 TextBox，保证输入法/光标正常）
    // =================================================================
    class RoundedField : Control
    {
        public Palette P;
        public float S;
        public TextBox Box;

        public RoundedField(Palette p, float s)
        {
            P = p; S = s;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = p.Bg;

            Box = new TextBox();
            Box.BorderStyle = BorderStyle.None;
            Box.BackColor = p.Field;
            Box.ForeColor = p.Text;
            Box.Font = Ui.Font(11f, false);
            Box.AutoSize = false;
            Box.GotFocus += delegate { Invalidate(); };
            Box.LostFocus += delegate { Invalidate(); };
            Controls.Add(Box);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            LayoutBox();
        }

        public void LayoutBox()
        {
            int inset = (int)Math.Round(14f * S);
            int h = Box.Font.Height + (int)Math.Round(6f * S);
            int y = (Height - h) / 2;
            if (y < 0) y = 0;
            Box.SetBounds(inset, y, Math.Max(10, Width - inset * 2), h);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(P.Bg);

            bool focused = Box.Focused;
            RectangleF r = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (GraphicsPath path = Ui.Round(r, 6f * S))
            {
                using (SolidBrush b = new SolidBrush(P.Field)) g.FillPath(b, path);
                using (Pen pen = new Pen(focused ? P.Accent : P.FieldBorder, focused ? 2f * S : Math.Max(1f, S)))
                    g.DrawPath(pen, path);
            }
        }
    }

    // =================================================================
    // 对话框基类：无边框圆角卡片 + 阴影 + 可拖动
    // =================================================================
    class BaseDialog : Form
    {
        protected Palette P;
        protected float S;

        [DllImport("gdi32.dll")]
        static extern IntPtr CreateRoundRectRgn(int l, int t, int r, int b, int w, int h);
        [DllImport("gdi32.dll")]
        static extern bool DeleteObject(IntPtr hObject);
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        public BaseDialog(Palette p, float s)
        {
            P = p; S = s;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            KeyPreview = true;
            AutoScaleMode = AutoScaleMode.None;
            DoubleBuffered = true;
            BackColor = p.Bg;
            Font = Ui.Font(9.5f, false);
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= 0x00020000;   // CS_DROPSHADOW
                return cp;
            }
        }

        protected void ApplyRoundRegion()
        {
            if (Width <= 0 || Height <= 0) return;
            int d = (int)Math.Round(18f * S);
            IntPtr h = CreateRoundRectRgn(0, 0, Width + 1, Height + 1, d, d);
            try
            {
                Region r = Region.FromHrgn(h);
                Region = r;
            }
            finally { DeleteObject(h); }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            ApplyRoundRegion();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(P.Bg);
            RectangleF r = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (GraphicsPath path = Ui.Round(r, 9f * S))
            using (Pen pen = new Pen(P.Border, Math.Max(1f, S)))
                g.DrawPath(pen, path);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0084)   // WM_NCHITTEST
            {
                base.WndProc(ref m);
                if ((int)m.Result == 1)   // HTCLIENT
                {
                    int lp = m.LParam.ToInt32();
                    int y = (short)((lp >> 16) & 0xFFFF);
                    if (y < (int)Math.Round(52f * S)) m.Result = (IntPtr)2;   // HTCAPTION
                }
                return;
            }
            base.WndProc(ref m);
        }

        protected void CenterOnCursorScreen()
        {
            Screen scr = Screen.FromPoint(Cursor.Position);
            Rectangle wa = scr.WorkingArea;
            int x = wa.X + (wa.Width - Width) / 2;
            int y = wa.Y + (int)((wa.Height - Height) * 0.45);
            Location = new Point(x, y);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Activate();
            BringToFront();
            SetForegroundWindow(Handle);
        }
    }

    // =================================================================
    // 命名对话框
    // =================================================================
    class NameDialog : BaseDialog
    {
        RoundedField _field;
        FlatButton _ok, _cancel;
        string _sub;
        string _path;
        Batch _batch;
        System.Windows.Forms.Timer _poll;
        bool _selStarted;
        public string Value;

        public NameDialog(Palette p, float s, string suggested, Batch batch)
            : this(p, s, suggested, batch.Folder, batch.Count)
        {
            _batch = batch;
        }

        // 调试渲染用（不带实时批）
        public NameDialog(Palette p, float s, string suggested, string folderPath, int count)
            : base(p, s)
        {
            _sub = Subtitle(count);
            _path = folderPath;

            Text = "新建文件夹并移入";
            ClientSize = new Size(Px(460), Px(218));

            _field = new RoundedField(p, s);
            _field.Text = suggested;
            _field.Box.Text = suggested;
            _field.SetBounds(Px(24), Px(77), Px(412), Px(38));
            Controls.Add(_field);
            _field.LayoutBox();

            _cancel = new FlatButton(p, s, "取消", false);
            _cancel.SetBounds(Px(250), Px(160), Px(84), Px(34));
            _cancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(_cancel);

            _ok = new FlatButton(p, s, "新建", true);
            _ok.SetBounds(Px(344), Px(160), Px(92), Px(34));
            _ok.Click += delegate { Accept(); };
            Controls.Add(_ok);

            CenterOnCursorScreen();
        }

        static string Subtitle(int n) { return "将 " + n + " 个项目移入新文件夹"; }

        int Px(float v) { return (int)Math.Round(v * S); }

        // 窗口开着的时候，Explorer 的其余调用还在往队列里追加 → 实时并进来
        void Poll()
        {
            if (_batch == null || IsDisposed) return;

            if (_batch.PollQueue()) RefreshSubtitle();
            if (!_selStarted && _batch.Count >= 15) StartSelectionRead();
        }

        void RefreshSubtitle()
        {
            _sub = Subtitle(_batch.Count);
            Loc.Log("批内计数 " + _batch.Count);
            Invalidate(new Rectangle(Px(24), Px(45), Px(412), Px(24)));
        }

        // 队列到 15 项 ≈ Explorer 的 15 项上限，说明有项没传进来；COM 读真实选中项（200-400ms）
        void StartSelectionRead()
        {
            _selStarted = true;
            string folder = _batch.Folder;
            Thread t = new Thread(delegate()
            {
                List<string> sel = ExplorerSelection.Get(folder);
                if (sel == null || sel.Count == 0) return;
                try
                {
                    BeginInvoke((MethodInvoker)delegate
                    {
                        if (_batch == null || IsDisposed) return;
                        if (_batch.Merge(sel)) RefreshSubtitle();
                    });
                }
                catch { }
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        void Accept()
        {
            string name = Sanitize(_field.Box.Text);
            if (name.Length == 0)
            {
                _field.Box.Focus();
                _field.Box.SelectAll();
                return;
            }
            Value = name;
            DialogResult = DialogResult.OK;
            Close();
        }

        public static string Sanitize(string s)
        {
            if (s == null) return "";
            s = s.Trim();
            char[] bad = new char[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };
            StringBuilder sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                sb.Append(Array.IndexOf(bad, c) >= 0 ? '_' : c);
            }
            string r = sb.ToString();
            int end = r.Length;
            while (end > 0 && (r[end - 1] == ' ' || r[end - 1] == '.')) end--;
            return r.Substring(0, end);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _field.Box.Focus();
            _field.Box.SelectAll();

            if (_batch == null) return;
            // 窗口立刻可见，后续项在打开期间陆续并入（不影响打字）
            _poll = new System.Windows.Forms.Timer();
            _poll.Interval = 40;
            _poll.Tick += delegate { Poll(); };
            _poll.Start();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (_poll != null) { _poll.Stop(); _poll.Dispose(); _poll = null; }
            _batch = null;
            base.OnFormClosed(e);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return true;
            }
            if (keyData == Keys.Enter)
            {
                Accept();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;

            using (Font ft = Ui.Font(13f, true))
                TextRenderer.DrawText(g, "新建文件夹并移入",
                    ft, new Rectangle(Px(24), Px(16), Px(412), Px(30)), P.Text, Ui.BaseFlags);

            using (Font fs = Ui.Font(9.5f, false))
                TextRenderer.DrawText(g, _sub,
                    fs, new Rectangle(Px(24), Px(47), Px(412), Px(20)), P.SubText, Ui.BaseFlags);

            using (Font fh = Ui.Font(9f, false))
                TextRenderer.DrawText(g, Ui.FitPath("位置：", _path, fh, Px(412)),
                    fh, new Rectangle(Px(24), Px(123), Px(412), Px(18)), P.SubText, Ui.BaseFlags);
        }
    }

    // =================================================================
    // 消息对话框（替换系统 MessageBox，风格一致）
    // =================================================================
    class MessageDialog : BaseDialog
    {
        string _title, _body;
        FlatButton _ok;
        int _bodyH;

        public MessageDialog(Palette p, float s, string title, string body)
            : base(p, s)
        {
            _title = title;
            _body = body;
            Text = title;

            using (Font fb = Ui.Font(10f, false))
            {
                Size sz = TextRenderer.MeasureText(body, fb,
                    new Size(Px(412), 10000), Ui.WrapFlags);
                _bodyH = Math.Min(sz.Height, Px(320));
            }

            int h = Px(56) + _bodyH + Px(16) + Px(34) + Px(24);
            ClientSize = new Size(Px(460), h);

            _ok = new FlatButton(p, s, "确定", true);
            _ok.SetBounds(Px(344), h - Px(24) - Px(34), Px(92), Px(34));
            _ok.Click += delegate { DialogResult = DialogResult.OK; Close(); };
            Controls.Add(_ok);

            CenterOnCursorScreen();
        }

        int Px(float v) { return (int)Math.Round(v * S); }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape || keyData == Keys.Enter)
            {
                DialogResult = DialogResult.OK;
                Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _ok.Focus();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;

            using (Font ft = Ui.Font(13f, true))
                TextRenderer.DrawText(g, _title,
                    ft, new Rectangle(Px(24), Px(16), Px(412), Px(30)), P.Text, Ui.BaseFlags);

            using (Font fb = Ui.Font(10f, false))
                TextRenderer.DrawText(g, _body,
                    fb, new Rectangle(Px(24), Px(56), Px(412), _bodyH), P.Text, Ui.WrapFlags);
        }
    }

    // =================================================================
    // 主流程
    // =================================================================
    static class Program
    {
        static float SystemScale()
        {
            try
            {
                using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
                    return g.DpiX / 96f;
            }
            catch { return 1f; }
        }

        [STAThread]
        static int Main(string[] rawArgs)
        {
            if (rawArgs.Length >= 2 && string.Equals(rawArgs[0], "-JobFile", StringComparison.OrdinalIgnoreCase))
                return JobRunner.Run(rawArgs[1]);

            // 无控制台的 winexe 里 stdout 默认是 ANSI，中文会变成 '?'；测试靠它读结果
            try { Console.OutputEncoding = new UTF8Encoding(false); } catch { }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 调试：把界面渲染成 PNG（不需要眼睛也能验证配色/排版）
            if (rawArgs.Length >= 2 && string.Equals(rawArgs[0], "-Shot", StringComparison.OrdinalIgnoreCase))
                return Shot(rawArgs);

            Stopwatch sw = Stopwatch.StartNew();

            List<string> incoming = new List<string>();
            for (int i = 0; i < rawArgs.Length; i++)
            {
                string t = (rawArgs[i] ?? "").Trim();
                if (t.Length >= 2 && t[0] == '"' && t[t.Length - 1] == '"')
                    t = t.Substring(1, t.Length - 2).Trim();
                if (t.Length > 0 && !Util.Contains(incoming, t)) incoming.Add(t);
            }
            if (incoming.Count == 0) return 0;

            Loc.Log("启动：" + incoming.Count + " 个参数 → " + string.Join(" | ", incoming.ToArray()));

            if (Store.DoneCovers(incoming))
            {
                Loc.Log("迟到的调用，上一轮已处理，跳过");
                return 0;
            }

            try { Directory.CreateDirectory(Loc.QDir); } catch { }
            for (int i = 0; i < incoming.Count; i++) Store.Append(incoming[i]);

            Store.CleanStaleLock();
            if (!Store.TryLock())
            {
                // 同批的其他进程可能正在收尾：等它结束，再确认本项是否已被移走。
                // 若锁一直没释放（窗口还开着），说明项已在队列里，owner 会取走，直接退出即可。
                Loc.Log("已有主进程，等待其收尾");
                if (!Store.WaitRelease(1500) || Store.DoneCovers(incoming))
                {
                    Loc.Log("已有主进程，本进程退出");
                    return 0;
                }
                Store.CleanStaleLock();
                if (!Store.TryLock())
                {
                    Loc.Log("已有主进程，本进程退出");
                    return 0;
                }
            }

            try { return RunOwner(sw, incoming); }
            finally { }
        }

        static int RunOwner(Stopwatch sw, List<string> incoming)
        {
            string folderPath = Util.ParentOf(incoming[0]);
            if (folderPath == null)
            {
                Loc.Log("无法确定父目录：" + incoming[0]);
                Store.Release();
                return 0;
            }

            bool quiet = Util.EnvFlag("NFS_QUIET");
            bool noElev = Util.EnvFlag("NFS_NOELEVATE");
            string envName = Util.Env("NFS_NAME");
            bool noUi = quiet || envName != null;

            Batch batch = new Batch(folderPath);
            batch.Merge(incoming);
            if (batch.Count == 0)
            {
                Loc.Log("一项都取不到，退出");
                Store.Release();
                return 0;
            }

            if (noUi)
            {
                // 无窗口：先问 Explorer 要真实选中项（唯一能知道“一共选了几项”的来源），
                // 拿不到再靠队列静默兜底
                List<string> sel = ExplorerSelection.Get(folderPath);
                if (sel != null && sel.Count >= incoming.Count && Util.Covers(sel, incoming))
                {
                    batch.Merge(sel);
                    Loc.Log("Explorer 选中项命中 " + sel.Count + " 项（" + sw.ElapsedMilliseconds + "ms）");
                }
                batch.Merge(Store.WaitForQueue(3000, 150));
                Loc.Log("队列 " + batch.Count + " 项（" + sw.ElapsedMilliseconds + "ms）");
            }
            // 有窗口时不在这里等：窗口立刻出现，NameDialog 每 40ms 轮询队列补入

            string suggested = SuggestName(batch.Snapshot());

            if (Util.EnvFlag("NFS_NOMOVE"))
            {
                Emit(suggested);
                Store.Release();
                return 0;
            }

            string newName;
            if (envName != null) newName = envName;
            else if (quiet) newName = suggested;
            else
            {
                Palette pal = Palette.Create();
                float s = SystemScale();
                using (NameDialog dlg = new NameDialog(pal, s, suggested, batch))
                {
                    Loc.Log("窗口显示（" + sw.ElapsedMilliseconds + "ms）");
                    if (dlg.ShowDialog() != DialogResult.OK) newName = null;
                    else newName = dlg.Value;
                }
            }

            if (newName == null)
            {
                Loc.Log("用户取消");
                Store.Release();
                return 0;
            }

            newName = NameDialog.Sanitize(newName);
            if (newName.Length == 0)
            {
                Loc.Log("名字为空，退出");
                Store.Release();
                return 0;
            }

            // 收尾：确认之后可能还有极少数迟到进程
            batch.Merge(Store.WaitForQueue(800, 120));

            // 队列到 15 项 ≈ Explorer 的 15 项上限 → 读真实选中项补齐（无窗口路径在这里同步补）
            if (batch.Count >= 15)
            {
                List<string> sel = ExplorerSelection.Get(folderPath);
                if (sel != null && sel.Count > 0 && batch.Merge(sel))
                    Loc.Log("Explorer 选中项补齐到 " + batch.Count + " 项（" + sw.ElapsedMilliseconds + "ms）");
            }

            List<string> items = batch.Snapshot();
            if (items.Count == 0)
            {
                Loc.Log("一项都取不到，退出");
                Store.Release();
                return 0;
            }
            Loc.Log("本批 " + items.Count + " 项（" + sw.ElapsedMilliseconds + "ms）");

            string target = Util.Join(folderPath, newName);
            int n = 2;
            while (Util.Exists(target)) { target = Util.Join(folderPath, newName + " (" + n + ")"); n++; }

            // ---- 建目录 ----
            bool dirReady = false;
            string dirErr = null;
            try { Directory.CreateDirectory(target); dirReady = true; }
            catch (Exception ex) { dirErr = ex.Message; Loc.Log("建目录失败：" + ex.Message); }

            List<MoveResult> results = new List<MoveResult>();

            if (!dirReady)
            {
                if (Elevator.IsAdmin() || noElev)
                {
                    Report(quiet, "新建文件夹并移入",
                        "无法在该位置新建文件夹（可能只读或受保护）。\n\n" + folderPath + "\n\n" + dirErr);
                    Store.Release();
                    return 0;
                }
                Loc.Log("需要管理员权限，提权建目录");
                ElevResult er = Elevator.Run(items, target);
                if (er.Cancelled)
                {
                    Report(quiet, "新建文件夹并移入", "已取消管理员授权，没有移动任何文件。");
                    Store.Release();
                    return 0;
                }
                results = er.Results;
                if (results.Count == 0)
                {
                    Report(quiet, "新建文件夹并移入",
                        "提权没有返回任何结果，文件没有移动。\n\n日志：" + Loc.LogFile);
                    Store.Release();
                    return 0;
                }
            }
            else
            {
                for (int i = 0; i < items.Count; i++) results.Add(Mover.MoveOne(items[i], target));

                // 只有“权限类”失败才提权重试；被占用提权没用，不打扰用户
                List<string> permFail = new List<string>();
                for (int i = 0; i < results.Count; i++)
                    if (!results[i].Ok && Mover.IsPermission(new Exception(results[i].Reason))) permFail.Add(results[i].Path);

                if (permFail.Count > 0 && !Elevator.IsAdmin() && !noElev)
                {
                    Loc.Log("权限类失败 " + permFail.Count + " 项，提权重试");
                    ElevResult er = Elevator.Run(permFail, target);
                    if (!er.Cancelled && er.Results.Count > 0)
                    {
                        Dictionary<string, MoveResult> map = new Dictionary<string, MoveResult>(StringComparer.OrdinalIgnoreCase);
                        for (int i = 0; i < er.Results.Count; i++) map[er.Results[i].Path] = er.Results[i];
                        for (int i = 0; i < results.Count; i++)
                        {
                            MoveResult r;
                            if (!results[i].Ok && map.TryGetValue(results[i].Path, out r)) results[i] = r;
                        }
                    }
                }
            }

            // ---- 汇总 ----
            List<string> okList = new List<string>();
            List<MoveResult> failList = new List<MoveResult>();
            for (int i = 0; i < results.Count; i++)
            {
                if (results[i].Ok) okList.Add(results[i].Path);
                else failList.Add(results[i]);
            }

            if (okList.Count > 0)
            {
                Store.WriteDone(target, okList);
                Loc.Log("成功 " + okList.Count + " 项 → " + target);
            }

            if (failList.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("已移入 ").Append(okList.Count).Append(" 项，失败 ").Append(failList.Count).Append(" 项。\n\n");
                int shown = Math.Min(failList.Count, 8);
                for (int i = 0; i < shown; i++)
                {
                    sb.Append("• ").Append(Util.Leaf(failList[i].Path)).Append('\n');
                    sb.Append("    ").Append(failList[i].Reason).Append('\n');
                }
                if (failList.Count > shown)
                    sb.Append("\n…另有 ").Append(failList.Count - shown).Append(" 项，详见日志：\n").Append(Loc.LogFile);
                else
                    sb.Append("\n位置：").Append(target);

                StringBuilder log = new StringBuilder();
                for (int i = 0; i < failList.Count; i++)
                    log.Append(failList[i].Path).Append(" :: ").Append(failList[i].Reason).Append(" | ");
                Loc.Log("失败 " + failList.Count + " 项：" + log.ToString());

                Report(quiet, "新建文件夹并移入", sb.ToString());
            }
            else if (results.Count > 0)
            {
                Loc.Log("全部成功：" + okList.Count + " 项 → " + target);
            }

            Store.Release();
            return 0;
        }

        // 调试：-Shot <png> [light|dark] [name|msg]
        static int Shot(string[] a)
        {
            string outPath = a[1];
            string mode = a.Length > 2 ? a[2] : "light";
            string kind = a.Length > 3 ? a[3] : "name";

            Palette p = Palette.Create(mode == "dark");
            float s = SystemScale();

            Form f;
            if (kind == "msg")
                f = new MessageDialog(p, s, "新建文件夹并移入",
                    "已移入 3 项，失败 2 项。\n\n• 季度报告.docx\n    文件被占用：另一个程序正在使用此文件。\n• 客户资料\n    权限不足：对路径的访问被拒绝。\n\n位置：D:\\资料\\2026");
            else
                f = new NameDialog(p, s, "图片 20261101-203000", @"D:\资料\2026\项目\客户交付\第一期", 5);

            f.StartPosition = FormStartPosition.Manual;
            f.Location = new Point(-30000, -30000);
            f.ShowInTaskbar = false;
            f.Show();
            Application.DoEvents();
            try
            {
                using (Bitmap bmp = new Bitmap(f.Width, f.Height))
                {
                    f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                    bmp.Save(outPath, System.Drawing.Imaging.ImageFormat.Png);
                }
                Console.WriteLine(f.Width + "x" + f.Height);
            }
            finally { f.Close(); }
            return 0;
        }

        static void Emit(string text)
        {
            // 测试/脚本靠这个文件读结果：winexe 的 stdout 是 ANSI，中文会被误解码
            try { File.WriteAllText(Path.Combine(Loc.QDir, "last-output.txt"), text, new UTF8Encoding(false)); }
            catch { }
            try { Console.WriteLine(text); } catch { }
        }

        static void Report(bool quiet, string title, string body)
        {
            if (quiet)
            {
                Emit(body);
                Loc.Log(body.Replace("\r", "").Replace("\n", " | "));
                return;
            }
            try
            {
                Palette pal = Palette.Create();
                float s = SystemScale();
                using (MessageDialog dlg = new MessageDialog(pal, s, title, body))
                    dlg.ShowDialog();
            }
            catch (Exception ex) { Loc.Log("弹提示失败：" + ex.Message); }
        }

        // ---- 建议名 ----
        static string SuggestName(List<string> items)
        {
            string stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            if (items.Count == 1)
            {
                string p = items[0];
                if (Directory.Exists(p)) return Util.Leaf(p);
                string name = Path.GetFileName(p);
                if (string.IsNullOrEmpty(name)) return "新建文件夹 " + stamp;
                string noExt = Path.GetFileNameWithoutExtension(name);
                return noExt.Length == 0 ? name : noExt;
            }

            List<string> exts = new List<string>();
            for (int i = 0; i < items.Count; i++)
            {
                if (Directory.Exists(items[i])) continue;
                string e = Path.GetExtension(items[i]).TrimStart('.').ToLowerInvariant();
                if (e.Length > 0 && !exts.Contains(e)) exts.Add(e);
            }
            if (exts.Count == 0) return "新建文件夹 " + stamp;

            string label = null;
            if (AllIn(exts, new string[] { "jpg", "jpeg", "png", "gif", "webp", "bmp", "ico", "tif", "tiff", "heic", "heif", "svg" })) label = "图片";
            else if (AllIn(exts, new string[] { "mp4", "mkv", "mov", "avi", "wmv", "flv", "webm", "m4v" })) label = "视频";
            else if (AllIn(exts, new string[] { "mp3", "wav", "flac", "aac", "ogg", "m4a", "wma" })) label = "音频";
            else if (AllIn(exts, new string[] { "pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "txt", "md", "csv" })) label = "文档";
            else if (AllIn(exts, new string[] { "zip", "rar", "7z", "tar", "gz" })) label = "压缩包";
            else if (AllIn(exts, new string[] { "exe", "msi", "bat", "cmd", "ps1" })) label = "程序";

            if (label == null)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < exts.Count; i++)
                {
                    if (i > 0) sb.Append('+');
                    sb.Append(exts[i].ToUpperInvariant());
                }
                label = sb.Append(" 文件").ToString();
            }
            return label + " " + stamp;
        }

        static bool AllIn(List<string> values, string[] allowed)
        {
            for (int i = 0; i < values.Count; i++)
                if (Array.IndexOf(allowed, values[i]) < 0) return false;
            return values.Count > 0;
        }
    }
}
