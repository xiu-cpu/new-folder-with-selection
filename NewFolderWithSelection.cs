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

// 声明目标框架：.NET Framework 4.6.2+ 默认打开长路径支持，不再依赖旁挂的
// NewFolderWithSelection.exe.config（那个文件丢了就退回 legacy 路径处理，
// >260 的路径会静默失败）。有了这一行，exe 单独拷到别处也是完整的。
[assembly: System.Runtime.Versioning.TargetFramework(".NETFramework,Version=v4.8", FrameworkDisplayName = ".NET Framework 4.8")]

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

                // Path.GetDirectoryName 对 >260 的路径会抛 PathTooLongException，
                // 加 \\?\ 前缀才肯算；算完再把前缀去掉
                bool prefixed = p.Length >= 240 && !p.StartsWith(@"\\?\");
                string parent = Path.GetDirectoryName(prefixed ? Full(p) : p);

                if (string.IsNullOrEmpty(parent))
                {
                    if (p.Length >= 2 && p[1] == ':') return p.Substring(0, 2) + "\\";
                    return null;
                }
                if (prefixed)
                {
                    if (parent.StartsWith(@"\\?\UNC\")) parent = @"\\" + parent.Substring(8);
                    else if (parent.StartsWith(@"\\?\")) parent = parent.Substring(4);
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

        // 超过 240 字符就加 \\?\ 前缀：本程序声明了 longPathAware，加上前缀后
        // 文件系统调用才能突破 260；短路径保持原样，避免改变既有行为
        public static string Full(string p)
        {
            if (string.IsNullOrEmpty(p)) return p;
            if (p.Length < 240) return p;
            if (p.StartsWith(@"\\?\")) return p;
            if (p.StartsWith(@"\\")) return @"\\?\UNC\" + p.Substring(2);
            return @"\\?\" + p;
        }

        public static bool Exists(string p)
        {
            if (string.IsNullOrEmpty(p)) return false;
            return File.Exists(Full(p)) || Directory.Exists(Full(p));
        }

        // 判断“是不是目录”必须走这里：没有 .exe.config 时 Directory.Exists 对 >260
        // 的路径返回 false，深目录会被当成文件去 File.Move
        public static bool DirExists(string p)
        {
            if (string.IsNullOrEmpty(p)) return false;
            try { return Directory.Exists(Full(p)); }
            catch { return false; }
        }

        public static string Clip(string s, int max)
        {
            if (string.IsNullOrEmpty(s) || s.Length <= max) return s;
            return s.Substring(0, max) + "…";
        }

        // 长路径下 Directory.CreateDirectory 需要前缀；万一运行库不认前缀就退回普通调用
        public static void CreateDirectory(string p)
        {
            try { Directory.CreateDirectory(Full(p)); }
            catch (ArgumentException) { Directory.CreateDirectory(p); }
            catch (NotSupportedException) { Directory.CreateDirectory(p); }
        }

        // 只删空目录：先确认真的空，非空一律不删（否则“目录非空”会被报成“空目录删不掉”）
        public static bool DeleteEmptyDir(string p)
        {
            try
            {
                if (!DirExists(p)) return false;
                if (Directory.GetFileSystemEntries(Full(p)).Length > 0) return false;
                Directory.Delete(Full(p), false);
                return true;
            }
            catch { return false; }
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

        // FileMode.Append 不是原子追加（先定位到当时的 EOF 再写），多个进程同时写会
        // 静默覆盖彼此的行。命名互斥体把“打开→写→关闭”串起来。
        const long LogMaxBytes = 2 * 1024 * 1024;
        static Mutex _logMutex;
        static readonly object _logMutexGate = new object();

        static Mutex LogMutex
        {
            get
            {
                lock (_logMutexGate)
                {
                    if (_logMutex == null)
                    {
                        try { _logMutex = new Mutex(false, @"Local\NewFolderWithSelection.log.v1"); }
                        catch { }
                    }
                    return _logMutex;
                }
            }
        }

        public static void Log(string msg)
        {
            Mutex m = LogMutex;
            bool owned = false;
            try
            {
                if (m != null)
                {
                    try { owned = m.WaitOne(2000); }
                    catch (AbandonedMutexException) { owned = true; }
                }

                Directory.CreateDirectory(QDir);

                // 只增不减的日志会到几十 MB：超过 2MB 就转成 .1（只留一份）
                try
                {
                    FileInfo fi = new FileInfo(LogFile);
                    if (fi.Exists && fi.Length > LogMaxBytes)
                    {
                        string bak = LogFile + ".1";
                        if (File.Exists(bak)) File.Delete(bak);
                        File.Move(LogFile, bak);
                    }
                }
                catch { }

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
            finally
            {
                if (owned) { try { m.ReleaseMutex(); } catch { } }
            }
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
            // 行格式：<写入时刻 UTC ticks>\t<escaped path>
            byte[] bytes = new UTF8Encoding(false).GetBytes(
                DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture) + "\t" + Esc(path) + "\r\n");

            Mutex m = QueueMutex;
            bool owned = false;
            bool written = false;
            try
            {
                if (m != null)
                {
                    try { owned = m.WaitOne(5000); }
                    catch (AbandonedMutexException) { owned = true; }
                }
                for (int i = 0; i < 60; i++)
                {
                    try
                    {
                        using (FileStream fs = new FileStream(QueueFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete))
                        { fs.Write(bytes, 0, bytes.Length); }
                        written = true;
                        return;
                    }
                    catch { Thread.Sleep(15); }
                }
            }
            finally
            {
                if (owned) { try { m.ReleaseMutex(); } catch { } }
                // 60 次重试都失败时这一项会无声消失，至少要在日志里留痕
                if (!written) Loc.Log("队列追加失败（重试 60 次）：" + path);
            }
        }

        // 队列互斥体：.NET Framework 的 FileMode.Append 只是“打开后定位到当前 EOF”，
        // 15 个进程同时写会互相覆盖（实测丢一半的行），必须串行化。
        static Mutex _queueMutex;
        static readonly object _queueMutexGate = new object();

        static Mutex QueueMutex
        {
            get
            {
                lock (_queueMutexGate)
                {
                    if (_queueMutex == null)
                    {
                        try { _queueMutex = new Mutex(false, @"Local\NewFolderWithSelection.queue.v1"); }
                        catch { }
                    }
                    return _queueMutex;
                }
            }
        }

        // 本进程只接受“启动前 2 秒之后”写入的队列行：上一次崩溃残留的旧行
        // 不会被当成这一批的选择（否则用户没选的旧文件也会被搬走）
        public static readonly long AcceptFromTicks =
            DateTime.UtcNow.Ticks - TimeSpan.TicksPerSecond * 2;

        public static bool TryParseLine(string line, out string path, out long ticks)
        {
            path = null;
            ticks = 0;
            if (line == null) return false;
            line = line.Trim();
            if (line.Length == 0) return false;
            int tab = line.IndexOf('\t');
            if (tab <= 0) return false;   // 旧格式/半截行：当作陈旧数据丢掉
            if (!long.TryParse(line.Substring(0, tab), NumberStyles.Integer, CultureInfo.InvariantCulture, out ticks))
                return false;
            string p = Unesc(line.Substring(tab + 1));
            if (p.Length == 0) return false;
            path = p;
            return true;
        }

        public static List<string> ReadQueue()
        {
            List<string> list = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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
                            string p;
                            long ticks;
                            if (!TryParseLine(line, out p, out ticks)) continue;
                            if (ticks < AcceptFromTicks) continue;
                            if (seen.Add(p)) list.Add(p);
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

        // 锁文件内容：<pid>|<UTC ticks>。owner 定期续租（心跳），回收只依据
        // “进程已不存在”或“心跳超过租期”，与对话框开多久无关。
        const int LeaseMs = 30000;      // 租期：心跳停 30 秒才算过期
        const int HeartbeatMs = 5000;   // 心跳间隔

        static System.Threading.Timer _lease;
        static volatile bool _leaseStopped = true;
        static string _ownerFolder;     // 当前 owner 正在处理的目录，写进锁里

        // pid|UTC ticks|owner 目录（'|' 不是合法文件名字符，可以安全地当分隔符）
        static string LeasePayload()
        {
            return Process.GetCurrentProcess().Id + "|" + DateTime.UtcNow.Ticks + "|" + (_ownerFolder ?? "");
        }

        static void TouchLease()
        {
            if (_leaseStopped) return;
            // 锁已被回收并重建（内容里的 pid 不是自己）：我们不再是 owner，
            // 不能再用 FileMode.Create 把别人的锁覆盖回来
            try
            {
                int pid; DateTime beat;
                if (TryReadLease(out pid, out beat) && pid != Process.GetCurrentProcess().Id)
                {
                    StopLease();
                    return;
                }
            }
            catch { }

            try
            {
                byte[] bytes = new UTF8Encoding(false).GetBytes(LeasePayload());
                using (FileStream fs = new FileStream(LockFile, FileMode.Create, FileAccess.Write,
                                                      FileShare.ReadWrite | FileShare.Delete))
                    fs.Write(bytes, 0, bytes.Length);
            }
            catch { }
        }

        // 拿到锁之后调用：立刻写一次心跳，然后每 HeartbeatMs 续一次
        public static void StartLease()
        {
            _leaseStopped = false;
            TouchLease();
            _lease = new System.Threading.Timer(delegate { TouchLease(); }, null, HeartbeatMs, HeartbeatMs);
        }

        static void StopLease()
        {
            _leaseStopped = true;
            System.Threading.Timer t = _lease;
            _lease = null;
            if (t != null) { try { t.Dispose(); } catch { } }
        }

        // 读锁内容；读不出（旧格式空文件 / 正被续租重写）返回 false，调用方退回按时间判断
        static bool TryReadLease(out int pid, out DateTime beat)
        {
            pid = 0;
            beat = DateTime.MinValue;
            try
            {
                string s;
                using (FileStream fs = new FileStream(LockFile, FileMode.Open, FileAccess.Read,
                                                      FileShare.ReadWrite | FileShare.Delete))
                using (StreamReader sr = new StreamReader(fs, Encoding.UTF8))
                    s = sr.ReadToEnd().Trim();

                int bar = s.IndexOf('|');
                if (bar <= 0) return false;

                long ticks;
                if (!int.TryParse(s.Substring(0, bar), NumberStyles.Integer, CultureInfo.InvariantCulture, out pid)) return false;
                if (!long.TryParse(s.Substring(bar + 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out ticks)) return false;
                if (ticks <= 0 || ticks > DateTime.MaxValue.Ticks) return false;

                beat = new DateTime(ticks, DateTimeKind.Utc);
                return true;
            }
            catch { return false; }
        }

        static bool ProcessAlive(int pid)
        {
            if (pid <= 0) return false;
            try
            {
                using (Process p = Process.GetProcessById(pid)) return !p.HasExited;
            }
            catch { return false; }
        }

        // 现 owner 正在处理的目录；读不出或旧格式（没有第三段）返回 null
        public static string ReadOwnerFolder()
        {
            try
            {
                string s;
                using (FileStream fs = new FileStream(LockFile, FileMode.Open, FileAccess.Read,
                                                      FileShare.ReadWrite | FileShare.Delete))
                using (StreamReader sr = new StreamReader(fs, Encoding.UTF8))
                    s = sr.ReadToEnd().Trim();

                string[] parts = s.Split('|');
                if (parts.Length < 3) return null;
                string f = parts[2].Trim();
                return f.Length == 0 ? null : f;
            }
            catch { return null; }
        }

        public static void CleanStaleLock()
        {
            try
            {
                if (!File.Exists(LockFile)) return;

                int pid; DateTime beat;
                if (TryReadLease(out pid, out beat))
                {
                    // owner 还活着且心跳新鲜 → 不动。窗口开多久都不会被误回收。
                    if (ProcessAlive(pid) && (DateTime.UtcNow - beat).TotalMilliseconds <= LeaseMs) return;
                    File.Delete(LockFile);
                    Loc.Log("回收陈旧锁：pid=" + pid + " 心跳=" + beat.ToString("o", CultureInfo.InvariantCulture));
                    return;
                }

                // 读不出内容（旧版本留下的空文件，或恰好读在续租中间）
                if ((DateTime.UtcNow - File.GetLastWriteTimeUtc(LockFile)).TotalSeconds > 30)
                    File.Delete(LockFile);
            }
            catch { }
        }

        public static bool TryLock(string ownerFolder)
        {
            try
            {
                _ownerFolder = ownerFolder;
                byte[] bytes = new UTF8Encoding(false).GetBytes(LeasePayload());
                using (FileStream fs = new FileStream(LockFile, FileMode.CreateNew, FileAccess.Write, FileShare.ReadWrite))
                    fs.Write(bytes, 0, bytes.Length);
                return true;
            }
            catch { return false; }
        }

        public static void Release()
        {
            StopLease();
            try { File.Delete(QueueFile); } catch { }
            try
            {
                // 锁里写的已经不是本进程：说明锁被回收后别人拿到了，别误删
                int pid; DateTime beat;
                if (File.Exists(LockFile) && TryReadLease(out pid, out beat)
                    && pid != Process.GetCurrentProcess().Id) return;
                File.Delete(LockFile);
            }
            catch { }
        }

        public static void WriteDone(string target, List<string> items)
        {
            // 先写临时文件再替换：peer 不会读到半截 done.txt 而重复建目录
            string tmp = DoneFile + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
                sb.AppendLine(Esc(target));
                for (int i = 0; i < items.Count; i++) sb.AppendLine(Esc(items[i]));
                File.WriteAllText(tmp, sb.ToString(), new UTF8Encoding(false));
                try { File.Replace(tmp, DoneFile, null, true); }
                catch
                {
                    try { File.Delete(DoneFile); } catch { }
                    File.Move(tmp, DoneFile);
                }
            }
            catch { try { File.Delete(tmp); } catch { } }
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
    // 队列增量读：记住已读字节偏移，只解析新追加的行。窗口开着时每 40ms
    // 全量重读 + 逐行线性查重是 O(n²)，大选择下会把 UI 线程拖住。
    // =================================================================
    class QueueTail
    {
        long _offset;
        readonly StringBuilder _carry = new StringBuilder();
        readonly Decoder _decoder = Encoding.UTF8.GetDecoder();

        public List<string> ReadNew()
        {
            List<string> list = new List<string>();
            try
            {
                using (FileStream fs = new FileStream(Store.QueueFile, FileMode.Open, FileAccess.Read,
                                                      FileShare.ReadWrite | FileShare.Delete))
                {
                    if (fs.Length < _offset) { _offset = 0; _carry.Length = 0; }   // 队列被释放后重建
                    fs.Seek(_offset, SeekOrigin.Begin);
                    byte[] buf = new byte[8192];
                    char[] chars = new char[Encoding.UTF8.GetMaxCharCount(buf.Length)];
                    int n;
                    while ((n = fs.Read(buf, 0, buf.Length)) > 0)
                    {
                        _offset += n;
                        // Decoder 跨块保持状态：块边界切开的多字节字符不会被拆坏
                        int cn = _decoder.GetChars(buf, 0, n, chars, 0);
                        _carry.Append(chars, 0, cn);
                        Drain(list);
                    }
                }
            }
            catch (FileNotFoundException) { }
            catch { }
            return list;
        }

        // 只处理完整的行；末尾没有换行的一段留在 _carry 里等下一次
        void Drain(List<string> into)
        {
            int start = 0;
            for (int i = 0; i < _carry.Length; i++)
            {
                if (_carry[i] != '\n') continue;
                string p;
                long ticks;
                if (Store.TryParseLine(_carry.ToString(start, i - start), out p, out ticks)
                    && ticks >= Store.AcceptFromTicks)
                    into.Add(p);
                start = i + 1;
            }
            if (start > 0) _carry.Remove(0, start);
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
        readonly HashSet<string> _seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        readonly QueueTail _tail = new QueueTail();

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
                    if (_seen.Contains(p)) continue;
                    if (!Util.Exists(p)) continue;
                    _seen.Add(p);
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

        // 轮询队列，把新到的项并进来（增量读，不重扫整个文件）
        public bool PollQueue() { return Merge(_tail.ReadNew()); }
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
            bool isDir = Util.DirExists(src);
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

        static string LongPath(string p) { return Util.Full(p); }

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
    // 提权子进程：NewFolderWithSelection.exe -JobFile <job.txt> -JobHash <sha256>
    //   job.txt   第 1 行 = 目标目录，第 2 行 = 结果文件，第 3 行起 = 待移动项
    //   -JobHash  父进程对 job.txt 原始字节的 SHA-256（十六进制）。子进程只认
    //             校验通过的那一份字节：写入后被人改过就整份拒绝，不移动任何文件。
    //   res.txt   每行 = ok\t<路径>\t<原因>
    // =================================================================
    static class JobRunner
    {
        public static string Hash(byte[] data)
        {
            using (System.Security.Cryptography.SHA256 sha = System.Security.Cryptography.SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", "");
        }

        public static int Run(string jobFile, string expectedHash)
        {
            string target = null;
            string resultFile = null;
            List<string> items = new List<string>();

            byte[] raw;
            try { raw = File.ReadAllBytes(jobFile); }
            catch { return 2; }

            // 先校验字节，再解析同一份字节：中间没有第二次读取，换不掉
            if (string.IsNullOrEmpty(expectedHash)
                || !string.Equals(Hash(raw), expectedHash, StringComparison.OrdinalIgnoreCase))
            {
                Loc.Log("提权任务校验失败，拒绝执行：" + jobFile);
                return 4;
            }

            try
            {
                string text = new UTF8Encoding(false).GetString(raw);
                string[] lines = text.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
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

            // 目标必须是一个有父目录的具体位置（不能是盘根，也不能是空串）
            string targetParent = Util.ParentOf(target);
            if (string.IsNullOrEmpty(targetParent))
            {
                Loc.Log("提权任务目标无效：" + target);
                return 2;
            }

            StringBuilder sb = new StringBuilder();
            bool createdHere = !Util.Exists(target);
            bool dirOk = true;
            try { Util.CreateDirectory(target); }
            catch (Exception ex)
            {
                dirOk = false;
                sb.AppendLine("0\t" + Store.Esc(target) + "\t" + Store.Esc("无法创建目标文件夹：" + ex.Message));
                Loc.Log("提权子进程建目录失败：" + ex.Message);
            }

            if (dirOk)
            {
                int moved = 0;
                for (int i = 0; i < items.Count; i++)
                {
                    // 只允许搬目标文件夹同级的项：跨目录/跨盘一律拒绝
                    if (!Util.SamePath(Util.ParentOf(items[i]), targetParent))
                    {
                        sb.AppendLine("0\t" + Store.Esc(items[i]) + "\t"
                                      + Store.Esc("拒绝处理：源位置与目标文件夹不在同一目录"));
                        Loc.Log("提权任务越界项已拒绝：" + items[i]);
                        continue;
                    }
                    MoveResult r = Mover.MoveOne(items[i], target);
                    if (r.Ok) moved++;
                    sb.AppendLine((r.Ok ? "1" : "0") + "\t" + Store.Esc(r.Path) + "\t" + Store.Esc(r.Reason));
                }
                // 一项都没搬成：把刚建的空目录撤掉。提权场景下父进程没权限删，
                // 只能由这个管理员子进程自己恢复。
                if (moved == 0 && items.Count > 0 && createdHere && Util.Exists(target))
                {
                    bool gone = Util.DeleteEmptyDir(target);
                    Loc.Log(gone ? "提权子进程全部失败，已回滚空目录：" + target
                                 : "提权子进程全部失败，空目录删不掉：" + target);
                }
                Loc.Log("提权子进程完成 " + moved + "/" + items.Count + " 项 → " + target);
            }

            try { File.WriteAllText(resultFile, sb.ToString(), new UTF8Encoding(false)); }
            catch (Exception ex) { Loc.Log("提权子进程写结果失败：" + ex.Message); return 3; }
            return 0;
        }
    }

    class ElevResult
    {
        public bool Cancelled;
        // null = 用户在 UAC 里拒绝了；非 null = 具体失败原因（写任务文件 / 启动进程 / 超时）
        public string Error;
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

            // 写到磁盘的字节和算哈希的字节是同一份内存缓冲：
            // 文件名不可预测（GUID），内容被改过子进程会发现哈希对不上
            byte[] jobBytes = new UTF8Encoding(false).GetBytes(sb.ToString());
            string jobHash = JobRunner.Hash(jobBytes);
            try
            {
                Directory.CreateDirectory(Loc.QDir);
                using (FileStream fs = new FileStream(job, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                    fs.Write(jobBytes, 0, jobBytes.Length);
            }
            catch (Exception ex)
            {
                // 磁盘满 / Temp 被锁 / 目录不可写 —— 这与“用户在 UAC 里点了否”是两件事
                Loc.Log("写提权任务失败（不是用户拒绝）：" + ex.Message);
                er.Cancelled = true;
                er.Error = "无法准备提权任务文件：" + ex.Message;
                return er;
            }

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = Application.ExecutablePath;
                psi.Arguments = "-JobFile \"" + job + "\" -JobHash " + jobHash;
                psi.Verb = "runas";
                psi.UseShellExecute = true;
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                psi.WorkingDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                using (Process p = Process.Start(psi))
                {
                    // 子进程卡住时不能无限干等（现在会没有任何窗口、没有任何提示）
                    if (p != null && !p.WaitForExit(30000))
                    {
                        Loc.Log("提权子进程 30 秒未退出，放弃等待");
                        try { p.Kill(); } catch { }
                        try { File.Delete(job); } catch { }
                        er.Cancelled = true;
                        er.Error = "管理员进程 30 秒没有响应，已停止等待。";
                        return er;
                    }
                }
            }
            catch (Exception ex)
            {
                System.ComponentModel.Win32Exception we = ex as System.ComponentModel.Win32Exception;
                if (we != null && we.NativeErrorCode == 1223)   // ERROR_CANCELLED：用户点了“否”
                {
                    Loc.Log("提权被用户取消（UAC 1223）");
                }
                else
                {
                    Loc.Log("提权启动失败（不是用户拒绝）：" + ex.Message);
                    er.Error = "无法启动管理员进程：" + ex.Message;
                }
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
                    // 声称成功的行回验源路径：结果文件被改过 / 陈旧时不再被当真
                    for (int i = 0; i < er.Results.Count; i++)
                    {
                        MoveResult r = er.Results[i];
                        if (r.Ok && Util.Exists(r.Path))
                        {
                            r.Ok = false;
                            r.Reason = "结果校验：源路径仍然存在";
                            Loc.Log("结果回验失败（声称成功但源还在）：" + r.Path);
                        }
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
        static string Pick(string[] candidates, float px, FontStyle style)
        {
            for (int i = 0; i < candidates.Length; i++)
            {
                try
                {
                    using (Font f = new Font(candidates[i], px, style, GraphicsUnit.Pixel))
                    {
                        if (string.Equals(f.FontFamily.Name, candidates[i], StringComparison.OrdinalIgnoreCase))
                            return candidates[i];
                    }
                }
                catch { }
            }
            return null;
        }

        // 字体统一按“像素”创建并乘上窗口所在显示器的缩放：Point 单位的 Font.Height
        // 由创建时的屏幕 DC（主屏/登录 DPI）决定，混合 DPI 下会出现“框按 A 屏放大、
        // 字按 B 屏画”的错位。
        static Font Create(float pt, bool semibold, float scale)
        {
            float px = pt * 96f / 72f * scale;
            if (px < 1f) px = 1f;

            string name;
            FontStyle style = FontStyle.Regular;
            if (semibold)
            {
                name = Pick(new string[] { "Segoe UI Variable Text Semibold", "Segoe UI Semibold" }, px, FontStyle.Regular);
                if (name == null) { name = Pick(new string[] { "Segoe UI" }, px, FontStyle.Bold); style = FontStyle.Bold; }
            }
            else
            {
                name = Pick(new string[] { "Segoe UI Variable Text", "Segoe UI" }, px, FontStyle.Regular);
            }
            if (name == null) return new Font(FontFamily.GenericSansSerif, px, style, GraphicsUnit.Pixel);
            return new Font(name, px, style, GraphicsUnit.Pixel);
        }

        // 控件自己持有 Font（Control.Font）时用这个：每次新建，避免被 Control.Dispose 影响
        public static Font OwnFont(float pt, bool semibold, float scale)
        {
            return Create(pt, semibold, scale);
        }

        // 重绘用：同一组参数只建一次，进程生命周期内复用
        static readonly Dictionary<string, Font> _fonts = new Dictionary<string, Font>();
        static readonly object _fontGate = new object();

        public static Font Font(float pt, bool semibold, float scale)
        {
            float px = pt * 96f / 72f * scale;
            if (px < 1f) px = 1f;
            string key = (semibold ? "b" : "r") + "#" + px.ToString("0.##", CultureInfo.InvariantCulture);
            lock (_fontGate)
            {
                Font f;
                if (_fonts.TryGetValue(key, out f)) return f;
                f = Create(pt, semibold, scale);
                _fonts[key] = f;
                return f;
            }
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
        // 结果按“宽度 + 字号 + 路径”缓存：逐字符 Measure 最坏要量 250 次
        static readonly Dictionary<string, string> _fitCache = new Dictionary<string, string>();

        public static string FitPath(string prefix, string path, Font f, int width)
        {
            string key = width.ToString(CultureInfo.InvariantCulture) + "|"
                       + f.Height.ToString(CultureInfo.InvariantCulture) + "|" + prefix + "|" + path;
            string cached;
            if (_fitCache.TryGetValue(key, out cached)) return cached;

            string result = FitPathCore(prefix, path, f, width);
            if (_fitCache.Count > 64) _fitCache.Clear();
            _fitCache[key] = result;
            return result;
        }

        static string FitPathCore(string prefix, string path, Font f, int width)
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
            Font = Ui.OwnFont(10f, false, S);
            Cursor = Cursors.Hand;
            TabStop = true;
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _press = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { if (e.Button == MouseButtons.Left) { _press = true; Invalidate(); } base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _press = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        // 只把 Enter/Space 当成自己的输入键；Tab/方向键要留给窗体做焦点导航
        protected override bool IsInputKey(Keys keyData)
        {
            return keyData == Keys.Enter || keyData == Keys.Space;
        }

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
            Box.Font = Ui.OwnFont(11f, false, s);
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
            Font = Ui.OwnFont(9.5f, false, s);
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
        bool _nameEdited;       // 用户动过输入框就不再自动改名
        bool _nameSuppress;     // 程序自己改输入框时不要当成用户编辑
        public string Value;

        public NameDialog(Palette p, float s, string suggested, Batch batch)
            : this(p, s, suggested, batch.Folder, batch.Count)
        {
            _batch = batch;
            _field.Box.TextChanged += delegate
            {
                if (!_nameSuppress) _nameEdited = true;
            };
            // 用户点进输入框（哪怕还没打字）就别再自动改名，否则光标会跳
            _field.Box.MouseDown += delegate { _nameEdited = true; };
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

            if (_batch.PollQueue()) OnBatchChanged();
            if (!_selStarted && _batch.Count >= 15) StartSelectionRead();
        }

        // 批次变了：更新“将 N 项”，并（在用户没动手改过名字时）按 README 的规则重算默认名
        void OnBatchChanged()
        {
            RefreshSubtitle();
            if (_nameEdited) return;
            string s = Program.SuggestName(_batch.Snapshot());
            if (string.IsNullOrEmpty(s) || string.Equals(s, _field.Box.Text, StringComparison.Ordinal)) return;
            _nameSuppress = true;
            try
            {
                _field.Box.Text = s;
                _field.Box.SelectAll();
                _field.Text = s;
            }
            finally { _nameSuppress = false; }
            Invalidate(new Rectangle(Px(24), Px(77), Px(412), Px(38)));
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
                        if (_batch.Merge(sel)) OnBatchChanged();
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

            Font ft = Ui.Font(13f, true, S);
            TextRenderer.DrawText(g, "新建文件夹并移入",
                ft, new Rectangle(Px(24), Px(16), Px(412), Px(30)), P.Text, Ui.BaseFlags);

            Font fs = Ui.Font(9.5f, false, S);
            TextRenderer.DrawText(g, _sub,
                fs, new Rectangle(Px(24), Px(47), Px(412), Px(20)), P.SubText, Ui.BaseFlags);

            Font fh = Ui.Font(9f, false, S);
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

            Font fb = Ui.Font(10f, false, s);
            Size sz = TextRenderer.MeasureText(body, fb,
                new Size(Px(412), 10000), Ui.WrapFlags);
            // 8 条失败明细 + 表头/位置行在 10pt 下约 320px，原来正好卡在上限被裁掉；
            // 放到 460 留出“另有 N 项、日志路径”那两行的位置
            _bodyH = Math.Min(sz.Height, Px(460));

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

            Font ft = Ui.Font(13f, true, S);
            TextRenderer.DrawText(g, _title,
                ft, new Rectangle(Px(24), Px(16), Px(412), Px(30)), P.Text, Ui.BaseFlags);

            Font fb = Ui.Font(10f, false, S);
            TextRenderer.DrawText(g, _body,
                fb, new Rectangle(Px(24), Px(56), Px(412), _bodyH), P.Text, Ui.WrapFlags);
        }
    }

    // =================================================================
    // 主流程
    // =================================================================
    static class Program
    {
        // 进程改成 PerMonitorV2 之后，屏幕 DC 只会给出主屏 DPI；
        // 窗口是跟着鼠标出现的，所以要问“鼠标所在那块屏”的 DPI。
        [StructLayout(LayoutKind.Sequential)]
        struct NativePoint { public int X; public int Y; }

        [DllImport("user32.dll")]
        static extern IntPtr MonitorFromPoint(NativePoint pt, uint flags);

        [DllImport("shcore.dll")]
        static extern int GetDpiForMonitor(IntPtr hmon, int dpiType, out uint dpiX, out uint dpiY);

        static float SystemScale()
        {
            try
            {
                Point cur = Cursor.Position;
                IntPtr mon = MonitorFromPoint(new NativePoint { X = cur.X, Y = cur.Y }, 2 /*DEFAULTTOPRIMARY*/);
                uint dx, dy;
                if (mon != IntPtr.Zero && GetDpiForMonitor(mon, 0 /*EFFECTIVE_DPI*/, out dx, out dy) == 0 && dx > 0)
                    return dx / 96f;
            }
            catch { }
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
            {
                // 没有 -JobHash 一律拒绝：调用方必须把内容哈希一起交过来
                string jobHash = null;
                if (rawArgs.Length >= 4 && string.Equals(rawArgs[2], "-JobHash", StringComparison.OrdinalIgnoreCase))
                    jobHash = rawArgs[3];
                return JobRunner.Run(rawArgs[1], jobHash);
            }

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
            string myFolder = Util.ParentOf(incoming[0]);
            if (!Store.TryLock(myFolder))
            {
                // owner 在处理别的目录时，本次调用会被队列的父目录过滤掉，等于“点了没反应”。
                // 立刻读一次 owner 目录就能判断，不必先干等 1.5 秒。
                string busy = Store.ReadOwnerFolder();
                if (busy != null && myFolder != null && !Util.SamePath(busy, myFolder))
                {
                    Loc.Log("被另一个目录的处理窗口挡住：" + busy);
                    ReportOtherFolder(busy);
                    return 0;
                }

                // 同批的其他进程可能正在收尾：等它结束，再确认本项是否已被移走。
                // 若锁一直没释放（窗口还开着），说明项已在队列里，owner 会取走，直接退出即可。
                Loc.Log("已有主进程，等待其收尾");
                if (!Store.WaitRelease(1500) || Store.DoneCovers(incoming))
                {
                    busy = Store.ReadOwnerFolder();
                    if (busy != null && myFolder != null && !Util.SamePath(busy, myFolder))
                    {
                        Loc.Log("窗口仍在处理另一个目录：" + busy);
                        ReportOtherFolder(busy);
                    }
                    else Loc.Log("已有主进程，本进程退出");
                    return 0;
                }
                Store.CleanStaleLock();
                if (!Store.TryLock(myFolder))
                {
                    Loc.Log("已有主进程，本进程退出");
                    return 0;
                }
            }

            // 异常路径也必须把锁和队列放掉，否则后续调用会被残留锁挡住
            try
            {
                // 拿到锁之后开始续租：窗口开多久锁都不会被判过期
                Store.StartLease();
                return RunOwner(sw, incoming);
            }
            finally { Store.Release(); }
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

            // 保留名称 / 控制字符：本该在建目录失败时给出“可能只读或受保护”这种误导提示，
            // 现在提前说清楚（也避免拿它去撞建目录失败、再误判成权限问题弹 UAC）
            string nameProblem = NameProblem(newName);
            if (nameProblem != null)
            {
                Loc.Log("名字不可用：" + NameDialog.Sanitize(newName));
                Store.Release();
                Report(quiet, "新建文件夹并移入", nameProblem + "\n\n这次没有移动任何文件。");
                return 0;
            }

            // NTFS 单个名字上限 255。超了就在这里说清楚，不要拿它去撞建目录失败、
            // 再误判成权限问题去弹 UAC。
            if (newName.Length > 255)
            {
                Loc.Log("名字过长：" + newName.Length + " 字符");
                Store.Release();
                Report(quiet, "新建文件夹并移入",
                    "名字太长：文件夹名最多 255 个字符，当前 " + newName.Length + " 个。\n\n"
                    + "请缩短后重试。这次没有移动任何文件。");
                return 0;
            }

            // 收尾：确认之后可能还有极少数迟到进程
            batch.Merge(Store.WaitForQueue(800, 120));

            // 只有在 Explorer 的 15 项上限被打满（确实可能有项没传进来）时才读 Shell COM：
            // 多选路径上无条件读会给“窗口已关、文件还没动”加 200–400ms。
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
            // createdHere 用来判断这个空目录是不是我们刚建的：全失败时只回滚自己建的
            bool createdHere = !Util.Exists(target);
            bool dirReady = false;
            string dirErr = null;
            Exception dirEx = null;
            try { Util.CreateDirectory(target); dirReady = true; }
            catch (Exception ex) { dirEx = ex; dirErr = ex.Message; Loc.Log("建目录失败：" + ex.Message); }

            List<MoveResult> results = new List<MoveResult>();

            if (!dirReady)
            {
                // 只有真正的权限问题才值得提权；名字非法、磁盘满之类弹 UAC 也没用
                if (Elevator.IsAdmin() || noElev || dirEx == null || !Mover.IsPermission(dirEx))
                {
                    Store.Release();
                    Report(quiet, "新建文件夹并移入",
                        "无法在该位置新建文件夹（可能只读或受保护）。\n\n" + folderPath + "\n\n" + dirErr);
                    return 0;
                }
                Loc.Log("需要管理员权限，提权建目录");
                ElevResult er = Elevator.Run(items, target);
                if (er.Cancelled)
                {
                    // 锁在弹提示框之前就放掉，用户放着提示框不管时同目录的后续操作不会被挡住
                    Store.Release();
                    if (er.Error == null)
                        Report(quiet, "新建文件夹并移入", "已取消管理员授权，没有移动任何文件。");
                    else
                        Report(quiet, "新建文件夹并移入",
                            er.Error + "\n\n没有移动任何文件。\n\n日志：" + Loc.LogFile);
                    return 0;
                }
                results = er.Results;
                if (results.Count == 0)
                {
                    Store.Release();
                    Report(quiet, "新建文件夹并移入",
                        "提权没有返回任何结果，文件没有移动。\n\n日志：" + Loc.LogFile);
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

            // 全失败时把刚建的空目录回滚掉：否则会留下一个空的“新建文件夹”，
            // 提权场景下这个空目录还在受保护位置，用户自己删不掉。
            bool rolledBack = false;
            if (okList.Count == 0 && createdHere && Util.Exists(target))
            {
                rolledBack = Util.DeleteEmptyDir(target);
                Loc.Log(rolledBack ? "全部失败，已回滚空目录：" + target
                                   : "全部失败，但空目录删不掉：" + target);
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
                    // 长异常信息（含整条路径）会被消息框高度上限裁掉，逐条截断
                    sb.Append("    ").Append(Util.Clip(failList[i].Reason, 100)).Append('\n');
                }
                if (failList.Count > shown)
                    sb.Append("\n…另有 ").Append(failList.Count - shown).Append(" 项。\n");

                // 回滚过就别再说“位置：”，否则同一段话里既说文件夹在那儿、又说已经删了
                if (!rolledBack) sb.Append("\n位置：").Append(target);
                sb.Append("\n日志：").Append(Loc.LogFile);
                if (rolledBack) sb.Append("\n\n空文件夹已回滚，没有留下空的“新建文件夹”。");

                StringBuilder log = new StringBuilder();
                for (int i = 0; i < failList.Count; i++)
                    log.Append(failList[i].Path).Append(" :: ").Append(failList[i].Reason).Append(" | ");
                Loc.Log("失败 " + failList.Count + " 项：" + log.ToString());

                // 移动已经结束、done.txt 也写好了，先放锁再弹模态框
                Store.Release();
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

        // 另一个目录的窗口正开着：文案必须说清“不是你点的这个文件夹”
        static void ReportOtherFolder(string busy)
        {
            Report(Util.EnvFlag("NFS_QUIET"), "新建文件夹并移入",
                "另一个窗口正在处理另一个文件夹里的选中项：\n\n" + busy
                + "\n\n请先完成或关闭那个窗口，再重试。这次没有移动任何文件。");
        }

        // 名字本身不可用（Windows 保留名称 / 控制字符）：在弹 UAC 之前给出可读提示
        static string NameProblem(string name)
        {
            for (int i = 0; i < name.Length; i++)
                if (name[i] < ' ')
                    return "名字里有不可见字符（控制字符），请删掉后重试。";

            string n = name;
            int dot = n.IndexOf('.');
            if (dot >= 0) n = n.Substring(0, dot);
            n = n.TrimEnd(' ');
            if (n.Length == 0) return null;

            string u = n.ToUpperInvariant();
            bool reserved = u == "CON" || u == "PRN" || u == "AUX" || u == "NUL"
                || (u.Length == 4 && (u.StartsWith("COM", StringComparison.Ordinal) || u.StartsWith("LPT", StringComparison.Ordinal))
                    && u[3] >= '1' && u[3] <= '9');
            if (!reserved) return null;

            return "“" + name + "”是 Windows 的保留名称（CON、NUL、COM1 等），不能用作文件夹名。\n\n"
                 + "请换一个名字重试。";
        }

        // ---- 建议名 ----
        public static string SuggestName(List<string> items)
        {
            string stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            if (items.Count == 1)
            {
                string p = items[0];
                if (Util.DirExists(p)) return Util.Leaf(p);
                string name = Path.GetFileName(p);
                if (string.IsNullOrEmpty(name)) return "新建文件夹 " + stamp;
                string noExt = Path.GetFileNameWithoutExtension(name);
                return noExt.Length == 0 ? name : noExt;
            }

            List<string> exts = new List<string>();
            for (int i = 0; i < items.Count; i++)
            {
                if (Util.DirExists(items[i])) continue;
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
