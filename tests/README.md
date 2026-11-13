# 回归测试

全部脚本用 `powershell -NoProfile -ExecutionPolicy Bypass -File <脚本>` 运行。
脚本按自己的位置定位工具目录，**换目录不用改脚本里的路径**。

## 安全约定（强制）

曾经发生过一件事故：`regression-ui.ps1` 用 `WScript.Shell.SendKeys` 往被测窗口打字，但焦点
不在被测窗口上，按键**落进了用户正在使用的 Trae 输入框**，污染了用户的输入。

所以：

1. **禁止**任何前台注入手段：`SendKeys`、`WScript.Shell.SendKeys`、`keybd_event`、
   `mouse_event`、`SetCursorPos`、`SetForegroundWindow` + 模拟输入、剪贴板读写
   （`Set-Clipboard` / `Get-Clipboard`）一律不许出现在测试里。
2. 需要发按键时**只能** `PostMessage(hwnd, WM_KEYDOWN/WM_KEYUP/WM_CHAR, ...)`，而且必须先校验：
   句柄存在、`GetWindowThreadProcessId` 等于被测进程、标题匹配；校验不过就 **SKIP 并记录**，
   绝不“试试看”。
3. 会往桌面弹窗的脚本默认**不跑**，必须显式 `$env:NFS_ALLOW_UI_TESTS=1`。
   跑完必须确认：没有残留进程、没有残留窗口、`%TEMP%` 里的临时目录已清理。
4. 能用离屏手段验证的优先离屏：`-Shot` 渲染 + `shot-analyze.ps1` 不需要桌面。

| 脚本 | 内容 | 耗时 |
|---|---|---|
| `regression-exe.ps1` | 35 项功能回归：多选/中文/空格/子文件夹/无效项、3 进程并发汇总、迟到调用、单文件/单文件夹/首个是文件夹、`-JobFile` 提权子进程、被占用文件、只读目录、建议名、重名/非法字符/空名字、锁无残留。全部 `NFS_QUIET=1`，**不弹窗** | ~20 秒 |
| `regression-ui.ps1` | 15 项真实窗口测试：窗口出现耗时（3 次）、尺寸、圆角（`PtInRegion` 验证角外/角内）、打字+回车、Esc 取消、空名字不关窗、**窗口开着时后续项实时并入同一批**。**默认 SKIP**，需 `NFS_ALLOW_UI_TESTS=1`；按键只走 PostMessage + 句柄校验 | ~15 秒 |
| `regression-over15.ps1` | 真实注册表菜单项 + Explorer 全选 20 项、只触发 15 次 → 验证 Shell COM 补齐，20 项全进同一个文件夹。**默认 SKIP**（会开 Explorer 窗口），选中用 Shell COM `SelectItem`，不用 SendKeys | ~20 秒 |
| `regression-elevate.ps1` | 真实提权链路：`C:\Program Files` 下普通用户不可写 → UAC 静默提权建目录并移动；测完用管理员权限清理 | ~5 秒 |
| `regression-concurrent.ps1` | **真并发**：15 个实例在同一个紧循环里同时启动（不是串行拉起）×10 轮，断言只建 1 个文件夹、15 项全部移入；另有崩溃残留队列行的污染用例。`-Exe <其它版本>` 可对比修复前后 | ~30 秒 |
| `capture-window.ps1` | 真实截图（`SetProcessDpiAwarenessContext(-4)`，150% 下也是 690x327 原生像素）。**默认 SKIP**；截图改用 `PrintWindow`，不抢前台、不怕遮挡。抓完**自检**输入框区域有没有画出字，没有就重抓（最多 12 次） | ~2 秒 |
| `shot-analyze.ps1` | 把 `-Shot` 生成的 PNG 转成 ASCII 色块图 + 关键区域包围盒，用于无眼睛检查配色/排版 | 即时 |

注意：

- `regression-ui.ps1` / `regression-over15.ps1` / `capture-window.ps1` 会**短暂弹窗 / 打开 Explorer 窗口**，
  默认 SKIP；只在确认没人在用这台机器、且显式设了 `NFS_ALLOW_UI_TESTS=1` 时才跑。
- `regression-elevate.ps1` 需要本机允许 UAC 静默提权（`ConsentPromptBehaviorAdmin=0`），否则会弹 UAC 框；当前已是管理员时会自动 SKIP。
- 提权建出来的目录继承受限 ACL，清理必须再提权（脚本里已处理）。
- `regression-over15.ps1` 会用 Shell COM 设置选中项，并先关掉上次残留的 Explorer 窗口。
- `regression-concurrent.ps1` 是队列原子性的守门测试：旧版本（`FileMode.Append` 直接写）在这台机器上
  15 并发会丢 45%–50% 的项，新版本 10 轮 0 丢项。

## 结果怎么读

- 工具本身不靠 stdout 传中文（`winexe` 的 stdout 是 ANSI，会乱码），测试读的是
  `%TEMP%\NewFolderWithSelection\last-output.txt`（UTF-8）。
- 日志：`%TEMP%\NewFolderWithSelection\last-run.log`（毫秒时间戳，每次追加）。
- 队列/锁：同目录下 `queue.txt`、`owner.lock`、`done.txt`。

## 参考数据（本机 Win11 22631 / 150% DPI / 非管理员）

| 指标 | 数值 |
|---|---|
| 窗口出现耗时（3 次均值） | 约 110–126 ms |
| 单次调用（无窗口，`NFS_QUIET`） | 约 0.5 秒（含 150ms 队列静默 + 120ms 收尾） |
| 20 项全选（真实菜单，只触发 15 次） | 约 0.8 秒，全部进同一个文件夹 |
| 真实提权（`C:\Program Files`） | 约 1.0 秒 |

## 界面自检

```powershell
# 离屏渲染（不用桌面，稳定）
.\NewFolderWithSelection.exe -Shot "$env:TEMP\shot.png" light name
.\NewFolderWithSelection.exe -Shot "$env:TEMP\shot-dark.png" dark msg
powershell -File .\tests\shot-analyze.ps1 -Path "$env:TEMP\shot.png"

# 真实截图（含窗口圆角/阴影裁切、真实字体渲染）
powershell -File .\tests\capture-window.ps1 -Dir "$env:TEMP\pics" -Count 5 -Out preview.png
```

`capture-window.ps1` 会等窗口可见 → 置前 → 抓屏 → 检查输入框区域有没有字，没有就重抓。
注意：**别用 `powershell -File script.ps1 -Files a,b,c` 传数组**，`-File` 只会取到第一个元素
（会让后面的参数错位），所以脚本改成用 `-Dir` + `-Count` 自己构造参数。
