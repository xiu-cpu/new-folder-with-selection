# 新建文件夹并移入（Windows 右键菜单扩展）

选中文件/文件夹 → 右键 → **显示更多选项 (Shift+F10)** → **新建文件夹并移入**。
弹出一个 Win11 风格的小卡片：确认名字后回车，所有选中项移进新文件夹；Esc 取消，什么都不动。

- 单文件 exe，无控制台窗口、不经过 `wscript`/`powershell`，**窗口约 0.11 秒出现**
- 界面全部自绘：圆角卡片、圆角输入框、强调色主按钮，跟随系统浅色/深色 + 系统强调色
- 支持 150% 等系统 DPI 缩放
- 目标位置不可写时自动 UAC 提权（不会丢文件）

预览（本机 150% DPI 实拍 / 离屏渲染）：

| 文件 | 内容 |
|---|---|
| `preview-light.png` | 真实截图（690x327，含圆角裁切） |
| `shot-light.png` | 浅色主题离屏渲染 |
| `shot-dark.png` | 深色主题离屏渲染 |
| `shot-msg.png` | 失败提示对话框 |

## 默认名字

| 选中内容 | 默认名 |
|---|---|
| 1 个文件 | 文件名（去扩展名） |
| 1 个文件夹 | 文件夹名（重名时自动加 ` (2)`） |
| 全是图片 / 视频 / 音频 / 文档 / 压缩包 / 程序 | `图片 20261101-183001`（类别 + 时间戳） |
| 混合类型 | `PNG+MP4+PDF 文件 时间戳` |
| 其他 | `新建文件夹 时间戳` |

名字会自动去掉首尾空格/句点，非法字符 `\ / : * ? " < > |` 换成 `_`；重名自动追加 ` (2)`、` (3)`。

## 一次选中很多项也只建一个文件夹

Explorer 的 shell 动词对**每个**选中项各调用一次本程序（`"%1"`），而且**最多只调用 15 次**。
所以：

1. 第一个进程抢到 `owner.lock` 后**立刻**弹窗，不等待；
2. 窗口开着的时候，每 40ms 轮询队列，把陆续到达的项并进同一批，副标题实时变成
   「将 N 个项目移入新文件夹」；
3. 队列攒到 15 项 ≈ Explorer 的 15 项上限，说明有项没传进来 → 后台 STA 线程用
   Shell COM 读 Explorer 的**真实选中项**补齐（约 200–400ms，不挡窗口、不影响打字）；
4. 点「新建」后再等约 120ms 队列静默，收掉最后几个迟到的进程，然后统一移动；
5. 移动完成后写 `done.txt`（10 秒内有效），迟到的调用看到自己已被覆盖就直接退出，
   **不会再多建一个文件夹**。

实测：20 项全选、只触发 15 次菜单 → 20 项全部移进同一个文件夹，耗时 0.79 秒。

## 权限：自动提权

| 情况 | 行为 |
|---|---|
| 目标位置不可写（`C:\Program Files`、只读目录、别人的目录） | 弹 UAC 提权，由管理员子进程建目录并移动 |
| 普通权限下个别文件被拒绝访问 | 只把**权限类**失败项交给管理员子进程重试 |
| 文件被占用 / 被其他程序锁住 | **不提权**（提权也解不开占用），逐条列出原因 |
| 用户拒绝 UAC | 什么都不动，提示已取消授权 |
| 已经是管理员 | 直接提示该位置受保护，不动任何文件 |

提权子进程和主程序是同一个 exe（`-JobFile` 协议），结果回传给前台统一汇总。
实测（`C:\Program Files` 下、普通用户不可写）：UAC 静默提权 → 文件夹建出、2 项全部移入，总耗时约 1.0 秒。

## 失败时

- 成功项照常移入；失败项**逐条列出文件名 + 原因**（最多显示 8 条）
- 完整细节写日志：`%TEMP%\NewFolderWithSelection\last-run.log`（每次追加，带毫秒时间戳）

| 提示 | 原因 | 处理 |
|---|---|---|
| 文件被占用 | 有程序正在读写该文件 | 关掉对应程序后重试 |
| 权限不足 | 目标/源位置需要管理员 | 在 UAC 里点「是」（或已经拒绝过） |
| 路径过长/无效 | 选项来自搜索位置、压缩包内等虚拟路径 | 先解压/定位到真实路径 |

## 文件

| 文件 | 用途 |
|---|---|
| `NewFolderWithSelection.cs` | 工具本体（自绘 UI + 批聚合 + 移动 + 提权），单文件 |
| `NewFolderWithSelection.exe` | 编译产物（`build.ps1` 生成），右键菜单直接调用它 |
| `app.manifest` | `asInvoker` + system DPI aware + longPathAware + Common-Controls v6 |
| `build.ps1` | 用系统 `csc.exe` 编译成 `winexe`（无控制台） |
| `install.ps1` | 注册菜单项（只写 `HKCU`，不需要管理员） |
| `uninstall.ps1` | 移除菜单项（不删工具） |
| `tests\` | 回归测试，见 `tests\README.md` |
| `LICENSE` | MIT |

## 安装 / 卸载

```powershell
# 安装（会先编译，再写 HKCU 注册表）
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1

# 卸载菜单项
powershell -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1

# 只重新编译
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
```

注册表写入位置：

```
HKCU\Software\Classes\AllFilesystemObjects\shell\NewFolderWithSelection
```

命令为 `"<目录>\NewFolderWithSelection.exe" "%1"`（`%1` 不能换成 `%*`：实测 Explorer 传 `%*` 是空串）。

## 调试环境变量

| 变量 | 作用 |
|---|---|
| `NFS_NAME=名字` | 跳过输入框，直接用这个名字 |
| `NFS_NOMOVE=1` | 只输出建议名，不建目录、不移动 |
| `NFS_QUIET=1` | 不弹任何窗口，结果写 `last-output.txt` / stdout |
| `NFS_NOELEVATE=1` | 不弹 UAC（测试用） |

界面自检（不需要眼睛）：`NewFolderWithSelection.exe -Shot out.png [light|dark] [name|msg]`
把窗口渲染成 PNG，再用 `tests\shot-analyze.ps1` 做像素分析。

## 已知限制

- Windows 11 下新版右键菜单不会显示这个项，要到「显示更多选项」里（或按 Shift+F10）。
- 只对文件系统对象生效；「此电脑」里的驱动器、回收站里的项目不行。
- 工具目录别移动：挪了位置要重新跑 `install.ps1`（注册表里是绝对路径）。
- 换 Windows 账户要重新跑 `install.ps1`。
- 搜索窗口 / 压缩包 / 库里的“虚拟”项目拿不到真实路径，只能按传给程序的那一项处理。
- 超过 15 项时，最后那几项依赖 Shell COM 读 Explorer 选中项补齐；如果此时用户已经改掉了选中状态，就只能处理已经传进来的项。

## 许可

MIT，见 [`LICENSE`](LICENSE)。
