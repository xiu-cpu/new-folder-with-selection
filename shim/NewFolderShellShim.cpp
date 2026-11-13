// NewFolderShellShim.cpp — 实验性 Windows 11 主右键菜单入口（当前未启用）
//
// 背景（2026-11-13 在本机实测结论）：
//   Win11 的主右键菜单只认「有包标识 + 在 AppxManifest 里声明 windows.fileExplorerContextMenus」
//   的条目；未打包的静态 verb、"显示更多选项" 里的 shellex 处理器、以及在经典 verb 上加
//   ExplorerCommandHandler 的写法，都不会出现在主菜单里。本机所有进主菜单的第三方项
//   （WinRAR / QQ / 火绒 / PowerToys）走的都是「稀疏包 + com:SurrogateServer + 包内原生 DLL」。
//
//   包内 com:Class 的 Path 必须是一个**原生** DLL（代理宿主直接 LoadLibrary 并取
//   DllGetClassObject），托管程序集不能放这里，所以这个薄壳必须是 C++。
//   本机没有任何 C/C++ 编译器（无 VS Build Tools、无 cl.exe、无 MinGW），因此这一步没有编译产物，
//   主菜单路径停在「代码与包定义就绪、缺一次原生编译」。
//
// 设计原则：这里不复制任何业务逻辑，Invoke 只把选中项的完整路径拼成命令行去启动
//   同目录的 NewFolderWithSelection.exe —— 与经典菜单里的 `command = "exe" "%1"` 等价，
//   队列、锁、默认名、失败回滚等全部仍由那个 exe 负责。
//
// 编译（有 MSVC 时）：
//   cl /LD /O2 /EHsc /DUNICODE /D_UNICODE NewFolderShellShim.cpp /link /OUT:NewFolderShellShim.dll ole32.lib shell32.lib shlwapi.lib
// 产物放置：与 NewFolderWithSelection.exe 同目录（即稀疏包的 ExternalLocation）。

#include <windows.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <string>
#include <vector>

// 与 package/AppxManifest.xml 里的 <desktop5:Verb Clsid=...> / <com:Class Id=...> 必须一致
static const GUID CLSID_NFWS_SHIM =
{ 0x3F1B9A62, 0x5C41, 0x4E7D, { 0x9A, 0x10, 0x7C, 0x2B, 0x4B, 0x6E, 0x1F, 0x88 } };

static const wchar_t* kTitle = L"\u65b0\u5efa\u6587\u4ef6\u5939\u5e76\u79fb\u5165 (&F)";
static const wchar_t* kExeName = L"NewFolderWithSelection.exe";

// DLL 所在目录 + exe 名；不写死任何绝对路径
static HMODULE g_self = NULL;   // 由 DllMain 记下：代理宿主里不能拿进程主模块

static std::wstring ExePath()
{
    // 这个 DLL 是被 dllhost.exe 之类的代理宿主加载的，GetModuleFileNameW(NULL)
    // 返回的是宿主的路径（拼出来会是 system32\NewFolderWithSelection.exe，必失败）。
    // 必须问“本模块”的句柄：优先用 DllMain 记下的，拿不到就按函数地址反查。
    HMODULE self = g_self;
    if (!self)
        GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                           reinterpret_cast<LPCWSTR>(&ExePath), &self);

    wchar_t buf[MAX_PATH * 2] = { 0 };
    if (!self || !GetModuleFileNameW(self, buf, _countof(buf)))
        return kExeName;   // 连自己都定位不到：只能给相对名，失败了在 Explorer 里看得见

    std::wstring p(buf);
    size_t pos = p.find_last_of(L"\\/");
    if (pos != std::wstring::npos) p = p.substr(0, pos + 1);
    else p.clear();
    return p + kExeName;
}

static HRESULT Dup(const std::wstring& s, LPWSTR* out)
{
    *out = (LPWSTR)CoTaskMemAlloc((s.size() + 1) * sizeof(wchar_t));
    if (!*out) return E_OUTOFMEMORY;
    wcscpy_s(*out, s.size() + 1, s.c_str());
    return S_OK;
}

// 收集选中项的文件系统路径；非文件系统项（搜索结果、压缩包内条目等）跳过
static std::vector<std::wstring> CollectPaths(IShellItemArray* items)
{
    std::vector<std::wstring> paths;
    if (!items) return paths;
    DWORD n = 0;
    if (FAILED(items->GetCount(&n))) return paths;
    for (DWORD i = 0; i < n; i++)
    {
        IShellItem* item = NULL;
        if (FAILED(items->GetItemAt(i, &item)) || !item) continue;
        LPWSTR name = NULL;
        if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &name)) && name)
        {
            paths.push_back(name);
            CoTaskMemFree(name);
        }
        item->Release();
    }
    return paths;
}

class ShimCommand : public IExplorerCommand
{
    LONG _ref;
public:
    ShimCommand() : _ref(1) {}

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IExplorerCommand)
        {
            *ppv = static_cast<IExplorerCommand*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&_ref); }
    IFACEMETHODIMP_(ULONG) Release()
    {
        ULONG c = InterlockedDecrement(&_ref);
        if (c == 0) delete this;
        return c;
    }

    // IExplorerCommand
    IFACEMETHODIMP GetTitle(IShellItemArray*, LPWSTR* name) { return Dup(kTitle, name); }

    IFACEMETHODIMP GetIcon(IShellItemArray*, LPWSTR* icon) { return Dup(ExePath(), icon); }

    IFACEMETHODIMP GetToolTip(IShellItemArray*, LPWSTR* tip)
    {
        return Dup(L"\u628a\u9009\u4e2d\u7684\u9879\u79fb\u5165\u4e00\u4e2a\u65b0\u5efa\u7684\u6587\u4ef6\u5939", tip);
    }

    IFACEMETHODIMP GetCanonicalName(GUID* guid) { *guid = CLSID_NFWS_SHIM; return S_OK; }

    IFACEMETHODIMP GetState(IShellItemArray* items, BOOL, EXPCMDSTATE* state)
    {
        // 只有拿得到文件系统路径时才出现，避免在“此电脑”等纯虚节点上给出会失败的入口
        std::vector<std::wstring> paths = CollectPaths(items);
        *state = paths.empty() ? ECS_HIDDEN : ECS_ENABLED;
        return S_OK;
    }

    IFACEMETHODIMP Invoke(IShellItemArray* items, IBindCtx*)
    {
        std::vector<std::wstring> paths = CollectPaths(items);
        if (paths.empty()) return E_INVALIDARG;

        std::wstring cmd = L"\"";
        cmd += ExePath();
        cmd += L"\"";
        for (size_t i = 0; i < paths.size(); i++)
        {
            cmd += L" \"";
            cmd += paths[i];
            cmd += L"\"";
        }

        STARTUPINFOW si = { sizeof(si) };
        PROCESS_INFORMATION pi = { 0 };
        std::vector<wchar_t> buf(cmd.begin(), cmd.end());
        buf.push_back(L'\0');
        if (!CreateProcessW(NULL, &buf[0], NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi))
            return HRESULT_FROM_WIN32(GetLastError());
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
        return S_OK;
    }

    IFACEMETHODIMP GetFlags(EXPCMDFLAGS* flags) { *flags = ECF_DEFAULT; return S_OK; }

    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** e)
    {
        *e = NULL;
        return E_NOTIMPL;
    }
};

class ShimFactory : public IClassFactory
{
    LONG _ref;
public:
    ShimFactory() : _ref(1) {}

    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IClassFactory)
        {
            *ppv = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&_ref); }
    IFACEMETHODIMP_(ULONG) Release()
    {
        ULONG c = InterlockedDecrement(&_ref);
        if (c == 0) delete this;
        return c;
    }
    IFACEMETHODIMP CreateInstance(IUnknown* outer, REFIID riid, void** ppv)
    {
        if (outer) return CLASS_E_NOAGGREGATION;
        ShimCommand* c = new (std::nothrow) ShimCommand();
        if (!c) return E_OUTOFMEMORY;
        HRESULT hr = c->QueryInterface(riid, ppv);
        c->Release();
        return hr;
    }
    IFACEMETHODIMP LockServer(BOOL) { return S_OK; }
};

STDAPI DllGetClassObject(REFCLSID clsid, REFIID riid, void** ppv)
{
    if (!IsEqualCLSID(clsid, CLSID_NFWS_SHIM)) return CLASS_E_CLASSNOTAVAILABLE;
    ShimFactory* f = new (std::nothrow) ShimFactory();
    if (!f) return E_OUTOFMEMORY;
    HRESULT hr = f->QueryInterface(riid, ppv);
    f->Release();
    return hr;
}

STDAPI DllCanUnloadNow() { return S_FALSE; }

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_self = hInstance;
        DisableThreadLibraryCalls(hInstance);
    }
    return TRUE;
}
