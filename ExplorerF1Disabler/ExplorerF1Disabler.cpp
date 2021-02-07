#include "framework.h"
#include "resource.h"

#include <shlobj.h>

#include <assert.h>
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/result.h>

#pragma comment(lib, "Version.lib")

#pragma comment(linker, "/manifestdependency:\"type='win32' \
    name='Microsoft.Windows.Common-Controls' \
    version='6.0.0.0' \
    processorArchitecture='*' \
    publicKeyToken='6595b64144ccf1df' \
    language='*'\"")

#define MAX_LOADSTRING 100
#define WM_NOTIFYICON (WM_USER + 100)

#define CATCH_SHOW_MSGBOX()                                                         \
    catch (const wil::ResultException &e)                                           \
    {                                                                               \
        wchar_t message[2048]{};                                                    \
        wil::GetFailureLogString(message, ARRAYSIZE(message), e.GetFailureInfo());  \
        MessageBox(hWnd, message, g_szTitle, MB_ICONERROR);                         \
    }


using namespace std::string_literals;
using namespace std::string_view_literals;

namespace my
{
    inline auto GetWindowThreadProcessId(HWND hwnd)
    {
        DWORD pid{};
        return std::make_tuple(::GetWindowThreadProcessId(hwnd, &pid), pid);
    }
    inline std::wstring GetClassName(HWND hwnd)
    {
        WCHAR className[256]{};
        unsigned len = ::GetClassName(hwnd, className, ARRAYSIZE(className));
        return { className, len };
    }
    inline std::wstring GetModuleFileName()
    {
        WCHAR fileName[256]{};
        unsigned len = ::GetModuleFileName(nullptr, fileName, ARRAYSIZE(fileName));
        return { fileName, len };
    }
} // namespace my

HINSTANCE g_hInst;
HHOOK g_hook;
WCHAR g_szTitle[MAX_LOADSTRING];
HWND g_hwnd;

LRESULT CALLBACK wndProc(HWND, UINT, WPARAM, LPARAM) noexcept;
INT_PTR CALLBACK about(HWND, UINT, WPARAM, LPARAM) noexcept;

constexpr auto NOTIFY_UID = 1;
constexpr auto szWindowClass = L"{B9EA841C-34B1-467B-8792-F7E2AE21D640}";

auto constexpr targetClassName{ L"CabinetWClass" };
// auto constexpr targetClassName{ L"Notepad" };

struct WindowInfo
{
    bool isExplorer;
    std::wstring filename;
};

WindowInfo getWindowInfo(HWND hWnd)
{
    WindowInfo info{};
    auto className{ my::GetClassName(hWnd) };
    auto [tid, pid] = my::GetWindowThreadProcessId(hWnd);
    if (wcscmp(className.c_str(), targetClassName) == 0)
    {
        info.isExplorer = true;
    }
    else if (pid != 0)
    {
        wil::unique_handle hProcess{ OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, pid) };
        FILETIME creationTime, exitTime, kernelTime, userTime{};
        GetProcessTimes(hProcess.get(), &creationTime, &exitTime, &kernelTime, &userTime);
        WCHAR filename[MAX_PATH]{};
        GetProcessImageFileName(hProcess.get(), filename, ARRAYSIZE(filename));
        auto sep = wcsrchr(filename, L'\\');

        info.filename = sep ? sep + 1 : filename;
    }
    return info;
}

LRESULT CALLBACK lowLevelKeyboardProc(int code, WPARAM wParam, LPARAM lParam) noexcept
{
    if (code < 0)
        return CallNextHookEx(NULL, code, wParam, lParam);

    auto pKbdll = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);

    if (!(pKbdll->flags & LLKHF_LOWER_IL_INJECTED) && pKbdll->vkCode == VK_F1)
    {
        auto hWnd{ GetFocus() };
        if (hWnd == nullptr)
            hWnd = GetForegroundWindow();

        if (hWnd)
        {
            auto info{ getWindowInfo(hWnd) };
            if (info.isExplorer || lstrcmpiW(info.filename.c_str(), L"excel.exe") == 0)
                return 1;
        }
    }

    return CallNextHookEx(NULL, code, wParam, lParam);
}

void uninstallHook()
{
    if (g_hook)
    {
        UnhookWindowsHookEx(g_hook);
        g_hook = nullptr;
    }
}

void installHook()
{
    if (g_hook)
        uninstallHook();

    g_hook = { SetWindowsHookEx(WH_KEYBOARD_LL, &lowLevelKeyboardProc, nullptr, 0 /*global*/) };
    THROW_LAST_ERROR_IF_NULL(g_hook);
}

bool isWorking()
{
    return g_hook;
}

BOOL addNotifyIcon(HWND hWnd, unsigned int uID)
{
    NOTIFYICONDATA nid{ sizeof(nid) };

    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.hWnd = hWnd;
    nid.uID = uID;
    nid.uCallbackMessage = WM_NOTIFYICON;
    nid.hIcon = LoadIcon(g_hInst, MAKEINTRESOURCE(IDI_SMALL));
    wcscpy_s(nid.szTip, g_szTitle);

    return Shell_NotifyIcon(NIM_ADD, &nid);
}

void deleteNotifyIcon(HWND hWnd, unsigned int uID)
{
    NOTIFYICONDATA nid = {};

    nid.cbSize = sizeof(nid);
    nid.hWnd = hWnd;
    nid.uID = uID;

    Shell_NotifyIcon(NIM_DELETE, &nid);
}

ATOM registerMyClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex{
        sizeof(WNDCLASSEX),
        CS_HREDRAW | CS_VREDRAW,
        &wndProc,
    };

    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));
    wcex.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));
    wcex.lpszClassName = szWindowClass;

    return RegisterClassExW(&wcex);
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine,
                      _In_ int nCmdShow)
{
    SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32);

    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);
    UNREFERENCED_PARAMETER(nCmdShow);

    g_hInst = hInstance;

    LoadStringW(hInstance, IDS_APP_TITLE, g_szTitle, MAX_LOADSTRING);

    wcscat_s(g_szTitle,
#if _M_X64
             L" (64bit)"
#else
             L" (32bit)"
#endif
    );

    wil::unique_mutex m{ CreateMutex(nullptr, FALSE, g_szTitle) };
    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        ::MessageBox(nullptr, L"Another instance is already running.", g_szTitle, MB_ICONEXCLAMATION);
        return 0;
    }
    auto hr{ CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE) };

    /*

    SHELLEXECUTEINFO sei{
        sizeof(SHELLEXECUTEINFO),
        SEE_MASK_WAITFORINPUTIDLE | SEE_MASK_NO_CONSOLE,
    };

    sei.lpFile = L"notepad.exe";
    sei.nShow = SW_SHOWNORMAL;
    ShellExecuteEx(&sei);
    //*/

    installHook();

    registerMyClass(hInstance);

    HWND hWnd = CreateWindowW(szWindowClass, g_szTitle, 0, CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr,
                              hInstance, nullptr);

    if (!hWnd)
    {
        return FALSE;
    }

    if (!addNotifyIcon(hWnd, NOTIFY_UID))
    {
        return FALSE;
    }

    MSG msg{};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    uninstallHook();

    return (int)msg.wParam;
}

void registerToShortcut(HWND hWnd)
try
{
    WCHAR targetPath[MAX_PATH]{};
    THROW_IF_FAILED(E_FAIL);
    THROW_IF_WIN32_BOOL_FALSE(SHGetSpecialFolderPath(hWnd, targetPath, CSIDL_STARTUP, FALSE));
    StringCchCat(targetPath, ARRAYSIZE(targetPath), (L"\\"s + g_szTitle + L".lnk"s).c_str());

    auto pShellLink{ wil::CoCreateInstance<IShellLink>(CLSID_ShellLink) };
    auto pPersistFile{ pShellLink.query<IPersistFile>() };

    THROW_IF_FAILED(pShellLink->SetPath(my::GetModuleFileName().c_str()));
    THROW_IF_FAILED(pPersistFile->Save(targetPath, TRUE));
}
CATCH_SHOW_MSGBOX()

LRESULT CALLBACK wndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    static UINT s_uTaskbarRestart;
    static HMENU s_menu = nullptr;

    switch (message)
    {
    case WM_CREATE:
        g_hwnd = hWnd;
        s_uTaskbarRestart = RegisterWindowMessage(L"TaskbarCreated");
        s_menu = LoadMenu(g_hInst, MAKEINTRESOURCE(IDR_MENU1));
        return DefWindowProc(hWnd, message, wParam, lParam);
    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case ID_ROOT_ABOUT:
            DialogBox(g_hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, &about);
            break;
        case ID_ROOT_REGISTERTOSTARTUPPROGRAM:
            registerToShortcut(hWnd);
            break;
        case ID_ROOT_EXIT:
            DestroyWindow(hWnd);
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    case WM_NOTIFYICON:
        switch (lParam)
        {
        case WM_RBUTTONDOWN:
        {
            POINT pt{};
            GetCursorPos(&pt);
            SetForegroundWindow(hWnd);
            TrackPopupMenu(GetSubMenu(s_menu, 0), TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, NULL);
        }
        break;
        default:
            ;
        }
        return 0;
    default:
        if (message == s_uTaskbarRestart)
            addNotifyIcon(hWnd, NOTIFY_UID);
        else
            return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

void getProductAndVersion(LPWSTR szCopyright, UINT uCopyrightLen, LPWSTR szProductVersion, UINT uProductVersionLen)
{
    WCHAR szFilename[MAX_PATH]{};
    GetModuleFileName(nullptr, szFilename, ARRAYSIZE(szFilename));

    DWORD dwHandle;
    DWORD dwSize = GetFileVersionInfoSize(szFilename, &dwHandle);
    if (dwSize == 0)
        std::abort();

    std::vector<BYTE> data(dwSize);
    if (!GetFileVersionInfo(szFilename, 0, dwSize, data.data()))
        std::abort();

    LPWSTR pvCopyright{}, pvProductVersion{};
    UINT iCopyrightLen{}, iProductVersionLen{};

    if (!VerQueryValue(data.data(), L"\\StringFileInfo\\040004b0\\LegalCopyright", (LPVOID*)&pvCopyright,
        &iCopyrightLen))
        std::abort();

    if (!VerQueryValue(data.data(), L"\\StringFileInfo\\040004b0\\ProductVersion", (LPVOID*)&pvProductVersion,
        &iProductVersionLen))
        std::abort();

    wcsncpy_s(szCopyright, uCopyrightLen, pvCopyright, iCopyrightLen);
    wcsncpy_s(szProductVersion, uProductVersionLen, pvProductVersion, iProductVersionLen);
}

INT_PTR CALLBACK about(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG: {
        WCHAR szCopyright[MAX_LOADSTRING]{};
        WCHAR szVersion[MAX_LOADSTRING]{};

        getProductAndVersion(szCopyright, ARRAYSIZE(szCopyright), szVersion, ARRAYSIZE(szVersion));

        SetWindowText(hDlg, g_szTitle);
        SetWindowText(GetDlgItem(hDlg, IDC_STATIC_COPYRIGHT), szCopyright);
        SetWindowText(GetDlgItem(hDlg, IDC_STATIC_VERSION), szVersion);

        return (INT_PTR)TRUE;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
