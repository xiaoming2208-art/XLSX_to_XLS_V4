#define UNICODE
#define _UNICODE
#include <windows.h>
#include <commdlg.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>

#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")

static const wchar_t* APP_TITLE = L"XLSX_to_XLS V4.0";
static const int ID_BTN_SELECT = 1001;
static const int ID_BTN_OPEN_OUTPUT = 1002;
static const int ID_LIST = 1003;
static const int ID_STATUS = 1004;

HWND g_list = nullptr;
HWND g_status = nullptr;
std::wstring g_lastOutput;

static std::wstring GetExeDir() {
    wchar_t path[MAX_PATH] = {};
    GetModuleFileNameW(nullptr, path, MAX_PATH);
    PathRemoveFileSpecW(path);
    return path;
}

static std::wstring Quote(const std::wstring& s) {
    return L"\"" + s + L"\"";
}

static void SetStatus(const std::wstring& s) {
    SetWindowTextW(g_status, s.c_str());
}

static void AppendList(const std::wstring& s) {
    SendMessageW(g_list, LB_ADDSTRING, 0, (LPARAM)s.c_str());
    int count = (int)SendMessageW(g_list, LB_GETCOUNT, 0, 0);
    if (count > 0) SendMessageW(g_list, LB_SETTOPINDEX, count - 1, 0);
}

static bool FileExists(const std::wstring& p) {
    DWORD a = GetFileAttributesW(p.c_str());
    return a != INVALID_FILE_ATTRIBUTES && !(a & FILE_ATTRIBUTE_DIRECTORY);
}

static bool DirExists(const std::wstring& p) {
    DWORD a = GetFileAttributesW(p.c_str());
    return a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_DIRECTORY);
}

static bool EnsureDir(const std::wstring& p) {
    if (DirExists(p)) return true;
    return CreateDirectoryW(p.c_str(), nullptr) != 0 || GetLastError() == ERROR_ALREADY_EXISTS;
}

static std::wstring ParentDir(const std::wstring& p) {
    wchar_t buf[MAX_PATH * 4] = {};
    wcsncpy_s(buf, p.c_str(), _TRUNCATE);
    PathRemoveFileSpecW(buf);
    return buf;
}

static std::wstring BaseNameNoExt(const std::wstring& p) {
    wchar_t name[MAX_PATH] = {};
    wcsncpy_s(name, PathFindFileNameW(p.c_str()), _TRUNCATE);
    PathRemoveExtensionW(name);
    return name;
}

static bool HasXlsxExt(const std::wstring& p) {
    const wchar_t* ext = PathFindExtensionW(p.c_str());
    return ext && _wcsicmp(ext, L".xlsx") == 0;
}

static bool IsOleBiff8(const std::wstring& p) {
    std::ifstream f(p, std::ios::binary);
    if (!f) return false;
    unsigned char b[8] = {};
    f.read((char*)b, 8);
    const unsigned char sig[8] = {0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1};
    return f.gcount() == 8 && memcmp(b, sig, 8) == 0;
}

static std::wstring FindLibreOfficeExe() {
    std::wstring exeDir = GetExeDir();
    std::vector<std::wstring> candidates = {
        exeDir + L"\\LibreOffice\\program\\soffice.exe",
        exeDir + L"\\program\\soffice.exe",
        L"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        L"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"
    };
    for (const auto& p : candidates) if (FileExists(p)) return p;
    return L"";
}

static bool RunLibreOfficeConvert(const std::wstring& soffice,
                                  const std::wstring& xlsx,
                                  const std::wstring& outdir,
                                  DWORD& exitCode) {
    std::wstring profile = GetExeDir() + L"\\LO_Profile";
    EnsureDir(profile);

    std::wstring profileArg = L"-env:UserInstallation=file:///" + profile;
    for (auto& c : profileArg) if (c == L'\\') c = L'/';

    std::wstring cmd =
        Quote(soffice) +
        L" --headless --nologo --nodefault --nolockcheck --nofirststartwizard "
        + Quote(profileArg) +
        L" --convert-to " + Quote(L"xls:MS Excel 97") +
        L" --outdir " + Quote(outdir) +
        L" " + Quote(xlsx);

    STARTUPINFOW si{};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi{};
    std::vector<wchar_t> mutableCmd(cmd.begin(), cmd.end());
    mutableCmd.push_back(L'\0');

    BOOL ok = CreateProcessW(
        nullptr, mutableCmd.data(),
        nullptr, nullptr, FALSE,
        CREATE_NO_WINDOW,
        nullptr, nullptr, &si, &pi
    );
    if (!ok) return false;

    WaitForSingleObject(pi.hProcess, INFINITE);
    exitCode = 999;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return true;
}

static std::wstring UniqueOutputPath(const std::wstring& outdir, const std::wstring& base) {
    std::wstring first = outdir + L"\\" + base + L".xls";
    if (!FileExists(first)) return first;
    for (int i = 1; i < 10000; ++i) {
        std::wstring p = outdir + L"\\" + base + L"_converted_" + std::to_wstring(i) + L".xls";
        if (!FileExists(p)) return p;
    }
    return first;
}

static bool MoveGeneratedIfNeeded(const std::wstring& expectedGenerated,
                                  const std::wstring& desired) {
    if (!FileExists(expectedGenerated)) return false;
    if (_wcsicmp(expectedGenerated.c_str(), desired.c_str()) == 0) return true;
    return MoveFileExW(expectedGenerated.c_str(), desired.c_str(),
                       MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING) != 0;
}

static void ConvertFiles(const std::vector<std::wstring>& files) {
    std::wstring soffice = FindLibreOfficeExe();
    if (soffice.empty()) {
        MessageBoxW(nullptr,
            L"找不到 LibreOffice 轉換引擎。\n\n"
            L"正式 Portable Release 應包含：\n"
            L"LibreOffice\\program\\soffice.exe\n\n"
            L"請使用 GitHub Actions 產生的正式 Release Artifact。",
            APP_TITLE, MB_ICONERROR);
        return;
    }

    int okCount = 0, failCount = 0;
    for (size_t i = 0; i < files.size(); ++i) {
        const std::wstring& src = files[i];
        if (!HasXlsxExt(src)) continue;

        std::wstring srcDir = ParentDir(src);
        std::wstring outdir = srcDir + L"\\XLS_轉換完成";
        if (!EnsureDir(outdir)) {
            AppendList(L"失敗：無法建立輸出資料夾 → " + src);
            ++failCount;
            continue;
        }

        g_lastOutput = outdir;
        std::wstring base = BaseNameNoExt(src);
        std::wstring generated = outdir + L"\\" + base + L".xls";
        std::wstring desired = UniqueOutputPath(outdir, base);

        // LibreOffice writes base.xls. If it already exists, temporarily rename it.
        std::wstring backup;
        if (FileExists(generated)) {
            backup = generated + L".v4bak";
            DeleteFileW(backup.c_str());
            MoveFileExW(generated.c_str(), backup.c_str(), MOVEFILE_REPLACE_EXISTING);
        }

        SetStatus(L"轉換中 " + std::to_wstring(i + 1) + L"/" +
                  std::to_wstring(files.size()) + L"：" + PathFindFileNameW(src.c_str()));

        DWORD ec = 999;
        bool launched = RunLibreOfficeConvert(soffice, src, outdir, ec);
        bool produced = launched && FileExists(generated) && ec == 0;

        if (produced) {
            if (_wcsicmp(generated.c_str(), desired.c_str()) != 0) {
                MoveGeneratedIfNeeded(generated, desired);
            }
            if (!backup.empty() && FileExists(backup)) {
                MoveFileExW(backup.c_str(), generated.c_str(), MOVEFILE_REPLACE_EXISTING);
            }

            if (FileExists(desired) && IsOleBiff8(desired)) {
                AppendList(L"完成：" + std::wstring(PathFindFileNameW(src.c_str())) +
                           L" → " + std::wstring(PathFindFileNameW(desired.c_str())));
                ++okCount;
            } else {
                AppendList(L"失敗：輸出不是有效 OLE/BIFF8 XLS → " + src);
                ++failCount;
            }
        } else {
            if (!backup.empty() && FileExists(backup)) {
                MoveFileExW(backup.c_str(), generated.c_str(), MOVEFILE_REPLACE_EXISTING);
            }
            AppendList(L"失敗：LibreOffice 轉換失敗，ExitCode=" + std::to_wstring(ec) + L" → " + src);
            ++failCount;
        }
    }

    SetStatus(L"完成：成功 " + std::to_wstring(okCount) + L"，失敗 " + std::to_wstring(failCount));
    MessageBoxW(nullptr,
        (L"批次轉換完成。\n\n成功：" + std::to_wstring(okCount) +
         L"\n失敗：" + std::to_wstring(failCount)).c_str(),
        APP_TITLE, failCount ? MB_ICONWARNING : MB_ICONINFORMATION);
}

static std::vector<std::wstring> SelectXlsxFiles(HWND hwnd) {
    std::vector<std::wstring> result;
    std::vector<wchar_t> buf(65536, 0);

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFilter = L"Excel XLSX (*.xlsx)\0*.xlsx\0All Files (*.*)\0*.*\0";
    ofn.lpstrFile = buf.data();
    ofn.nMaxFile = (DWORD)buf.size();
    ofn.lpstrTitle = L"選擇要轉換的 XLSX（可多選）";
    ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT;

    if (!GetOpenFileNameW(&ofn)) return result;

    std::wstring first = buf.data();
    const wchar_t* p = buf.data() + first.size() + 1;

    if (*p == L'\0') {
        result.push_back(first);
        return result;
    }

    std::wstring dir = first;
    while (*p) {
        std::wstring name = p;
        result.push_back(dir + L"\\" + name);
        p += name.size() + 1;
    }
    return result;
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE: {
        HFONT font = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        HWND hTitle = CreateWindowW(L"STATIC",
            L"XLSX → XLS V4.0  Native Windows Portable",
            WS_CHILD | WS_VISIBLE,
            20, 18, 600, 28, hwnd, nullptr, nullptr, nullptr);
        SendMessageW(hTitle, WM_SETFONT, (WPARAM)font, TRUE);

        HWND hInfo = CreateWindowW(L"STATIC",
            L"免 Python｜免 .NET｜免 Excel COM｜產線端免 BUILD",
            WS_CHILD | WS_VISIBLE,
            20, 48, 650, 22, hwnd, nullptr, nullptr, nullptr);
        SendMessageW(hInfo, WM_SETFONT, (WPARAM)font, TRUE);

        HWND b1 = CreateWindowW(L"BUTTON", L"選擇 XLSX 並開始轉換",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            20, 82, 220, 38, hwnd, (HMENU)ID_BTN_SELECT, nullptr, nullptr);
        SendMessageW(b1, WM_SETFONT, (WPARAM)font, TRUE);

        HWND b2 = CreateWindowW(L"BUTTON", L"開啟最後輸出資料夾",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            252, 82, 180, 38, hwnd, (HMENU)ID_BTN_OPEN_OUTPUT, nullptr, nullptr);
        SendMessageW(b2, WM_SETFONT, (WPARAM)font, TRUE);

        g_list = CreateWindowW(L"LISTBOX", nullptr,
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_NOINTEGRALHEIGHT,
            20, 135, 720, 315, hwnd, (HMENU)ID_LIST, nullptr, nullptr);
        SendMessageW(g_list, WM_SETFONT, (WPARAM)font, TRUE);

        g_status = CreateWindowW(L"STATIC", L"就緒",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            20, 465, 720, 24, hwnd, (HMENU)ID_STATUS, nullptr, nullptr);
        SendMessageW(g_status, WM_SETFONT, (WPARAM)font, TRUE);
        return 0;
    }
    case WM_COMMAND:
        if (LOWORD(wp) == ID_BTN_SELECT) {
            auto files = SelectXlsxFiles(hwnd);
            if (!files.empty()) ConvertFiles(files);
            return 0;
        }
        if (LOWORD(wp) == ID_BTN_OPEN_OUTPUT) {
            if (!g_lastOutput.empty() && DirExists(g_lastOutput)) {
                ShellExecuteW(hwnd, L"open", g_lastOutput.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            } else {
                MessageBoxW(hwnd, L"目前還沒有輸出資料夾。", APP_TITLE, MB_ICONINFORMATION);
            }
            return 0;
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, PWSTR, int nCmdShow) {
    const wchar_t CLASS_NAME[] = L"XLSXtoXLSV40NativeWindow";

    WNDCLASSW wc{};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(
        0, CLASS_NAME, APP_TITLE,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 780, 545,
        nullptr, nullptr, hInst, nullptr);

    if (!hwnd) return 1;
    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return (int)msg.wParam;
}
