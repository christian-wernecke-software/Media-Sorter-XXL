#define _WIN32_WINNT 0x0600
#include <windows.h>
#include <gdiplus.h>
#include <winhttp.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <commctrl.h>
#include "resource.h"
#include <string>
#include <filesystem>
#include <thread>
#include <atomic>
#include <mutex>
#include <vector>
#include <map>
#include <set>
#include <sstream>
#include <iomanip>
#include <iostream>
#include <random>
#include <memory>
#include <queue>
#include <condition_variable>

// --- THREAD-SAFE QUEUE ---
template<typename T>
class SafeQueue {
private:
    std::queue<T> m_queue;
    std::mutex m_mutex;
    std::condition_variable m_cond;
    bool m_finished = false;

public:
    void push(T item) {
        {
            std::lock_guard<std::mutex> lock(m_mutex);
            m_queue.push(std::move(item));
        }
        m_cond.notify_one();
    }

    bool pop(T& item) {
        std::unique_lock<std::mutex> lock(m_mutex);
        m_cond.wait(lock, [this]() { return !m_queue.empty() || m_finished; });
        if (m_queue.empty()) return false;
        item = std::move(m_queue.front());
        m_queue.pop();
        return true;
    }

    void set_finished() {
        {
            std::lock_guard<std::mutex> lock(m_mutex);
            m_finished = true;
        }
        m_cond.notify_all();
    }

    size_t size() {
        std::lock_guard<std::mutex> lock(m_mutex);
        return m_queue.size();
    }
};

// Link against necessary libraries (MSVC directives)
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "msimg32.lib")

// Enable Visual Styles
#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

using namespace Gdiplus;
namespace fs = std::filesystem;

// GLOBAL STATE
std::wstring g_SourcePath;
std::wstring g_TargetPath;
std::atomic<bool> g_Running(false);
std::atomic<bool> g_StopRequested(false);
HWND g_hBtnStart = NULL;
HWND g_hBtnStop = NULL;
HWND g_hBtnBrowseSource = NULL;
HWND g_hBtnBrowseTarget = NULL;
HWND g_hBtnHelp = NULL;
HWND g_hEditSource = NULL;
HWND g_hEditTarget = NULL;
HWND g_hProgress = NULL;
HWND g_hStatus = NULL;
HWND g_hWnd = NULL;

// --- UI THEME COLORS (RGB) ---
const COLORREF CLR_BG_DARK      = RGB(5, 5, 8);       // Black
const COLORREF CLR_BG_LIGHTER   = RGB(15, 15, 20);    // Very Dark Gray
const COLORREF CLR_CARD_BG      = RGB(25, 25, 30);    // Card Background
const COLORREF CLR_ACCENT_ORANGE = RGB(255, 120, 0);   // Primary Accent
const COLORREF CLR_ACCENT_BLUE   = RGB(0, 135, 255);   // Secondary Accent/Complementary

const COLORREF CLR_TEXT_WHITE   = RGB(248, 248, 252);
const COLORREF CLR_TEXT_GRAY    = RGB(170, 170, 185);

const COLORREF CLR_EDIT_BG      = RGB(35, 35, 45);
const COLORREF CLR_STATUS_BG    = RGB(10, 10, 15);
const COLORREF CLR_PROGRESS_BG  = RGB(40, 40, 45);

// Button Palette
const COLORREF CLR_BTN_START_A  = CLR_ACCENT_ORANGE;
const COLORREF CLR_BTN_START_B  = RGB(220, 90, 0);
const COLORREF CLR_BTN_STOP_A   = RGB(80, 80, 90);
const COLORREF CLR_BTN_STOP_B   = RGB(60, 60, 70);
const COLORREF CLR_BTN_BROWSE_A = CLR_ACCENT_BLUE;
const COLORREF CLR_BTN_BROWSE_B = RGB(0, 100, 220);

// GDI handles
HBRUSH g_hBrushBg = NULL;
HBRUSH g_hBrushCard = NULL;
HBRUSH g_hBrushEdit = NULL;
HBRUSH g_hBrushStatus = NULL;
HFONT g_hFontHeader = NULL;
HFONT g_hFontTagline = NULL;
HFONT g_hFontLabel = NULL;
HFONT g_hFontButton = NULL;
HFONT g_hFontStatus = NULL;

// Stats
std::atomic<int> g_ProcessedCount(0);
std::atomic<int> g_SuccessCount(0);
std::atomic<int> g_SkippedCount(0);
std::atomic<int> g_TotalFiles(0);

// Cache for Geocoding (Lat,Lon -> City Name)
std::map<std::wstring, std::wstring> g_LocationCache;
std::mutex g_LocationCacheMutex;
std::mutex g_NetworkMutex; // Ensure 1 search at a time

// Initializer for GDI+
ULONG_PTR g_gdiplusToken;

// Helper for RAII GDI+ Image cleanup
struct GdiPlusImageDeleter {
    void operator()(Gdiplus::Image* image) {
        if (image) delete image;
    }
};

// --- UI HELPERS ---

void CreateThemeBrushes() {
    g_hBrushBg     = CreateSolidBrush(CLR_BG_DARK);
    g_hBrushCard   = CreateSolidBrush(CLR_CARD_BG);
    g_hBrushEdit   = CreateSolidBrush(CLR_EDIT_BG);
    g_hBrushStatus = CreateSolidBrush(CLR_STATUS_BG);
}

void CreateThemeFonts() {
    g_hFontHeader  = CreateFontW(28, 0, 0, 0, FW_BOLD,   FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
    g_hFontTagline = CreateFontW(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
    g_hFontLabel   = CreateFontW(15, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
    g_hFontButton  = CreateFontW(15, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
    g_hFontStatus  = CreateFontW(13, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
}

void DestroyThemeResources() {
    if (g_hBrushBg)     DeleteObject(g_hBrushBg);
    if (g_hBrushCard)   DeleteObject(g_hBrushCard);
    if (g_hBrushEdit)   DeleteObject(g_hBrushEdit);
    if (g_hBrushStatus) DeleteObject(g_hBrushStatus);
    if (g_hFontHeader)  DeleteObject(g_hFontHeader);
    if (g_hFontTagline) DeleteObject(g_hFontTagline);
    if (g_hFontLabel)   DeleteObject(g_hFontLabel);
    if (g_hFontButton)  DeleteObject(g_hFontButton);
    if (g_hFontStatus)  DeleteObject(g_hFontStatus);
}

void PaintGradientRect(HDC hdc, RECT rc, COLORREF clrTop, COLORREF clrBottom) {
    TRIVERTEX vert[2];
    GRADIENT_RECT gRect;
    vert[0].x = rc.left;  vert[0].y = rc.top;
    vert[0].Red   = (COLOR16)(GetRValue(clrTop) << 8);
    vert[0].Green = (COLOR16)(GetGValue(clrTop) << 8);
    vert[0].Blue  = (COLOR16)(GetBValue(clrTop) << 8);
    vert[0].Alpha = 0;
    vert[1].x = rc.right; vert[1].y = rc.bottom;
    vert[1].Red   = (COLOR16)(GetRValue(clrBottom) << 8);
    vert[1].Green = (COLOR16)(GetGValue(clrBottom) << 8);
    vert[1].Blue  = (COLOR16)(GetBValue(clrBottom) << 8);
    vert[1].Alpha = 0;
    gRect.UpperLeft = 0;
    gRect.LowerRight = 1;
    GradientFill(hdc, vert, 2, &gRect, 1, GRADIENT_FILL_RECT_V);
}

void DrawOwnerButton(LPDRAWITEMSTRUCT dis, COLORREF clrA, COLORREF clrB) {
    HDC hdc = dis->hDC;
    RECT rc = dis->rcItem;

    // Draw rounded rect with GDI+
    Gdiplus::Graphics g(hdc);
    g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);

    int r = 6; // corner radius
    Gdiplus::GraphicsPath path;
    path.AddArc(rc.left, rc.top, r*2, r*2, 180, 90);
    path.AddArc(rc.right - r*2 - 1, rc.top, r*2, r*2, 270, 90);
    path.AddArc(rc.right - r*2 - 1, rc.bottom - r*2 - 1, r*2, r*2, 0, 90);
    path.AddArc(rc.left, rc.bottom - r*2 - 1, r*2, r*2, 90, 90);
    path.CloseFigure();

    // Gradient fill
    Gdiplus::LinearGradientBrush brush(
        Gdiplus::Rect(rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top),
        Gdiplus::Color(255, GetRValue(clrA), GetGValue(clrA), GetBValue(clrA)),
        Gdiplus::Color(255, GetRValue(clrB), GetGValue(clrB), GetBValue(clrB)),
        Gdiplus::LinearGradientModeVertical);
    g.FillPath(&brush, &path);

    // Pressed state: slight offset
    if (dis->itemState & ODS_SELECTED) {
        Gdiplus::SolidBrush overlay(Gdiplus::Color(40, 0, 0, 0));
        g.FillPath(&overlay, &path);
    }

    // Text
    wchar_t text[128];
    GetWindowTextW(dis->hwndItem, text, 128);
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, CLR_TEXT_WHITE);
    HFONT oldFont = (HFONT)SelectObject(hdc, g_hFontButton);
    DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, oldFont);
}

// --- UTILITIES ---

std::wstring GetIniPath() {
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    std::wstring path = exePath;
    return path.substr(0, path.find_last_of(L".")) + L".ini";
}

void LoadSettings() {
    std::wstring ini = GetIniPath();
    wchar_t buf[MAX_PATH];
    GetPrivateProfileStringW(L"Settings", L"Source", L"", buf, MAX_PATH, ini.c_str());
    g_SourcePath = buf;
    GetPrivateProfileStringW(L"Settings", L"Target", L"", buf, MAX_PATH, ini.c_str());
    g_TargetPath = buf;
}

void SaveSettings() {
    std::wstring ini = GetIniPath();
    WritePrivateProfileStringW(L"Settings", L"Source", g_SourcePath.c_str(), ini.c_str());
    WritePrivateProfileStringW(L"Settings", L"Target", g_TargetPath.c_str(), ini.c_str());
}

void Log(const std::wstring& msg) {
    if (g_hStatus) {
        SetWindowTextW(g_hStatus, msg.c_str());
    }
}

// Generate random temp folder name
std::wstring GenerateTempSubfolderName() {
    static const char alphanum[] = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<> dis(0, sizeof(alphanum) - 2);
    std::wstring s = L"_temp_";
    for (int i = 0; i < 8; ++i) {
        s += (wchar_t)alphanum[dis(gen)];
    }
    return s;
}

// Execute command hidden
bool RunCommand(const std::wstring& cmd, const std::wstring& args) {
    SHELLEXECUTEINFOW sei = { sizeof(SHELLEXECUTEINFOW) };
    sei.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
    sei.lpVerb = L"open";
    sei.lpFile = cmd.c_str();
    sei.lpParameters = args.c_str();
    sei.nShow = SW_HIDE;
    
    if (ShellExecuteExW(&sei)) {
        WaitForSingleObject(sei.hProcess, INFINITE);
        DWORD exitCode = 0;
        GetExitCodeProcess(sei.hProcess, &exitCode);
        CloseHandle(sei.hProcess);
        return exitCode == 0;
    }
    return false;
}

// --- OPEN FOLDER DIALOG ---

static int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData) {
    if (uMsg == BFFM_INITIALIZED) {
        if (lpData != NULL) {
            SendMessageW(hwnd, BFFM_SETSELECTIONW, TRUE, lpData);
        }
    }
    return 0;
}

bool SelectFolder(HWND hWnd, std::wstring& path, const wchar_t* title) {
    BROWSEINFOW bi = { 0 };
    bi.hwndOwner = hWnd;
    bi.lpszTitle = title;
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
    bi.lpfn = BrowseCallbackProc;
    bi.lParam = (LPARAM)path.c_str();

    LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
    if (pidl != 0) {
        wchar_t buffer[MAX_PATH];
        if (SHGetPathFromIDListW(pidl, buffer)) {
            path = buffer;
            CoTaskMemFree(pidl);
            return true;
        }
        CoTaskMemFree(pidl);
    }
    return false;
}

// --- METADATA & IMAGE PROCESSING ---

// Helper to convert rational to double
double RationalToDouble(PropertyItem* item, int index = 0) {
    if (item->type != PropertyTagTypeRational) return 0.0;
    long* rational = (long*)item->value;
    long num = rational[index * 2];
    long den = rational[index * 2 + 1];
    if (den == 0) return 0.0;
    return (double)num / (double)den;
}

// Helper to get GPS coordinate
double GetGPSCoordinate(PropertyItem* itemInfo, PropertyItem* itemRef) {
    // Info has 3 rationals: deg, min, sec
    double deg = RationalToDouble(itemInfo, 0);
    double min = RationalToDouble(itemInfo, 1);
    double sec = RationalToDouble(itemInfo, 2);
    double result = deg + min / 60.0 + sec / 3600.0;

    // Ref is a string "N", "S", "E", "W"
    if (itemRef && itemRef->value) {
        char ref = ((char*)itemRef->value)[0];
        if (ref == 'S' || ref == 'W') {
            result *= -1.0;
        }
    }
    return result;
}

std::wstring ReverseGeocode(double lat, double lon) {
    // Limit precision to avoid hammering API
    std::wstringstream keyInfo;
    keyInfo << std::fixed << std::setprecision(3) << lat << L"_" << lon;
    std::wstring key = keyInfo.str();

    {
        std::lock_guard<std::mutex> lock(g_LocationCacheMutex);
        if (g_LocationCache.find(key) != g_LocationCache.end()) {
            return g_LocationCache[key];
        }
    }

    std::wstring result = L""; 

    // WinHTTP Request (Synchronized to 1 request at a time)
    {
        std::lock_guard<std::mutex> networkLock(g_NetworkMutex);
        
        HINTERNET hSession = WinHttpOpen(L"MediaSorter/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (hSession) {
            HINTERNET hConnect = WinHttpConnect(hSession, L"nominatim.openstreetmap.org", INTERNET_DEFAULT_HTTPS_PORT, 0);
            if (hConnect) {
                std::wstring path = L"/reverse?format=json&lat=" + std::to_wstring(lat) + L"&lon=" + std::to_wstring(lon) + L"&zoom=10"; 
                HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(), NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
                if (hRequest) {
                    if (WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)) {
                        if (WinHttpReceiveResponse(hRequest, NULL)) {
                            DWORD dwSize = 0;
                            DWORD dwDownloaded = 0;
                            std::string response;
                            do {
                                dwSize = 0;
                                if (!WinHttpQueryDataAvailable(hRequest, &dwSize)) break;
                                if (dwSize == 0) break;
                                char* tempBuffer = new char[dwSize + 1];
                                if (!tempBuffer) break;
                                if (WinHttpReadData(hRequest, tempBuffer, dwSize, &dwDownloaded)) {
                                    tempBuffer[dwDownloaded] = 0;
                                    response += tempBuffer;
                                }
                                delete[] tempBuffer;
                            } while (dwSize > 0);

                            auto extract = [&](const std::string& key) -> std::wstring {
                                std::string search = "\"" + key + "\":\"";
                                size_t pos = response.find(search);
                                if (pos != std::string::npos) {
                                    size_t end = response.find("\"", pos + search.length());
                                    if (end != std::string::npos) {
                                        std::string val = response.substr(pos + search.length(), end - (pos + search.length()));
                                        int len = MultiByteToWideChar(CP_UTF8, 0, val.c_str(), -1, NULL, 0);
                                        if (len > 0) {
                                            std::vector<wchar_t> wbuf(len);
                                            MultiByteToWideChar(CP_UTF8, 0, val.c_str(), -1, &wbuf[0], len);
                                            return std::wstring(&wbuf[0]);
                                        }
                                    }
                                }
                                return L"";
                            };

                            result = extract("city");
                            if (result.empty()) result = extract("town");
                            if (result.empty()) result = extract("village");
                            if (result.empty()) result = extract("municipality");
                        }
                    }
                    WinHttpCloseHandle(hRequest);
                }
                WinHttpCloseHandle(hConnect);
            }
            WinHttpCloseHandle(hSession);
        }
        
        // Respect Nominatim Usage Policy (max 1 req/sec)
        std::this_thread::sleep_for(std::chrono::milliseconds(1100));
    }

    {
        std::lock_guard<std::mutex> lock(g_LocationCacheMutex);
        g_LocationCache[key] = result;
    }
    return result;
}

struct FileMetadata {
    SYSTEMTIME date;
    bool hasDate = false;
    std::wstring location = L"";
};

FileMetadata GetFileMetadata(const std::wstring& path) {
    FileMetadata meta;
    memset(&meta.date, 0, sizeof(SYSTEMTIME));

    try {
        // Default to File Modification Time
        HANDLE hFile = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE) {
            FILETIME ftWrite;
            if (GetFileTime(hFile, NULL, NULL, &ftWrite)) {
                 FileTimeToSystemTime(&ftWrite, &meta.date);
            }
            CloseHandle(hFile);
        }

        // Try GDI+ for Images
        std::unique_ptr<Gdiplus::Image, GdiPlusImageDeleter> image(new Gdiplus::Image(path.c_str()));
        if (image && image->GetLastStatus() == Gdiplus::Ok) {
            // 1. Date (PropertyTagExifDTOrig = 0x9003)
            UINT size = image->GetPropertyItemSize(0x9003);
            if (size > 0) {
                PropertyItem* item = (PropertyItem*)malloc(size);
                if (item) {
                    if (image->GetPropertyItem(0x9003, size, item) == Gdiplus::Ok) {
                       char* dateStr = (char*)item->value;
                       if (dateStr && strlen(dateStr) >= 19) {
                           try {
                               meta.date.wYear = std::stoi(std::string(dateStr, 4));
                               meta.date.wMonth = std::stoi(std::string(dateStr + 5, 2));
                               meta.date.wDay = std::stoi(std::string(dateStr + 8, 2));
                               meta.date.wHour = std::stoi(std::string(dateStr + 11, 2));
                               meta.date.wMinute = std::stoi(std::string(dateStr + 14, 2));
                               meta.date.wSecond = std::stoi(std::string(dateStr + 17, 2));
                               meta.hasDate = true;
                           } catch (...) {}
                       }
                    }
                    free(item);
                }
            }

            // 2. GPS
            UINT latSize = image->GetPropertyItemSize(0x0002);
            UINT latRefSize = image->GetPropertyItemSize(0x0001);
            UINT lonSize = image->GetPropertyItemSize(0x0004);
            UINT lonRefSize = image->GetPropertyItemSize(0x0003);

            if (latSize > 0 && latRefSize > 0 && lonSize > 0 && lonRefSize > 0) {
                PropertyItem* latItem = (PropertyItem*)malloc(latSize);
                PropertyItem* latRefItem = (PropertyItem*)malloc(latRefSize);
                PropertyItem* lonItem = (PropertyItem*)malloc(lonSize);
                PropertyItem* lonRefItem = (PropertyItem*)malloc(lonRefSize);

                if (latItem && latRefItem && lonItem && lonRefItem) {
                    bool ok = (image->GetPropertyItem(0x0002, latSize, latItem) == Ok) &&
                              (image->GetPropertyItem(0x0001, latRefSize, latRefItem) == Ok) &&
                              (image->GetPropertyItem(0x0004, lonSize, lonItem) == Ok) &&
                              (image->GetPropertyItem(0x0003, lonRefSize, lonRefItem) == Ok);

                    if (ok) {
                        double lat = GetGPSCoordinate(latItem, latRefItem);
                        double lon = GetGPSCoordinate(lonItem, lonRefItem);
                        try {
                            meta.location = ReverseGeocode(lat, lon);
                        } catch (...) {}
                    }
                }
                if(latItem) free(latItem); 
                if(latRefItem) free(latRefItem); 
                if(lonItem) free(lonItem); 
                if(lonRefItem) free(lonRefItem);
            }
        }
    } catch (...) {
        // Log(L"Error reading metadata");
    }

    return meta;
}

// Forward declaration
void ProcessDirectory(const fs::path& dir);

void ProcessZip(const fs::path& zipPath) {
    try {
        std::wstring tempName = GenerateTempSubfolderName();
        fs::path tempDir = fs::path(g_TargetPath) / tempName;

        fs::create_directories(tempDir);
        Log((L"Extracting ZIP: " + zipPath.filename().wstring()).c_str());

        // Extract using tar
        // tar -xf "zipfile" -C "tempdir"
        std::wstring args = L"-xf \"" + zipPath.wstring() + L"\" -C \"" + tempDir.wstring() + L"\"";
        
        // Note: tar should be in path on Windows 10/11
        if (RunCommand(L"tar.exe", args)) {
             ProcessDirectory(tempDir);
        } else {
             Log(L"Failed to extract ZIP.");
        }

        // Cleanup
        std::error_code ec;
        fs::remove_all(tempDir, ec);

    } catch (...) {
        Log(L"ZIP Processing Error");
    }
}

void ProcessFile(const fs::path& filePath) {
    if (g_StopRequested) return;

    try {
        std::wstring filename = filePath.filename().wstring();
        std::wstring ext = filePath.extension().wstring();
        
        // Check for ZIP
        if (ext == L".zip" || ext == L".ZIP") {
            ProcessZip(filePath);
            return;
        }

        g_ProcessedCount++;
        SendMessage(g_hProgress, PBM_SETPOS, g_ProcessedCount, 0);
        Log((L"Processing: " + filename).c_str());

        FileMetadata meta = GetFileMetadata(filePath.wstring());

        // Build Target Path (V2)
        // Target/YYYY/YYYY-MM/
        std::wstringstream ssPath;
        ssPath << g_TargetPath << L"\\" << meta.date.wYear << L"\\" 
               << meta.date.wYear << L"-" << std::setw(2) << std::setfill(L'0') << meta.date.wMonth << L"\\";

        // Filename: YYYY-MM-DD HH-mm-ss [Location].ext
        std::wstringstream ssName;
        ssName << meta.date.wYear << L"-"
               << std::setw(2) << std::setfill(L'0') << meta.date.wMonth << L"-"
               << std::setw(2) << std::setfill(L'0') << meta.date.wDay << L" "
               << std::setw(2) << std::setfill(L'0') << meta.date.wHour << L"-"
               << std::setw(2) << std::setfill(L'0') << meta.date.wMinute << L"-"
               << std::setw(2) << std::setfill(L'0') << meta.date.wSecond;

        if (!meta.location.empty()) {
            ssName << L" " << meta.location;
        }

        std::wstring baseName = ssName.str();
        
        fs::path targetDir = ssPath.str();
        
        fs::create_directories(targetDir);

        fs::path targetFile = targetDir / (baseName + ext);
        int dup = 0;
        bool isDuplicate = false;

        while (fs::exists(targetFile)) {
             std::error_code ec;
             if (fs::file_size(targetFile, ec) == fs::file_size(filePath, ec)) {
                  g_SkippedCount++; 
                  isDuplicate = true;
                  break;
             }
             dup++;
             targetFile = targetDir / (baseName + L"_" + std::to_wstring(dup) + ext);
        }

        if (!isDuplicate) {
            fs::copy_file(filePath, targetFile);
            g_SuccessCount++;
        }
    } catch (const std::exception& e) {
        g_SkippedCount++;
        std::wstring err = L"Error: ";
        std::string what = e.what();
        err += std::wstring(what.begin(), what.end());
        Log(err);
    } catch (...) {
        g_SkippedCount++;
        Log(L"Unknown error processing file.");
    }
}

void ProcessDirectory(const fs::path& dir) {
    try {
        for (const auto& entry : fs::recursive_directory_iterator(dir)) {
            if (g_StopRequested) break;
            if (entry.is_regular_file()) {
                ProcessFile(entry.path());
            }
        }
    } catch (...) {
        // Directory access error
    }
}

// --- CUSTOM SUMMARY DIALOG ---

LRESULT CALLBACK SummaryWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE: {
        int y = 20;
        CreateWindowW(L"STATIC", L"Sorting Process Completed!", WS_VISIBLE | WS_CHILD | SS_LEFT, 20, y, 320, 30, hWnd, (HMENU)301, NULL, NULL);
        
        y += 45;
        int labelX = 30;
        int valueX = 240;
        int rowH = 24;

        auto addRow = [&](const wchar_t* label, int val, int id) {
            CreateWindowW(L"STATIC", label, WS_VISIBLE | WS_CHILD | SS_LEFT, labelX, y, 200, rowH, hWnd, (HMENU)(size_t)(400 + id), NULL, NULL);
            CreateWindowW(L"STATIC", std::to_wstring(val).c_str(), WS_VISIBLE | WS_CHILD | SS_RIGHT, valueX, y, 60, rowH, hWnd, (HMENU)(size_t)(500 + id), NULL, NULL);
            y += rowH;
        };

        addRow(L"\u2022 Total Files Found:", g_TotalFiles, 1);
        addRow(L"\u2022 Successfully Copied:", g_SuccessCount, 2);
        addRow(L"\u2022 Skipped (Duplicates):", g_SkippedCount, 3);
        addRow(L"\u2022 Processed Total:", g_ProcessedCount, 4);

        y += 20;
        CreateWindowW(L"STATIC", L"Your media is now organized and ready.", WS_VISIBLE | WS_CHILD | SS_LEFT, 20, y, 320, 20, hWnd, (HMENU)302, NULL, NULL);

        // Standard OK Button
        CreateWindowW(L"BUTTON", L"OK", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 125, y + 40, 110, 30, hWnd, (HMENU)IDOK, NULL, NULL);

        HFONT hFontBold = CreateFontW(22, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS,
            CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, VARIABLE_PITCH, L"Segoe UI");

        EnumChildWindows(hWnd, [](HWND hChild, LPARAM lParam) -> BOOL {
            HFONT hBold = (HFONT)lParam;
            int id = GetDlgCtrlID(hChild);
            if (id == 301) SendMessage(hChild, WM_SETFONT, (WPARAM)hBold, TRUE);
            else SendMessage(hChild, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);
            return TRUE;
        }, (LPARAM)hFontBold);
        break;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) DestroyWindow(hWnd);
        break;
    case WM_CTLCOLORSTATIC: {
        HDC hdcStatic = (HDC)wParam;
        SetBkMode(hdcStatic, TRANSPARENT);
        return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
    }
    default: return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

void ShowSummaryDialog(HWND hParent) {
    std::wstring className = L"SummaryDlgClass_" + std::to_wstring(GetTickCount());
    WNDCLASSW wc = { 0 };
    wc.lpfnWndProc = SummaryWndProc;
    wc.hInstance = GetModuleHandle(NULL);
    wc.lpszClassName = className.c_str();
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    RegisterClassW(&wc);

    int w = 360, h = 300;
    RECT pr; GetWindowRect(hParent, &pr);
    int x = pr.left + (pr.right - pr.left - w) / 2;
    int y = pr.top + (pr.bottom - pr.top - h) / 2;

    HWND hDlg = CreateWindowExW(WS_EX_DLGMODALFRAME, wc.lpszClassName, L"Process Complete",
        WS_POPUP | WS_CAPTION | WS_SYSMENU, x, y, w, h, hParent, NULL, wc.hInstance, NULL);

    EnableWindow(hParent, FALSE);
    ShowWindow(hDlg, SW_SHOW);
    
    MSG msg;
    while (IsWindow(hDlg) && GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    EnableWindow(hParent, TRUE);
    SetForegroundWindow(hParent);
    UnregisterClassW(className.c_str(), wc.hInstance);
}

void WorkerThread(SafeQueue<fs::path>& queue) {
    fs::path filePath;
    while (queue.pop(filePath)) {
        if (g_StopRequested) break;
        ProcessFile(filePath);
    }
}

void ScanningThread() {
    if (g_SourcePath.empty() || g_TargetPath.empty()) {
        MessageBoxW(g_hWnd, L"Please select Source and Target folders.", L"Error", MB_ICONERROR);
        g_Running = false;
        EnableWindow(g_hBtnStart, TRUE);
        EnableWindow(g_hBtnStop, FALSE);
        return;
    }

    // Reset stats
    g_ProcessedCount = 0;
    g_SuccessCount = 0;
    g_SkippedCount = 0;
    g_TotalFiles = 0;

    Log(L"Counting files...");
    
    std::vector<fs::path> rootFiles;
    try {
        for (const auto& entry : fs::recursive_directory_iterator(g_SourcePath)) {
            if (g_StopRequested) break;
            if (entry.is_regular_file()) {
                rootFiles.push_back(entry.path());
            }
        }
    } catch (...) {
        MessageBoxW(g_hWnd, L"Error reading source directory.", L"Error", MB_ICONERROR);
        g_Running = false;
        return;
    }

    if (rootFiles.empty()) {
        Log(L"No files found.");
        g_Running = false;
        EnableWindow(g_hBtnStart, TRUE);
        EnableWindow(g_hBtnStop, FALSE);
        return;
    }

    g_TotalFiles = (int)rootFiles.size();
    SendMessage(g_hProgress, PBM_SETRANGE, 0, MAKELPARAM(0, g_TotalFiles));
    SendMessage(g_hProgress, PBM_SETPOS, 0, 0);

    SafeQueue<fs::path> queue;
    int numThreads = std::thread::hardware_concurrency();
    if (numThreads < 1) numThreads = 2;
    if (numThreads > 8) numThreads = 8; // Don't overwhelm IO

    std::vector<std::thread> workers;
    for (int i = 0; i < numThreads; ++i) {
        workers.emplace_back(WorkerThread, std::ref(queue));
    }

    Log(L"Processing in parallel...");
    for (const auto& filePath : rootFiles) {
        if (g_StopRequested) break;
        queue.push(filePath);
    }
    queue.set_finished();

    for (auto& t : workers) {
        t.join();
    }

    Log(L"Finished.");
    ShowWindow(g_hProgress, SW_HIDE);

    ShowSummaryDialog(g_hWnd);

    g_Running = false;
    g_StopRequested = false;
    EnableWindow(g_hBtnStart, TRUE);
    EnableWindow(g_hBtnStop, FALSE);
}

// --- WINDOW PROCEDURE ---

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE: {
        CreateThemeBrushes();
        CreateThemeFonts();

        // --- Layout constants ---
        int margin = 20;
        int headerH = 70;
        int cardTop = headerH + 10;
        int cardLeft = margin;
        int cardW = 540 - 2 * margin;
        int innerLeft = cardLeft + 16;
        int editW = 340;
        int btnBrowseW = 40;
        int editLeft = innerLeft + 110;
        int btnBrowseLeft = editLeft + editW + 8;

        // Row 1: Source
        int row1Y = cardTop + 20;
        CreateWindowW(L"STATIC", L"Source Folder:", WS_VISIBLE | WS_CHILD | SS_LEFT, innerLeft, row1Y + 2, 105, 20, hWnd, (HMENU)201, NULL, NULL);
        g_hEditSource = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, editLeft, row1Y, editW, 24, hWnd, NULL, NULL, NULL);
        g_hBtnBrowseSource = CreateWindowW(L"BUTTON", L"\u2026", WS_VISIBLE | WS_CHILD | BS_OWNERDRAW, btnBrowseLeft, row1Y, btnBrowseW, 24, hWnd, (HMENU)101, NULL, NULL);

        // Row 2: Target
        int row2Y = row1Y + 38;
        CreateWindowW(L"STATIC", L"Target Folder:", WS_VISIBLE | WS_CHILD | SS_LEFT, innerLeft, row2Y + 2, 105, 20, hWnd, (HMENU)202, NULL, NULL);
        g_hEditTarget = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, editLeft, row2Y, editW, 24, hWnd, NULL, NULL, NULL);
        g_hBtnBrowseTarget = CreateWindowW(L"BUTTON", L"\u2026", WS_VISIBLE | WS_CHILD | BS_OWNERDRAW, btnBrowseLeft, row2Y, btnBrowseW, 24, hWnd, (HMENU)102, NULL, NULL);

        // Row 3: Buttons
        int row3Y = row2Y + 44;
        g_hBtnStart = CreateWindowW(L"BUTTON", L"\u25B6  Start Sorting", WS_VISIBLE | WS_CHILD | BS_OWNERDRAW, innerLeft, row3Y, 160, 34, hWnd, (HMENU)103, NULL, NULL);
        g_hBtnStop  = CreateWindowW(L"BUTTON", L"\u25A0  Stop", WS_VISIBLE | WS_CHILD | BS_OWNERDRAW | WS_DISABLED, innerLeft + 170, row3Y, 100, 34, hWnd, (HMENU)104, NULL, NULL);

        // Help button in header
        g_hBtnHelp = CreateWindowW(L"BUTTON", L"?", WS_VISIBLE | WS_CHILD | BS_OWNERDRAW, 540 - 45, 20, 30, 30, hWnd, (HMENU)105, NULL, NULL);

        // Progress bar (initially hidden)
        int progY = row3Y + 50;
        g_hProgress = CreateWindowW(PROGRESS_CLASSW, NULL, WS_CHILD, innerLeft, progY, cardW - 32, 14, hWnd, NULL, NULL, NULL);
        SendMessage(g_hProgress, PBM_SETBARCOLOR, 0, (LPARAM)CLR_ACCENT_ORANGE);
        SendMessage(g_hProgress, PBM_SETBKCOLOR, 0, (LPARAM)CLR_PROGRESS_BG);

        // Status bar at the very bottom
        g_hStatus = CreateWindowW(L"STATIC", L"Ready.", WS_VISIBLE | WS_CHILD | SS_LEFT, 20, 0, 560, 28, hWnd, (HMENU)205, NULL, NULL);

        // Apply fonts to all children
        EnumChildWindows(hWnd, [](HWND hChild, LPARAM lParam) -> BOOL {
            SendMessage(hChild, WM_SETFONT, (WPARAM)lParam, TRUE);
            return TRUE;
        }, (LPARAM)g_hFontLabel);

        // Override status font
        SendMessage(g_hStatus, WM_SETFONT, (WPARAM)g_hFontStatus, TRUE);

        LoadSettings();
        SetWindowTextW(g_hEditSource, g_SourcePath.c_str());
        SetWindowTextW(g_hEditTarget, g_TargetPath.c_str());
        break;
    }

    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        RECT rc;
        GetClientRect(hWnd, &rc);

        // Full background gradient
        PaintGradientRect(hdc, rc, CLR_BG_DARK, CLR_BG_LIGHTER);

        // Header banner (top 70px)
        RECT rcHeader = { 0, 0, rc.right, 70 };
        PaintGradientRect(hdc, rcHeader, RGB(20, 20, 38), RGB(35, 35, 58));

        // Global accent line below header
        HPEN hPen = CreatePen(PS_SOLID, 2, CLR_ACCENT_ORANGE);
        HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
        MoveToEx(hdc, 0, 70, NULL);
        LineTo(hdc, rc.right, 70);
        SelectObject(hdc, hOldPen);
        DeleteObject(hPen);

        // Header text
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, CLR_TEXT_WHITE);
        HFONT oldFont = (HFONT)SelectObject(hdc, g_hFontHeader);
        RECT rcTitle = { 20, 10, rc.right - 20, 48 };
        DrawTextW(hdc, L"Media Sorter XXL", -1, &rcTitle, DT_LEFT | DT_SINGLELINE);
        SelectObject(hdc, oldFont);

        // Tagline
        SetTextColor(hdc, CLR_TEXT_GRAY);
        oldFont = (HFONT)SelectObject(hdc, g_hFontTagline);
        RECT rcTagline = { 22, 44, rc.right - 20, 65 };
        DrawTextW(hdc, L"Organize your media. Automatically.", -1, &rcTagline, DT_LEFT | DT_SINGLELINE);
        SelectObject(hdc, oldFont);

        // Card background (rounded rect behind the controls)
        int cardTop = 90;
        int cardLeft = 20;
        int cardRight = rc.right - 20;
        
        // Calculate progress bar Y position for margin consistency
        // header(70) + gap(10) + row1(20) + row2(38) + row3(44) + gap(50) = 232 (if cardTop was 80)
        // Since cardTop is 90 here, we add the offset.
        int progY = cardTop + 20 + 38 + 44 + 50; 
        int cardBottom = progY + 14 + 16; 

        Gdiplus::Graphics g(hdc);
        g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
        int cr = 8;
        Gdiplus::GraphicsPath cardPath;
        cardPath.AddArc(cardLeft, cardTop, cr*2, cr*2, 180, 90);
        cardPath.AddArc(cardRight - cr*2, cardTop, cr*2, cr*2, 270, 90);
        cardPath.AddArc(cardRight - cr*2, cardBottom - cr*2, cr*2, cr*2, 0, 90);
        cardPath.AddArc(cardLeft, cardBottom - cr*2, cr*2, cr*2, 90, 90);
        cardPath.CloseFigure();
        Gdiplus::SolidBrush cardBrush(Gdiplus::Color(255, 25, 25, 30)); // Solid card bg
        g.FillPath(&cardBrush, &cardPath);

        // Card border (Orange accent)
        Gdiplus::Pen cardPen(Gdiplus::Color(100, 255, 120, 0), 1.0f);
        g.DrawPath(&cardPen, &cardPath);

        // Status bar background
        RECT rcStatus = { 0, rc.bottom - 28, rc.right, rc.bottom };
        HBRUSH hStatusBrush = CreateSolidBrush(CLR_STATUS_BG);
        FillRect(hdc, &rcStatus, hStatusBrush);
        DeleteObject(hStatusBrush);

        EndPaint(hWnd, &ps);
        break;
    }

    case WM_ERASEBKGND:
        return 1; // We handle painting in WM_PAINT

    case WM_CTLCOLORSTATIC: {
        HDC hdcStatic = (HDC)wParam;
        HWND hCtrl = (HWND)lParam;
        int ctrlId = GetDlgCtrlID(hCtrl);

        SetBkMode(hdcStatic, TRANSPARENT);

        if (ctrlId == 205) { // Status label
            SetTextColor(hdcStatic, CLR_ACCENT_ORANGE);
            return (LRESULT)g_hBrushStatus;
        }

        // Default labels
        SetTextColor(hdcStatic, CLR_TEXT_WHITE);
        return (LRESULT)GetStockObject(NULL_BRUSH);
    }

    case WM_CTLCOLOREDIT: {
        HDC hdcEdit = (HDC)wParam;
        SetTextColor(hdcEdit, CLR_TEXT_WHITE);
        SetBkColor(hdcEdit, CLR_EDIT_BG);
        return (LRESULT)g_hBrushEdit;
    }

    case WM_DRAWITEM: {
        LPDRAWITEMSTRUCT dis = (LPDRAWITEMSTRUCT)lParam;
        int id = (int)dis->CtlID;

        // First fill background behind button to avoid artifacts
        FillRect(dis->hDC, &dis->rcItem, g_hBrushBg);

        if (id == 103) { // Start
            DrawOwnerButton(dis, CLR_BTN_START_A, CLR_BTN_START_B);
        } else if (id == 104) { // Stop
            if (IsWindowEnabled(dis->hwndItem)) {
                DrawOwnerButton(dis, CLR_BTN_STOP_A, CLR_BTN_STOP_B);
            } else {
                DrawOwnerButton(dis, RGB(80, 80, 100), RGB(60, 60, 80));
            }
        } else if (id == 101 || id == 102) { // Browse
            DrawOwnerButton(dis, CLR_BTN_BROWSE_A, CLR_BTN_BROWSE_B);
        } else if (id == 105) { // Help
            DrawOwnerButton(dis, CLR_ACCENT_BLUE, RGB(0, 100, 200));
        }
        return TRUE;
    }

    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id == 101) { // Select Source
            if (SelectFolder(hWnd, g_SourcePath, L"Select Source Folder")) {
                SetWindowTextW(g_hEditSource, g_SourcePath.c_str());
                SaveSettings();
            }
        }
        else if (id == 102) { // Select Target
            if (SelectFolder(hWnd, g_TargetPath, L"Select Target Folder")) {
                SetWindowTextW(g_hEditTarget, g_TargetPath.c_str());
                SaveSettings();
            }
        }
        else if (id == 103) { // Start
            wchar_t buf[MAX_PATH];
            GetWindowTextW(g_hEditSource, buf, MAX_PATH);
            g_SourcePath = buf;
            GetWindowTextW(g_hEditTarget, buf, MAX_PATH);
            g_TargetPath = buf;

            if (g_SourcePath.empty() || g_TargetPath.empty()) {
                MessageBoxW(hWnd, L"Please select both Source and Target folders.", L"Media Sorter XXL", MB_ICONERROR);
                break;
            }

            // Check if paths exist and are directories
            try {
                if (!fs::exists(g_SourcePath) || !fs::is_directory(g_SourcePath)) {
                    MessageBoxW(hWnd, L"Source folder is invalid or does not exist.", L"Media Sorter XXL", MB_ICONERROR);
                    break;
                }
                if (!fs::exists(g_TargetPath) || !fs::is_directory(g_TargetPath)) {
                    MessageBoxW(hWnd, L"Target folder is invalid or does not exist.", L"Media Sorter XXL", MB_ICONERROR);
                    break;
                }
                if (fs::equivalent(g_SourcePath, g_TargetPath)) {
                    MessageBoxW(hWnd, L"Source and Target folders must not be identical.", L"Media Sorter XXL", MB_ICONERROR);
                    break;
                }
            } catch (const std::exception& e) {
                std::string what = e.what();
                std::wstring msg = L"Error while verifying folders: ";
                msg += std::wstring(what.begin(), what.end());
                MessageBoxW(hWnd, msg.c_str(), L"Media Sorter XXL", MB_ICONERROR);
                break;
            } catch (...) {
                MessageBoxW(hWnd, L"An unknown error occurred during folder verification.", L"Media Sorter XXL", MB_ICONERROR);
                break;
            }

            if (g_Running) break;
            g_Running = true;
            g_StopRequested = false;
            ShowWindow(g_hProgress, SW_SHOW);
            EnableWindow(g_hBtnStart, FALSE);
            EnableWindow(g_hBtnStop, TRUE);
            InvalidateRect(g_hBtnStart, NULL, TRUE);
            InvalidateRect(g_hBtnStop, NULL, TRUE);
            std::thread(ScanningThread).detach();
        }
        else if (id == 104) { // Stop
            if (g_Running) {
                g_StopRequested = true;
                Log(L"Stopping...");
            }
        } else if (id == 105) { // Help
            const wchar_t* helpText = L"How your files are organized and renamed:\n\n"
                L"1. Sorting into Folders:\n"
                L"Files are moved to the target folder into a date-based structure:\n"
                L"Target / [Year] / [Year-Month] /\n"
                L"Example: Target / 2023 / 2023-10 /\n\n"
                L"2. Renaming Files:\n"
                L"Each file is renamed using its creation date and location (if available):\n"
                L"Format: YYYY-MM-DD HH-mm-ss [Location].ext\n"
                L"Example: 2023-10-15 14-30-05 Paris.jpg\n\n"
                L"3. Duplicate Handling:\n"
                L"If a file with the same name exists, a suffix (_1, _2, etc.) is added.\n"
                L"Exact duplicates (same name and size) are skipped automatically.";
            
            MessageBoxW(hWnd, helpText, L"Quick Help - Media Sorter XXL", MB_OK | MB_ICONINFORMATION);
        }
        break;
    }

    case WM_SIZE: {
        int w = LOWORD(lParam);
        int h = HIWORD(lParam);
        if (g_hStatus) {
            SetWindowPos(g_hStatus, NULL, 20, h - 28, w - 40, 28, SWP_NOZORDER);
        }
        break;
    }

    case WM_DESTROY:
        SaveSettings();
        DestroyThemeResources();
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {
    // Initialize GDI+
    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&g_gdiplusToken, &gdiplusStartupInput, NULL);

    // Initialize Common Controls (for progress bar)
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_PROGRESS_CLASS;
    InitCommonControlsEx(&icex);

    WNDCLASSW wc = { 0 };
    wc.lpszClassName = L"MediaSorterXXLClass";
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hIcon = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_APP_ICON));
    wc.hbrBackground = NULL; // We paint our own background
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    RegisterClassW(&wc);

    // Center window on screen
    int winW = 580, winH = 360;
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    int posX = (screenW - winW) / 2;
    int posY = (screenH - winH) / 2;

    g_hWnd = CreateWindowW(L"MediaSorterXXLClass", L"Media Sorter XXL",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        posX, posY, winW, winH, NULL, NULL, hInstance, NULL);

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    GdiplusShutdown(g_gdiplusToken);
    return (int)msg.wParam;
}
