#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <shellapi.h>
#include <wrl.h>
#include <WebView2.h>

#include "resource.h"
#include "point_file_validation.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <string>
#include <vector>

using Microsoft::WRL::Callback;
using Microsoft::WRL::ComPtr;

namespace {

constexpr int ID_URL = 101;
constexpr int ID_GO = 102;
constexpr int ID_BACK = 103;
constexpr int ID_STATUS = 104;
constexpr UINT WM_DOWNLOAD_RESULT = WM_APP + 1;
HWND main_window = nullptr;
HWND url_edit = nullptr;
HWND status_label = nullptr;
HWND browser_host = nullptr;
ComPtr<ICoreWebView2Controller> controller;
ComPtr<ICoreWebView2> webview;
std::filesystem::path app_root;
std::filesystem::path staging;
std::filesystem::path inbox;
std::wstring initial_url;

std::wstring control_text(HWND control) {
    const int size = GetWindowTextLengthW(control);
    std::wstring result(static_cast<std::size_t>(std::max(size, 0)) + 1, L'\0');
    GetWindowTextW(control, result.data(), size + 1);
    result.resize(static_cast<std::size_t>(std::max(size, 0)));
    return result;
}

std::wstring lower(std::wstring value) {
    std::transform(value.begin(), value.end(), value.begin(), towlower);
    return value;
}

bool allowed_extension(const std::filesystem::path& path) {
    const std::wstring ext = lower(path.extension().wstring());
    return ext == L".csv" || ext == L".xlsx" || ext == L".xls" || ext == L".xlsm";
}

bool validate_download(const std::filesystem::path& path, std::wstring& reason) {
    return point_validation::validate(path, path.filename(), reason);
}

std::filesystem::path unique_path(const std::filesystem::path& directory,
                                  std::filesystem::path filename) {
    filename = filename.filename();
    if (filename.empty() || filename == L"." || filename == L"..") filename = L"download.xlsx";
    auto candidate = directory / filename;
    for (unsigned int number = 1; std::filesystem::exists(candidate); ++number) {
        candidate = directory /
            (filename.stem().wstring() + L" (" + std::to_wstring(number) + L")" +
             filename.extension().wstring());
    }
    return candidate;
}

void resize_controls(HWND window) {
    RECT client{};
    GetClientRect(window, &client);
    const int width = std::max(320L, client.right);
    MoveWindow(GetDlgItem(window, ID_BACK), 10, 10, 70, 30, TRUE);
    MoveWindow(url_edit, 90, 10, std::max(100, width - 200), 30, TRUE);
    MoveWindow(GetDlgItem(window, ID_GO), width - 100, 10, 90, 30, TRUE);
    MoveWindow(status_label, 10, 46, width - 20, 24, TRUE);
    MoveWindow(browser_host, 0, 74, width, std::max(100L, client.bottom - 74), TRUE);
    if (controller) {
        RECT host{};
        GetClientRect(browser_host, &host);
        controller->put_Bounds(host);
    }
}

void navigate() {
    if (!webview) return;
    std::wstring url = control_text(url_edit);
    if (url.rfind(L"https://", 0) != 0) {
        MessageBoxW(main_window, L"Point Browser Fetcher accepts HTTPS websites only.",
                    L"Secure navigation blocked", MB_OK | MB_ICONWARNING);
        return;
    }
    webview->Navigate(url.c_str());
}

void initialize_webview(HINSTANCE instance) {
    const std::filesystem::path profile = app_root / L"BrowserFetcher" / L"Profile";
    std::filesystem::create_directories(profile);
    const HRESULT started = CreateCoreWebView2EnvironmentWithOptions(
        nullptr, profile.c_str(), nullptr,
        Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
            [instance](HRESULT result, ICoreWebView2Environment* environment) -> HRESULT {
                if (FAILED(result) || !environment) {
                    MessageBoxW(main_window, L"Microsoft Edge WebView2 Runtime is required.",
                                L"Point Browser Fetcher", MB_OK | MB_ICONERROR);
                    return result;
                }
                return environment->CreateCoreWebView2Controller(
                    browser_host,
                    Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                        [](HRESULT controller_result, ICoreWebView2Controller* created) -> HRESULT {
                            if (FAILED(controller_result) || !created) return controller_result;
                            controller = created;
                            controller->get_CoreWebView2(&webview);
                            ComPtr<ICoreWebView2Settings> settings;
                            webview->get_Settings(&settings);
                            settings->put_AreDevToolsEnabled(FALSE);
                            settings->put_AreDefaultContextMenusEnabled(FALSE);
                            settings->put_IsStatusBarEnabled(TRUE);
                            settings->put_IsZoomControlEnabled(TRUE);
                            ComPtr<ICoreWebView2Settings4> settings4;
                            if (SUCCEEDED(settings.As(&settings4))) {
                                settings4->put_IsPasswordAutosaveEnabled(FALSE);
                                settings4->put_IsGeneralAutofillEnabled(FALSE);
                            }

                            EventRegistrationToken token{};
                            ComPtr<ICoreWebView2_4> webview4;
                            if (FAILED(webview.As(&webview4))) return E_NOINTERFACE;
                            EventRegistrationToken navigation_token{};
                            webview->add_NavigationStarting(
                                Callback<ICoreWebView2NavigationStartingEventHandler>(
                                    [](ICoreWebView2*, ICoreWebView2NavigationStartingEventArgs* args) -> HRESULT {
                                        LPWSTR uri_raw = nullptr;
                                        args->get_Uri(&uri_raw);
                                        const std::wstring uri = uri_raw ? uri_raw : L"";
                                        CoTaskMemFree(uri_raw);
                                        if (uri.rfind(L"https://", 0) != 0 && uri != L"about:blank") {
                                            args->put_Cancel(TRUE);
                                            SetWindowTextW(status_label,
                                                L"Blocked navigation: Point Browser Fetcher permits HTTPS only.");
                                        }
                                        return S_OK;
                                    }).Get(), &navigation_token);
                            webview4->add_DownloadStarting(
                                Callback<ICoreWebView2DownloadStartingEventHandler>(
                                    [](ICoreWebView2*, ICoreWebView2DownloadStartingEventArgs* args) -> HRESULT {
                                        LPWSTR proposed_raw = nullptr;
                                        args->get_ResultFilePath(&proposed_raw);
                                        std::filesystem::path proposed = proposed_raw ? proposed_raw : L"download";
                                        CoTaskMemFree(proposed_raw);
                                        if (!allowed_extension(proposed)) {
                                            args->put_Cancel(TRUE);
                                            SetWindowTextW(status_label, L"Blocked download: unsupported file type.");
                                            return S_OK;
                                        }
                                        const auto destination = unique_path(staging, proposed.filename());
                                        args->put_ResultFilePath(destination.c_str());
                                        args->put_Handled(TRUE);

                                        ComPtr<ICoreWebView2DownloadOperation> operation;
                                        args->get_DownloadOperation(&operation);
                                        LPWSTR source_raw = nullptr;
                                        operation->get_Uri(&source_raw);
                                        const std::wstring source = source_raw ? source_raw : L"";
                                        CoTaskMemFree(source_raw);
                                        if (source.rfind(L"https://", 0) != 0) {
                                            args->put_Cancel(TRUE);
                                            SetWindowTextW(status_label,
                                                L"Blocked download: the final download URL is not HTTPS.");
                                            return S_OK;
                                        }
                                        EventRegistrationToken state_token{};
                                        operation->add_StateChanged(
                                            Callback<ICoreWebView2StateChangedEventHandler>(
                                                [operation, destination](ICoreWebView2DownloadOperation*, IUnknown*) -> HRESULT {
                                                    COREWEBVIEW2_DOWNLOAD_STATE state{};
                                                    operation->get_State(&state);
                                                    if (state == COREWEBVIEW2_DOWNLOAD_STATE_COMPLETED) {
                                                        std::wstring reason;
                                                        if (validate_download(destination, reason)) {
                                                            std::error_code copy_error;
                                                            const auto final_path = inbox / destination.filename();
                                                            std::filesystem::copy_file(destination, final_path,
                                                                std::filesystem::copy_options::overwrite_existing,
                                                                copy_error);
                                                            if (!copy_error) {
                                                                std::error_code cleanup_error;
                                                                std::filesystem::remove(destination, cleanup_error);
                                                                const UINT notification = RegisterWindowMessageW(
                                                                    L"Point.InboxChanged.v1");
                                                                if (notification != 0)
                                                                    PostMessageW(HWND_BROADCAST, notification, 0, 0);
                                                            }
                                                            std::wstring message_text;
                                                            if (copy_error) {
                                                                message_text = L"Validated, but Inbox copy failed (error " +
                                                                    std::to_wstring(copy_error.value()) + L").";
                                                            } else {
                                                                message_text = L"Validated and added to Point Inbox: " +
                                                                    final_path.filename().wstring();
                                                            }
                                                            auto* message = new std::wstring(std::move(message_text));
                                                            PostMessageW(main_window, WM_DOWNLOAD_RESULT,
                                                                         copy_error ? 0 : 1,
                                                                         reinterpret_cast<LPARAM>(message));
                                                        } else {
                                                            std::error_code ignored;
                                                            std::filesystem::remove(destination, ignored);
                                                            auto* message = new std::wstring(L"Rejected and deleted: " + reason);
                                                            PostMessageW(main_window, WM_DOWNLOAD_RESULT, 0,
                                                                         reinterpret_cast<LPARAM>(message));
                                                        }
                                                    } else if (state == COREWEBVIEW2_DOWNLOAD_STATE_INTERRUPTED) {
                                                        std::error_code ignored;
                                                        std::filesystem::remove(destination, ignored);
                                                        auto* message = new std::wstring(L"Download was interrupted; partial file deleted.");
                                                        PostMessageW(main_window, WM_DOWNLOAD_RESULT, 0,
                                                                     reinterpret_cast<LPARAM>(message));
                                                    }
                                                    return S_OK;
                                                }).Get(), &state_token);
                                        SetWindowTextW(status_label, L"Downloading to secure staging area...");
                                        return S_OK;
                                    }).Get(), &token);
                            resize_controls(main_window);
                            SetWindowTextW(status_label,
                                L"Open an HTTPS page. Sign in on the website if required, then download a CSV or Excel file.");
                            if (!initial_url.empty()) {
                                SetWindowTextW(url_edit, initial_url.c_str());
                                webview->Navigate(initial_url.c_str());
                            }
                            return S_OK;
                        }).Get());
            }).Get());
    if (FAILED(started)) {
        MessageBoxW(main_window, L"WebView2 initialization could not start.",
                    L"Point Browser Fetcher", MB_OK | MB_ICONERROR);
    }
}

LRESULT CALLBACK window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
    switch (message) {
    case WM_CREATE: {
        auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
        CreateWindowW(L"BUTTON", L"Back", WS_CHILD | WS_VISIBLE,
                      0, 0, 0, 0, window,
                      reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_BACK)),
                      create->hInstance, nullptr);
        url_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"https://",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 0, 0, 0, 0, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_URL)),
            create->hInstance, nullptr);
        CreateWindowW(L"BUTTON", L"Open", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                      0, 0, 0, 0, window,
                      reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_GO)),
                      create->hInstance, nullptr);
        status_label = CreateWindowW(L"STATIC", L"Starting secure browser...",
            WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_STATUS)),
            create->hInstance, nullptr);
        browser_host = CreateWindowW(L"STATIC", nullptr, WS_CHILD | WS_VISIBLE,
            0, 0, 0, 0, window, nullptr, create->hInstance, nullptr);
        initialize_webview(create->hInstance);
        return 0;
    }
    case WM_SIZE:
        resize_controls(window);
        return 0;
    case WM_COMMAND:
        if (LOWORD(wparam) == ID_GO) navigate();
        else if (LOWORD(wparam) == ID_BACK && webview) webview->GoBack();
        return 0;
    case WM_DOWNLOAD_RESULT: {
        std::unique_ptr<std::wstring> result(reinterpret_cast<std::wstring*>(lparam));
        SetWindowTextW(status_label, result->c_str());
        MessageBoxW(window, result->c_str(), L"Point Browser Fetcher",
                    MB_OK | (wparam ? MB_ICONINFORMATION : MB_ICONWARNING));
        return 0;
    }
    case WM_DESTROY:
        if (controller) controller->Close();
        webview.Reset();
        controller.Reset();
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(window, message, wparam, lparam);
}

}  // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int show) {
    if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) return 1;
    int argument_count = 0;
    LPWSTR* arguments = CommandLineToArgvW(GetCommandLineW(), &argument_count);
    if (arguments) {
        if (argument_count >= 2) {
            const std::wstring candidate = arguments[1];
            if (candidate.size() <= 8192 && candidate.rfind(L"https://", 0) == 0)
                initial_url = candidate;
        }
        LocalFree(arguments);
    }
    wchar_t executable[MAX_PATH]{};
    GetModuleFileNameW(nullptr, executable, MAX_PATH);
    app_root = std::filesystem::path(executable).parent_path();
    staging = app_root / L"BrowserFetcher" / L"Staging";
    inbox = app_root / L"Inbox";
    std::filesystem::create_directories(staging);
    std::filesystem::create_directories(inbox);

    const HICON icon = LoadIconW(instance, MAKEINTRESOURCEW(IDI_POINT));
    WNDCLASSEXW type{sizeof(type)};
    type.lpfnWndProc = window_proc;
    type.hInstance = instance;
    type.lpszClassName = L"PointBrowserFetcherWindow";
    type.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    type.hIcon = icon;
    type.hIconSm = icon;
    type.hbrBackground = reinterpret_cast<HBRUSH>(
        static_cast<INT_PTR>(COLOR_WINDOW + 1));
    if (!RegisterClassExW(&type)) return 1;
    main_window = CreateWindowExW(0, type.lpszClassName,
        L"Point Browser Fetcher — Secure Interactive Download",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 1100, 760,
        nullptr, nullptr, instance, nullptr);
    if (!main_window) return 1;
    ShowWindow(main_window, show);
    UpdateWindow(main_window);
    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
    CoUninitialize();
    return static_cast<int>(message.wParam);
}
