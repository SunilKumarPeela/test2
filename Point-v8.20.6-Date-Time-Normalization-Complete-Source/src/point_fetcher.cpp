#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <winhttp.h>
#include <wincred.h>
#include <commctrl.h>
#include <shellapi.h>

#include "resource.h"
#include "point_file_validation.h"

#include <algorithm>
#include <cstdint>
#include <ctime>
#include <cwctype>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <memory>
#include <sstream>
#include <stdexcept>
#include <string>
#include <thread>
#include <vector>

namespace {

constexpr int ID_URL = 101;
constexpr int ID_FILENAME = 102;
constexpr int ID_AUTH = 103;
constexpr int ID_USER_HEADER = 104;
constexpr int ID_SECRET = 105;
constexpr int ID_INTERVAL = 106;
constexpr int ID_SAVE = 107;
constexpr int ID_REMOVE = 108;
constexpr int ID_RUN = 109;
constexpr int ID_LIST = 110;
constexpr int ID_STATUS = 111;
constexpr int ID_OPEN_STAGING = 112;
constexpr UINT_PTR TIMER_ID = 1;
constexpr UINT WM_FETCH_COMPLETE = WM_APP + 1;

enum class AuthMode { Public = 0, Basic = 1, Bearer = 2, ApiKey = 3 };

struct DownloadSchedule {
    std::wstring url;
    std::wstring filename;
    AuthMode auth = AuthMode::Public;
    std::wstring user_or_header;
    std::wstring credential_target;
    int interval_hours = 1;
    std::int64_t next_run = 0;
};

struct FetchResult {
    std::size_t index = 0;
    bool success = false;
    bool browser_fallback = false;
    std::wstring fallback_url;
    std::wstring message;
};

HWND main_window = nullptr;
HWND url_edit = nullptr;
HWND filename_edit = nullptr;
HWND auth_combo = nullptr;
HWND user_header_edit = nullptr;
HWND secret_edit = nullptr;
HWND interval_combo = nullptr;
HWND schedule_list = nullptr;
HWND status_text = nullptr;
std::filesystem::path root;
std::filesystem::path staging;
std::filesystem::path inbox;
std::vector<DownloadSchedule> schedules;
std::thread fetch_thread;
bool fetch_running = false;

std::filesystem::path publish_to_inbox(
        const std::filesystem::path& validated_file,
        const std::wstring& filename) {
    // Publish only a complete validated file. Copying to a temporary Inbox
    // name followed by an atomic rename prevents Point from indexing partial
    // content while a fetch is still being committed.
    std::filesystem::create_directories(inbox);
    const auto inbox_temporary =
        inbox / (L".point-fetch-" + filename + L".tmp");
    const auto inbox_destination = inbox / filename;
    std::error_code publish_error;
    std::filesystem::copy_file(
        validated_file, inbox_temporary,
        std::filesystem::copy_options::overwrite_existing, publish_error);
    if (publish_error) {
        std::filesystem::remove(inbox_temporary, publish_error);
        throw std::runtime_error(
            "Validated download could not be copied to Point Inbox");
    }
    if (!MoveFileExW(
            inbox_temporary.c_str(), inbox_destination.c_str(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        std::filesystem::remove(inbox_temporary, publish_error);
        throw std::runtime_error(
            "Validated download could not be activated in Point Inbox. "
            "Close the existing workbook in Excel and try again");
    }
    const UINT notification =
        RegisterWindowMessageW(L"Point.InboxChanged.v1");
    if (notification != 0)
        PostMessageW(HWND_BROADCAST, notification, 0, 0);
    return inbox_destination;
}

std::wstring text(HWND control) {
    const int length = GetWindowTextLengthW(control);
    std::wstring value(static_cast<std::size_t>(std::max(0, length)) + 1, L'\0');
    GetWindowTextW(control, value.data(), length + 1);
    value.resize(static_cast<std::size_t>(std::max(0, length)));
    return value;
}

std::int64_t now_epoch() {
    return static_cast<std::int64_t>(std::time(nullptr));
}

std::wstring next_text(std::int64_t epoch) {
    const std::time_t value = static_cast<std::time_t>(epoch);
    std::tm local{};
    if (localtime_s(&local, &value) != 0) return L"Unknown";
    wchar_t buffer[64]{};
    if (!wcsftime(buffer, 64, L"%Y-%m-%d %I:%M %p", &local)) return L"Unknown";
    return buffer;
}

std::wstring auth_name(AuthMode mode) {
    switch (mode) {
    case AuthMode::Basic: return L"Basic username/password";
    case AuthMode::Bearer: return L"Bearer token";
    case AuthMode::ApiKey: return L"API key";
    default: return L"Public HTTPS";
    }
}

std::wstring credential_id(const std::wstring& url, const std::wstring& name) {
    std::uint64_t hash = 1469598103934665603ULL;
    for (const wchar_t ch : url + L"|" + name) {
        hash ^= static_cast<std::uint16_t>(ch);
        hash *= 1099511628211ULL;
    }
    std::wstringstream value;
    value << L"PointFetcher:" << std::hex << hash;
    return value.str();
}

void save_credential(
        const std::wstring& target, const std::wstring& username,
        const std::wstring& secret) {
    if (secret.empty()) throw std::runtime_error("Authentication secret is required");
    CREDENTIALW credential{};
    credential.Type = CRED_TYPE_GENERIC;
    credential.TargetName = const_cast<LPWSTR>(target.c_str());
    credential.UserName = const_cast<LPWSTR>(username.c_str());
    credential.CredentialBlobSize = static_cast<DWORD>(
        secret.size() * sizeof(wchar_t));
    credential.CredentialBlob = reinterpret_cast<LPBYTE>(
        const_cast<wchar_t*>(secret.data()));
    credential.Persist = CRED_PERSIST_LOCAL_MACHINE;
    if (!CredWriteW(&credential, 0))
        throw std::runtime_error("Windows Credential Manager rejected the secret");
}

std::pair<std::wstring, std::wstring> read_credential(
        const std::wstring& target) {
    PCREDENTIALW credential = nullptr;
    if (!CredReadW(target.c_str(), CRED_TYPE_GENERIC, 0, &credential) ||
        !credential)
        throw std::runtime_error("Saved credential was not found");
    std::unique_ptr<CREDENTIALW, decltype(&CredFree)> guard(credential, CredFree);
    const std::wstring username = credential->UserName
        ? credential->UserName : L"";
    const auto chars = credential->CredentialBlobSize / sizeof(wchar_t);
    const std::wstring secret(
        reinterpret_cast<wchar_t*>(credential->CredentialBlob), chars);
    return {username, secret};
}

bool supported_name(const std::wstring& name) {
    auto extension = std::filesystem::path(name).extension().wstring();
    std::transform(extension.begin(), extension.end(), extension.begin(),
        [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    return extension == L".csv" || extension == L".xlsx" ||
           extension == L".xls" || extension == L".xlsm";
}

bool valid_api_header(const std::wstring& value) {
    if (value.empty()) return false;
    return std::all_of(value.begin(), value.end(), [](wchar_t ch) {
        return iswalnum(ch) || ch == L'-';
    });
}

void save_config() {
    std::ofstream output(
        root / "point-fetcher.dat", std::ios::binary | std::ios::trunc);
    if (!output) throw std::runtime_error("Cannot save Point Fetcher configuration");
    const char signature[] = "POINT_FETCHER_V1";
    output.write(signature, sizeof(signature));
    const auto count = static_cast<std::uint32_t>(schedules.size());
    output.write(reinterpret_cast<const char*>(&count), sizeof(count));
    const auto write_text = [&](const std::wstring& value) {
        const auto length = static_cast<std::uint32_t>(value.size());
        output.write(reinterpret_cast<const char*>(&length), sizeof(length));
        output.write(reinterpret_cast<const char*>(value.data()),
            static_cast<std::streamsize>(length * sizeof(wchar_t)));
    };
    for (const auto& schedule : schedules) {
        const int auth = static_cast<int>(schedule.auth);
        output.write(reinterpret_cast<const char*>(&auth), sizeof(auth));
        output.write(reinterpret_cast<const char*>(&schedule.interval_hours),
            sizeof(schedule.interval_hours));
        output.write(reinterpret_cast<const char*>(&schedule.next_run),
            sizeof(schedule.next_run));
        write_text(schedule.url);
        write_text(schedule.filename);
        write_text(schedule.user_or_header);
        write_text(schedule.credential_target);
    }
    if (!output) throw std::runtime_error("Could not complete Fetcher configuration");
}

void load_config() {
    std::ifstream input(root / "point-fetcher.dat", std::ios::binary);
    if (!input) return;
    char signature[17]{};
    input.read(signature, sizeof(signature));
    if (!input || std::string(signature) != "POINT_FETCHER_V1") return;
    std::uint32_t count = 0;
    input.read(reinterpret_cast<char*>(&count), sizeof(count));
    if (!input || count > 10'000) return;
    const auto read_text = [&](std::wstring& value) -> bool {
        std::uint32_t length = 0;
        input.read(reinterpret_cast<char*>(&length), sizeof(length));
        if (!input || length > 32'768) return false;
        value.resize(length);
        input.read(reinterpret_cast<char*>(value.data()),
            static_cast<std::streamsize>(length * sizeof(wchar_t)));
        return static_cast<bool>(input);
    };
    for (std::uint32_t index = 0; index < count; ++index) {
        DownloadSchedule schedule;
        int auth = 0;
        input.read(reinterpret_cast<char*>(&auth), sizeof(auth));
        input.read(reinterpret_cast<char*>(&schedule.interval_hours),
            sizeof(schedule.interval_hours));
        input.read(reinterpret_cast<char*>(&schedule.next_run),
            sizeof(schedule.next_run));
        if (!input || !read_text(schedule.url) ||
            !read_text(schedule.filename) ||
            !read_text(schedule.user_or_header) ||
            !read_text(schedule.credential_target))
            return;
        if (auth < 0 || auth > 3 || schedule.interval_hours < 1 ||
            schedule.interval_hours > 24 || !supported_name(schedule.filename))
            continue;
        schedule.auth = static_cast<AuthMode>(auth);
        schedules.push_back(schedule);
    }
}

void refresh_list() {
    SendMessageW(schedule_list, LB_RESETCONTENT, 0, 0);
    for (const auto& schedule : schedules) {
        std::wstringstream line;
        line << schedule.filename << L" | " << auth_name(schedule.auth)
             << L" | Every " << schedule.interval_hours
             << L"h | Next: " << next_text(schedule.next_run)
             << L" | " << schedule.url;
        const auto value = line.str();
        SendMessageW(schedule_list, LB_ADDSTRING, 0,
            reinterpret_cast<LPARAM>(value.c_str()));
    }
    if (schedules.empty())
        SendMessageW(schedule_list, LB_ADDSTRING, 0,
            reinterpret_cast<LPARAM>(L"No website/API downloads configured"));
}

std::wstring download(const DownloadSchedule& schedule) {
    URL_COMPONENTSW parts{sizeof(parts)};
    parts.dwSchemeLength = static_cast<DWORD>(-1);
    parts.dwHostNameLength = static_cast<DWORD>(-1);
    parts.dwUrlPathLength = static_cast<DWORD>(-1);
    parts.dwExtraInfoLength = static_cast<DWORD>(-1);
    if (!WinHttpCrackUrl(schedule.url.c_str(), 0, 0, &parts) ||
        parts.nScheme != INTERNET_SCHEME_HTTPS)
        throw std::runtime_error("Only valid HTTPS URLs are allowed");
    const std::wstring host(parts.lpszHostName, parts.dwHostNameLength);
    std::wstring resource(parts.lpszUrlPath, parts.dwUrlPathLength);
    if (parts.dwExtraInfoLength)
        resource.append(parts.lpszExtraInfo, parts.dwExtraInfoLength);
    if (resource.empty()) resource = L"/";

    HINTERNET session = WinHttpOpen(L"PointFetcher/1.0",
        WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);
    if (!session) throw std::runtime_error("WinHTTP session could not start");
    auto close_session = std::unique_ptr<void, decltype(&WinHttpCloseHandle)>(
        session, WinHttpCloseHandle);
    WinHttpSetTimeouts(session, 15000, 15000, 30000, 30000);
    HINTERNET connection = WinHttpConnect(
        session, host.c_str(), parts.nPort, 0);
    if (!connection) throw std::runtime_error("Could not connect to HTTPS host");
    auto close_connection = std::unique_ptr<void, decltype(&WinHttpCloseHandle)>(
        connection, WinHttpCloseHandle);
    HINTERNET request = WinHttpOpenRequest(connection, L"GET", resource.c_str(),
        nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
        WINHTTP_FLAG_SECURE);
    if (!request) throw std::runtime_error("Could not create HTTPS request");
    auto close_request = std::unique_ptr<void, decltype(&WinHttpCloseHandle)>(
        request, WinHttpCloseHandle);

    std::wstring headers;
    if (schedule.auth != AuthMode::Public) {
        const auto [username, secret] = read_credential(schedule.credential_target);
        if (schedule.auth == AuthMode::Basic) {
            if (!WinHttpSetCredentials(request, WINHTTP_AUTH_TARGET_SERVER,
                    WINHTTP_AUTH_SCHEME_BASIC, username.c_str(),
                    secret.c_str(), nullptr))
                throw std::runtime_error("Basic authentication could not be applied");
        } else if (schedule.auth == AuthMode::Bearer) {
            if (secret.find_first_of(L"\r\n") != std::wstring::npos)
                throw std::runtime_error("Bearer token contains invalid characters");
            headers = L"Authorization: Bearer " + secret + L"\r\n";
        } else {
            if (!valid_api_header(schedule.user_or_header))
                throw std::runtime_error("API header name is invalid");
            if (secret.find_first_of(L"\r\n") != std::wstring::npos)
                throw std::runtime_error("API key contains invalid characters");
            headers = schedule.user_or_header + L": " + secret + L"\r\n";
        }
    }
    if (!WinHttpSendRequest(request,
            headers.empty() ? WINHTTP_NO_ADDITIONAL_HEADERS : headers.c_str(),
            headers.empty() ? 0 : static_cast<DWORD>(-1),
            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
        !WinHttpReceiveResponse(request, nullptr))
        throw std::runtime_error("HTTPS request failed");
    DWORD status = 0;
    DWORD status_size = sizeof(status);
    WinHttpQueryHeaders(request,
        WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
        WINHTTP_HEADER_NAME_BY_INDEX, &status, &status_size,
        WINHTTP_NO_HEADER_INDEX);
    if (status < 200 || status >= 300)
        throw std::runtime_error("Server returned HTTP status " +
            std::to_string(status));

    const auto temporary = staging / (schedule.filename + L".download");
    std::ofstream output(temporary, std::ios::binary | std::ios::trunc);
    if (!output) throw std::runtime_error("Cannot create temporary download");
    std::uint64_t total = 0;
    for (;;) {
        DWORD available = 0;
        if (!WinHttpQueryDataAvailable(request, &available))
            throw std::runtime_error("Could not read HTTPS response");
        if (!available) break;
        std::vector<char> buffer(std::min<DWORD>(available, 64 * 1024));
        DWORD received = 0;
        if (!WinHttpReadData(request, buffer.data(),
                static_cast<DWORD>(buffer.size()), &received))
            throw std::runtime_error("HTTPS download was interrupted");
        total += received;
        if (total > 2ULL * 1024ULL * 1024ULL * 1024ULL)
            throw std::runtime_error("Download exceeds the 2 GiB safety limit");
        output.write(buffer.data(), received);
    }
    output.close();
    std::wstring validation_reason;
    if (!point_validation::validate(temporary, schedule.filename,
                                    validation_reason)) {
        std::error_code ignored;
        std::filesystem::remove(temporary, ignored);
        const int length = WideCharToMultiByte(CP_UTF8, 0,
            validation_reason.c_str(), -1, nullptr, 0, nullptr, nullptr);
        std::string reason(static_cast<std::size_t>(std::max(1, length)), '\0');
        WideCharToMultiByte(CP_UTF8, 0, validation_reason.c_str(), -1,
            reason.data(), length, nullptr, nullptr);
        if (!reason.empty() && reason.back() == '\0') reason.pop_back();
        throw std::runtime_error(reason);
    }
    const auto destination = staging / schedule.filename;
    if (!MoveFileExW(temporary.c_str(), destination.c_str(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
        throw std::runtime_error("Validated download could not be published");

    return publish_to_inbox(destination, schedule.filename).wstring();
}

void start_fetch(std::size_t index, bool allow_browser_fallback) {
    if (fetch_running || index >= schedules.size()) return;
    if (fetch_thread.joinable()) fetch_thread.join();
    fetch_running = true;
    EnableWindow(GetDlgItem(main_window, ID_SAVE), FALSE);
    EnableWindow(GetDlgItem(main_window, ID_REMOVE), FALSE);
    EnableWindow(GetDlgItem(main_window, ID_RUN), FALSE);
    SetWindowTextW(status_text, L"Downloading and validating...");
    fetch_thread = std::thread([index, allow_browser_fallback] {
        auto result = std::make_unique<FetchResult>();
        result->index = index;
        try {
            result->message =
                L"Validated and added to Point Imported Files: " +
                download(schedules[index]);
            result->success = true;
        } catch (const std::exception& ex) {
            const int size = MultiByteToWideChar(CP_UTF8, 0, ex.what(), -1,
                nullptr, 0);
            std::wstring message(static_cast<std::size_t>(std::max(1, size)), L'\0');
            MultiByteToWideChar(CP_UTF8, 0, ex.what(), -1, message.data(), size);
            if (!message.empty() && message.back() == L'\0') message.pop_back();
            result->message = L"Download failed: " + message;
            if (allow_browser_fallback &&
                message.find(L"website returned HTML") != std::wstring::npos) {
                result->browser_fallback = true;
                result->fallback_url = schedules[index].url;
            }
        }
        if (PostMessageW(main_window, WM_FETCH_COMPLETE, 0,
                reinterpret_cast<LPARAM>(result.get()))) result.release();
    });
}

void add_schedule() {
    DownloadSchedule schedule;
    schedule.url = text(url_edit);
    schedule.filename = text(filename_edit);
    schedule.user_or_header = text(user_header_edit);
    schedule.auth = static_cast<AuthMode>(SendMessageW(
        auth_combo, CB_GETCURSEL, 0, 0));
    const int interval_index = static_cast<int>(SendMessageW(
        interval_combo, CB_GETCURSEL, 0, 0));
    schedule.interval_hours = static_cast<int>(SendMessageW(
        interval_combo, CB_GETITEMDATA, interval_index, 0));
    if (schedule.url.rfind(L"https://", 0) != 0)
        throw std::runtime_error("Enter an HTTPS URL");
    if (schedule.filename.empty() ||
        std::filesystem::path(schedule.filename).filename().wstring() !=
            schedule.filename ||
        !supported_name(schedule.filename))
        throw std::runtime_error("Enter a safe .csv, .xlsx, .xls, or .xlsm filename");
    if (schedule.auth == AuthMode::Basic && schedule.user_or_header.empty())
        throw std::runtime_error("Basic authentication requires a username");
    if (schedule.auth == AuthMode::ApiKey &&
        !valid_api_header(schedule.user_or_header))
        throw std::runtime_error("Enter a valid API header such as X-API-Key");
    schedule.credential_target = credential_id(schedule.url, schedule.filename);
    if (schedule.auth != AuthMode::Public)
        save_credential(schedule.credential_target,
            schedule.auth == AuthMode::Basic ? schedule.user_or_header : L"",
            text(secret_edit));
    schedule.next_run = now_epoch() + schedule.interval_hours * 3600LL;
    schedules.push_back(schedule);
    save_config();
    refresh_list();
    SendMessageW(schedule_list, LB_SETCURSEL, schedules.size() - 1, 0);
    SetWindowTextW(secret_edit, L"");
    SetWindowTextW(status_text, L"Download schedule saved securely.");
}

void layout(HWND window) {
    RECT rect{};
    GetClientRect(window, &rect);
    const int width = rect.right;
    const int height = rect.bottom;
    MoveWindow(url_edit, 150, 20, std::max(240, width - 170), 26, TRUE);
    MoveWindow(filename_edit, 150, 54, 260, 26, TRUE);
    MoveWindow(auth_combo, 150, 88, 260, 180, TRUE);
    MoveWindow(user_header_edit, 580, 88, std::max(180, width - 600), 26, TRUE);
    MoveWindow(secret_edit, 150, 122, 260, 26, TRUE);
    MoveWindow(interval_combo, 580, 122, 190, 180, TRUE);
    MoveWindow(schedule_list, 20, 210, std::max(300, width - 40),
        std::max(140, height - 270), TRUE);
    MoveWindow(status_text, 20, height - 36, width - 40, 24, TRUE);
}

LRESULT CALLBACK window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
    switch (message) {
    case WM_CREATE: {
        CreateWindowW(L"STATIC", L"HTTPS URL", WS_CHILD | WS_VISIBLE,
            20, 23, 120, 22, window, nullptr, nullptr, nullptr);
        url_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 150, 20, 700, 26,
            window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_URL)), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Output filename", WS_CHILD | WS_VISIBLE,
            20, 57, 120, 22, window, nullptr, nullptr, nullptr);
        filename_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 150, 54, 260, 26,
            window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_FILENAME)), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Authentication", WS_CHILD | WS_VISIBLE,
            20, 91, 120, 22, window, nullptr, nullptr, nullptr);
        auth_combo = CreateWindowW(L"COMBOBOX", nullptr,
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, 150, 88, 260, 180,
            window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_AUTH)), nullptr, nullptr);
        const wchar_t* modes[] = {L"Public HTTPS", L"Basic username/password",
            L"Bearer token", L"API key"};
        for (const auto mode : modes) SendMessageW(auth_combo, CB_ADDSTRING, 0,
            reinterpret_cast<LPARAM>(mode));
        SendMessageW(auth_combo, CB_SETCURSEL, 0, 0);
        CreateWindowW(L"STATIC", L"Username / API header", WS_CHILD | WS_VISIBLE,
            425, 91, 150, 22, window, nullptr, nullptr, nullptr);
        user_header_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 580, 88, 270, 26,
            window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_USER_HEADER)), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Password / token / key", WS_CHILD | WS_VISIBLE,
            20, 125, 120, 22, window, nullptr, nullptr, nullptr);
        secret_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_PASSWORD, 150, 122,
            260, 26, window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SECRET)), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Download interval", WS_CHILD | WS_VISIBLE,
            425, 125, 150, 22, window, nullptr, nullptr, nullptr);
        interval_combo = CreateWindowW(L"COMBOBOX", nullptr,
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, 580, 122, 190, 180,
            window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_INTERVAL)), nullptr, nullptr);
        const int intervals[] = {1, 2, 3, 4, 6, 8, 12, 24};
        for (const int hours : intervals) {
            const auto label = L"Every " + std::to_wstring(hours) + L" hour(s)";
            const LRESULT index = SendMessageW(interval_combo, CB_ADDSTRING, 0,
                reinterpret_cast<LPARAM>(label.c_str()));
            SendMessageW(interval_combo, CB_SETITEMDATA, index, hours);
        }
        SendMessageW(interval_combo, CB_SETCURSEL, 0, 0);
        CreateWindowW(L"BUTTON", L"Save Schedule", WS_CHILD | WS_VISIBLE,
            20, 164, 150, 34, window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SAVE)), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Remove", WS_CHILD | WS_VISIBLE,
            182, 164, 120, 34, window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_REMOVE)), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Download Now", WS_CHILD | WS_VISIBLE,
            314, 164, 150, 34, window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_RUN)), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Open Staging Folder", WS_CHILD | WS_VISIBLE,
            476, 164, 180, 34, window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_OPEN_STAGING)), nullptr, nullptr);
        schedule_list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", nullptr,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | LBS_NOINTEGRALHEIGHT,
            20, 210, 830, 300, window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_LIST)), nullptr, nullptr);
        SendMessageW(schedule_list, LB_SETHORIZONTALEXTENT, 1800, 0);
        status_text = CreateWindowW(L"STATIC", L"Ready — secrets use Windows Credential Manager.",
            WS_CHILD | WS_VISIBLE, 20, 530, 830, 24, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_STATUS)), nullptr, nullptr);
        refresh_list();
        SetTimer(window, TIMER_ID, 30000, nullptr);
        return 0;
    }
    case WM_SIZE: layout(window); return 0;
    case WM_COMMAND:
        switch (LOWORD(wparam)) {
        case ID_SAVE:
            try { add_schedule(); }
            catch (const std::exception& ex) {
                MessageBoxA(window, ex.what(), "Point Fetcher", MB_ICONWARNING);
            }
            return 0;
        case ID_REMOVE: {
            const int selected = static_cast<int>(SendMessageW(schedule_list,
                LB_GETCURSEL, 0, 0));
            if (selected >= 0 && static_cast<std::size_t>(selected) < schedules.size()) {
                if (!schedules[selected].credential_target.empty())
                    CredDeleteW(schedules[selected].credential_target.c_str(),
                        CRED_TYPE_GENERIC, 0);
                schedules.erase(schedules.begin() + selected);
                save_config();
                refresh_list();
            }
            return 0;
        }
        case ID_RUN: {
            const int selected = static_cast<int>(SendMessageW(schedule_list,
                LB_GETCURSEL, 0, 0));
            if (selected >= 0)
                start_fetch(static_cast<std::size_t>(selected), true);
            return 0;
        }
        case ID_OPEN_STAGING:
            ShellExecuteW(window, L"open", staging.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            return 0;
        }
        break;
    case WM_TIMER:
        if (wparam == TIMER_ID && !fetch_running) {
            const auto now = now_epoch();
            for (std::size_t index = 0; index < schedules.size(); ++index) {
                if (schedules[index].next_run <= now) {
                    start_fetch(index, false);
                    break;
                }
            }
        }
        return 0;
    case WM_FETCH_COMPLETE: {
        std::unique_ptr<FetchResult> result(reinterpret_cast<FetchResult*>(lparam));
        if (fetch_thread.joinable()) fetch_thread.join();
        fetch_running = false;
        EnableWindow(GetDlgItem(window, ID_SAVE), TRUE);
        EnableWindow(GetDlgItem(window, ID_REMOVE), TRUE);
        EnableWindow(GetDlgItem(window, ID_RUN), TRUE);
        if (result && result->index < schedules.size()) {
            schedules[result->index].next_run = now_epoch() +
                schedules[result->index].interval_hours * 3600LL;
            save_config();
            refresh_list();
            SetWindowTextW(status_text, result->message.c_str());
            if (result->browser_fallback) {
                const auto browser = root.parent_path() / L"PointBrowserFetcher.exe";
                const std::wstring parameters = L"\"" + result->fallback_url + L"\"";
                const auto launched = reinterpret_cast<INT_PTR>(ShellExecuteW(
                    window, L"open", browser.c_str(), parameters.c_str(),
                    root.parent_path().c_str(), SW_SHOWNORMAL));
                if (launched > 32) {
                    SetWindowTextW(status_text,
                        L"This website requires a browser. Opened Point Browser Fetcher automatically.");
                } else {
                    MessageBoxW(window,
                        L"This website requires Point Browser Fetcher, but it could not be opened.",
                        L"Point Fetcher", MB_ICONWARNING);
                }
            } else if (!result->success) {
                MessageBoxW(window, result->message.c_str(), L"Point Fetcher", MB_ICONWARNING);
            }
        }
        return 0;
    }
    case WM_DESTROY:
        if (fetch_thread.joinable()) fetch_thread.join();
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(window, message, wparam, lparam);
}

}  // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int show) {
    wchar_t executable[MAX_PATH]{};
    GetModuleFileNameW(nullptr, executable, MAX_PATH);
    const auto app_root = std::filesystem::path(executable).parent_path();
    root = app_root / "Fetcher";
    staging = root / "Staging";
    inbox = app_root / "Inbox";
    std::filesystem::create_directories(staging);
    std::filesystem::create_directories(inbox);
    load_config();
    INITCOMMONCONTROLSEX controls{sizeof(controls), ICC_STANDARD_CLASSES};
    InitCommonControlsEx(&controls);
    const HICON point_icon = LoadIconW(
        instance, MAKEINTRESOURCEW(IDI_POINT));
    const HICON point_small_icon = reinterpret_cast<HICON>(LoadImageW(
        instance, MAKEINTRESOURCEW(IDI_POINT), IMAGE_ICON,
        GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON),
        LR_DEFAULTCOLOR));
    if (!point_icon || !point_small_icon) {
        MessageBoxW(nullptr,
            L"Point Fetcher's embedded application icon could not be loaded.",
            L"Point Fetcher", MB_OK | MB_ICONERROR);
        return 1;
    }
    WNDCLASSEXW type{sizeof(type)};
    type.lpfnWndProc = window_proc;
    type.hInstance = instance;
    type.lpszClassName = L"PointFetcherWindow";
    type.hCursor = LoadCursor(nullptr, IDC_ARROW);
    type.hIcon = point_icon;
    type.hIconSm = point_small_icon;
    type.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    if (!RegisterClassExW(&type)) return 1;
    main_window = CreateWindowExW(0, type.lpszClassName,
        L"Point Fetcher — Secure HTTPS Report Downloader",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 920, 650,
        nullptr, nullptr, instance, nullptr);
    if (!main_window) return 1;
    SendMessageW(main_window, WM_SETICON, ICON_BIG,
        reinterpret_cast<LPARAM>(point_icon));
    SendMessageW(main_window, WM_SETICON, ICON_SMALL,
        reinterpret_cast<LPARAM>(point_small_icon));
    ShowWindow(main_window, show == SW_HIDE ? SW_SHOWNORMAL : show);
    UpdateWindow(main_window);
    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
    return static_cast<int>(message.wParam);
}
