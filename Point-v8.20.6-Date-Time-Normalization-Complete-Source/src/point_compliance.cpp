#include "point_compliance.h"

#define NOMINMAX
#include <windows.h>
#include <aclapi.h>
#include <sddl.h>
#include <wincrypt.h>

#include <algorithm>
#include <chrono>
#include <cctype>
#include <fstream>
#include <sstream>
#include <stdexcept>
#include <vector>

namespace point::compliance {
namespace {

std::string trim_copy(std::string value) {
    const auto first = value.find_first_not_of(" \t\r\n");
    if (first == std::string::npos) return {};
    const auto last = value.find_last_not_of(" \t\r\n");
    return value.substr(first, last - first + 1);
}

bool parse_bool(const std::string& value) {
    std::string lower = trim_copy(value);
    std::transform(lower.begin(), lower.end(), lower.begin(),
        [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
    return lower == "true" || lower == "yes" || lower == "1";
}

std::wstring widen_utf8(const std::string& value) {
    if (value.empty()) return {};
    const int count = MultiByteToWideChar(
        CP_UTF8, MB_ERR_INVALID_CHARS, value.data(),
        static_cast<int>(value.size()), nullptr, 0);
    if (count <= 0) throw std::runtime_error("Invalid UTF-8 in security policy");
    std::wstring result(static_cast<std::size_t>(count), L'\0');
    MultiByteToWideChar(
        CP_UTF8, MB_ERR_INVALID_CHARS, value.data(),
        static_cast<int>(value.size()), result.data(), count);
    return result;
}

bool current_token_is_member(const std::wstring& account_name) {
    DWORD sid_bytes = 0;
    DWORD domain_chars = 0;
    SID_NAME_USE use{};
    LookupAccountNameW(
        nullptr, account_name.c_str(), nullptr, &sid_bytes,
        nullptr, &domain_chars, &use);
    if (GetLastError() != ERROR_INSUFFICIENT_BUFFER) return false;
    std::vector<unsigned char> sid(sid_bytes);
    std::wstring domain(domain_chars, L'\0');
    if (!LookupAccountNameW(
            nullptr, account_name.c_str(), sid.data(), &sid_bytes,
            domain.data(), &domain_chars, &use)) {
        return false;
    }
    BOOL member = FALSE;
    return CheckTokenMembership(nullptr, sid.data(), &member) && member;
}

std::vector<unsigned char> read_binary(const std::filesystem::path& path) {
    std::ifstream input(path, std::ios::binary);
    if (!input) throw std::runtime_error("Unable to open protected file");
    input.seekg(0, std::ios::end);
    const auto length = input.tellg();
    if (length < 0 || length > 64 * 1024 * 1024)
        throw std::runtime_error("Protected file exceeds safety limit");
    input.seekg(0, std::ios::beg);
    std::vector<unsigned char> bytes(static_cast<std::size_t>(length));
    if (!bytes.empty())
        input.read(reinterpret_cast<char*>(bytes.data()), length);
    if (!input) throw std::runtime_error("Protected file read failed");
    return bytes;
}

void write_binary_atomic(
        const std::filesystem::path& path,
        const unsigned char* data, std::size_t size) {
    auto temporary = path;
    temporary += L".secure.tmp";
    {
        std::ofstream output(
            temporary, std::ios::binary | std::ios::trunc);
        if (!output) throw std::runtime_error("Unable to write protected file");
        output.write(
            reinterpret_cast<const char*>(data),
            static_cast<std::streamsize>(size));
        if (!output) throw std::runtime_error("Protected file write failed");
    }
    if (!MoveFileExW(
            temporary.c_str(), path.c_str(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        std::error_code ignored;
        std::filesystem::remove(temporary, ignored);
        throw std::runtime_error("Unable to publish protected file");
    }
}

void delete_expired_files(
        const std::filesystem::path& directory, int retention_days) {
    if (retention_days < 1) return;
    std::error_code ec;
    const auto cutoff = std::filesystem::file_time_type::clock::now() -
        std::chrono::hours(24LL * retention_days);
    for (const auto& entry :
         std::filesystem::directory_iterator(directory, ec)) {
        if (ec) break;
        if (!entry.is_regular_file(ec) || ec) continue;
        // This DPAPI-protected policy is durable application configuration,
        // not an expiring saved query or generated workspace artifact.
        if (entry.path().filename() == L"field-synonyms.dat") continue;
        const auto modified = entry.last_write_time(ec);
        if (!ec && modified < cutoff)
            std::filesystem::remove(entry.path(), ec);
        ec.clear();
    }
}

}  // namespace

Policy load_policy(const std::filesystem::path& root) {
    Policy policy;
    std::ifstream input(root / "point-security.conf");
    if (!input)
        throw std::runtime_error(
            "Missing point-security.conf. Run configure_compliance.bat first.");
    std::string line;
    while (std::getline(input, line)) {
        line = trim_copy(line);
        if (line.empty() || line[0] == '#') continue;
        const auto separator = line.find('=');
        if (separator == std::string::npos)
            throw std::runtime_error("Invalid point-security.conf entry");
        const auto key = trim_copy(line.substr(0, separator));
        const auto value = trim_copy(line.substr(separator + 1));
        if (key == "enforce_windows_groups")
            policy.enforce_windows_groups = parse_bool(value);
        else if (key == "allowed_windows_groups")
            policy.allowed_groups = widen_utf8(value);
        else if (key == "export_windows_groups")
            policy.export_groups = widen_utf8(value);
        else if (key == "export_retention_days")
            policy.export_retention_days = std::stoi(value);
        else if (key == "workspace_retention_days")
            policy.workspace_retention_days = std::stoi(value);
        else if (key == "log_retention_days")
            policy.log_retention_days = std::stoi(value);
        else
            throw std::runtime_error("Unknown point-security.conf setting");
    }
    if (policy.enforce_windows_groups && policy.allowed_groups.empty())
        throw std::runtime_error("At least one allowed Windows group is required");
    if (policy.export_groups.empty())
        throw std::runtime_error("At least one export Windows group is required");
    for (int days : {
            policy.export_retention_days,
            policy.workspace_retention_days,
            policy.log_retention_days}) {
        if (days < 1 || days > 3650)
            throw std::runtime_error("Retention days must be from 1 to 3650");
    }
    return policy;
}

void authorize_current_user(const Policy& policy) {
    if (!policy.enforce_windows_groups) return;
    std::wstringstream groups(policy.allowed_groups);
    std::wstring group;
    while (std::getline(groups, group, L';')) {
        const auto first = group.find_first_not_of(L" \t");
        const auto last = group.find_last_not_of(L" \t");
        if (first == std::wstring::npos) continue;
        group = group.substr(first, last - first + 1);
        if (current_token_is_member(group)) return;
    }
    throw std::runtime_error(
        "Access denied: the Windows user is not in an allowed Point group");
}

bool current_user_can_export(const Policy& policy) {
    // A per-user installation is already isolated by its owner-only ACL and
    // does not require machine-local deployment groups. Company policies that
    // enable group enforcement retain the stricter exporter-group check.
    if (!policy.enforce_windows_groups) return true;
    std::wstringstream groups(policy.export_groups);
    std::wstring group;
    while (std::getline(groups, group, L';')) {
        const auto first = group.find_first_not_of(L" \t");
        const auto last = group.find_last_not_of(L" \t");
        if (first == std::wstring::npos) continue;
        group = group.substr(first, last - first + 1);
        if (current_token_is_member(group)) return true;
    }
    return false;
}

void harden_data_directories(const std::filesystem::path& root) {
    PSECURITY_DESCRIPTOR descriptor = nullptr;
    if (!ConvertStringSecurityDescriptorToSecurityDescriptorW(
            L"D:P(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICI;FA;;;OW)",
            SDDL_REVISION_1, &descriptor, nullptr)) {
        throw std::runtime_error("Unable to create the Point directory ACL");
    }
    for (const wchar_t* name : {
            L"Inbox", L"Workspace", L"Exports", L"Logs"}) {
        const auto path = root / name;
        if (!SetFileSecurityW(
                path.c_str(), DACL_SECURITY_INFORMATION, descriptor)) {
            LocalFree(descriptor);
            throw std::runtime_error("Unable to protect Point data directories");
        }
    }
    LocalFree(descriptor);
}

void enforce_retention(
        const std::filesystem::path& root, const Policy& policy) {
    delete_expired_files(
        root / "Exports", policy.export_retention_days);
    delete_expired_files(
        root / "Workspace", policy.workspace_retention_days);
    delete_expired_files(
        root / "Logs", policy.log_retention_days);
}

void protect_file_for_current_user(const std::filesystem::path& path) {
    auto plain = read_binary(path);
    DATA_BLOB input{
        static_cast<DWORD>(plain.size()), plain.data()};
    DATA_BLOB output{};
    if (!CryptProtectData(
            &input, L"Point protected workspace", nullptr, nullptr, nullptr,
            CRYPTPROTECT_UI_FORBIDDEN, &output)) {
        throw std::runtime_error("Windows DPAPI could not protect the file");
    }
    try {
        write_binary_atomic(path, output.pbData, output.cbData);
    } catch (...) {
        SecureZeroMemory(plain.data(), plain.size());
        LocalFree(output.pbData);
        throw;
    }
    SecureZeroMemory(plain.data(), plain.size());
    LocalFree(output.pbData);
}

std::string read_user_protected_file(const std::filesystem::path& path) {
    auto protected_bytes = read_binary(path);
    if (protected_bytes.size() >= 11 &&
        std::equal(
            protected_bytes.begin(), protected_bytes.begin() + 11,
            reinterpret_cast<const unsigned char*>("POINT_VIEW_"))) {
        throw std::runtime_error(
            "Unencrypted legacy view rejected; save it again with Point v8");
    }
    DATA_BLOB input{
        static_cast<DWORD>(protected_bytes.size()),
        protected_bytes.data()};
    DATA_BLOB output{};
    if (!CryptUnprotectData(
            &input, nullptr, nullptr, nullptr, nullptr,
            CRYPTPROTECT_UI_FORBIDDEN, &output)) {
        throw std::runtime_error(
            "Workspace cannot be decrypted by the current Windows user");
    }
    std::string result(
        reinterpret_cast<const char*>(output.pbData), output.cbData);
    SecureZeroMemory(output.pbData, output.cbData);
    LocalFree(output.pbData);
    return result;
}

}  // namespace point::compliance
