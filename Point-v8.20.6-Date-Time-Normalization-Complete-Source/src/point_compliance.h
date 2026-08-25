#pragma once

#include <filesystem>
#include <string>

namespace point::compliance {

struct Policy {
    bool enforce_windows_groups = true;
    std::wstring allowed_groups =
        L"Point Users;Point Administrators";
    std::wstring export_groups =
        L"Point Exporters;Point Administrators";
    int export_retention_days = 30;
    int workspace_retention_days = 30;
    int log_retention_days = 365;
};

Policy load_policy(const std::filesystem::path& root);
void authorize_current_user(const Policy& policy);
bool current_user_can_export(const Policy& policy);
void harden_data_directories(const std::filesystem::path& root);
void enforce_retention(
    const std::filesystem::path& root, const Policy& policy);
void protect_file_for_current_user(const std::filesystem::path& path);
std::string read_user_protected_file(const std::filesystem::path& path);

}  // namespace point::compliance
