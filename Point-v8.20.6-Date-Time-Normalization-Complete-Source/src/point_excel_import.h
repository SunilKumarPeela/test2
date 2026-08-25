#pragma once

#include <filesystem>
#include <functional>
#include <string>
#include <vector>

namespace point {

struct ExcelImportResult {
    std::vector<std::filesystem::path> csv_sources;
    std::vector<std::string> issues;
    std::size_t workbook_count = 0;
    std::size_t worksheet_count = 0;
    bool cache_reused = false;
};

ExcelImportResult prepare_import_sources(
    const std::filesystem::path& inbox,
    const std::filesystem::path& cache,
    const std::function<void(
        std::size_t, std::size_t,
        const std::filesystem::path&, bool)>& progress = {},
    const std::function<bool()>& cancelled = {});

}  // namespace point
