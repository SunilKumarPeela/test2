#pragma once

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <string>
#include <vector>

namespace point_validation {

inline std::uint16_t u16(const unsigned char* value) {
    return static_cast<std::uint16_t>(value[0]) |
           (static_cast<std::uint16_t>(value[1]) << 8U);
}

inline std::uint32_t u32(const unsigned char* value) {
    return static_cast<std::uint32_t>(value[0]) |
           (static_cast<std::uint32_t>(value[1]) << 8U) |
           (static_cast<std::uint32_t>(value[2]) << 16U) |
           (static_cast<std::uint32_t>(value[3]) << 24U);
}

inline std::wstring lower_extension(const std::filesystem::path& name) {
    std::wstring extension = name.extension().wstring();
    std::transform(extension.begin(), extension.end(), extension.begin(),
        [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    return extension;
}

inline bool looks_like_html(std::string text) {
    std::transform(text.begin(), text.end(), text.begin(),
        [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
    return text.find("<!doctype html") != std::string::npos ||
           text.find("<html") != std::string::npos ||
           text.find("<script") != std::string::npos ||
           text.find("access denied") != std::string::npos ||
           text.find("sign in") != std::string::npos;
}

inline bool valid_office_zip(const std::filesystem::path& path) {
    std::ifstream input(path, std::ios::binary);
    if (!input) return false;
    input.seekg(0, std::ios::end);
    const auto end_position = input.tellg();
    if (end_position < 22) return false;
    const auto file_size = static_cast<std::uint64_t>(
        static_cast<std::streamoff>(end_position));
    const std::uint64_t tail_size = std::min<std::uint64_t>(file_size, 65'557U);
    std::vector<unsigned char> tail(static_cast<std::size_t>(tail_size));
    input.seekg(static_cast<std::streamoff>(file_size - tail_size), std::ios::beg);
    input.read(reinterpret_cast<char*>(tail.data()),
               static_cast<std::streamsize>(tail.size()));
    if (!input) return false;

    std::size_t eocd = tail.size();
    for (std::size_t position = tail.size() - 22; ; --position) {
        if (u32(tail.data() + position) == 0x06054B50U) {
            eocd = position;
            break;
        }
        if (position == 0) break;
    }
    if (eocd == tail.size()) return false;
    const std::uint16_t entries = u16(tail.data() + eocd + 10);
    const std::uint32_t central_size = u32(tail.data() + eocd + 12);
    const std::uint32_t central_offset = u32(tail.data() + eocd + 16);
    if (entries == 0 || entries == 0xFFFFU || central_size == 0xFFFFFFFFU ||
        central_offset == 0xFFFFFFFFU ||
        static_cast<std::uint64_t>(central_offset) + central_size > file_size)
        return false;

    input.clear();
    input.seekg(static_cast<std::streamoff>(central_offset), std::ios::beg);
    bool content_types = false;
    bool root_relationships = false;
    bool workbook = false;
    for (std::uint32_t index = 0; index < entries; ++index) {
        unsigned char header[46]{};
        input.read(reinterpret_cast<char*>(header), sizeof(header));
        if (!input || u32(header) != 0x02014B50U) return false;
        const std::uint16_t flags = u16(header + 8);
        const std::uint16_t filename_length = u16(header + 28);
        const std::uint16_t extra_length = u16(header + 30);
        const std::uint16_t comment_length = u16(header + 32);
        if ((flags & 0x0001U) != 0 || filename_length == 0 || filename_length > 4096)
            return false;
        std::string filename(filename_length, '\0');
        input.read(filename.data(), filename_length);
        if (!input) return false;
        std::replace(filename.begin(), filename.end(), '\\', '/');
        if (filename.find("../") != std::string::npos ||
            (!filename.empty() && filename.front() == '/'))
            return false;
        std::string normalized = filename;
        std::transform(normalized.begin(), normalized.end(), normalized.begin(),
            [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
        content_types = content_types || normalized == "[content_types].xml";
        root_relationships = root_relationships || normalized == "_rels/.rels";
        workbook = workbook || normalized == "xl/workbook.xml";
        input.seekg(static_cast<std::streamoff>(extra_length) + comment_length,
                    std::ios::cur);
        if (!input) return false;
    }
    return content_types && root_relationships && workbook;
}

inline bool validate(const std::filesystem::path& path,
                     const std::filesystem::path& expected_name,
                     std::wstring& reason) {
    std::error_code error;
    if (!std::filesystem::is_regular_file(path, error) || error) {
        reason = L"The completed download is not a regular file.";
        return false;
    }
    const auto size = std::filesystem::file_size(path, error);
    if (error || size == 0 || size > 2ULL * 1024ULL * 1024ULL * 1024ULL) {
        reason = L"The downloaded file is empty or exceeds the 2 GiB limit.";
        return false;
    }
    const std::wstring extension = lower_extension(expected_name);
    if (extension != L".csv" && extension != L".xlsx" &&
        extension != L".xls" && extension != L".xlsm") {
        reason = L"Only CSV, XLS, XLSX, and XLSM downloads are accepted.";
        return false;
    }
    std::ifstream input(path, std::ios::binary);
    if (!input) {
        reason = L"The downloaded file cannot be opened for validation.";
        return false;
    }
    std::vector<unsigned char> prefix(8192);
    input.read(reinterpret_cast<char*>(prefix.data()),
               static_cast<std::streamsize>(prefix.size()));
    prefix.resize(static_cast<std::size_t>(input.gcount()));
    if (prefix.empty() || looks_like_html(std::string(prefix.begin(), prefix.end()))) {
        reason = L"The website returned HTML, a login page, or an access-denied page.";
        return false;
    }
    if (extension == L".xlsx" || extension == L".xlsm") {
        if (prefix.size() < 4 || prefix[0] != 0x50 || prefix[1] != 0x4B ||
            prefix[2] != 0x03 || prefix[3] != 0x04 || !valid_office_zip(path)) {
            reason = L"The response is not a structurally valid Excel workbook.";
            return false;
        }
    } else if (extension == L".xls") {
        const unsigned char signature[] =
            {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
        if (prefix.size() < sizeof(signature) ||
            !std::equal(std::begin(signature), std::end(signature), prefix.begin())) {
            reason = L"The response does not have a valid legacy Excel signature.";
            return false;
        }
    } else {
        if (std::find(prefix.begin(), prefix.end(), 0) != prefix.end()) {
            reason = L"The CSV contains binary data.";
            return false;
        }
        const std::string sample(prefix.begin(), prefix.end());
        if (sample.find('\n') == std::string::npos ||
            (sample.find(',') == std::string::npos &&
             sample.find(';') == std::string::npos &&
             sample.find('\t') == std::string::npos)) {
            reason = L"The response does not look like structured CSV data.";
            return false;
        }
    }
    return true;
}

}  // namespace point_validation
