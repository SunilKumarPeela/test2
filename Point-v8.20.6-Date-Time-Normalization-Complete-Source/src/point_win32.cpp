#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <shobjidl.h>

#include "point_core.h"
#include "point_compliance.h"
#include "point_excel_import.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <ctime>
#include <cwctype>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <map>
#include <memory>
#include <numeric>
#include <optional>
#include <set>
#include <sstream>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <thread>

namespace {

constexpr int ID_REFRESH = 101;
constexpr int ID_SEARCH = 102;
constexpr int ID_EXPORT = 103;
constexpr int ID_STATUS = 104;
constexpr int ID_MODE = 105;
constexpr int ID_GRID_SCROLL = 106;
constexpr int ID_COPY_ALL = 107;
constexpr int ID_HELP = 108;
constexpr int ID_GRID_HSCROLL = 109;
constexpr int ID_REFRESH_PROGRESS = 110;
constexpr int ID_TAB_WORKSPACE = 115;
constexpr int ID_TAB_INPUT = 116;
constexpr int ID_INPUT_DROP_ZONE = 117;
constexpr int ID_INPUT_UPLOAD = 118;
constexpr int ID_INPUT_FILE_LIST = 119;
constexpr int ID_INPUT_REMOVE = 120;
constexpr int ID_INPUT_FIELD_SEARCH = 121;
constexpr int ID_INPUT_ARCHIVE_SELECT = 122;
constexpr int ID_INPUT_ARCHIVE_LIST = 123;
constexpr int ID_TAB_SCHEDULE = 124;
constexpr int ID_SCHEDULE_ADD = 125;
constexpr int ID_SCHEDULE_INTERVAL = 126;
constexpr int ID_SCHEDULE_LIST = 127;
constexpr int ID_SCHEDULE_REMOVE = 128;
constexpr int ID_SCHEDULE_RUN_NOW = 129;
constexpr int ID_INPUT_ARCHIVE_TOGGLE = 130;
constexpr int ID_SCHEDULE_FETCHER = 131;
constexpr int ID_MENU_IMPORT_SUMMARY = 201;
constexpr int ID_MENU_CLEAR_RESULTS = 202;
constexpr int ID_MENU_NEXT_SOURCE = 203;
constexpr int ID_FIND_TEXT = 204;
constexpr int ID_FIND_NEXT = 205;
constexpr int ID_MENU_DYNAMIC_CHART = 206;
constexpr int ID_MENU_SAVE_VIEW = 207;
constexpr int ID_MENU_LOAD_VIEW = 208;
constexpr int ID_MENU_RISK_WATCHLIST = 209;
constexpr int ID_MENU_SET_CHANGE_BASELINE = 210;
constexpr int ID_MENU_FIELD_SYNONYMS = 211;
constexpr int ID_MENU_RELATIONSHIP_DIAGNOSTICS = 212;
constexpr int ID_MENU_EXPLAIN_RESULT = 213;
constexpr int ID_DATA_FILTER_EQUALS = 214;
constexpr int ID_DATA_FILTER_CONTAINS = 215;
constexpr int ID_DATA_FILTER_STARTS = 216;
constexpr int ID_DATA_FILTER_ENDS = 217;
constexpr int ID_DATA_FILTER_BLANK = 218;
constexpr int ID_DATA_FILTER_NOT_BLANK = 219;
constexpr int ID_DATA_FILTER_CLEAR = 220;
constexpr int ID_DATA_SPLIT_SLASH = 221;
constexpr int ID_DATA_SPLIT_BACKSLASH = 222;
constexpr int ID_DATA_SPLIT_PIPE = 223;
constexpr int ID_DATA_SPLIT_COMMA = 224;
constexpr int ID_DATA_SPLIT_SEMICOLON = 225;
constexpr int ID_DATA_SPLIT_COLON = 226;
constexpr int ID_DATA_SPLIT_SPACE = 227;
constexpr int ID_DATA_SPLIT_CUSTOM = 228;
constexpr int ID_QUICK_TEXT_TRANSFORM = 229;
constexpr int ID_QUICK_KEEP_LEFT = 230;
constexpr int ID_QUICK_KEEP_RIGHT = 231;
constexpr int ID_TRANSFORM_PATTERN = 232;
constexpr int ID_SYNONYM_SAVE = 12002;
constexpr int ID_SYNONYM_CANCEL = 12003;
constexpr int ID_SYNONYM_HEADER_BASE = 12100;
constexpr int ID_SYNONYM_CELL_BASE = 12200;
constexpr int ID_SYNONYM_HSCROLL = 12300;
constexpr int ID_SYNONYM_VSCROLL = 12301;
constexpr int SYNONYM_FIELD_COLUMNS = 64;
constexpr int SYNONYM_ROWS = 64;
constexpr int SYNONYM_VISIBLE_COLUMNS = 8;
constexpr int SYNONYM_VISIBLE_ROWS = 12;
constexpr int ID_CHART_FILTER = 10001;
constexpr int ID_CHART_TYPE = 10002;
constexpr int ID_CHART_TOP = 10003;
constexpr int ID_CHART_SORT = 10004;
constexpr int ID_HEADER_BASE = 300;
constexpr int ID_CELL_BASE = 1000;
// Point keeps 256 logical columns available while only showing the columns
// that fit in the window. This removes the old six-column/five-field limit
// without allocating controls for 2,000,000 rows.
constexpr int GRID_COLUMNS = 256;
constexpr int MIN_COLUMN_WIDTH = 180;
// Keep enough reusable edit controls for a full-HD maximized window without
// approaching the Windows per-process USER-object limit. Layout shows only
// the rows that fit and stretches them slightly when extra height remains.
constexpr int GRID_CONTROL_ROWS = 32;
constexpr int LOGICAL_ROWS = 2'000'000;
constexpr std::uint64_t COUNT_DETAIL_LIMIT = 50;
constexpr UINT_PTR SUGGESTION_TIMER_ID = 1;
constexpr UINT_PTR AUTO_SCHEDULE_TIMER_ID = 2;
constexpr UINT_PTR INBOX_NOTIFICATION_TIMER_ID = 3;
constexpr UINT WM_POINT_REFRESH_PROGRESS = WM_APP + 20;
constexpr UINT WM_POINT_REFRESH_COMPLETE = WM_APP + 21;
constexpr UINT WM_POINT_ARCHIVE_SCAN_COMPLETE = WM_APP + 22;
UINT inbox_changed_message = 0;

std::unique_ptr<point::Engine> engine;
std::unique_ptr<point::Engine> previous_engine;
point::QueryResult last_result;
std::filesystem::path app_root;
point::compliance::Policy compliance_policy;
bool export_authorized = false;
HWND main_window = nullptr;
HWND status_text = nullptr;
HWND refresh_progress = nullptr;
HWND help_text = nullptr;
HWND suggestion_list = nullptr;
HWND suggestion_target = nullptr;
HWND grid_scrollbar = nullptr;
HWND grid_hscrollbar = nullptr;
HWND input_drop_zone = nullptr;
HWND input_upload_button = nullptr;
HWND input_file_list = nullptr;
HWND input_remove_button = nullptr;
HWND input_field_search = nullptr;
HWND input_imported_label = nullptr;
HWND input_archive_label = nullptr;
HWND input_archive_select_button = nullptr;
HWND input_archive_list = nullptr;
HWND input_archive_toggle_button = nullptr;
HWND schedule_help_text = nullptr;
HWND schedule_add_button = nullptr;
HWND schedule_interval_combo = nullptr;
HWND schedule_list = nullptr;
HWND schedule_remove_button = nullptr;
HWND schedule_run_now_button = nullptr;
HWND input_columns_popup = nullptr;
HWND input_columns_popup_list = nullptr;
HWND input_field_tooltip = nullptr;
std::wstring input_field_tooltip_text;
std::string input_field_tooltip_key;
std::unordered_map<std::string, std::wstring> input_field_sample_cache;
std::unordered_map<std::string, std::size_t> input_common_field_counts;
std::vector<std::filesystem::path> input_archive_files;
std::filesystem::path input_archive_root;
std::thread input_archive_scan_thread;
std::atomic_bool input_archive_scan_running{false};
std::atomic_bool input_archive_scan_cancel{false};
bool input_tab_active = false;
bool schedule_tab_active = false;
bool input_archive_collapsed = false;

struct AutoImportSchedule {
    std::filesystem::path source;
    int interval_hours = 1;
    std::int64_t next_run_epoch = 0;
};
std::vector<AutoImportSchedule> auto_import_schedules;
std::vector<HWND> header_cells;
std::vector<std::vector<HWND>> data_cells;
std::unordered_map<std::uint64_t, std::string> cell_store;
std::unordered_set<std::uint64_t> pending_identity_resolution_cells;
std::array<std::vector<int>, GRID_COLUMNS> cell_rows_by_column;
std::array<bool, GRID_COLUMNS> cell_row_index_needs_sort{};
std::array<bool, GRID_COLUMNS> cell_row_index_needs_rebuild{};
std::set<int> universal_missing_rows;
std::set<int> universal_duplicate_rows;
std::set<int> universal_duplicate_dark_rows;
bool universal_results_displayed = false;
std::set<int> universal_pending_lookup_rows;
std::vector<std::string> universal_result_headers;
std::vector<std::string> universal_lookup_history;
int first_visible_row = 0;
int first_visible_column = 0;
int visible_grid_columns = 1;
int visible_grid_rows = 15;
std::vector<int> column_widths(
    static_cast<std::size_t>(GRID_COLUMNS), MIN_COLUMN_WIDTH);
int grid_row_height = 28;
bool resizing_column = false;
int resizing_column_index = -1;
int resize_start_screen_x = 0;
int resize_start_width = MIN_COLUMN_WIDTH;
bool resizing_rows = false;
int resize_start_screen_y = 0;
int resize_start_row_height = 28;
int vertical_wheel_remainder = 0;
int horizontal_wheel_remainder = 0;
bool internal_cell_update = false;
std::atomic_bool refresh_running{false};
std::atomic_bool refresh_cancel_requested{false};
std::thread refresh_thread;

struct RefreshProgressUpdate {
    int percent = 0;
    std::wstring text;
};

struct RefreshCompletion {
    std::unique_ptr<point::Engine> refreshed;
    point::ExcelImportResult imported;
    std::string error;
    bool cancelled = false;
    bool preserve_workspace = false;
};
bool narrow_mode = false;
bool count_mode = false;
bool compare_mode = false;
bool analyze_mode = false;
bool insight_mode = false;
bool chart_mode = false;
bool change_mode = false;
int generated_count_column = -1;
bool generated_count_details = false;
int count_detail_first_column = -1;
int count_detail_column_count = 0;
bool generated_analysis_headers = false;
bool generated_compare_header = false;
std::string compare_identity_field_cache;
std::vector<std::string> compare_fields_cache;
std::vector<std::string> compare_inputs_cache;
std::vector<std::string> analysis_key_fields_cache;
std::vector<std::string> analysis_key_values_cache;
bool generated_insight_headers = false;
std::vector<std::string> insight_fields_cache;
bool generated_chart_headers = false;
std::vector<std::string> chart_fields_cache;
std::vector<std::string> chart_filter_values_cache;
bool generated_change_headers = false;
std::vector<std::string> change_key_fields_cache;
int selected_source_index = -1;
HWND find_text = nullptr;
HWND quick_text_transform_button = nullptr;
HWND quick_keep_left_button = nullptr;
HWND quick_keep_right_button = nullptr;
HWND transform_pattern_text = nullptr;
HWND transform_target_cell = nullptr;
int last_find_row = -1;
int last_find_column = -1;
int last_sort_column = -1;
bool sort_ascending = true;
std::string last_import_issue;
std::optional<point::QueryResult> data_tools_unfiltered_result;
std::vector<point::FieldSynonymGroup> field_synonym_groups;
HWND synonym_window = nullptr;
std::vector<HWND> synonym_headers;
std::vector<std::vector<HWND>> synonym_cells;
std::vector<HWND> synonym_column_labels;
std::vector<HWND> synonym_row_labels;
std::vector<std::string> synonym_model_headers(SYNONYM_FIELD_COLUMNS);
std::vector<std::vector<std::string>> synonym_model_cells(
    SYNONYM_ROWS,
    std::vector<std::string>(SYNONYM_FIELD_COLUMNS));
int synonym_first_column = 0;
int synonym_first_row = 0;
HWND synonym_hscrollbar = nullptr;
HWND synonym_vscrollbar = nullptr;
bool synonym_view_loading = false;

struct ChartBar {
    std::wstring group;
    std::wstring label;
    std::wstring series;
    double value = 0.0;
    std::vector<point::QueryCondition> conditions;
    RECT bounds{};
};
std::vector<ChartBar> chart_bars;
std::vector<ChartBar> chart_all_bars;
std::wstring chart_title;
HWND chart_window = nullptr;
HWND chart_filter = nullptr;
enum class ChartView { Auto, Bar, Column, Pie, Stacked };
ChartView chart_view = ChartView::Auto;
int chart_top_limit = 20;  // 0 means all.
bool chart_descending = true;
int chart_selected_index = -1;

struct GridPosition {
    int row = 0;  // -1 is the heading row.
    int column = 0;
};

GridPosition selection_anchor{};
GridPosition selection_end{};
bool selection_candidate = false;
bool selection_active = false;

struct LearnedTextTransformation {
    enum class Operation {
        RemoveHighlighted,
        SkipLeftKeepRight,
        SkipRightKeepLeft
    };
    bool available = false;
    int column = -1;
    std::string selected_text;
    point::TextMatchPosition position = point::TextMatchPosition::Anywhere;
    Operation operation = Operation::RemoveHighlighted;
    std::vector<std::pair<int, std::string>> undo_values;
};
LearnedTextTransformation learned_text_transformation;
HWND quick_transform_cell = nullptr;
DWORD quick_transform_start = 0;
DWORD quick_transform_end = 0;
HBRUSH selected_cell_brush = nullptr;
HBRUSH duplicate_light_brush = nullptr;
HBRUSH duplicate_dark_brush = nullptr;
HBRUSH criteria_cell_brush = nullptr;
HBRUSH changed_row_brush = nullptr;
HBRUSH synonym_missing_brush = nullptr;
HBRUSH synonym_duplicate_brush = nullptr;
HWND pending_suggestion_target = nullptr;
int pending_suggestion_id = 0;

bool universal_mode_active() {
    return !narrow_mode && !count_mode && !compare_mode &&
           !analyze_mode && !insight_mode && !chart_mode &&
           !change_mode;
}

bool same_normalized_headers(
        const std::vector<std::string>& left,
        const std::vector<std::string>& right) {
    if (left.size() != right.size()) return false;
    for (std::size_t index = 0; index < left.size(); ++index) {
        if (point::normalize_name(left[index]) !=
            point::normalize_name(right[index])) {
            return false;
        }
    }
    return true;
}

void open_dynamic_chart();
void drill_down_chart_bar(
    const std::wstring& group, const std::wstring& label,
    const std::wstring& series);
void clear_count_detail_columns();
void layout(HWND window);
void update_mode_ui();
void begin_chart_heading_edit(int column, const std::string& typed_value);
void refresh_engine(bool preserve_workspace = false);
void import_input_files(const std::vector<std::filesystem::path>& files);
void set_schedule_tab(bool active);
void set_input_tab(bool active);
void refresh_input_file_list();
void show_input_file_columns();
std::vector<std::wstring> input_file_headers(const std::wstring& filename);
void show_input_file_columns();
void hide_input_columns_popup();
void remove_selected_input_file();
void load_visible_cells();
void repaint_grid_cells();
void refresh_synonym_grid_colors();
std::string narrow(const std::wstring& text);
std::wstring control_text(HWND control);
void set_control_text(HWND control, const std::string& value);

std::vector<point::FieldSynonymGroup> parse_field_synonyms(
        const std::string& text) {
    std::vector<point::FieldSynonymGroup> groups;
    std::unordered_map<std::string, std::size_t> canonical_rows;
    std::istringstream input(text);
    std::string line;
    std::size_t line_number = 0;
    while (std::getline(input, line)) {
        ++line_number;
        if (!line.empty() && line.back() == '\r') line.pop_back();
        line = point::trim(line);
        if (line.empty() || line[0] == '#') continue;
        const auto separator = line.find('=');
        if (separator == std::string::npos ||
            line.find('=', separator + 1) != std::string::npos) {
            throw std::runtime_error(
                "Line " + std::to_string(line_number) +
                " must use: Canonical Field = Synonym, Synonym");
        }
        const auto canonical = point::trim(line.substr(0, separator));
        if (canonical.empty())
            throw std::runtime_error(
                "Line " + std::to_string(line_number) +
                " has a blank canonical field");
        const auto canonical_key = point::normalize_name(canonical);
        if (canonical_rows.contains(canonical_key))
            throw std::runtime_error(
                "Canonical field '" + canonical + "' is listed twice");
        point::FieldSynonymGroup group;
        group.canonical_field = canonical;
        std::istringstream synonyms(line.substr(separator + 1));
        std::string synonym;
        while (std::getline(synonyms, synonym, ',')) {
            synonym = point::trim(synonym);
            if (!synonym.empty()) group.synonyms.push_back(synonym);
        }
        if (group.synonyms.empty())
            throw std::runtime_error(
                "Line " + std::to_string(line_number) +
                " requires at least one synonym");
        canonical_rows[canonical_key] = groups.size();
        groups.push_back(std::move(group));
    }
    return groups;
}

std::string serialize_field_synonyms(
        const std::vector<point::FieldSynonymGroup>& groups) {
    std::ostringstream output;
    output << "POINT_FIELD_SYNONYMS_V1\n";
    for (const auto& group : groups) {
        output << group.canonical_field << " = ";
        for (std::size_t index = 0; index < group.synonyms.size(); ++index) {
            if (index) output << ", ";
            output << group.synonyms[index];
        }
        output << '\n';
    }
    return output.str();
}

std::vector<std::string> split_synonym_cell(const std::string& text) {
    std::vector<std::string> names;
    std::string name;
    for (const char ch : text) {
        if (ch == ',' || ch == ';') {
            name = point::trim(name);
            if (!name.empty()) names.push_back(name);
            name.clear();
        } else {
            name.push_back(ch);
        }
    }
    name = point::trim(name);
    if (!name.empty()) names.push_back(name);
    return names;
}

void commit_synonym_view() {
    if (synonym_view_loading ||
        synonym_headers.size() !=
            static_cast<std::size_t>(SYNONYM_VISIBLE_COLUMNS) ||
        synonym_cells.size() !=
            static_cast<std::size_t>(SYNONYM_VISIBLE_ROWS)) {
        return;
    }
    for (int visible_column = 0;
         visible_column < SYNONYM_VISIBLE_COLUMNS; ++visible_column) {
        const int logical_column =
            synonym_first_column + visible_column;
        synonym_model_headers[static_cast<std::size_t>(logical_column)] =
            narrow(control_text(
                synonym_headers[static_cast<std::size_t>(visible_column)]));
        for (int visible_row = 0;
             visible_row < SYNONYM_VISIBLE_ROWS; ++visible_row) {
            const int logical_row = synonym_first_row + visible_row;
            synonym_model_cells[static_cast<std::size_t>(logical_row)]
                               [static_cast<std::size_t>(logical_column)] =
                narrow(control_text(
                    synonym_cells[static_cast<std::size_t>(visible_row)]
                                  [static_cast<std::size_t>(visible_column)]));
        }
    }
}

void update_synonym_scrollbars() {
    if (IsWindow(synonym_hscrollbar)) {
        SCROLLINFO info{sizeof(info)};
        info.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
        info.nMin = 0;
        info.nMax = SYNONYM_FIELD_COLUMNS - 1;
        info.nPage = SYNONYM_VISIBLE_COLUMNS;
        info.nPos = synonym_first_column;
        SetScrollInfo(synonym_hscrollbar, SB_CTL, &info, TRUE);
    }
    if (IsWindow(synonym_vscrollbar)) {
        SCROLLINFO info{sizeof(info)};
        info.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
        info.nMin = 0;
        info.nMax = SYNONYM_ROWS - 1;
        info.nPage = SYNONYM_VISIBLE_ROWS;
        info.nPos = synonym_first_row;
        SetScrollInfo(synonym_vscrollbar, SB_CTL, &info, TRUE);
    }
}

void load_synonym_view() {
    synonym_view_loading = true;
    for (int visible_column = 0;
         visible_column < SYNONYM_VISIBLE_COLUMNS; ++visible_column) {
        const int logical_column =
            synonym_first_column + visible_column;
        set_control_text(
            synonym_headers[static_cast<std::size_t>(visible_column)],
            synonym_model_headers[
                static_cast<std::size_t>(logical_column)]);
        if (visible_column <
            static_cast<int>(synonym_column_labels.size())) {
            const auto label =
                L"Field " + std::to_wstring(logical_column + 1);
            SetWindowTextW(
                synonym_column_labels[
                    static_cast<std::size_t>(visible_column)],
                label.c_str());
        }
        for (int visible_row = 0;
             visible_row < SYNONYM_VISIBLE_ROWS; ++visible_row) {
            const int logical_row = synonym_first_row + visible_row;
            set_control_text(
                synonym_cells[static_cast<std::size_t>(visible_row)]
                              [static_cast<std::size_t>(visible_column)],
                synonym_model_cells[static_cast<std::size_t>(logical_row)]
                                   [static_cast<std::size_t>(logical_column)]);
        }
    }
    for (int visible_row = 0;
         visible_row < SYNONYM_VISIBLE_ROWS &&
         visible_row < static_cast<int>(synonym_row_labels.size());
         ++visible_row) {
        const auto label =
            std::to_wstring(synonym_first_row + visible_row + 1);
        SetWindowTextW(
            synonym_row_labels[static_cast<std::size_t>(visible_row)],
            label.c_str());
    }
    synonym_view_loading = false;
    update_synonym_scrollbars();
    refresh_synonym_grid_colors();
}

std::vector<point::FieldSynonymGroup> field_synonyms_from_grid() {
    commit_synonym_view();
    std::vector<point::FieldSynonymGroup> groups;
    for (int column = 0; column < SYNONYM_FIELD_COLUMNS; ++column) {
        const auto canonical = point::trim(
            synonym_model_headers[static_cast<std::size_t>(column)]);
        std::vector<std::string> synonyms;
        for (int row = 0; row < SYNONYM_ROWS; ++row) {
            const auto cell_text = point::trim(
                synonym_model_cells[static_cast<std::size_t>(row)]
                                   [static_cast<std::size_t>(column)]);
            const auto cell_synonyms = split_synonym_cell(cell_text);
            synonyms.insert(
                synonyms.end(),
                cell_synonyms.begin(), cell_synonyms.end());
        }
        if (canonical.empty() && synonyms.empty()) continue;
        if (canonical.empty())
            throw std::runtime_error(
                "A synonym column has values but no canonical field heading");
        if (synonyms.empty())
            throw std::runtime_error(
                "Canonical field '" + canonical +
                "' requires at least one synonym underneath it");
        groups.push_back({canonical, std::move(synonyms)});
    }
    return groups;
}

std::vector<HWND> synonym_grid_controls() {
    std::vector<HWND> controls;
    controls.reserve(
        SYNONYM_VISIBLE_COLUMNS * (SYNONYM_VISIBLE_ROWS + 1));
    controls.insert(
        controls.end(), synonym_headers.begin(), synonym_headers.end());
    for (const auto& row : synonym_cells)
        controls.insert(controls.end(), row.begin(), row.end());
    return controls;
}

std::size_t synonym_grid_occurrences(const std::string& normalized) {
    if (normalized.empty()) return 0;
    std::size_t occurrences = 0;
    for (int column = 0; column < SYNONYM_FIELD_COLUMNS; ++column) {
        for (const auto& name : split_synonym_cell(
                 synonym_model_headers[static_cast<std::size_t>(column)])) {
            if (point::normalize_name(name) == normalized) ++occurrences;
        }
        for (int row = 0; row < SYNONYM_ROWS; ++row) {
            for (const auto& name : split_synonym_cell(
                     synonym_model_cells[static_cast<std::size_t>(row)]
                                        [static_cast<std::size_t>(column)])) {
                if (point::normalize_name(name) == normalized) ++occurrences;
            }
        }
    }
    return occurrences;
}

bool imported_field_exists(const std::string& normalized) {
    if (normalized.empty() || !engine) return false;
    for (const auto& dataset : engine->datasets()) {
        if (std::any_of(
                dataset.headers.begin(), dataset.headers.end(),
                [&](const auto& header) {
                    return point::normalize_name(header) == normalized;
                })) {
            return true;
        }
    }
    return false;
}

void refresh_synonym_grid_colors() {
    for (const auto control : synonym_grid_controls()) {
        if (IsWindow(control))
            InvalidateRect(control, nullptr, TRUE);
    }
}

void load_field_synonyms() {
    const auto path = app_root / "Workspace" / "field-synonyms.dat";
    if (!std::filesystem::exists(path)) {
        field_synonym_groups.clear();
        return;
    }
    const auto protected_text =
        point::compliance::read_user_protected_file(path);
    static const std::string signature = "POINT_FIELD_SYNONYMS_V1\n";
    if (protected_text.rfind(signature, 0) != 0)
        throw std::runtime_error("Field synonym configuration is invalid");
    field_synonym_groups =
        parse_field_synonyms(protected_text.substr(signature.size()));
    point::Engine validator;
    validator.set_field_synonyms(field_synonym_groups);
}

void save_field_synonyms(
        const std::vector<point::FieldSynonymGroup>& groups) {
    const auto path = app_root / "Workspace" / "field-synonyms.dat";
    const auto temporary = app_root / "Workspace" / "field-synonyms.tmp";
    {
        std::ofstream output(temporary, std::ios::binary | std::ios::trunc);
        if (!output) throw std::runtime_error(
            "Unable to write the field synonym configuration");
        const auto content = serialize_field_synonyms(groups);
        output.write(content.data(), static_cast<std::streamsize>(content.size()));
        output.flush();
        if (!output) throw std::runtime_error(
            "Unable to complete the field synonym configuration");
    }
    point::compliance::protect_file_for_current_user(temporary);
    if (!MoveFileExW(
            temporary.c_str(), path.c_str(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        std::error_code ignored;
        std::filesystem::remove(temporary, ignored);
        throw std::runtime_error(
            "Unable to publish the field synonym configuration");
    }
}

void open_field_synonyms() {
    if (refresh_running.load()) {
        MessageBoxW(main_window,
            L"Point is still indexing the Inbox. Wait for Refresh to finish "
            L"so the Synonym Manager can use every current sheet heading.",
            L"Field Synonym Manager", MB_ICONINFORMATION);
        return;
    }
    if (IsWindow(synonym_window)) {
        SetForegroundWindow(synonym_window);
        return;
    }
    if (field_synonym_groups.size() >
        static_cast<std::size_t>(SYNONYM_FIELD_COLUMNS)) {
        MessageBoxW(
            main_window,
            L"The saved configuration exceeds the manager's 64-field "
            L"capacity. The existing mappings remain "
            L"active and have not been changed.",
            L"Field Synonym Manager", MB_ICONERROR);
        return;
    }
    synonym_window = CreateWindowExW(
        WS_EX_DLGMODALFRAME, L"PointSynonymWindow",
        L"Point — Field Synonym Manager",
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT, 1240, 640,
        main_window, nullptr,
        reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(
            main_window, GWLP_HINSTANCE)), nullptr);
}

void hide_suggestions();
void schedule_suggestions(HWND target, int id);
void accept_suggestion();

LRESULT CALLBACK synonym_edit_subclass(
        HWND control, UINT message, WPARAM wparam, LPARAM lparam,
        UINT_PTR subclass_id, DWORD_PTR reference_data) {
    (void)reference_data;
    if (message == WM_KEYDOWN && wparam == VK_DOWN &&
        suggestion_target == control &&
        IsWindowVisible(suggestion_list)) {
        SetFocus(suggestion_list);
        if (SendMessageW(suggestion_list, LB_GETCURSEL, 0, 0) == LB_ERR)
            SendMessageW(suggestion_list, LB_SETCURSEL, 0, 0);
        return 0;
    }
    if (message == WM_KEYDOWN && wparam == VK_ESCAPE &&
        suggestion_target == control) {
        hide_suggestions();
        return 0;
    }
    if (message == WM_KEYDOWN && wparam == VK_RETURN &&
        suggestion_target == control &&
        IsWindowVisible(suggestion_list)) {
        if (SendMessageW(suggestion_list, LB_GETCURSEL, 0, 0) == LB_ERR)
            SendMessageW(suggestion_list, LB_SETCURSEL, 0, 0);
        accept_suggestion();
        return 0;
    }
    if (message == WM_KEYDOWN && wparam == VK_RETURN) {
        const int id = GetDlgCtrlID(control);
        const bool reverse =
            (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        int target_row = -1;
        int target_column = -1;
        if (id >= ID_SYNONYM_HEADER_BASE &&
            id < ID_SYNONYM_HEADER_BASE + SYNONYM_VISIBLE_COLUMNS) {
            const int visible_column = id - ID_SYNONYM_HEADER_BASE;
            const int logical_column =
                synonym_first_column + visible_column;
            if (reverse && logical_column > 0) {
                target_column = logical_column - 1;
                target_row = SYNONYM_ROWS - 1;
            } else {
                target_column = logical_column;
                target_row = 0;
            }
        } else if (id >= ID_SYNONYM_CELL_BASE &&
                   id < ID_SYNONYM_CELL_BASE +
                       SYNONYM_VISIBLE_COLUMNS * SYNONYM_VISIBLE_ROWS) {
            const int index = id - ID_SYNONYM_CELL_BASE;
            const int visible_row = index / SYNONYM_VISIBLE_COLUMNS;
            const int visible_column = index % SYNONYM_VISIBLE_COLUMNS;
            const int logical_row = synonym_first_row + visible_row;
            const int logical_column =
                synonym_first_column + visible_column;
            if (reverse) {
                target_column = logical_column;
                target_row = logical_row - 1;
            } else if (logical_row + 1 < SYNONYM_ROWS) {
                target_column = logical_column;
                target_row = logical_row + 1;
            } else if (logical_column + 1 < SYNONYM_FIELD_COLUMNS) {
                target_column = logical_column + 1;
                target_row = -1;
            }
        }
        commit_synonym_view();
        if (target_column >= 0 &&
            target_column < SYNONYM_FIELD_COLUMNS &&
            target_row >= -1 && target_row < SYNONYM_ROWS) {
            if (target_column < synonym_first_column)
                synonym_first_column = target_column;
            else if (target_column >=
                     synonym_first_column + SYNONYM_VISIBLE_COLUMNS)
                synonym_first_column = target_column -
                    SYNONYM_VISIBLE_COLUMNS + 1;
            if (target_row >= 0) {
                if (target_row < synonym_first_row)
                    synonym_first_row = target_row;
                else if (target_row >=
                         synonym_first_row + SYNONYM_VISIBLE_ROWS)
                    synonym_first_row = target_row -
                        SYNONYM_VISIBLE_ROWS + 1;
            }
            load_synonym_view();
            HWND target = target_row < 0
                ? synonym_headers[static_cast<std::size_t>(
                    target_column - synonym_first_column)]
                : synonym_cells[static_cast<std::size_t>(
                    target_row - synonym_first_row)]
                               [static_cast<std::size_t>(
                    target_column - synonym_first_column)];
            if (IsWindow(target)) {
                SetFocus(target);
                SendMessageW(target, EM_SETSEL, 0, -1);
            }
        } else if (!reverse) {
            HWND save = GetDlgItem(synonym_window, ID_SYNONYM_SAVE);
            if (IsWindow(save)) SetFocus(save);
        }
        return 0;
    }
    if (message == WM_CHAR && wparam == VK_RETURN) return 0;
    if (message == WM_NCDESTROY)
        RemoveWindowSubclass(control, synonym_edit_subclass, subclass_id);
    return DefSubclassProc(control, message, wparam, lparam);
}

bool is_chart_series_marker(const std::string& value) {
    const auto marker = point::normalize_name(value);
    return marker == "series" || marker == "@series";
}

std::wstring widen(const std::string& text) {
    if (text.empty()) return {};
    int size = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
                                   text.data(), static_cast<int>(text.size()),
                                   nullptr, 0);
    if (!size) {
        size = MultiByteToWideChar(CP_ACP, 0, text.data(),
                                   static_cast<int>(text.size()), nullptr, 0);
        std::wstring output(static_cast<std::size_t>(size), L'\0');
        MultiByteToWideChar(CP_ACP, 0, text.data(),
                            static_cast<int>(text.size()), output.data(), size);
        return output;
    }
    std::wstring output(static_cast<std::size_t>(size), L'\0');
    MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(),
                        static_cast<int>(text.size()), output.data(), size);
    return output;
}

std::string narrow(const std::wstring& text) {
    if (text.empty()) return {};
    const int size = WideCharToMultiByte(
        CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
        nullptr, 0, nullptr, nullptr);
    std::string output(static_cast<std::size_t>(size), '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.data(),
                        static_cast<int>(text.size()), output.data(), size,
                        nullptr, nullptr);
    return output;
}

std::wstring control_text(HWND control) {
    const int length = GetWindowTextLengthW(control);
    if (length <= 0) return {};
    std::wstring text(static_cast<std::size_t>(length) + 1, L'\0');
    GetWindowTextW(control, text.data(), length + 1);
    text.resize(static_cast<std::size_t>(length));
    return text;
}

void set_control_text(HWND control, const std::string& value) {
    const auto wide = widen(value);
    if (control_text(control) == wide) return;
    SetWindowTextW(control, wide.c_str());
}

void set_status(const std::wstring& value) {
    SetWindowTextW(status_text, value.c_str());
}

void show_refresh_progress(const std::wstring& status, int percent) {
    set_status(status);
    if (refresh_progress) {
        SendMessageW(refresh_progress, PBM_SETPOS,
                     static_cast<WPARAM>(std::clamp(percent, 0, 100)), 0);
        ShowWindow(refresh_progress, SW_SHOW);
        UpdateWindow(refresh_progress);
    }
    if (status_text) UpdateWindow(status_text);
}

// Searches use UI-owned controls and storage. Service the Windows queue at
// bounded checkpoints so the window keeps painting and remains responsive.
void pump_search_messages() {
    MSG message{};
    while (PeekMessageW(&message, nullptr, 0, 0, PM_REMOVE)) {
        if (message.message == WM_QUIT) {
            PostQuitMessage(static_cast<int>(message.wParam));
            return;
        }
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
}

std::uint64_t cell_key(int row, int column) {
    return static_cast<std::uint64_t>(row) *
           static_cast<std::uint64_t>(GRID_COLUMNS) +
           static_cast<std::uint64_t>(column);
}

std::string stored_cell(int row, int column) {
    const auto found = cell_store.find(cell_key(row, column));
    return found == cell_store.end() ? std::string{} : found->second;
}

void store_cell(int row, int column, const std::string& value) {
    const auto key = cell_key(row, column);
    const auto existing = cell_store.find(key);
    if (value.empty()) {
        if (existing != cell_store.end()) {
            cell_store.erase(existing);
            cell_row_index_needs_rebuild[static_cast<std::size_t>(column)] =
                true;
        }
    } else {
        if (existing == cell_store.end()) {
            auto& rows =
                cell_rows_by_column[static_cast<std::size_t>(column)];
            if (!rows.empty() && row < rows.back())
                cell_row_index_needs_sort[
                    static_cast<std::size_t>(column)] = true;
            rows.push_back(row);
        }
        cell_store[key] = value;
    }
}

void invalidate_cell_row_indexes() {
    cell_row_index_needs_rebuild.fill(true);
}

void clear_cell_store() {
    cell_store.clear();
    pending_identity_resolution_cells.clear();
    for (auto& rows : cell_rows_by_column) rows.clear();
    cell_row_index_needs_sort.fill(false);
    cell_row_index_needs_rebuild.fill(false);
}

const std::vector<int>& indexed_rows_for_column(int column) {
    const auto index = static_cast<std::size_t>(column);
    auto& rows = cell_rows_by_column[index];
    if (cell_row_index_needs_rebuild[index]) {
        rows.clear();
        rows.reserve(cell_store.size() / 4 + 1);
        for (const auto& [key, value] : cell_store) {
            (void)value;
            if (static_cast<int>(
                    key % static_cast<std::uint64_t>(GRID_COLUMNS)) == column) {
                rows.push_back(static_cast<int>(
                    key / static_cast<std::uint64_t>(GRID_COLUMNS)));
            }
        }
        cell_row_index_needs_rebuild[index] = false;
        cell_row_index_needs_sort[index] = true;
    }
    if (cell_row_index_needs_sort[index]) {
        std::sort(rows.begin(), rows.end());
        rows.erase(std::unique(rows.begin(), rows.end()), rows.end());
        cell_row_index_needs_sort[index] = false;
    }
    return rows;
}

bool resolve_identity_name_cell(
        int row, int column, bool block_ambiguous) {
    if (!engine || column < 0 || column >= GRID_COLUMNS) return false;
    const auto field = point::trim(narrow(control_text(
        header_cells[static_cast<std::size_t>(column)])));
    const auto entered = point::trim(stored_cell(row, column));
    if (field.empty() || entered.empty()) return false;
    const auto resolution =
        engine->resolve_identity_from_name(field, entered);
    if (resolution.status ==
        point::IdentityResolutionStatus::Unique) {
        if (resolution.value != entered) {
            store_cell(row, column, resolution.value);
            return true;
        }
        return false;
    }
    if (resolution.status ==
            point::IdentityResolutionStatus::Ambiguous &&
        block_ambiguous) {
        throw std::runtime_error(
            "Full name '" + entered + "' matches multiple people for " +
            field + ". Point did not guess. Enter a unique Employee ID, "
            "Username, or Email value.");
    }
    return false;
}

std::size_t resolve_all_identity_name_inputs() {
    std::vector<std::uint64_t> cells(
        pending_identity_resolution_cells.begin(),
        pending_identity_resolution_cells.end());
    std::size_t converted = 0;
    int last_percent = -1;
    for (std::size_t index = 0; index < cells.size(); ++index) {
        const auto key = cells[index];
        const int row = static_cast<int>(
            key / static_cast<std::uint64_t>(GRID_COLUMNS));
        const int column = static_cast<int>(
            key % static_cast<std::uint64_t>(GRID_COLUMNS));
        if (resolve_identity_name_cell(row, column, true)) ++converted;
        pending_identity_resolution_cells.erase(key);
        if ((index & 31U) == 31U || index + 1 == cells.size()) {
            const int percent = 3 + static_cast<int>(
                9 * (index + 1) /
                std::max<std::size_t>(1, cells.size()));
            if (percent != last_percent) {
                last_percent = percent;
                std::wstringstream progress;
                progress << L"Preparing " << (index + 1) << L" of "
                         << cells.size() << L" entered value(s)...";
                show_refresh_progress(progress.str(), percent);
            }
            pump_search_messages();
        }
    }
    if (converted) {
        load_visible_cells();
        point::append_audit(
            app_root, "NAME_IDENTITY_RESOLVED",
            std::to_string(converted) +
            " unique full-name input(s) converted");
    }
    return converted;
}

int highest_stored_row() {
    int highest = -1;
    for (const auto& entry : cell_store) {
        const int row = static_cast<int>(
            entry.first / static_cast<std::uint64_t>(GRID_COLUMNS));
        highest = std::max(highest, row);
    }
    return highest;
}

void erase_rows_from(int first_row) {
    for (auto iterator = cell_store.begin();
         iterator != cell_store.end();) {
        const int row = static_cast<int>(
            iterator->first /
            static_cast<std::uint64_t>(GRID_COLUMNS));
        if (row >= first_row)
            iterator = cell_store.erase(iterator);
        else
            ++iterator;
    }
    invalidate_cell_row_indexes();
}

void commit_visible_cells() {
    if (internal_cell_update) return;
    const int last_column = std::min(
        GRID_COLUMNS, first_visible_column + visible_grid_columns);
    for (int visible_row = 0; visible_row < visible_grid_rows; ++visible_row) {
        const int logical_row = first_visible_row + visible_row;
        if (logical_row >= LOGICAL_ROWS) break;
        for (int column = first_visible_column;
             column < last_column; ++column) {
            store_cell(
                logical_row, column,
                narrow(control_text(
                    data_cells[static_cast<std::size_t>(visible_row)]
                              [static_cast<std::size_t>(column)])));
        }
    }
}

void load_visible_cells() {
    internal_cell_update = true;
    const int last_column = std::min(
        GRID_COLUMNS, first_visible_column + visible_grid_columns);
    for (int visible_row = 0; visible_row < visible_grid_rows; ++visible_row) {
        const int logical_row = first_visible_row + visible_row;
        for (int column = first_visible_column;
             column < last_column; ++column) {
            const auto value = logical_row < LOGICAL_ROWS
                ? stored_cell(logical_row, column) : std::string{};
            set_control_text(
                data_cells[static_cast<std::size_t>(visible_row)]
                          [static_cast<std::size_t>(column)],
                value);
        }
    }
    internal_cell_update = false;
    // Text may be unchanged while selection, duplicate, missing, criteria,
    // or change-state coloring changed. Repaint only the visible viewport.
    repaint_grid_cells();
}

void update_scrollbar() {
    SCROLLINFO info{sizeof(info)};
    info.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    info.nMin = 0;
    info.nMax = LOGICAL_ROWS - 1;
    info.nPage = static_cast<UINT>(std::max(1, visible_grid_rows));
    info.nPos = first_visible_row;
    SetScrollInfo(grid_scrollbar, SB_CTL, &info, TRUE);
}

void update_column_scrollbar() {
    if (!grid_hscrollbar) return;
    SCROLLINFO info{sizeof(info)};
    info.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    info.nMin = 0;
    info.nMax = GRID_COLUMNS - 1;
    info.nPage = static_cast<UINT>(
        std::max(1, visible_grid_columns));
    const int maximum =
        std::max(0, GRID_COLUMNS - visible_grid_columns);
    first_visible_column =
        std::clamp(first_visible_column, 0, maximum);
    info.nPos = first_visible_column;
    SetScrollInfo(grid_hscrollbar, SB_CTL, &info, TRUE);
}

void scroll_to(int requested_row) {
    const int maximum = LOGICAL_ROWS - visible_grid_rows;
    const int target = std::clamp(requested_row, 0, maximum);
    if (target == first_visible_row) return;
    commit_visible_cells();
    first_visible_row = target;
    load_visible_cells();
    update_scrollbar();
    std::wstringstream message;
    message << L"Rows " << first_visible_row + 1 << L"–"
            << first_visible_row + visible_grid_rows
            << L" of " << LOGICAL_ROWS;
    set_status(message.str());
}

void hide_suggestions() {
    if (main_window) KillTimer(main_window, SUGGESTION_TIMER_ID);
    pending_suggestion_target = nullptr;
    pending_suggestion_id = 0;
    ShowWindow(suggestion_list, SW_HIDE);
    suggestion_target = nullptr;
}

void schedule_suggestions(HWND target, int id) {
    if (internal_cell_update || !main_window) return;
    pending_suggestion_target = target;
    pending_suggestion_id = id;
    KillTimer(main_window, SUGGESTION_TIMER_ID);
    SetTimer(main_window, SUGGESTION_TIMER_ID, 25, nullptr);
}

bool prefix_matches(const std::string& value, const std::string& prefix) {
    const auto normalized_value = point::normalize_name(value);
    const auto normalized_prefix = point::normalize_name(prefix);
    return normalized_prefix.empty() ||
           normalized_value.rfind(normalized_prefix, 0) == 0;
}

bool synonym_search_matches(
        const std::string& value, const std::string& typed) {
    const auto normalized_value = point::normalize_name(value);
    const auto normalized_typed = point::normalize_name(typed);
    return !normalized_typed.empty() &&
           normalized_value.find(normalized_typed) != std::string::npos;
}

std::vector<std::string> current_synonym_field_candidates() {
    std::map<std::string, std::string> unique;
    if (engine) {
        for (const auto& dataset : engine->datasets()) {
            for (const auto& field : dataset.headers) {
                const auto normalized = point::normalize_name(field);
                if (!normalized.empty()) unique.try_emplace(normalized, field);
            }
        }
    }
    for (const auto& group : field_synonym_groups) {
        const auto canonical = point::normalize_name(group.canonical_field);
        if (!canonical.empty())
            unique.try_emplace(canonical, group.canonical_field);
        for (const auto& synonym : group.synonyms) {
            const auto normalized = point::normalize_name(synonym);
            if (!normalized.empty()) unique.try_emplace(normalized, synonym);
        }
    }
    // These defaults are only a startup fallback. Once Inbox reports are
    // indexed, their real headings drive the Synonym Manager suggestions.
    if (unique.empty()) {
        static const std::vector<std::string> defaults{
            "Employee ID", "User ID", "Username", "Email",
            "User Principal Name", "Display Name", "Full Name",
            "First Name", "Last Name", "Computer ID", "Computer Name",
            "Device ID", "Device Name", "Asset ID", "Asset Tag",
            "Serial Number", "Ticket ID", "Ticket Status",
            "Account Status", "Group Name", "Department", "Manager"};
        for (const auto& field : defaults)
            unique.emplace(point::normalize_name(field), field);
    }
    std::vector<std::string> fields;
    fields.reserve(unique.size());
    for (const auto& [normalized, field] : unique) {
        (void)normalized;
        fields.push_back(field);
    }
    return fields;
}

int cell_column_from_id(int id) {
    if (id >= ID_HEADER_BASE && id < ID_HEADER_BASE + GRID_COLUMNS)
        return id - ID_HEADER_BASE;
    if (id >= ID_CELL_BASE &&
        id < ID_CELL_BASE + GRID_CONTROL_ROWS * GRID_COLUMNS)
        return (id - ID_CELL_BASE) % GRID_COLUMNS;
    return -1;
}

bool is_header_id(int id) {
    return id >= ID_HEADER_BASE && id < ID_HEADER_BASE + GRID_COLUMNS;
}

bool is_synonym_header_id(int id) {
    return id >= ID_SYNONYM_HEADER_BASE &&
           id < ID_SYNONYM_HEADER_BASE + SYNONYM_VISIBLE_COLUMNS;
}

bool is_synonym_cell_id(int id) {
    return id >= ID_SYNONYM_CELL_BASE &&
           id < ID_SYNONYM_CELL_BASE +
               SYNONYM_VISIBLE_COLUMNS * SYNONYM_VISIBLE_ROWS;
}

bool grid_position_from_control(HWND control, GridPosition& position) {
    if (!control) return false;
    const int id = GetDlgCtrlID(control);
    if (is_header_id(id)) {
        position.row = -1;
        position.column = id - ID_HEADER_BASE;
        return true;
    }
    if (id >= ID_CELL_BASE &&
        id < ID_CELL_BASE + GRID_CONTROL_ROWS * GRID_COLUMNS) {
        const int visible_index = id - ID_CELL_BASE;
        position.row =
            first_visible_row + visible_index / GRID_COLUMNS;
        position.column = visible_index % GRID_COLUMNS;
        return position.row < LOGICAL_ROWS;
    }
    return false;
}

bool position_in_selection(const GridPosition& position) {
    if (!selection_active) return false;
    const int first_row =
        std::min(selection_anchor.row, selection_end.row);
    const int last_row =
        std::max(selection_anchor.row, selection_end.row);
    const int first_column =
        std::min(selection_anchor.column, selection_end.column);
    const int last_column =
        std::max(selection_anchor.column, selection_end.column);
    return position.row >= first_row && position.row <= last_row &&
           position.column >= first_column &&
           position.column <= last_column;
}

void repaint_grid_cells() {
    const int last_column = std::min(
        GRID_COLUMNS, first_visible_column + visible_grid_columns);
    for (int column = first_visible_column;
         column < last_column; ++column) {
        InvalidateRect(
            header_cells[static_cast<std::size_t>(column)], nullptr, TRUE);
        for (int row = 0; row < visible_grid_rows; ++row) {
            InvalidateRect(
                data_cells[static_cast<std::size_t>(row)]
                          [static_cast<std::size_t>(column)],
                nullptr, TRUE);
        }
    }
}

bool pointer_near_right_edge(HWND control, LPARAM lparam) {
    RECT area{};
    GetClientRect(control, &area);
    const int x = static_cast<short>(LOWORD(lparam));
    return x >= area.right - 6;
}

bool pointer_near_bottom_edge(HWND control, LPARAM lparam) {
    RECT area{};
    GetClientRect(control, &area);
    const int y = static_cast<short>(HIWORD(lparam));
    return y >= area.bottom - 5;
}

void commit_grid_control(HWND control) {
    GridPosition position{};
    if (!grid_position_from_control(control, position) ||
        position.row < 0) {
        return;
    }
    store_cell(
        position.row, position.column,
        narrow(control_text(control)));
}

void show_grid_column(int column) {
    if (column < first_visible_column) {
        commit_visible_cells();
        first_visible_column = column;
        layout(main_window);
        load_visible_cells();
        return;
    }
    if (column >= first_visible_column + visible_grid_columns) {
        commit_visible_cells();
        first_visible_column = column;
        layout(main_window);
        load_visible_cells();
    }
}

void focus_grid_position(
        HWND current, int target_row, int target_column,
        bool extend_selection = false) {
    commit_grid_control(current);
    target_row = std::clamp(target_row, -1, LOGICAL_ROWS - 1);
    target_column = std::clamp(target_column, 0, GRID_COLUMNS - 1);

    if (target_row >= 0 &&
        (target_row < first_visible_row ||
         target_row >= first_visible_row + visible_grid_rows)) {
        const int requested =
            target_row < first_visible_row
            ? target_row
            : target_row - visible_grid_rows + 1;
        scroll_to(requested);
    }
    show_grid_column(target_column);

    GridPosition target{target_row, target_column};
    if (extend_selection) {
        if (!selection_active) {
            GridPosition current_position{};
            if (grid_position_from_control(current, current_position))
                selection_anchor = current_position;
        }
        selection_end = target;
        selection_active = true;
    } else {
        selection_anchor = target;
        selection_end = target;
        selection_active = false;
    }
    repaint_grid_cells();

    HWND target_control = nullptr;
    if (target_row == -1) {
        target_control =
            header_cells[static_cast<std::size_t>(target_column)];
    } else {
        const int visible_row = target_row - first_visible_row;
        if (visible_row >= 0 && visible_row < visible_grid_rows) {
            target_control =
                data_cells[static_cast<std::size_t>(visible_row)]
                          [static_cast<std::size_t>(target_column)];
        }
    }
    if (target_control) {
        SetFocus(target_control);
        SendMessageW(target_control, EM_SETSEL, 0, -1);
    }
}

void autofit_column(int column) {
    if (column < 0 || column >= GRID_COLUMNS) return;
    std::size_t longest = control_text(
        header_cells[static_cast<std::size_t>(column)]).size();
    for (int row = 0; row < visible_grid_rows; ++row) {
        longest = std::max(
            longest,
            control_text(
                data_cells[static_cast<std::size_t>(row)]
                          [static_cast<std::size_t>(column)]).size());
    }
    column_widths[static_cast<std::size_t>(column)] =
        std::clamp(
            28 + static_cast<int>(longest) * 8,
            70, 600);
    layout(main_window);
}

void show_suggestions_for(HWND target, int id);
void accept_suggestion();
void sort_results_by_column(int column);
void clear_selected_range();
std::vector<std::string> selected_headers(
    int column_limit = GRID_COLUMNS);
void copy_selected_range();
void copy_active_cell(HWND focus);
void paste_at_cell(HWND focus);
void select_used_grid(HWND focus);
void split_selected_column(const std::string& delimiter);

void hide_quick_text_transform() {
    quick_transform_cell = nullptr;
    quick_transform_start = 0;
    quick_transform_end = 0;
}

void offer_quick_text_transform(HWND control) {
    GridPosition position{};
    if (!grid_position_from_control(control, position) || position.row < 0) {
        hide_quick_text_transform();
        return;
    }
    DWORD start = 0, end = 0;
    SendMessageW(
        control, EM_GETSEL,
        reinterpret_cast<WPARAM>(&start),
        reinterpret_cast<LPARAM>(&end));
    const DWORD length = static_cast<DWORD>(GetWindowTextLengthW(control));
    if (start == end || end > length || end - start > 512) {
        hide_quick_text_transform();
        return;
    }
    quick_transform_cell = control;
    transform_target_cell = control;
    quick_transform_start = start;
    quick_transform_end = end;
    EnableWindow(quick_keep_left_button, TRUE);
    EnableWindow(quick_text_transform_button, TRUE);
    EnableWindow(quick_keep_right_button, TRUE);
    SetWindowTextW(
        status_text,
        L"Use << to keep the left side, Cut to remove the highlighted "
        L"text, or >> to keep the right side. Right-click for more options.");
}

bool transformation_matches_column(HWND control) {
    GridPosition position{};
    return learned_text_transformation.available &&
        grid_position_from_control(control, position) &&
        position.row >= 0 &&
        position.column == learned_text_transformation.column;
}

std::optional<std::string> apply_learned_rule_to_value(
        const std::string& value,
        const LearnedTextTransformation& rule) {
    switch (rule.operation) {
    case LearnedTextTransformation::Operation::RemoveHighlighted:
        return point::remove_text_pattern(
            value, rule.selected_text, rule.position);
    case LearnedTextTransformation::Operation::SkipLeftKeepRight:
        return point::keep_text_side(
            value, rule.selected_text, point::TextSide::Right);
    case LearnedTextTransformation::Operation::SkipRightKeepLeft:
        return point::keep_text_side(
            value, rule.selected_text, point::TextSide::Left);
    }
    return std::nullopt;
}

void apply_quick_text_transform(
        LearnedTextTransformation::Operation operation) {
    HWND target = quick_transform_cell;
    if (!target || !IsWindow(target)) target = transform_target_cell;
    if (!target || !IsWindow(target)) {
        MessageBoxW(
            main_window,
            L"Select a data cell first, then highlight text or type the "
            L"delimiter/pattern in the transform box.",
            L"Quick Transformation", MB_ICONINFORMATION);
        return;
    }
    GridPosition position{};
    if (!grid_position_from_control(target, position) ||
        position.row < 0) {
        hide_quick_text_transform();
        return;
    }
    const auto wide_value = control_text(target);
    const auto typed_pattern = control_text(transform_pattern_text);
    std::string selected;
    DWORD selected_start = quick_transform_start;
    DWORD selected_end = quick_transform_end;
    if (!typed_pattern.empty()) {
        if (typed_pattern.size() > 32) {
            MessageBoxW(main_window,
                L"The transform pattern must be between 1 and 32 characters.",
                L"Quick Transformation", MB_ICONWARNING);
            return;
        }
        selected = narrow(typed_pattern);
        const auto found = wide_value.find(typed_pattern);
        if (found != std::wstring::npos) {
            selected_start = static_cast<DWORD>(found);
            selected_end = static_cast<DWORD>(found + typed_pattern.size());
        } else {
            selected_start = 1;
            selected_end = 1;
        }
    } else {
        if (selected_end > wide_value.size() ||
            selected_start >= selected_end) {
            MessageBoxW(
                main_window,
                L"Highlight text inside the selected cell or type a "
                L"delimiter/pattern in the small box beside Find Next.",
                L"Quick Transformation", MB_ICONINFORMATION);
            return;
        }
        selected = narrow(wide_value.substr(
            selected_start, selected_end - selected_start));
    }
    const auto position_rule = selected_start == 0
        ? point::TextMatchPosition::Prefix
        : (selected_end == wide_value.size()
            ? point::TextMatchPosition::Suffix
            : point::TextMatchPosition::Anywhere);
    const auto before = narrow(wide_value);
    LearnedTextTransformation candidate;
    candidate.available = true;
    candidate.column = position.column;
    candidate.selected_text = selected;
    candidate.position = position_rule;
    candidate.operation = operation;
    const auto after = apply_learned_rule_to_value(before, candidate);
    if (!after) {
        MessageBoxW(
            main_window,
            L"The selected cell does not contain that delimiter/pattern. "
            L"Point did not change the value.",
            L"Quick Transformation", MB_ICONINFORMATION);
        return;
    }
    learned_text_transformation = {};
    learned_text_transformation = candidate;
    learned_text_transformation.undo_values.push_back(
        {position.row, before});
    store_cell(position.row, position.column, *after);
    set_control_text(target, *after);
    transform_target_cell = target;
    hide_quick_text_transform();
    set_status(
        L"Transformation learned and applied to this cell. Right-click a "
        L"cell in this field to preview, apply to selected rows/all rows, "
        L"or undo.");
    point::append_audit(
        app_root, "TEXT_TRANSFORM_LEARNED",
        "column=" + std::to_string(position.column) +
        "; token_length=" + std::to_string(selected.size()) +
        "; operation=" +
        (operation == LearnedTextTransformation::Operation::SkipLeftKeepRight
            ? "skip_left" :
         operation == LearnedTextTransformation::Operation::SkipRightKeepLeft
            ? "skip_right" : "remove_highlighted") +
        "; scope=single_cell");
}

void update_transform_lens() {
    if (!transform_target_cell || !IsWindow(transform_target_cell)) return;
    GridPosition position{};
    if (!grid_position_from_control(transform_target_cell, position) ||
        position.row < 0) return;
    const auto pattern_wide = control_text(transform_pattern_text);
    if (pattern_wide.empty()) {
        set_status(
            L"Transform Lens ready: enter text or a delimiter, then use "
            L"<<, ✂, or >>. Highlighted text also works.");
        return;
    }
    if (pattern_wide.size() > 32) {
        set_status(L"Transform Lens: pattern is limited to 32 characters.");
        return;
    }
    const auto value = narrow(control_text(transform_target_cell));
    const auto pattern = narrow(pattern_wide);
    const auto left = point::keep_text_side(
        value, pattern, point::TextSide::Left);
    const auto cut = point::remove_text_pattern(
        value, pattern, point::TextMatchPosition::Anywhere);
    const auto right = point::keep_text_side(
        value, pattern, point::TextSide::Right);
    if (!left && !cut && !right) {
        set_status(
            L"Transform Lens: the selected cell does not contain “" +
            pattern_wide + L"”. No value will be changed.");
        return;
    }
    auto preview = [](const std::optional<std::string>& value) {
        if (!value) return std::wstring(L"no match");
        auto text = widen(*value);
        if (text.size() > 30) text = text.substr(0, 27) + L"...";
        return text.empty() ? std::wstring(L"(blank)") : text;
    };
    set_status(
        L"Transform Lens  << “" + preview(left) +
        L"”   ✂ “" + preview(cut) +
        L"”   >> “" + preview(right) + L"”");
}

struct TransformationPreview {
    std::size_t affected = 0;
    std::size_t unmatched = 0;
    std::vector<std::pair<std::string, std::string>> examples;
    std::vector<int> matching_rows;
};

TransformationPreview preview_learned_transformation(
        bool selected_rows_only) {
    commit_visible_cells();
    TransformationPreview preview;
    int first_row = 0;
    int last_row = LOGICAL_ROWS - 1;
    if (selected_rows_only) {
        if (!selection_active)
            throw std::runtime_error(
                "Select one or more grid rows before using selected-row scope");
        first_row = std::max(0, std::min(
            selection_anchor.row, selection_end.row));
        last_row = std::max(
            selection_anchor.row, selection_end.row);
    }
    const auto& indexed_rows = indexed_rows_for_column(
        learned_text_transformation.column);
    const auto first = std::lower_bound(
        indexed_rows.begin(), indexed_rows.end(), first_row);
    const auto last = std::upper_bound(
        indexed_rows.begin(), indexed_rows.end(), last_row);
    for (auto iterator = first; iterator != last; ++iterator) {
        const int row = *iterator;
        const auto before = stored_cell(
            row, learned_text_transformation.column);
        if (point::trim(before).empty()) continue;
        const auto after = apply_learned_rule_to_value(
            before, learned_text_transformation);
        if (after) {
            ++preview.affected;
            preview.matching_rows.push_back(row);
            if (preview.examples.size() < 5)
                preview.examples.push_back({before, *after});
        } else {
            ++preview.unmatched;
        }
    }
    return preview;
}

bool confirm_transformation_preview(
        const TransformationPreview& preview,
        bool selected_rows_only,
        bool ask_to_apply = true) {
    std::wstringstream message;
    message << L"Transformation: ";
    if (learned_text_transformation.operation ==
            LearnedTextTransformation::Operation::SkipLeftKeepRight) {
        message << L"skip everything left of '"
                << widen(learned_text_transformation.selected_text)
                << L"' and keep the right side";
    } else if (learned_text_transformation.operation ==
            LearnedTextTransformation::Operation::SkipRightKeepLeft) {
        message << L"skip everything right of '"
                << widen(learned_text_transformation.selected_text)
                << L"' and keep the left side";
    } else {
        message << L"remove '"
                << widen(learned_text_transformation.selected_text)
                << L"' from "
                << (learned_text_transformation.position ==
                        point::TextMatchPosition::Prefix
                    ? L"the beginning" :
                    learned_text_transformation.position ==
                        point::TextMatchPosition::Suffix
                    ? L"the end" : L"the matching position");
    }
    message
            << L".\r\n\r\n"
            << preview.affected << L" row(s) will change.\r\n"
            << preview.unmatched
            << L" populated row(s) do not match and will remain unchanged."
            << L"\r\n\r\nPreview:\r\n";
    for (const auto& [before, after] : preview.examples)
        message << L"• " << widen(before) << L"  →  "
                << widen(after) << L"\r\n";
    if (ask_to_apply) {
        message << L"\r\nApply to "
                << (selected_rows_only
                    ? L"the selected rows" : L"this field")
                << L"?";
    }
    const int choice = MessageBoxW(
        main_window, message.str().c_str(),
        L"Preview Learned Transformation",
        ask_to_apply
            ? MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2
            : MB_OK | MB_ICONINFORMATION);
    return ask_to_apply && choice == IDYES;
}

void apply_learned_transformation(bool selected_rows_only) {
    if (!learned_text_transformation.available)
        throw std::runtime_error(
            "Highlight text inside a cell and click the ✂ action first");
    const auto preview = preview_learned_transformation(selected_rows_only);
    if (preview.affected == 0)
        throw std::runtime_error(
            "No additional populated rows match the learned transformation");
    if (!confirm_transformation_preview(preview, selected_rows_only)) return;
    for (const int row : preview.matching_rows) {
        const auto before = stored_cell(
            row, learned_text_transformation.column);
        const auto after = apply_learned_rule_to_value(
            before, learned_text_transformation);
        if (!after) continue;
        learned_text_transformation.undo_values.push_back({row, before});
        store_cell(row, learned_text_transformation.column, *after);
    }
    load_visible_cells();
    set_status(
        L"Transformation applied to " +
        std::to_wstring(preview.affected) +
        L" row(s); unmatched rows were preserved. Right-click to undo.");
    point::append_audit(
        app_root, "TEXT_TRANSFORM_APPLIED",
        "column=" + std::to_string(learned_text_transformation.column) +
        "; affected=" + std::to_string(preview.affected) +
        "; unmatched=" + std::to_string(preview.unmatched) +
        "; operation=" +
        (learned_text_transformation.operation ==
                LearnedTextTransformation::Operation::SkipLeftKeepRight
            ? "skip_left" :
         learned_text_transformation.operation ==
                LearnedTextTransformation::Operation::SkipRightKeepLeft
            ? "skip_right" : "remove_highlighted") +
        (selected_rows_only ? "; scope=selected_rows" : "; scope=field"));
}

void undo_learned_transformation() {
    if (!learned_text_transformation.available ||
        learned_text_transformation.undo_values.empty())
        throw std::runtime_error("There is no text transformation to undo");
    for (const auto& [row, value] :
         learned_text_transformation.undo_values) {
        store_cell(row, learned_text_transformation.column, value);
    }
    const auto restored = learned_text_transformation.undo_values.size();
    learned_text_transformation.undo_values.clear();
    load_visible_cells();
    set_status(
        L"Last transformation undone; " +
        std::to_wstring(restored) + L" row(s) restored.");
    point::append_audit(
        app_root, "TEXT_TRANSFORM_UNDO",
        "restored=" + std::to_string(restored));
}

LRESULT CALLBACK suggestion_list_subclass(
        HWND control, UINT message, WPARAM wparam, LPARAM lparam,
        UINT_PTR, DWORD_PTR) {
    if (message == WM_LBUTTONUP) {
        const LRESULT result =
            DefSubclassProc(control, message, wparam, lparam);
        accept_suggestion();
        return result;
    }
    if (message == WM_KEYDOWN && wparam == VK_RETURN) {
        accept_suggestion();
        return 0;
    }
    return DefSubclassProc(control, message, wparam, lparam);
}

LRESULT CALLBACK grid_edit_subclass(
        HWND control, UINT message, WPARAM wparam, LPARAM lparam,
        UINT_PTR, DWORD_PTR) {
    switch (message) {
    case WM_LBUTTONDBLCLK:
        if (is_header_id(GetDlgCtrlID(control))) {
            if (pointer_near_right_edge(control, lparam)) {
                autofit_column(
                    GetDlgCtrlID(control) - ID_HEADER_BASE);
                return 0;
            }
            sort_results_by_column(
                GetDlgCtrlID(control) - ID_HEADER_BASE);
            return 0;
        }
        if (pointer_near_bottom_edge(control, lparam)) {
            grid_row_height = 28;
            layout(main_window);
            return 0;
        }
        break;
    case WM_SETCURSOR: {
        POINT pointer{};
        GetCursorPos(&pointer);
        ScreenToClient(control, &pointer);
        RECT area{};
        GetClientRect(control, &area);
        if (is_header_id(GetDlgCtrlID(control)) &&
            pointer.x >= area.right - 6) {
            SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
            return TRUE;
        }
        if (!is_header_id(GetDlgCtrlID(control)) &&
            pointer.y >= area.bottom - 5) {
            SetCursor(LoadCursor(nullptr, IDC_SIZENS));
            return TRUE;
        }
        break;
    }
    case WM_CONTEXTMENU: {
        constexpr UINT MENU_COPY = 1;
        constexpr UINT MENU_CUT = 2;
        constexpr UINT MENU_PASTE = 3;
        constexpr UINT MENU_CLEAR = 4;
        constexpr UINT MENU_SELECT_USED = 5;
        constexpr UINT MENU_AUTOFIT = 6;
        constexpr UINT MENU_APPLY_FIELD = 7;
        constexpr UINT MENU_APPLY_SELECTED = 8;
        constexpr UINT MENU_PREVIEW_TRANSFORM = 9;
        constexpr UINT MENU_UNDO_TRANSFORM = 10;
        constexpr UINT MENU_REMOVE_SELECTION = 11;
        constexpr UINT MENU_SKIP_LEFT = 12;
        constexpr UINT MENU_SKIP_RIGHT = 13;
        constexpr UINT MENU_SPLIT_SELECTION = 14;
        DWORD text_selection_start = 0;
        DWORD text_selection_end = 0;
        SendMessageW(
            control, EM_GETSEL,
            reinterpret_cast<WPARAM>(&text_selection_start),
            reinterpret_cast<LPARAM>(&text_selection_end));
        const bool has_text_selection =
            !is_header_id(GetDlgCtrlID(control)) &&
            text_selection_start < text_selection_end;
        HMENU menu = CreatePopupMenu();
        if (!menu) return 0;
        AppendMenuW(menu, MF_STRING, MENU_COPY, L"Copy\tCtrl+C");
        AppendMenuW(menu, MF_STRING, MENU_CUT, L"Cut\tCtrl+X");
        AppendMenuW(menu, MF_STRING, MENU_PASTE, L"Paste\tCtrl+V");
        AppendMenuW(menu, MF_STRING, MENU_CLEAR, L"Clear\tDelete");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(
            menu, MF_STRING, MENU_SELECT_USED,
            L"Select Used Grid\tCtrl+A");
        if (has_text_selection) {
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(
                menu, MF_STRING, MENU_REMOVE_SELECTION,
                L"✂  Remove Highlighted Text and Learn Rule");
            AppendMenuW(
                menu, MF_STRING, MENU_SPLIT_SELECTION,
                L"Split into Part Columns (preserve original)");
            AppendMenuW(
                menu, MF_STRING, MENU_SKIP_LEFT,
                L"Skip Left — Keep Text on Right and Learn Rule");
            AppendMenuW(
                menu, MF_STRING, MENU_SKIP_RIGHT,
                L"Skip Right — Keep Text on Left and Learn Rule");
        }
        if (is_header_id(GetDlgCtrlID(control))) {
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(
                menu, MF_STRING, MENU_AUTOFIT,
                L"AutoFit Column");
        } else if (transformation_matches_column(control)) {
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(
                menu, MF_STRING, MENU_PREVIEW_TRANSFORM,
                L"Preview Learned Transformation...");
            AppendMenuW(
                menu, MF_STRING, MENU_APPLY_FIELD,
                L"Apply This Transformation to Field...");
            AppendMenuW(
                menu,
                selection_active ? MF_STRING : MF_GRAYED,
                MENU_APPLY_SELECTED,
                L"Apply This Transformation to Selected Rows...");
            AppendMenuW(
                menu,
                learned_text_transformation.undo_values.empty()
                    ? MF_GRAYED : MF_STRING,
                MENU_UNDO_TRANSFORM,
                L"Undo Last Transformation");
        }
        POINT pointer{
            GET_X_LPARAM(lparam),
            GET_Y_LPARAM(lparam)};
        if (pointer.x == -1 && pointer.y == -1) {
            RECT area{};
            GetWindowRect(control, &area);
            pointer = {area.left + 12, area.top + 12};
        }
        const UINT command = TrackPopupMenu(
            menu,
            TPM_RETURNCMD | TPM_RIGHTBUTTON,
            pointer.x, pointer.y, 0,
            main_window, nullptr);
        DestroyMenu(menu);
        SetFocus(control);
        switch (command) {
        case MENU_COPY:
            if (selection_active) copy_selected_range();
            else copy_active_cell(control);
            break;
        case MENU_CUT:
            if (selection_active) {
                copy_selected_range();
                clear_selected_range();
            } else {
                SendMessageW(control, WM_CUT, 0, 0);
            }
            break;
        case MENU_PASTE:
            paste_at_cell(control);
            break;
        case MENU_CLEAR:
            if (selection_active)
                clear_selected_range();
            else
                SetWindowTextW(control, L"");
            break;
        case MENU_SELECT_USED:
            select_used_grid(control);
            break;
        case MENU_AUTOFIT:
            autofit_column(
                GetDlgCtrlID(control) - ID_HEADER_BASE);
            break;
        case MENU_REMOVE_SELECTION:
            quick_transform_cell = control;
            quick_transform_start = text_selection_start;
            quick_transform_end = text_selection_end;
            apply_quick_text_transform(
                LearnedTextTransformation::Operation::RemoveHighlighted);
            break;
        case MENU_SKIP_LEFT:
            quick_transform_cell = control;
            quick_transform_start = text_selection_start;
            quick_transform_end = text_selection_end;
            apply_quick_text_transform(
                LearnedTextTransformation::Operation::SkipLeftKeepRight);
            break;
        case MENU_SKIP_RIGHT:
            quick_transform_cell = control;
            quick_transform_start = text_selection_start;
            quick_transform_end = text_selection_end;
            apply_quick_text_transform(
                LearnedTextTransformation::Operation::SkipRightKeepLeft);
            break;
        case MENU_SPLIT_SELECTION: {
            const auto text = control_text(control);
            if (text_selection_start < text_selection_end &&
                text_selection_end <= text.size()) {
                split_selected_column(narrow(text.substr(
                    text_selection_start,
                    text_selection_end - text_selection_start)));
            }
            break;
        }
        case MENU_PREVIEW_TRANSFORM:
            try {
                const auto preview = preview_learned_transformation(false);
                confirm_transformation_preview(preview, false, false);
            } catch (const std::exception& ex) {
                MessageBoxW(main_window, widen(ex.what()).c_str(),
                            L"Text Transformation", MB_ICONWARNING);
            }
            break;
        case MENU_APPLY_FIELD:
            try {
                apply_learned_transformation(false);
            } catch (const std::exception& ex) {
                MessageBoxW(main_window, widen(ex.what()).c_str(),
                            L"Text Transformation", MB_ICONWARNING);
            }
            break;
        case MENU_APPLY_SELECTED:
            try {
                apply_learned_transformation(true);
            } catch (const std::exception& ex) {
                MessageBoxW(main_window, widen(ex.what()).c_str(),
                            L"Text Transformation", MB_ICONWARNING);
            }
            break;
        case MENU_UNDO_TRANSFORM:
            try {
                undo_learned_transformation();
            } catch (const std::exception& ex) {
                MessageBoxW(main_window, widen(ex.what()).c_str(),
                            L"Text Transformation", MB_ICONWARNING);
            }
            break;
        }
        return 0;
    }
    case WM_KEYDOWN: {
        GridPosition current{};
        if (!grid_position_from_control(control, current)) break;
        const bool control_down =
            (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool shift_down =
            (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        int target_row = current.row;
        int target_column = current.column;
        bool handled = true;
        switch (wparam) {
        case VK_TAB:
            target_column += shift_down ? -1 : 1;
            if (target_column < 0) {
                target_column = GRID_COLUMNS - 1;
                target_row = std::max(-1, target_row - 1);
            } else if (target_column >= GRID_COLUMNS) {
                target_column = 0;
                target_row = std::min(
                    LOGICAL_ROWS - 1, target_row + 1);
            }
            break;
        case VK_RETURN:
            target_row += shift_down ? -1 : 1;
            if (current.row == -1 && shift_down) target_row = -1;
            break;
        case VK_UP:
            target_row = control_down
                ? (current.row <= 0 ? -1 : 0)
                : current.row - 1;
            break;
        case VK_DOWN:
            target_row = control_down
                ? std::max(0, highest_stored_row())
                : current.row + 1;
            break;
        case VK_LEFT: {
            DWORD start = 0, end = 0;
            SendMessageW(
                control, EM_GETSEL,
                reinterpret_cast<WPARAM>(&start),
                reinterpret_cast<LPARAM>(&end));
            if (!control_down && start != 0) {
                handled = false;
                break;
            }
            target_column = control_down ? 0 : current.column - 1;
            break;
        }
        case VK_RIGHT: {
            DWORD start = 0, end = 0;
            SendMessageW(
                control, EM_GETSEL,
                reinterpret_cast<WPARAM>(&start),
                reinterpret_cast<LPARAM>(&end));
            if (!control_down &&
                end != static_cast<DWORD>(
                    GetWindowTextLengthW(control))) {
                handled = false;
                break;
            }
            const int last_column = std::max(
                0, static_cast<int>(selected_headers().size()) - 1);
            target_column = control_down
                ? last_column : current.column + 1;
            break;
        }
        case VK_PRIOR:
            target_row -= visible_grid_rows;
            break;
        case VK_NEXT:
            target_row += visible_grid_rows;
            break;
        case VK_HOME:
            if (control_down) target_row = -1;
            target_column = 0;
            break;
        case VK_END:
            target_column = std::max(
                0, static_cast<int>(selected_headers().size()) - 1);
            if (control_down)
                target_row = std::max(0, highest_stored_row());
            break;
        case VK_DELETE:
            if (selection_active) {
                clear_selected_range();
                return 0;
            }
            handled = false;
            break;
        default:
            handled = false;
            break;
        }
        if (handled) {
            hide_suggestions();
            focus_grid_position(
                control, target_row, target_column,
                shift_down && wparam != VK_TAB &&
                wparam != VK_RETURN);
            return 0;
        }
        break;
    }
    case WM_CHAR: {
        const LRESULT result =
            DefSubclassProc(control, message, wparam, lparam);
        if (!internal_cell_update) {
            KillTimer(main_window, SUGGESTION_TIMER_ID);
            pending_suggestion_target = nullptr;
            pending_suggestion_id = 0;
            show_suggestions_for(control, GetDlgCtrlID(control));
        }
        return result;
    }
    case WM_LBUTTONDOWN: {
        hide_quick_text_transform();
        hide_suggestions();
        if (is_header_id(GetDlgCtrlID(control)) &&
            pointer_near_right_edge(control, lparam)) {
            resizing_column = true;
            resizing_column_index =
                GetDlgCtrlID(control) - ID_HEADER_BASE;
            POINT pointer{};
            GetCursorPos(&pointer);
            resize_start_screen_x = pointer.x;
            resize_start_width = column_widths[
                static_cast<std::size_t>(resizing_column_index)];
            SetCapture(control);
            return 0;
        }
        if (!is_header_id(GetDlgCtrlID(control)) &&
            pointer_near_bottom_edge(control, lparam)) {
            resizing_rows = true;
            POINT pointer{};
            GetCursorPos(&pointer);
            resize_start_screen_y = pointer.y;
            resize_start_row_height = grid_row_height;
            SetCapture(control);
            return 0;
        }
        GridPosition clicked{};
        if (grid_position_from_control(control, clicked)) {
            if (clicked.row >= 0) transform_target_cell = control;
            selection_anchor = clicked;
            selection_end = clicked;
            selection_candidate = true;
            selection_active = false;
            repaint_grid_cells();
        }
        break;
    }
    case WM_MOUSEMOVE:
        if (resizing_column &&
            resizing_column_index >= 0) {
            POINT pointer{};
            GetCursorPos(&pointer);
            column_widths[
                static_cast<std::size_t>(
                    resizing_column_index)] =
                std::clamp<int>(
                    resize_start_width +
                    static_cast<int>(pointer.x) -
                    resize_start_screen_x,
                    60, 600);
            layout(main_window);
            return 0;
        }
        if (resizing_rows) {
            POINT pointer{};
            GetCursorPos(&pointer);
            grid_row_height = std::clamp<int>(
                resize_start_row_height +
                static_cast<int>(pointer.y) -
                resize_start_screen_y,
                20, 80);
            layout(main_window);
            return 0;
        }
        if (selection_candidate && (wparam & MK_LBUTTON) != 0) {
            POINT point{};
            GetCursorPos(&point);
            HWND beneath = WindowFromPoint(point);
            GridPosition hovered{};
            if (grid_position_from_control(beneath, hovered) &&
                (hovered.row != selection_anchor.row ||
                 hovered.column != selection_anchor.column)) {
                selection_end = hovered;
                selection_active = true;
                repaint_grid_cells();
                return 0;
            }
        }
        break;
    case WM_LBUTTONUP:
        if (resizing_column || resizing_rows) {
            resizing_column = false;
            resizing_column_index = -1;
            resizing_rows = false;
            ReleaseCapture();
            layout(main_window);
            return 0;
        }
        selection_candidate = false;
        if (selection_active) {
            ReleaseCapture();
            repaint_grid_cells();
            return 0;
        }
        {
            const LRESULT result =
                DefSubclassProc(control, message, wparam, lparam);
            offer_quick_text_transform(control);
            return result;
        }
    case WM_CAPTURECHANGED:
        resizing_column = false;
        resizing_column_index = -1;
        resizing_rows = false;
        selection_candidate = false;
        break;
    }
    return DefSubclassProc(control, message, wparam, lparam);
}

std::vector<std::string> matching_fields(const std::string& typed) {
    std::vector<std::string> matches;
    if (!engine) return matches;
    const auto normalized = point::normalize_name(typed);
    for (const auto& field : engine->all_fields()) {
        const auto candidate = point::normalize_name(field);
        if (candidate == normalized) return {field};
        if (!normalized.empty() && candidate.rfind(normalized, 0) == 0)
            matches.push_back(field);
    }
    return matches;
}

void position_suggestions(HWND target) {
    RECT target_rect{};
    GetWindowRect(target, &target_rect);
    const int width = static_cast<int>(target_rect.right - target_rect.left);
    SetWindowPos(
        suggestion_list, HWND_TOPMOST,
        target_rect.left, target_rect.bottom,
        std::max(180, width), 160,
        SWP_NOACTIVATE | SWP_SHOWWINDOW);
}

void show_suggestions_for(HWND target, int id) {
    if (internal_cell_update) return;
    const bool synonym_target =
        is_synonym_header_id(id) || is_synonym_cell_id(id);
    if (!engine && !synonym_target) return;
    const auto prefix = narrow(control_text(target));
    std::vector<std::string> suggestions;

    if ((is_synonym_header_id(id) || is_synonym_cell_id(id)) &&
        point::trim(prefix).empty()) {
        hide_suggestions();
        return;
    }

    if (is_synonym_header_id(id)) {
        for (const auto& field : current_synonym_field_candidates()) {
            if (synonym_search_matches(field, prefix))
                suggestions.push_back(field);
        }
        const auto normalized_prefix = point::normalize_name(prefix);
        std::stable_sort(suggestions.begin(), suggestions.end(),
            [&](const auto& left, const auto& right) {
                const bool left_starts =
                    point::normalize_name(left).rfind(normalized_prefix, 0) == 0;
                const bool right_starts =
                    point::normalize_name(right).rfind(normalized_prefix, 0) == 0;
                if (left_starts != right_starts) return left_starts;
                return point::normalize_name(left) < point::normalize_name(right);
            });
        if (suggestions.size() > 100) suggestions.resize(100);
    } else if (is_synonym_cell_id(id)) {
        for (const auto& field : current_synonym_field_candidates())
            if (synonym_search_matches(field, prefix))
                suggestions.push_back(field);
        std::sort(suggestions.begin(), suggestions.end());
        if (suggestions.size() > 100) suggestions.resize(100);
    } else if (is_header_id(id)) {
        for (const auto& field : engine->all_fields()) {
            if (prefix_matches(field, prefix)) suggestions.push_back(field);
            if (suggestions.size() >= 50) break;
        }
    } else {
        const int column = cell_column_from_id(id);
        if (column < 0) return;
        auto field = narrow(control_text(
            header_cells[static_cast<std::size_t>(column)]));
        const auto matches = matching_fields(field);
        if (matches.size() == 1) field = matches.front();
        suggestions = engine->values_for_field(field, prefix, 50);
    }

    SendMessageW(suggestion_list, LB_RESETCONTENT, 0, 0);
    for (const auto& suggestion : suggestions) {
        if (suggestion.size() > 1024) continue;
        const auto label = widen(suggestion);
        SendMessageW(suggestion_list, LB_ADDSTRING, 0,
                     reinterpret_cast<LPARAM>(label.c_str()));
    }
    if (suggestions.empty()) {
        hide_suggestions();
        return;
    }
    SendMessageW(suggestion_list, LB_SETCURSEL, 0, 0);
    suggestion_target = target;
    position_suggestions(target);
    UpdateWindow(suggestion_list);
    std::wstringstream message;
    message << suggestions.size() << L" suggestion(s) for \""
            << control_text(target) << L"\".";
    set_status(message.str());
}

void accept_suggestion() {
    if (!suggestion_target) return;
    const LRESULT selected =
        SendMessageW(suggestion_list, LB_GETCURSEL, 0, 0);
    if (selected == LB_ERR) return;
    const LRESULT length = SendMessageW(
        suggestion_list, LB_GETTEXTLEN,
        static_cast<WPARAM>(selected), 0);
    if (length == LB_ERR || length < 0 || length > 2048) {
        hide_suggestions();
        return;
    }
    std::wstring buffer(static_cast<std::size_t>(length) + 1, L'\0');
    SendMessageW(suggestion_list, LB_GETTEXT,
                 static_cast<WPARAM>(selected),
                 reinterpret_cast<LPARAM>(buffer.data()));
    buffer.resize(static_cast<std::size_t>(length));
    internal_cell_update = true;
    SetWindowTextW(suggestion_target, buffer.c_str());
    internal_cell_update = false;
    SetFocus(suggestion_target);
    const int end = GetWindowTextLengthW(suggestion_target);
    SendMessageW(
        suggestion_target, EM_SETSEL,
        static_cast<WPARAM>(end), static_cast<LPARAM>(end));
    hide_suggestions();
}

std::vector<std::string> selected_headers(
    int column_limit) {
    std::vector<std::string> entered;
    int last_used = -1;
    column_limit = std::clamp(column_limit, 0, GRID_COLUMNS);
    for (int column = 0; column < column_limit; ++column) {
        auto field = point::trim(narrow(control_text(
            header_cells[static_cast<std::size_t>(column)])));
        if (!field.empty()) last_used = column;
        entered.push_back(std::move(field));
    }
    if (last_used < 0) return {};
    return std::vector<std::string>(
        entered.begin(), entered.begin() + last_used + 1);
}

std::vector<std::string> resolved_headers(
    int column_limit = GRID_COLUMNS) {
    auto headers = selected_headers(column_limit);
    if (headers.empty())
        throw std::runtime_error(
            "Type at least one field name in the heading row");

    internal_cell_update = true;
    try {
        for (std::size_t column = 0; column < headers.size(); ++column) {
            if (point::trim(headers[column]).empty())
                throw std::runtime_error(
                    "Do not leave blank headings between selected columns");
            const auto matches = matching_fields(headers[column]);
            if (matches.empty())
                throw std::runtime_error(
                    "Unknown field: " + headers[column]);
            if (matches.size() > 1) {
                std::string message =
                    "'" + headers[column] + "' matches ";
                for (std::size_t i = 0; i < matches.size(); ++i) {
                    if (i) {
                        message += i + 1 == matches.size()
                            ? " and " : ", ";
                    }
                    message += matches[i];
                    if (i == 4 && matches.size() > 5) {
                        message += ", and more";
                        break;
                    }
                }
                message +=
                    ". Type more letters or choose a suggestion.";
                throw std::runtime_error(message);
            }
            headers[column] = matches.front();
            set_control_text(header_cells[column], headers[column]);
        }
    } catch (...) {
        internal_cell_update = false;
        throw;
    }
    internal_cell_update = false;
    return headers;
}

int first_result_row_for_mode() {
    if (compare_mode) return generated_compare_header ? 0 : 2;
    if (narrow_mode || count_mode) return 1;
    return 0;
}

point::QueryResult snapshot_grid() {
    commit_visible_cells();
    point::QueryResult snapshot;
    snapshot.headers = selected_headers();
    const int first_result_row =
        first_result_row_for_mode();
    const int last_stored_row = highest_stored_row();
    for (int row = first_result_row;
         row <= last_stored_row; ++row) {
        std::vector<std::string> values;
        bool any_value = false;
        for (std::size_t column = 0;
             column < snapshot.headers.size(); ++column) {
            const auto value = stored_cell(
                row, static_cast<int>(column));
            if (!point::trim(value).empty()) any_value = true;
            values.push_back(value);
        }
        if (any_value) snapshot.rows.push_back(std::move(values));
    }
    snapshot.explanation = "Editable Point grid snapshot.";
    return snapshot;
}

void clear_count_detail_columns() {
    if (!generated_count_details ||
        count_detail_first_column < 0 ||
        count_detail_column_count <= 0) {
        return;
    }

    const int first_column = count_detail_first_column;
    const int last_column = std::min(
        GRID_COLUMNS,
        first_column + count_detail_column_count);
    internal_cell_update = true;
    for (int column = first_column;
         column < last_column; ++column) {
        SetWindowTextW(
            header_cells[static_cast<std::size_t>(column)], L"");
    }
    for (auto iterator = cell_store.begin();
         iterator != cell_store.end();) {
        const int column = static_cast<int>(
            iterator->first %
            static_cast<std::uint64_t>(GRID_COLUMNS));
        if (column >= first_column && column < last_column)
            iterator = cell_store.erase(iterator);
        else
            ++iterator;
    }
    invalidate_cell_row_indexes();
    generated_count_details = false;
    count_detail_first_column = -1;
    count_detail_column_count = 0;
    internal_cell_update = false;
}

void clear_grid() {
    hide_quick_text_transform();
    transform_target_cell = nullptr;
    learned_text_transformation = {};
    internal_cell_update = true;
    for (HWND header : header_cells) SetWindowTextW(header, L"");
    clear_cell_store();
    universal_missing_rows.clear();
    universal_duplicate_rows.clear();
    universal_duplicate_dark_rows.clear();
    universal_results_displayed = false;
    universal_pending_lookup_rows.clear();
    universal_result_headers.clear();
    universal_lookup_history.clear();
    generated_count_column = -1;
    generated_count_details = false;
    count_detail_first_column = -1;
    count_detail_column_count = 0;
    generated_analysis_headers = false;
    generated_compare_header = false;
    compare_identity_field_cache.clear();
    compare_fields_cache.clear();
    compare_inputs_cache.clear();
    analysis_key_fields_cache.clear();
    analysis_key_values_cache.clear();
    generated_insight_headers = false;
    insight_fields_cache.clear();
    generated_chart_headers = false;
    chart_fields_cache.clear();
    chart_filter_values_cache.clear();
    generated_change_headers = false;
    change_key_fields_cache.clear();
    first_visible_row = 0;
    first_visible_column = 0;
    internal_cell_update = false;
    load_visible_cells();
    update_scrollbar();
    hide_suggestions();
    last_result = {};
}

void display_data_tools_result(const point::QueryResult& result) {
    if (result.headers.size() > static_cast<std::size_t>(GRID_COLUMNS))
        throw std::runtime_error(
            "Data Tools output exceeds the 256 visible workspace columns");
    if (result.rows.size() > static_cast<std::size_t>(LOGICAL_ROWS))
        throw std::runtime_error(
            "Data Tools output exceeds the 2,000,000-row workspace limit");
    // A transformed grid is a standalone editable result. Leaving Narrow,
    // Count, or Chart active would make later snapshots skip setup rows or
    // reinterpret Part columns as mode inputs.
    narrow_mode = false;
    count_mode = false;
    compare_mode = false;
    analyze_mode = false;
    insight_mode = false;
    chart_mode = false;
    change_mode = false;
    clear_grid();
    internal_cell_update = true;
    for (std::size_t column = 0;
         column < result.headers.size(); ++column) {
        set_control_text(header_cells[column], result.headers[column]);
    }
    for (std::size_t row = 0; row < result.rows.size(); ++row) {
        for (std::size_t column = 0;
             column < result.rows[row].size() &&
             column < static_cast<std::size_t>(GRID_COLUMNS); ++column) {
            store_cell(
                static_cast<int>(row), static_cast<int>(column),
                result.rows[row][column]);
        }
    }
    last_result = result;
    internal_cell_update = false;
    first_visible_row = 0;
    first_visible_column = 0;
    load_visible_cells();
    update_scrollbar();
    update_column_scrollbar();
    update_mode_ui();
    layout(main_window);
}

std::size_t selected_data_tools_column(
        const point::QueryResult& source) {
    const int selected = std::clamp(
        selection_end.column, 0, GRID_COLUMNS - 1);
    if (selected >= static_cast<int>(source.headers.size()))
        throw std::runtime_error(
            "Select a cell or heading in the column first");
    return static_cast<std::size_t>(selected);
}

void apply_data_filter(point::RowFilterOperator operation) {
    try {
        const auto current = snapshot_grid();
        if (current.headers.empty())
            throw std::runtime_error("There is no workspace data to filter");
        if (!data_tools_unfiltered_result)
            data_tools_unfiltered_result = current;
        const auto& source = *data_tools_unfiltered_result;
        const auto column = selected_data_tools_column(source);
        std::string value;
        if (operation != point::RowFilterOperator::IsBlank &&
            operation != point::RowFilterOperator::IsNotBlank) {
            if (selection_end.row >= 0)
                value = stored_cell(selection_end.row,
                                    static_cast<int>(column));
            if (point::trim(value).empty())
                value = narrow(control_text(find_text));
        }
        const auto filtered = point::filter_query_result(
            source, column, operation, value);
        display_data_tools_result(filtered);
        set_status(
            L"Filter applied: " +
            std::to_wstring(filtered.rows.size()) + L" of " +
            std::to_wstring(source.rows.size()) + L" row(s) shown. "
            L"Use Data Tools > Clear Filter to restore all rows.");
        point::append_audit(
            app_root, "DATA_FILTER_APPLIED",
            source.headers[column] + " retained " +
            std::to_string(filtered.rows.size()) + " of " +
            std::to_string(source.rows.size()) + " rows");
    } catch (const std::exception& ex) {
        MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Data Tools — Filter", MB_ICONWARNING);
    }
}

void clear_data_filter() {
    if (!data_tools_unfiltered_result) {
        set_status(L"No Data Tools filter is active.");
        return;
    }
    const auto original = *data_tools_unfiltered_result;
    data_tools_unfiltered_result.reset();
    display_data_tools_result(original);
    set_status(L"Filter cleared; all original rows restored.");
    point::append_audit(app_root, "DATA_FILTER_CLEARED", "all rows restored");
}

void split_selected_column(const std::string& delimiter) {
    try {
        auto source = snapshot_grid();
        if (source.headers.empty())
            throw std::runtime_error("There is no workspace data to split");
        const auto column = selected_data_tools_column(source);
        bool selected_column_has_data = false;
        for (const auto& row : source.rows) {
            if (column < row.size() &&
                !point::trim(row[column]).empty()) {
                selected_column_has_data = true;
                break;
            }
        }
        bool recovered_heading_value = false;
        if (!selected_column_has_data &&
            source.headers[column].find(delimiter) != std::string::npos) {
            // New users sometimes paste their first value into Point's
            // heading row.  When there is no data beneath it and the heading
            // itself contains the requested delimiter, recover it as row 1
            // instead of generating empty "value Part 1" headings.
            const auto pasted_value = source.headers[column];
            source.headers[column] = "Original Value";
            std::vector<std::string> recovered(source.headers.size());
            recovered[column] = pasted_value;
            source.rows.push_back(std::move(recovered));
            recovered_heading_value = true;
        }
        const auto transformed = point::split_query_result_column(
            source, column, delimiter);
        if (transformed.headers.size() >
            static_cast<std::size_t>(GRID_COLUMNS)) {
            throw std::runtime_error(
                "Split would exceed the 256 visible workspace columns");
        }

        std::wstringstream preview;
        preview << L"Split '" << widen(source.headers[column])
                << L"' using '" << widen(delimiter) << L"'.\r\n\r\n"
                << L"The original column will be preserved. New Part columns "
                   L"will be appended.\r\n\r\nPreview:\r\n";
        if (recovered_heading_value) {
            preview << L"Point detected that the value was entered in the "
                       L"heading row and will move it into the first data "
                       L"row.\r\n\r\n";
        }
        const std::size_t first_new = source.headers.size();
        for (std::size_t row = 0;
             row < std::min<std::size_t>(3, transformed.rows.size()); ++row) {
            preview << L"• "
                    << widen(column < source.rows[row].size()
                        ? source.rows[row][column] : std::string{})
                    << L"  →  ";
            for (std::size_t part = first_new;
                 part < transformed.rows[row].size(); ++part) {
                if (part > first_new) preview << L" | ";
                preview << widen(transformed.rows[row][part]);
            }
            preview << L"\r\n";
        }
        preview << L"\r\nApply this split?";
        if (MessageBoxW(
                main_window, preview.str().c_str(),
                L"Data Tools — Split Column Preview",
                MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2) != IDYES) {
            return;
        }
        data_tools_unfiltered_result.reset();
        display_data_tools_result(transformed);
        set_status(
            L"Column split completed. Original values were preserved and "
            L"Part columns were appended.");
        point::append_audit(
            app_root, "DATA_COLUMN_SPLIT",
            source.headers[column] + " delimiter=" + delimiter +
            (recovered_heading_value ? "; recovered_heading_value=true" : ""));
    } catch (const std::exception& ex) {
        MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Data Tools — Split Column", MB_ICONWARNING);
    }
}

// Chart results temporarily replace the user's setup headings with the
// generated category/count headings.  If a generated heading is edited, move
// back to setup state immediately so the newly typed field is not overwritten
// by the old cached field on the next Search.
void begin_chart_heading_edit(
    int column, const std::string& typed_value) {
    if (!chart_mode || !generated_chart_headers ||
        internal_cell_update || column < 0 || column >= GRID_COLUMNS) {
        return;
    }

    internal_cell_update = true;
    for (HWND header : header_cells) SetWindowTextW(header, L"");
    for (std::size_t i = 0;
         i < chart_fields_cache.size() && i < header_cells.size(); ++i) {
        set_control_text(header_cells[i], chart_fields_cache[i]);
    }
    set_control_text(
        header_cells[static_cast<std::size_t>(column)], typed_value);

    clear_cell_store();
    for (std::size_t i = 0;
         i < chart_filter_values_cache.size() &&
         i < static_cast<std::size_t>(GRID_COLUMNS); ++i) {
        if (!chart_filter_values_cache[i].empty()) {
            store_cell(0, static_cast<int>(i),
                       chart_filter_values_cache[i]);
        }
    }
    generated_chart_headers = false;
    first_visible_row = 0;
    last_result = {};
    internal_cell_update = false;
    load_visible_cells();
    update_scrollbar();
    if (chart_window) DestroyWindow(chart_window);
    set_status(L"Chart field changed. Choose any filters, then click Search.");
    repaint_grid_cells();
}

void update_mode_ui() {
    if (!main_window) return;
    SetWindowTextW(
        GetDlgItem(main_window, ID_MODE),
        change_mode ? L"Mode: Change" :
        chart_mode ? L"Mode: Chart" :
        insight_mode ? L"Mode: Insights" :
        analyze_mode ? L"Mode: Analyze" :
        compare_mode ? L"Mode: Compare" :
        count_mode ? L"Mode: Count" :
        narrow_mode ? L"Mode: Narrow" : L"Mode: Universal");
    if (help_text) {
        SetWindowTextW(
            help_text,
            change_mode
            ? L"Change: type 1-3 keys; changed rows are shown in red."
            : chart_mode
            ? L"Chart: blank=graph, @series=colored series, other value=filter."
            : insight_mode
            ? L"Insights: leave headings blank for all fields, or select fields."
            : analyze_mode
            ? L"Analyze: type 1-3 key fields; optionally filter row 1."
            : compare_mode
            ? L"Compare: identity first, then any needed fields; users in rows 1-2."
            : count_mode
            ? L"Count: affected object IDs fill automatically when Count <= 50."
            : narrow_mode
            ? L"Narrow: enter exact AND conditions in the yellow first row."
            : L"Universal: heading=desired output; paste any known value below.");
    }
    repaint_grid_cells();
}

void toggle_query_mode() {
    hide_suggestions();
    commit_visible_cells();
    clear_count_detail_columns();
    internal_cell_update = true;
    if (generated_count_column >= 0 &&
        generated_count_column < GRID_COLUMNS) {
        HWND generated_header =
            header_cells[
                static_cast<std::size_t>(generated_count_column)];
        if (point::normalize_name(
                narrow(control_text(generated_header))) == "count") {
            SetWindowTextW(generated_header, L"");
        }
        generated_count_column = -1;
    }
    if (generated_analysis_headers) {
        for (HWND header : header_cells) SetWindowTextW(header, L"");
        for (std::size_t i = 0;
             i < analysis_key_fields_cache.size() &&
             i < header_cells.size(); ++i) {
            set_control_text(
                header_cells[i], analysis_key_fields_cache[i]);
        }
        generated_analysis_headers = false;
    }
    if (generated_insight_headers) {
        for (HWND header : header_cells) SetWindowTextW(header, L"");
        for (std::size_t i = 0;
             i < insight_fields_cache.size() &&
             i < header_cells.size(); ++i) {
            set_control_text(header_cells[i], insight_fields_cache[i]);
        }
        generated_insight_headers = false;
    }
    if (generated_chart_headers) {
        for (HWND header : header_cells) SetWindowTextW(header, L"");
        for (std::size_t i = 0;
             i < chart_fields_cache.size() &&
             i < header_cells.size(); ++i) {
            set_control_text(header_cells[i], chart_fields_cache[i]);
        }
        generated_chart_headers = false;
    }
    if (generated_change_headers) {
        for (HWND header : header_cells) SetWindowTextW(header, L"");
        for (std::size_t i = 0;
             i < change_key_fields_cache.size() &&
             i < header_cells.size(); ++i) {
            set_control_text(
                header_cells[i], change_key_fields_cache[i]);
        }
        generated_change_headers = false;
    }
    if (generated_compare_header &&
        !compare_identity_field_cache.empty()) {
        for (HWND header : header_cells) SetWindowTextW(header, L"");
        for (std::size_t i = 0;
             i < compare_fields_cache.size() &&
             i < header_cells.size(); ++i) {
            set_control_text(header_cells[i], compare_fields_cache[i]);
        }
        generated_compare_header = false;
    }
    if (!narrow_mode && !count_mode &&
        !compare_mode && !analyze_mode && !insight_mode &&
        !chart_mode && !change_mode) {
        narrow_mode = true;
    } else if (narrow_mode) {
        narrow_mode = false;
        count_mode = true;
    } else if (count_mode) {
        count_mode = false;
        compare_mode = true;
    } else if (compare_mode) {
        compare_mode = false;
        analyze_mode = true;
    } else if (analyze_mode) {
        analyze_mode = false;
        insight_mode = true;
    } else if (insight_mode) {
        insight_mode = false;
        chart_mode = true;
    } else if (chart_mode) {
        chart_mode = false;
        change_mode = true;
    } else {
        change_mode = false;
    }
    clear_cell_store();
    universal_missing_rows.clear();
    universal_duplicate_rows.clear();
    universal_duplicate_dark_rows.clear();
    universal_results_displayed = false;
    universal_pending_lookup_rows.clear();
    universal_result_headers.clear();
    universal_lookup_history.clear();
    first_visible_row = 0;
    last_result = {};
    selection_active = false;
    selection_candidate = false;
    internal_cell_update = false;
    load_visible_cells();
    update_scrollbar();
    update_mode_ui();
    set_status(
        change_mode
        ? L"Change: type 1-3 key headings; compares the retained baseline."
        : chart_mode
        ? L"Chart: blank=graph (1-4), @series=shared series, other=exact filter."
        : insight_mode
        ? L"Insight Agent: optionally select fields, then choose Search."
        : analyze_mode
        ? L"Analyze: type 1-3 key headings; optionally enter exact values "
          L"in row 1."
        : compare_mode
        ? L"Compare mode: use two headings, then enter User A and User B "
          L"down the first column."
        : count_mode
        ? L"Count mode: filled yellow cells are exact filters; blank cells "
          L"are grouped; affected object IDs are automatic for Count <= 50."
        : narrow_mode
        ? L"Narrow mode: filled cells in the yellow row are exact AND "
          L"conditions."
        : L"Universal Lookup: headings are desired outputs; paste any exact "
          L"known value down the first column.");
}

std::string selected_source_name() {
    if (!engine || selected_source_index < 0 ||
        selected_source_index >=
            static_cast<int>(engine->datasets().size())) {
        return {};
    }
    return engine->datasets()[
        static_cast<std::size_t>(selected_source_index)].name;
}

void refresh_engine(bool preserve_workspace) {
    HWND refresh_button = GetDlgItem(main_window, ID_REFRESH);
    if (refresh_running.load()) {
        refresh_cancel_requested.store(true);
        SetWindowTextW(refresh_button, L"Cancelling...");
        EnableWindow(refresh_button, FALSE);
        set_status(L"Cancelling refresh after the current file...");
        return;
    }
    if (refresh_thread.joinable()) refresh_thread.join();
    if (preserve_workspace) commit_visible_cells();
    refresh_running.store(true);
    refresh_cancel_requested.store(false);
    SetWindowTextW(refresh_button, L"Cancel");
    EnableWindow(GetDlgItem(main_window, ID_SEARCH), FALSE);
    EnableWindow(GetDlgItem(main_window, ID_MODE), FALSE);
    EnableWindow(GetDlgItem(main_window, ID_COPY_ALL), FALSE);
    EnableWindow(GetDlgItem(main_window, ID_EXPORT), FALSE);
    show_refresh_progress(L"Refreshing — checking source files...", 5);
    const auto root = app_root;
    const auto synonyms = field_synonym_groups;
    const point::Engine* incremental_source = engine.get();
    refresh_thread = std::thread([
            root, synonyms, incremental_source, preserve_workspace,
            window = main_window]() {
        auto completion = std::make_unique<RefreshCompletion>();
        completion->preserve_workspace = preserve_workspace;
        auto post_progress = [window](int percent, const std::wstring& text) {
            auto update = std::make_unique<RefreshProgressUpdate>();
            update->percent = percent;
            update->text = text;
            if (PostMessageW(
                    window, WM_POINT_REFRESH_PROGRESS, 0,
                    reinterpret_cast<LPARAM>(update.get())))
                update.release();
        };
        try {
            completion->imported = point::prepare_import_sources(
                root / "Inbox", root / "Workspace" / "excel-cache",
                [&](std::size_t current, std::size_t total,
                    const std::filesystem::path& file, bool cached) {
                    const int percent = 10 + static_cast<int>(
                        35 * current / std::max<std::size_t>(1, total));
                    post_progress(
                        percent,
                        std::wstring(cached ? L"Using cached workbook: "
                                            : L"Importing workbook: ") +
                            file.filename().wstring());
                }, [] { return refresh_cancel_requested.load(); });
            if (refresh_cancel_requested.load())
                throw std::runtime_error("Refresh cancelled");
            completion->refreshed = std::make_unique<point::Engine>();
            completion->refreshed->set_field_synonyms(synonyms);
            completion->refreshed->load_files_incremental(
                completion->imported.csv_sources, incremental_source,
                [&](std::size_t current, std::size_t total,
                    const std::filesystem::path& file, bool reused) {
                    const int percent = 50 + static_cast<int>(
                        40 * current / std::max<std::size_t>(1, total));
                    post_progress(
                        percent,
                        std::wstring(reused ? L"Reusing index: "
                                            : L"Indexing file: ") +
                            file.filename().wstring());
                }, [] { return refresh_cancel_requested.load(); });
            post_progress(95, L"Preparing workspace...");
        } catch (const std::exception& ex) {
            completion->cancelled = refresh_cancel_requested.load() ||
                std::string(ex.what()) == "Refresh cancelled";
            if (!completion->cancelled) completion->error = ex.what();
        }
        if (PostMessageW(
                window, WM_POINT_REFRESH_COMPLETE, 0,
                reinterpret_cast<LPARAM>(completion.get())))
            completion.release();
    });
}

void complete_refresh(std::unique_ptr<RefreshCompletion> completion) {
    if (refresh_thread.joinable()) refresh_thread.join();
    refresh_running.store(false);
    HWND refresh_button = GetDlgItem(main_window, ID_REFRESH);
    SetWindowTextW(refresh_button, L"Refresh");
    EnableWindow(refresh_button, TRUE);
    EnableWindow(GetDlgItem(main_window, ID_SEARCH), TRUE);
    EnableWindow(GetDlgItem(main_window, ID_MODE), TRUE);
    EnableWindow(GetDlgItem(main_window, ID_COPY_ALL), TRUE);
    EnableWindow(GetDlgItem(main_window, ID_EXPORT), TRUE);
    ShowWindow(refresh_progress, SW_HIDE);
    if (completion->cancelled) {
        set_status(L"Refresh cancelled; existing data was preserved.");
        return;
    }
    if (!completion->error.empty()) {
        last_import_issue = completion->error;
        set_status(L"Refresh failed; existing data was preserved.");
        MessageBoxW(main_window, widen(completion->error).c_str(),
                    L"Point", MB_ICONERROR);
        return;
    }
    auto& imported = completion->imported;
    last_import_issue = imported.issues.empty() ? std::string{} :
        imported.issues.front();
    auto refreshed = std::move(completion->refreshed);
    try {
        if (!previous_engine && engine &&
            !engine->datasets().empty())
            previous_engine = std::move(engine);
        engine = std::move(refreshed);
        selected_source_index = -1;
        if (completion->preserve_workspace) {
            load_visible_cells();
            update_scrollbar();
        } else {
            clear_grid();
        }
        refresh_input_file_list();
        const auto fields = engine->all_fields();
        std::wstringstream message;
        message << L"Ready — " << engine->datasets().size()
                << L" report(s), " << fields.size() << L" field(s), "
                << engine->relationships().size() << L" relationship(s), "
                << LOGICAL_ROWS << L" workspace rows";
        if (imported.workbook_count)
            message << L", " << imported.workbook_count
                    << L" workbook(s)/" << imported.worksheet_count
                    << L" sheet(s)";
        if (!engine->issues().empty())
            message << L", " << engine->issues().size() << L" rejected";
        if (!imported.issues.empty())
            message << L", " << imported.issues.size()
                    << L" Excel import issue(s)";
        if (previous_engine)
            message << L", Change baseline ready";
        if (imported.cache_reused)
            message << L", unchanged Excel cache reused";
        set_status(message.str());
        point::append_audit(
            app_root, "REFRESH",
            std::to_string(engine->datasets().size()) + " accepted, " +
            std::to_string(engine->issues().size()) + " rejected, " +
            std::to_string(imported.workbook_count) + " workbooks, " +
            std::to_string(imported.issues.size()) + " Excel issues");
        if (!imported.issues.empty()) {
            MessageBoxW(
                main_window, widen(imported.issues.front()).c_str(),
                L"Excel import issue", MB_ICONWARNING);
        }
    } catch (const std::exception& ex) {
        last_import_issue = ex.what();
        set_status(L"Refresh failed");
        MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Point", MB_ICONERROR);
    }
}

bool supported_input_file(const std::filesystem::path& path) {
    // Microsoft Office creates an owner/lock file beside an open workbook.
    // It is not a workbook and is commonly locked by Excel, so never expose it
    // as an importable Point input.
    const auto filename = path.filename().wstring();
    if (filename.rfind(L"~$", 0) == 0) return false;
    auto extension = path.extension().wstring();
    std::transform(extension.begin(), extension.end(), extension.begin(),
        [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    return extension == L".csv" || extension == L".xlsx" ||
           extension == L".xls" || extension == L".xlsm";
}

struct ArchiveScanCompletion {
    std::filesystem::path root;
    std::vector<std::filesystem::path> files;
    std::wstring error;
    bool limited = false;
};

bool input_file_already_imported(const std::filesystem::path& source) {
    std::error_code error;
    const auto destination = app_root / "Inbox" / source.filename();
    return std::filesystem::exists(destination, error);
}

void begin_archive_scan(const std::filesystem::path& root) {
    if (input_archive_scan_running.exchange(true)) {
        MessageBoxW(main_window,
            L"Wait for the current folder or drive scan to finish.",
            L"Available Files", MB_ICONINFORMATION);
        return;
    }
    if (input_archive_scan_thread.joinable())
        input_archive_scan_thread.join();
    input_archive_scan_cancel.store(false);
    input_archive_root = root;
    input_archive_files.clear();
    SendMessageW(input_archive_list, LB_RESETCONTENT, 0, 0);
    SendMessageW(input_archive_list, LB_ADDSTRING, 0,
        reinterpret_cast<LPARAM>(L"Scanning folder and subfolders..."));
    EnableWindow(input_archive_select_button, FALSE);
    set_status(L"Scanning available CSV and Excel files in the background...");

    input_archive_scan_thread = std::thread([root]() {
        auto completion = std::make_unique<ArchiveScanCompletion>();
        completion->root = root;
        try {
            constexpr std::size_t maximum_files = 10'000;
            std::error_code error;
            std::filesystem::recursive_directory_iterator iterator(
                root,
                std::filesystem::directory_options::skip_permission_denied,
                error);
            const std::filesystem::recursive_directory_iterator end;
            while (iterator != end && !input_archive_scan_cancel.load()) {
                if (!error && iterator->is_regular_file(error) &&
                    supported_input_file(iterator->path())) {
                    completion->files.push_back(iterator->path());
                    if (completion->files.size() >= maximum_files) {
                        completion->limited = true;
                        break;
                    }
                }
                iterator.increment(error);
                if (error) error.clear();
            }
            std::sort(completion->files.begin(), completion->files.end(),
                [](const auto& left, const auto& right) {
                    return _wcsicmp(left.wstring().c_str(),
                        right.wstring().c_str()) < 0;
                });
        } catch (const std::exception& ex) {
            completion->error = widen(ex.what());
        }
        if (main_window && IsWindow(main_window)) {
            PostMessageW(main_window, WM_POINT_ARCHIVE_SCAN_COMPLETE, 0,
                reinterpret_cast<LPARAM>(completion.release()));
        }
    });
}

void choose_archive_folder_or_drive() {
    if (input_archive_scan_running.load()) return;
    IFileOpenDialog* dialog = nullptr;
    const HRESULT created = CoCreateInstance(
        CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&dialog));
    if (FAILED(created) || !dialog) {
        MessageBoxW(main_window, L"Point could not open the folder picker.",
            L"Available Files", MB_ICONERROR);
        return;
    }
    DWORD options = 0;
    dialog->GetOptions(&options);
    dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM |
        FOS_PATHMUSTEXIST);
    dialog->SetTitle(L"Select a folder or drive to scan");
    if (SUCCEEDED(dialog->Show(main_window))) {
        IShellItem* item = nullptr;
        if (SUCCEEDED(dialog->GetResult(&item)) && item) {
            PWSTR selected = nullptr;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &selected)) &&
                selected) {
                begin_archive_scan(std::filesystem::path(selected));
                CoTaskMemFree(selected);
            }
            item->Release();
        }
    }
    dialog->Release();
}

void draw_archive_file_item(const DRAWITEMSTRUCT& item) {
    if (item.itemID == static_cast<UINT>(-1)) return;
    const bool valid = item.itemID < input_archive_files.size();
    const bool selected = (item.itemState & ODS_SELECTED) != 0;
    HBRUSH background = CreateSolidBrush(selected
        ? GetSysColor(COLOR_HIGHLIGHT) : RGB(255, 255, 255));
    FillRect(item.hDC, &item.rcItem, background);
    DeleteObject(background);
    SetBkMode(item.hDC, TRANSPARENT);
    SetTextColor(item.hDC, selected
        ? GetSysColor(COLOR_HIGHLIGHTTEXT) : RGB(30, 30, 30));
    RECT text_bounds = item.rcItem;
    text_bounds.left += 8;
    text_bounds.right -= valid ? 62 : 8;
    std::wstring text;
    if (valid) {
        const auto& path = input_archive_files[item.itemID];
        text = path.filename().wstring() + L"   —   " +
            path.parent_path().wstring();
    } else {
        const LRESULT length = SendMessageW(
            item.hwndItem, LB_GETTEXTLEN, item.itemID, 0);
        if (length > 0) {
            text.resize(static_cast<std::size_t>(length) + 1);
            SendMessageW(item.hwndItem, LB_GETTEXT, item.itemID,
                reinterpret_cast<LPARAM>(text.data()));
            text.resize(static_cast<std::size_t>(length));
        }
    }
    DrawTextW(item.hDC, text.c_str(), -1, &text_bounds,
        DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);
    if (valid) {
        RECT action = item.rcItem;
        action.left = action.right - 54;
        InflateRect(&action, -5, -5);
        const bool imported = input_file_already_imported(
            input_archive_files[item.itemID]);
        HBRUSH action_brush = CreateSolidBrush(imported
            ? RGB(220, 240, 220) : RGB(200, 225, 255));
        FillRect(item.hDC, &action, action_brush);
        DeleteObject(action_brush);
        FrameRect(item.hDC, &action,
            static_cast<HBRUSH>(GetStockObject(GRAY_BRUSH)));
        SetTextColor(item.hDC, imported ? RGB(30, 105, 45) : RGB(10, 65, 135));
        DrawTextW(item.hDC, imported ? L"✓" : L"↑", -1, &action,
            DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_NOPREFIX);
    }
    if (item.itemState & ODS_FOCUS) DrawFocusRect(item.hDC, &item.rcItem);
}

LRESULT CALLBACK input_archive_list_subclass(
        HWND window, UINT message, WPARAM wparam, LPARAM lparam,
        UINT_PTR, DWORD_PTR) {
    if (message == WM_LBUTTONUP) {
        POINT point{GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        const LRESULT hit = SendMessageW(window, LB_ITEMFROMPOINT, 0,
            MAKELPARAM(point.x, point.y));
        if (!HIWORD(hit)) {
            const int index = LOWORD(hit);
            RECT bounds{};
            if (index >= 0 &&
                static_cast<std::size_t>(index) < input_archive_files.size() &&
                SendMessageW(window, LB_GETITEMRECT, index,
                    reinterpret_cast<LPARAM>(&bounds)) != LB_ERR &&
                point.x >= bounds.right - 62) {
                import_input_files({input_archive_files[
                    static_cast<std::size_t>(index)]});
                InvalidateRect(window, &bounds, TRUE);
                return 0;
            }
        }
    }
    return DefSubclassProc(window, message, wparam, lparam);
}

std::int64_t current_epoch_seconds() {
    return static_cast<std::int64_t>(std::time(nullptr));
}

std::wstring schedule_next_run_text(std::int64_t epoch) {
    const std::time_t value = static_cast<std::time_t>(epoch);
    std::tm local{};
    if (localtime_s(&local, &value) != 0) return L"Unknown";
    wchar_t buffer[64]{};
    if (!wcsftime(buffer, std::size(buffer), L"%Y-%m-%d %I:%M %p", &local))
        return L"Unknown";
    return buffer;
}

void refresh_schedule_list() {
    if (!IsWindow(schedule_list)) return;
    SendMessageW(schedule_list, LB_RESETCONTENT, 0, 0);
    for (const auto& schedule : auto_import_schedules) {
        std::wstringstream line;
        line << schedule.source.filename().wstring() << L"  |  Every "
             << schedule.interval_hours << L" hour(s)  |  Next: "
             << schedule_next_run_text(schedule.next_run_epoch) << L"  |  "
             << schedule.source.parent_path().wstring();
        const auto text = line.str();
        SendMessageW(schedule_list, LB_ADDSTRING, 0,
            reinterpret_cast<LPARAM>(text.c_str()));
    }
    if (auto_import_schedules.empty())
        SendMessageW(schedule_list, LB_ADDSTRING, 0,
            reinterpret_cast<LPARAM>(L"No automatic imports scheduled"));
}

void save_auto_import_schedules() {
    const auto path = app_root / "Workspace" / "auto-import-schedules.dat";
    const auto temporary = app_root / "Workspace" / "auto-import-schedules.tmp";
    std::ofstream output(temporary, std::ios::binary | std::ios::trunc);
    if (!output) throw std::runtime_error("Unable to save auto-import schedules");
    output << "POINT_AUTO_IMPORT_SCHEDULES_V1\n";
    for (const auto& schedule : auto_import_schedules) {
        output << schedule.interval_hours << ' '
               << schedule.next_run_epoch << ' '
               << std::quoted(narrow(schedule.source.wstring())) << '\n';
    }
    output.flush();
    if (!output) throw std::runtime_error("Unable to complete schedule file");
    output.close();
    point::compliance::protect_file_for_current_user(temporary);
    if (!MoveFileExW(temporary.c_str(), path.c_str(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        std::error_code ignored;
        std::filesystem::remove(temporary, ignored);
        throw std::runtime_error("Unable to publish auto-import schedules");
    }
}

void load_auto_import_schedules() {
    auto_import_schedules.clear();
    const auto path = app_root / "Workspace" / "auto-import-schedules.dat";
    if (!std::filesystem::exists(path)) return;
    const auto protected_text = point::compliance::read_user_protected_file(path);
    std::istringstream input(protected_text);
    std::string signature;
    std::getline(input, signature);
    if (signature != "POINT_AUTO_IMPORT_SCHEDULES_V1")
        throw std::runtime_error("Auto-import schedule configuration is invalid");
    AutoImportSchedule schedule;
    std::string path_utf8;
    while (input >> schedule.interval_hours >> schedule.next_run_epoch >>
           std::quoted(path_utf8)) {
        if (schedule.interval_hours < 1 || schedule.interval_hours > 24)
            continue;
        schedule.source = std::filesystem::path(widen(path_utf8));
        if (!supported_input_file(schedule.source)) continue;
        auto_import_schedules.push_back(schedule);
    }
}

void add_auto_import_schedules() {
    IFileOpenDialog* dialog = nullptr;
    if (FAILED(CoCreateInstance(CLSID_FileOpenDialog, nullptr,
            CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog))) || !dialog) {
        MessageBoxW(main_window, L"Point could not open the file picker.",
            L"Auto Schedule", MB_ICONERROR);
        return;
    }
    DWORD options = 0;
    dialog->GetOptions(&options);
    dialog->SetOptions(options | FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST |
        FOS_ALLOWMULTISELECT);
    const COMDLG_FILTERSPEC filters[] = {
        {L"CSV and Excel files", L"*.csv;*.xlsx;*.xls;*.xlsm"},
        {L"All files", L"*.*"}};
    dialog->SetFileTypes(static_cast<UINT>(std::size(filters)), filters);
    dialog->SetTitle(L"Select files for automatic Inbox import");
    if (SUCCEEDED(dialog->Show(main_window))) {
        IShellItemArray* results = nullptr;
        if (SUCCEEDED(dialog->GetResults(&results)) && results) {
            DWORD count = 0;
            results->GetCount(&count);
            const int selected = static_cast<int>(SendMessageW(
                schedule_interval_combo, CB_GETCURSEL, 0, 0));
            const int interval = static_cast<int>(SendMessageW(
                schedule_interval_combo, CB_GETITEMDATA,
                selected == CB_ERR ? 0 : selected, 0));
            const auto now = current_epoch_seconds();
            std::size_t filename_conflicts = 0;
            for (DWORD index = 0; index < count; ++index) {
                IShellItem* item = nullptr;
                if (FAILED(results->GetItemAt(index, &item)) || !item) continue;
                PWSTR selected_path = nullptr;
                if (SUCCEEDED(item->GetDisplayName(
                        SIGDN_FILESYSPATH, &selected_path)) && selected_path) {
                    const std::filesystem::path source(selected_path);
                    if (supported_input_file(source)) {
                        const auto destination_conflict = std::find_if(
                            auto_import_schedules.begin(),
                            auto_import_schedules.end(),
                            [&](const auto& existing) {
                                return _wcsicmp(
                                    existing.source.filename().wstring().c_str(),
                                    source.filename().wstring().c_str()) == 0 &&
                                    _wcsicmp(existing.source.wstring().c_str(),
                                        source.wstring().c_str()) != 0;
                            });
                        if (destination_conflict !=
                                auto_import_schedules.end()) {
                            ++filename_conflicts;
                            CoTaskMemFree(selected_path);
                            item->Release();
                            continue;
                        }
                        const auto duplicate = std::find_if(
                            auto_import_schedules.begin(),
                            auto_import_schedules.end(),
                            [&](const auto& existing) {
                                return _wcsicmp(existing.source.wstring().c_str(),
                                    source.wstring().c_str()) == 0;
                            });
                        if (duplicate == auto_import_schedules.end()) {
                            auto_import_schedules.push_back({
                                source, std::clamp(interval, 1, 24),
                                now + std::clamp(interval, 1, 24) * 3600LL});
                        } else {
                            duplicate->interval_hours =
                                std::clamp(interval, 1, 24);
                            duplicate->next_run_epoch = now +
                                duplicate->interval_hours * 3600LL;
                        }
                    }
                    CoTaskMemFree(selected_path);
                }
                item->Release();
            }
            results->Release();
            try {
                save_auto_import_schedules();
                refresh_schedule_list();
                set_status(L"Automatic import schedule saved.");
                if (filename_conflicts) {
                    MessageBoxW(main_window,
                        L"One or more files were skipped because another "
                        L"scheduled source uses the same filename. Inbox "
                        L"destinations must remain unambiguous.",
                        L"Auto Schedule", MB_ICONWARNING);
                }
            } catch (const std::exception& ex) {
                MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Auto Schedule", MB_ICONERROR);
            }
        }
    }
    dialog->Release();
}

void remove_selected_auto_schedules() {
    if (auto_import_schedules.empty()) return;
    std::vector<int> selected;
    const int count = static_cast<int>(SendMessageW(
        schedule_list, LB_GETSELCOUNT, 0, 0));
    if (count <= 0) return;
    selected.resize(static_cast<std::size_t>(count));
    SendMessageW(schedule_list, LB_GETSELITEMS, count,
        reinterpret_cast<LPARAM>(selected.data()));
    std::sort(selected.rbegin(), selected.rend());
    for (const int index : selected) {
        if (index >= 0 && static_cast<std::size_t>(index) <
                auto_import_schedules.size())
            auto_import_schedules.erase(auto_import_schedules.begin() + index);
    }
    save_auto_import_schedules();
    refresh_schedule_list();
}

void run_auto_import_schedules(bool selected_only) {
    if (refresh_running.load() || auto_import_schedules.empty()) return;
    std::unordered_set<int> selected;
    if (selected_only) {
        const int count = static_cast<int>(SendMessageW(
            schedule_list, LB_GETSELCOUNT, 0, 0));
        if (count <= 0) return;
        std::vector<int> indexes(static_cast<std::size_t>(count));
        SendMessageW(schedule_list, LB_GETSELITEMS, count,
            reinterpret_cast<LPARAM>(indexes.data()));
        selected.insert(indexes.begin(), indexes.end());
    }
    const auto now = current_epoch_seconds();
    std::size_t imported = 0;
    for (std::size_t index = 0; index < auto_import_schedules.size(); ++index) {
        auto& schedule = auto_import_schedules[index];
        if ((selected_only && !selected.contains(static_cast<int>(index))) ||
            (!selected_only && schedule.next_run_epoch > now))
            continue;
        std::error_code error;
        if (!std::filesystem::exists(schedule.source, error)) {
            schedule.next_run_epoch = now + schedule.interval_hours * 3600LL;
            continue;
        }
        const auto destination = app_root / "Inbox" / schedule.source.filename();
        if (_wcsicmp(schedule.source.wstring().c_str(),
                destination.wstring().c_str()) != 0) {
            std::filesystem::copy_file(schedule.source, destination,
                std::filesystem::copy_options::overwrite_existing, error);
            if (error) {
                schedule.next_run_epoch = now + schedule.interval_hours * 3600LL;
                continue;
            }
        }
        ++imported;
        schedule.next_run_epoch = now + schedule.interval_hours * 3600LL;
    }
    save_auto_import_schedules();
    refresh_schedule_list();
    if (imported) {
        point::append_audit(app_root, "AUTO_IMPORT",
            "files=" + std::to_string(imported));
        set_status(std::to_wstring(imported) +
            L" scheduled file(s) copied to Inbox; refreshing...");
        refresh_engine();
    }
}

std::string input_canonical_field(const std::string& field) {
    const auto normalized = point::normalize_name(field);
    for (const auto& group : field_synonym_groups) {
        const auto canonical = point::normalize_name(group.canonical_field);
        if (normalized == canonical) return canonical;
        for (const auto& synonym : group.synonyms)
            if (normalized == point::normalize_name(synonym))
                return canonical;
    }
    return normalized;
}

void refresh_input_file_list() {
    if (!IsWindow(input_file_list)) return;
    hide_input_columns_popup();
    input_field_sample_cache.clear();
    input_field_tooltip_key.clear();
    if (IsWindow(input_field_tooltip))
        ShowWindow(input_field_tooltip, SW_HIDE);
    SendMessageW(input_file_list, LB_RESETCONTENT, 0, 0);
    std::vector<std::filesystem::path> files;
    std::error_code error;
    const auto inbox = app_root / "Inbox";
    if (std::filesystem::exists(inbox, error)) {
        for (const auto& entry : std::filesystem::directory_iterator(inbox, error)) {
            if (error) break;
            if (entry.is_regular_file(error) && supported_input_file(entry.path()))
                files.push_back(entry.path().filename());
        }
    }
    std::sort(files.begin(), files.end());
    input_common_field_counts.clear();
    if (engine) {
        for (const auto& file : files) {
            std::unordered_set<std::string> fields_in_file;
            for (const auto& header : input_file_headers(file.wstring()))
                fields_in_file.insert(input_canonical_field(narrow(header)));
            for (const auto& field : fields_in_file)
                if (!field.empty()) ++input_common_field_counts[field];
        }
    }
    const auto search = point::normalize_name(
        input_field_search ? narrow(control_text(input_field_search)) : "");
    std::size_t visible_files = 0;
    for (const auto& file : files) {
        if (!search.empty()) {
            bool matches = point::normalize_name(narrow(file.wstring())).find(search) !=
                std::string::npos;
            if (!matches) {
                const auto headers = input_file_headers(file.wstring());
                matches = std::any_of(headers.begin(), headers.end(),
                    [&](const std::wstring& header) {
                        return point::normalize_name(narrow(header)).find(search) !=
                            std::string::npos;
                    });
            }
            if (!matches) continue;
        }
        SendMessageW(input_file_list, LB_ADDSTRING, 0,
            reinterpret_cast<LPARAM>(file.c_str()));
        ++visible_files;
    }
    if (visible_files == 0)
        SendMessageW(input_file_list, LB_ADDSTRING, 0,
            reinterpret_cast<LPARAM>(files.empty()
                ? L"No input files added yet"
                : L"No files match this field search"));
    InvalidateRect(input_file_list, nullptr, TRUE);
}

std::wstring input_file_cache_component(const std::filesystem::path& file) {
    std::wstring result = file.stem().wstring();
    for (wchar_t& ch : result) {
        if (ch < 32 || ch == L'<' || ch == L'>' || ch == L':' ||
            ch == L'"' || ch == L'/' || ch == L'\\' || ch == L'|' ||
            ch == L'?' || ch == L'*')
            ch = L'_';
    }
    while (!result.empty() &&
           (result.back() == L' ' || result.back() == L'.'))
        result.pop_back();
    if (result.empty()) result = L"Sheet";
    if (result.size() > 80) result.resize(80);
    return result;
}

bool dataset_belongs_to_input(
        const std::wstring& filename, const point::DataSet& dataset) {
    const std::filesystem::path input(filename);
    auto extension = input.extension().wstring();
    std::transform(extension.begin(), extension.end(), extension.begin(),
        [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    const auto dataset_filename = dataset.path.filename().wstring();
    if (extension == L".csv")
        return _wcsicmp(
            dataset_filename.c_str(), input.filename().c_str()) == 0;
    const std::wstring marker =
        L"__" + input_file_cache_component(input) + L"__";
    return dataset_filename.find(marker) != std::wstring::npos;
}

std::vector<std::wstring> input_file_headers(const std::wstring& filename) {
    std::vector<std::wstring> headers;
    if (!engine || filename == L"No input files added yet") return headers;
    std::unordered_set<std::string> seen;
    for (const auto& dataset : engine->datasets()) {
        if (!dataset_belongs_to_input(filename, dataset)) continue;
        for (const auto& header : dataset.headers) {
            const auto normalized = input_canonical_field(header);
            if (!normalized.empty() && seen.insert(normalized).second)
                headers.push_back(widen(header));
        }
    }
    return headers;
}

std::pair<std::size_t, std::size_t> input_file_size_summary(
        const std::wstring& filename) {
    std::size_t rows = 0;
    std::size_t sheets = 0;
    if (!engine) return {rows, sheets};
    for (const auto& dataset : engine->datasets()) {
        if (!dataset_belongs_to_input(filename, dataset)) continue;
        rows += dataset.rows.size();
        ++sheets;
    }
    return {rows, sheets};
}

bool input_header_is_common(const std::wstring& header) {
    if (!engine) return false;
    const auto wanted = input_canonical_field(narrow(header));
    const auto found = input_common_field_counts.find(wanted);
    return found != input_common_field_counts.end() && found->second >= 2;
}

std::wstring input_field_sample_text(
        const std::wstring& filename, const std::wstring& header) {
    const std::string canonical = input_canonical_field(narrow(header));
    const std::string cache_key = narrow(filename) + '\x1f' + canonical;
    if (const auto found = input_field_sample_cache.find(cache_key);
        found != input_field_sample_cache.end())
        return found->second;

    std::wstringstream text;
    text << header << L"\r\n";
    if (point::is_highly_sensitive_field(narrow(header))) {
        text << L"Sample preview hidden for a sensitive field.";
        return input_field_sample_cache.emplace(cache_key, text.str())
            .first->second;
    }

    std::vector<std::wstring> samples;
    std::unordered_set<std::string> seen;
    if (engine) {
        for (const auto& dataset : engine->datasets()) {
            if (!dataset_belongs_to_input(filename, dataset)) continue;
            std::size_t column = dataset.headers.size();
            for (std::size_t index = 0; index < dataset.headers.size(); ++index) {
                if (input_canonical_field(dataset.headers[index]) == canonical) {
                    column = index;
                    break;
                }
            }
            if (column >= dataset.headers.size()) continue;
            for (const auto& row : dataset.rows) {
                if (column >= row.size()) continue;
                const auto value = point::trim(row[column]);
                const auto normalized = point::normalize_name(value);
                if (value.empty() || !seen.insert(normalized).second) continue;
                std::wstring sample = widen(value);
                if (sample.size() > 70) sample = sample.substr(0, 67) + L"...";
                samples.push_back(std::move(sample));
                if (samples.size() == 5) break;
            }
            if (samples.size() == 5) break;
        }
    }
    if (samples.empty()) {
        text << L"No non-empty sample values found.";
    } else {
        text << L"Sample values:";
        for (const auto& sample : samples) text << L"\r\n• " << sample;
    }
    return input_field_sample_cache.emplace(cache_key, text.str())
        .first->second;
}

bool input_badge_at_point(
        HDC context, const RECT& item_bounds,
        const std::vector<std::wstring>& headers, POINT point,
        std::wstring& header) {
    int left = std::min(item_bounds.right, item_bounds.left + 228) + 8;
    const int right_limit = item_bounds.right - 8;
    const int top = item_bounds.top + 9;
    const int bottom = item_bounds.bottom - 9;
    for (const auto& candidate : headers) {
        SIZE size{};
        GetTextExtentPoint32W(context, candidate.c_str(),
            static_cast<int>(candidate.size()), &size);
        const int width = std::clamp(static_cast<int>(size.cx) + 18, 62, 170);
        if (left + width + 58 > right_limit) break;
        RECT bounds{left, top, left + width, bottom};
        if (PtInRect(&bounds, point)) {
            header = candidate;
            return true;
        }
        left += width + 6;
    }
    return false;
}

void hide_input_field_tooltip() {
    input_field_tooltip_key.clear();
    if (IsWindow(input_field_tooltip))
        ShowWindow(input_field_tooltip, SW_HIDE);
}

void show_input_field_tooltip(
        const std::wstring& filename, const std::wstring& header,
        POINT client_point) {
    const std::string key = narrow(filename) + '\x1f' +
        input_canonical_field(narrow(header));
    if (key == input_field_tooltip_key) return;
    input_field_tooltip_key = key;
    input_field_tooltip_text = input_field_sample_text(filename, header);
    if (!IsWindow(input_field_tooltip)) {
        input_field_tooltip = CreateWindowExW(
            WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
            L"STATIC", nullptr,
            WS_POPUP | WS_BORDER | SS_LEFT,
            0, 0, 360, 140,
            main_window, nullptr, GetModuleHandleW(nullptr), nullptr);
        if (!input_field_tooltip) return;
        SendMessageW(input_field_tooltip, WM_SETFONT,
            SendMessageW(input_file_list, WM_GETFONT, 0, 0), TRUE);
    }
    SetWindowTextW(input_field_tooltip, input_field_tooltip_text.c_str());

    HDC context = GetDC(input_field_tooltip);
    RECT measured{0, 0, 420, 0};
    HFONT old_font = nullptr;
    if (context) {
        const auto font = reinterpret_cast<HFONT>(
            SendMessageW(input_field_tooltip, WM_GETFONT, 0, 0));
        if (font) old_font = reinterpret_cast<HFONT>(
            SelectObject(context, font));
        DrawTextW(context, input_field_tooltip_text.c_str(), -1, &measured,
            DT_CALCRECT | DT_LEFT | DT_WORDBREAK | DT_NOPREFIX);
        if (old_font) SelectObject(context, old_font);
        ReleaseDC(input_field_tooltip, context);
    }
    const int width = std::clamp(
        static_cast<int>(measured.right - measured.left) + 24, 240, 440);
    const int height = std::clamp(
        static_cast<int>(measured.bottom - measured.top) + 20, 70, 220);
    ClientToScreen(input_file_list, &client_point);
    POINT anchor{client_point.x + 16, client_point.y + 22};
    MONITORINFO monitor{sizeof(monitor)};
    GetMonitorInfoW(
        MonitorFromPoint(anchor, MONITOR_DEFAULTTONEAREST), &monitor);
    if (anchor.x + width > monitor.rcWork.right)
        anchor.x = std::max(monitor.rcWork.left, client_point.x - width - 12);
    if (anchor.y + height > monitor.rcWork.bottom)
        anchor.y = std::max(monitor.rcWork.top, client_point.y - height - 12);
    SetWindowPos(input_field_tooltip, HWND_TOPMOST,
        anchor.x, anchor.y, width, height,
        SWP_NOACTIVATE | SWP_SHOWWINDOW);
    RedrawWindow(input_field_tooltip, nullptr, nullptr,
        RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
}

struct InputMoreBadge {
    RECT bounds{};
    std::size_t first_hidden = 0;
    bool visible = false;
};

InputMoreBadge input_more_badge(
        HDC context, const RECT& item_bounds,
        const std::vector<std::wstring>& headers) {
    InputMoreBadge result;
    int left = std::min(item_bounds.right, item_bounds.left + 228) + 8;
    const int right_limit = item_bounds.right - 8;
    for (const auto& header : headers) {
        SIZE size{};
        GetTextExtentPoint32W(context, header.c_str(),
            static_cast<int>(header.size()), &size);
        const int width = std::clamp(
            static_cast<int>(size.cx) + 18, 62, 170);
        if (left + width + 58 > right_limit) break;
        left += width + 6;
        ++result.first_hidden;
    }
    if (result.first_hidden < headers.size() && left + 52 <= right_limit) {
        result.bounds = {
            left, item_bounds.top + 9,
            std::min(left + 52, right_limit), item_bounds.bottom - 9};
        result.visible = true;
    }
    return result;
}

void hide_input_columns_popup() {
    if (IsWindow(input_columns_popup) &&
        IsWindowVisible(input_columns_popup)) {
        ShowWindow(input_columns_popup, SW_HIDE);
        if (IsWindow(input_file_list)) {
            InvalidateRect(input_file_list, nullptr, TRUE);
            UpdateWindow(input_file_list);
        }
    }
}

void open_imported_input_file(const std::wstring& filename) {
    const std::filesystem::path name(filename);
    if (name != name.filename() || !supported_input_file(name)) {
        MessageBoxW(
            main_window, L"Select an imported CSV or Excel file first.",
            L"Open File", MB_OK | MB_ICONINFORMATION);
        return;
    }
    const auto path = app_root / "Inbox" / name;
    const DWORD attributes = GetFileAttributesW(path.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES ||
        (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 ||
        (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
        MessageBoxW(
            main_window,
            L"Point could not open this file because it is missing or is "
            L"not a regular imported file.",
            L"Open File", MB_OK | MB_ICONWARNING);
        return;
    }
    const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(
        main_window, L"open", path.c_str(), nullptr,
        path.parent_path().c_str(), SW_SHOWNORMAL));
    if (result <= 32) {
        MessageBoxW(
            main_window,
            L"Windows could not open this file. Confirm that Excel or "
            L"another application is associated with this file type.",
            L"Open File", MB_OK | MB_ICONERROR);
        point::append_audit(
            app_root, "INPUT_FILE_OPEN_BLOCKED",
            "extension=" + narrow(name.extension().wstring()) +
            "; shell_error=" + std::to_string(result));
        return;
    }
    set_status(L"Opened imported file: " + name.filename().wstring());
    point::append_audit(
        app_root, "INPUT_FILE_OPENED",
        "extension=" + narrow(name.extension().wstring()));
}

LRESULT CALLBACK input_popup_subclass(
        HWND window, UINT message, WPARAM wparam, LPARAM lparam,
        UINT_PTR, DWORD_PTR) {
    if (message == WM_KEYDOWN && wparam == VK_ESCAPE) {
        hide_input_columns_popup();
        SetFocus(input_file_list);
        return 0;
    }
    if (message == WM_KILLFOCUS)
        hide_input_columns_popup();
    return DefSubclassProc(window, message, wparam, lparam);
}

void show_remaining_input_columns(
        const std::wstring& filename, const std::vector<std::wstring>& headers,
        std::size_t first_hidden, RECT badge_bounds) {
    if (first_hidden >= headers.size()) return;
    if (!IsWindow(input_columns_popup)) {
        input_columns_popup = CreateWindowExW(
            WS_EX_TOOLWINDOW | WS_EX_TOPMOST, L"STATIC", nullptr,
            WS_POPUP | WS_CAPTION | WS_BORDER | WS_CLIPCHILDREN,
            0, 0, 320, 240, main_window, nullptr, nullptr, nullptr);
        if (!input_columns_popup) return;
        input_columns_popup_list = CreateWindowExW(
            WS_EX_CLIENTEDGE, L"LISTBOX", nullptr,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL |
                LBS_NOINTEGRALHEIGHT,
            8, 8, 296, 200, input_columns_popup, nullptr, nullptr, nullptr);
        if (!input_columns_popup_list) {
            DestroyWindow(input_columns_popup);
            input_columns_popup = nullptr;
            return;
        }
        SendMessageW(input_columns_popup_list, WM_SETFONT,
            SendMessageW(input_file_list, WM_GETFONT, 0, 0), TRUE);
        SetWindowSubclass(
            input_columns_popup_list, input_popup_subclass, 1, 0);
    }
    SetWindowTextW(input_columns_popup,
        (L"Remaining columns — " + filename).c_str());
    SendMessageW(input_columns_popup_list, LB_RESETCONTENT, 0, 0);
    int horizontal_extent = 320;
    HDC context = GetDC(input_columns_popup_list);
    for (std::size_t index = first_hidden; index < headers.size(); ++index) {
        const std::wstring line =
            (input_header_is_common(headers[index]) ? L"●  " : L"   ") +
            headers[index];
        SendMessageW(input_columns_popup_list, LB_ADDSTRING, 0,
            reinterpret_cast<LPARAM>(line.c_str()));
        if (context) {
            SIZE size{};
            GetTextExtentPoint32W(context, line.c_str(),
                static_cast<int>(line.size()), &size);
            horizontal_extent = std::max(
                horizontal_extent, static_cast<int>(size.cx) + 28);
        }
    }
    if (context) ReleaseDC(input_columns_popup_list, context);
    SendMessageW(input_columns_popup_list, LB_SETHORIZONTALEXTENT,
        static_cast<WPARAM>(horizontal_extent), 0);

    POINT anchor{badge_bounds.right + 4, badge_bounds.top};
    ClientToScreen(input_file_list, &anchor);
    const int remaining = static_cast<int>(headers.size() - first_hidden);
    const int width = std::clamp(horizontal_extent + 24, 280, 520);
    const int height = std::clamp(remaining * 26 + 44, 110, 360);
    MONITORINFO monitor{sizeof(monitor)};
    GetMonitorInfoW(MonitorFromPoint(anchor, MONITOR_DEFAULTTONEAREST), &monitor);
    if (anchor.x + width > monitor.rcWork.right) {
        POINT left_anchor{badge_bounds.left - width - 4, badge_bounds.top};
        ClientToScreen(input_file_list, &left_anchor);
        anchor.x = std::max(monitor.rcWork.left, left_anchor.x);
    }
    if (anchor.y + height > monitor.rcWork.bottom)
        anchor.y = std::max(monitor.rcWork.top, monitor.rcWork.bottom - height);
    SetWindowPos(input_columns_popup, HWND_TOPMOST,
        anchor.x, anchor.y, width, height,
        SWP_SHOWWINDOW);
    RECT client{};
    GetClientRect(input_columns_popup, &client);
    MoveWindow(input_columns_popup_list, 8, 8,
        std::max(80L, client.right - 16),
        std::max(50L, client.bottom - 16), TRUE);
    RedrawWindow(input_columns_popup, nullptr, nullptr,
        RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    SetFocus(input_columns_popup_list);
}

LRESULT CALLBACK input_file_list_subclass(
        HWND window, UINT message, WPARAM wparam, LPARAM lparam,
        UINT_PTR, DWORD_PTR) {
    if (message == WM_CONTEXTMENU) {
        hide_input_field_tooltip();
        hide_input_columns_popup();
        POINT screen_point{
            GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        int index = static_cast<int>(SendMessageW(
            window, LB_GETCARETINDEX, 0, 0));
        if (screen_point.x != -1 || screen_point.y != -1) {
            POINT client_point = screen_point;
            ScreenToClient(window, &client_point);
            const LRESULT hit = SendMessageW(
                window, LB_ITEMFROMPOINT, 0,
                MAKELPARAM(client_point.x, client_point.y));
            if (HIWORD(hit)) return 0;
            index = LOWORD(hit);
        } else if (index != LB_ERR) {
            RECT bounds{};
            if (SendMessageW(window, LB_GETITEMRECT, index,
                    reinterpret_cast<LPARAM>(&bounds)) != LB_ERR) {
                screen_point = {bounds.left + 16, bounds.top + 16};
                ClientToScreen(window, &screen_point);
            }
        }
        if (index == LB_ERR) return 0;
        const LRESULT length = SendMessageW(
            window, LB_GETTEXTLEN, index, 0);
        if (length <= 0 || length > 32'767) return 0;
        std::wstring filename(
            static_cast<std::size_t>(length) + 1, L'\0');
        SendMessageW(window, LB_GETTEXT, index,
            reinterpret_cast<LPARAM>(filename.data()));
        filename.resize(static_cast<std::size_t>(length));
        if (!supported_input_file(std::filesystem::path(filename))) return 0;
        SendMessageW(window, LB_SETCARETINDEX, index, FALSE);

        constexpr UINT MENU_OPEN_FILE = 1;
        constexpr UINT MENU_VIEW_COLUMNS = 2;
        HMENU menu = CreatePopupMenu();
        if (!menu) return 0;
        AppendMenuW(menu, MF_STRING, MENU_OPEN_FILE, L"Open File");
        AppendMenuW(menu, MF_STRING, MENU_VIEW_COLUMNS, L"View Columns");
        SetForegroundWindow(main_window);
        const UINT command = TrackPopupMenu(
            menu, TPM_RETURNCMD | TPM_RIGHTBUTTON,
            screen_point.x, screen_point.y, 0, main_window, nullptr);
        DestroyMenu(menu);
        if (command == MENU_OPEN_FILE)
            open_imported_input_file(filename);
        else if (command == MENU_VIEW_COLUMNS)
            show_input_file_columns();
        return 0;
    }
    if (message == WM_MOUSEMOVE) {
        TRACKMOUSEEVENT tracking{sizeof(tracking), TME_LEAVE, window, 0};
        TrackMouseEvent(&tracking);
        POINT point{GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        const LRESULT hit = SendMessageW(window, LB_ITEMFROMPOINT, 0,
            MAKELPARAM(point.x, point.y));
        bool shown = false;
        if (!HIWORD(hit)) {
            const int index = LOWORD(hit);
            RECT item_bounds{};
            if (SendMessageW(window, LB_GETITEMRECT, index,
                    reinterpret_cast<LPARAM>(&item_bounds)) != LB_ERR) {
                const LRESULT length = SendMessageW(window, LB_GETTEXTLEN, index, 0);
                if (length > 0) {
                    std::wstring filename(static_cast<std::size_t>(length) + 1, L'\0');
                    SendMessageW(window, LB_GETTEXT, index,
                        reinterpret_cast<LPARAM>(filename.data()));
                    filename.resize(static_cast<std::size_t>(length));
                    const auto headers = input_file_headers(filename);
                    HDC context = GetDC(window);
                    HFONT old_font = nullptr;
                    if (context) {
                        const auto font = reinterpret_cast<HFONT>(
                            SendMessageW(window, WM_GETFONT, 0, 0));
                        if (font) old_font = reinterpret_cast<HFONT>(
                            SelectObject(context, font));
                        std::wstring header;
                        if (input_badge_at_point(
                                context, item_bounds, headers, point, header)) {
                            show_input_field_tooltip(filename, header, point);
                            shown = true;
                        }
                        if (old_font) SelectObject(context, old_font);
                        ReleaseDC(window, context);
                    }
                }
            }
        }
        if (!shown) hide_input_field_tooltip();
    } else if (message == WM_MOUSELEAVE) {
        hide_input_field_tooltip();
    } else if (message == WM_LBUTTONUP) {
        hide_input_field_tooltip();
        const LRESULT normal = DefSubclassProc(window, message, wparam, lparam);
        POINT point{GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        const LRESULT hit = SendMessageW(window, LB_ITEMFROMPOINT, 0,
            MAKELPARAM(point.x, point.y));
        if (!HIWORD(hit)) {
            const int index = LOWORD(hit);
            RECT item_bounds{};
            if (SendMessageW(window, LB_GETITEMRECT, index,
                    reinterpret_cast<LPARAM>(&item_bounds)) != LB_ERR) {
                const LRESULT length = SendMessageW(
                    window, LB_GETTEXTLEN, index, 0);
                if (length > 0) {
                    std::wstring filename(
                        static_cast<std::size_t>(length) + 1, L'\0');
                    SendMessageW(window, LB_GETTEXT, index,
                        reinterpret_cast<LPARAM>(filename.data()));
                    filename.resize(static_cast<std::size_t>(length));
                    const auto headers = input_file_headers(filename);
                    HDC context = GetDC(window);
                    HFONT old_font = nullptr;
                    if (context) {
                        const auto font = reinterpret_cast<HFONT>(
                            SendMessageW(window, WM_GETFONT, 0, 0));
                        if (font) old_font = reinterpret_cast<HFONT>(
                            SelectObject(context, font));
                    }
                    const auto badge = context
                        ? input_more_badge(context, item_bounds, headers)
                        : InputMoreBadge{};
                    if (context && old_font) SelectObject(context, old_font);
                    if (context) ReleaseDC(window, context);
                    if (badge.visible && PtInRect(&badge.bounds, point)) {
                        show_remaining_input_columns(
                            filename, headers, badge.first_hidden, badge.bounds);
                        return normal;
                    }
                }
            }
        }
        hide_input_columns_popup();
        return normal;
    }
    if (message == WM_KEYDOWN && wparam == VK_ESCAPE)
        hide_input_columns_popup();
    return DefSubclassProc(window, message, wparam, lparam);
}

void show_input_file_columns() {
    const LRESULT index = SendMessageW(input_file_list, LB_GETCARETINDEX, 0, 0);
    if (index == LB_ERR) return;
    const LRESULT length = SendMessageW(input_file_list, LB_GETTEXTLEN, index, 0);
    if (length <= 0) return;
    std::wstring filename(static_cast<std::size_t>(length) + 1, L'\0');
    SendMessageW(input_file_list, LB_GETTEXT, index,
        reinterpret_cast<LPARAM>(filename.data()));
    filename.resize(static_cast<std::size_t>(length));
    if (!supported_input_file(std::filesystem::path(filename))) return;
    const auto headers = input_file_headers(filename);
    const auto [rows, sheets] = input_file_size_summary(filename);
    std::wstringstream report;
    report << filename << L"\r\n\r\n" << rows << L" row(s), "
           << headers.size() << L" field(s)";
    if (sheets > 1) report << L", " << sheets << L" worksheet(s)";
    report << L"\r\n\r\nColumns:\r\n";
    const std::size_t limit = std::min<std::size_t>(headers.size(), 100);
    for (std::size_t i = 0; i < limit; ++i)
        report << (input_header_is_common(headers[i]) ? L"● " : L"  ")
               << headers[i] << L"\r\n";
    if (headers.size() > limit)
        report << L"...and " << headers.size() - limit << L" more";
    if (headers.empty()) report << L"Headers unavailable";
    report << L"\r\n\r\n● = common field in multiple files";
    MessageBoxW(main_window, report.str().c_str(), L"Input File Columns",
        MB_OK | MB_ICONINFORMATION);
}

void draw_input_file_item(const DRAWITEMSTRUCT& item) {
    if (item.itemID == static_cast<UINT>(-1)) return;
    std::wstring filename;
    const LRESULT length = SendMessageW(
        item.hwndItem, LB_GETTEXTLEN, item.itemID, 0);
    if (length > 0) {
        filename.assign(static_cast<std::size_t>(length) + 1, L'\0');
        SendMessageW(item.hwndItem, LB_GETTEXT, item.itemID,
            reinterpret_cast<LPARAM>(filename.data()));
        filename.resize(static_cast<std::size_t>(length));
    }
    const bool selected = (item.itemState & ODS_SELECTED) != 0;
    const COLORREF row_background = selected
        ? GetSysColor(COLOR_HIGHLIGHT) : RGB(255, 255, 255);
    HBRUSH background = CreateSolidBrush(row_background);
    FillRect(item.hDC, &item.rcItem, background);
    DeleteObject(background);
    SetBkMode(item.hDC, TRANSPARENT);
    SetTextColor(item.hDC, selected
        ? GetSysColor(COLOR_HIGHLIGHTTEXT) : RGB(25, 25, 25));
    RECT name_rect = item.rcItem;
    name_rect.left += 8;
    name_rect.right = std::min(name_rect.right, name_rect.left + 220);
    name_rect.bottom = name_rect.top + 23;
    DrawTextW(item.hDC, filename.c_str(), -1, &name_rect,
        DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);

    if (supported_input_file(std::filesystem::path(filename))) {
        const auto headers = input_file_headers(filename);
        const auto [rows, sheets] = input_file_size_summary(filename);
        std::wstring summary = std::to_wstring(rows) + L" rows • " +
            std::to_wstring(headers.size()) + L" fields";
        if (sheets > 1)
            summary += L" • " + std::to_wstring(sheets) + L" sheets";
        RECT summary_rect = item.rcItem;
        summary_rect.left += 8;
        summary_rect.right = name_rect.right;
        summary_rect.top += 22;
        SetTextColor(item.hDC, selected
            ? RGB(235, 242, 252) : RGB(95, 95, 95));
        DrawTextW(item.hDC, summary.c_str(), -1, &summary_rect,
            DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        int left = name_rect.right + 8;
        const int right_limit = item.rcItem.right - 8;
        const int top = item.rcItem.top + 9;
        const int bottom = item.rcItem.bottom - 9;
        std::size_t shown = 0;
        for (const auto& header : headers) {
            SIZE size{};
            GetTextExtentPoint32W(item.hDC, header.c_str(),
                static_cast<int>(header.size()), &size);
            const int badge_width = std::clamp(
                static_cast<int>(size.cx) + 18, 62, 170);
            if (left + badge_width + 58 > right_limit) break;
            RECT badge{left, top, left + badge_width, bottom};
            const bool common = input_header_is_common(header);
            HBRUSH badge_brush = CreateSolidBrush(common
                ? RGB(185, 218, 255) : RGB(235, 235, 235));
            FillRect(item.hDC, &badge, badge_brush);
            DeleteObject(badge_brush);
            FrameRect(item.hDC, &badge,
                static_cast<HBRUSH>(GetStockObject(GRAY_BRUSH)));
            SetTextColor(item.hDC,
                common ? RGB(10, 65, 135) : RGB(70, 70, 70));
            InflateRect(&badge, -7, 0);
            DrawTextW(item.hDC, header.c_str(), -1, &badge,
                DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);
            left += badge_width + 6;
            ++shown;
        }
        if (shown < headers.size() && left + 52 <= right_limit) {
            const std::wstring more =
                L"+" + std::to_wstring(headers.size() - shown);
            RECT badge{left, top, std::min(left + 52, right_limit), bottom};
            HBRUSH badge_brush = CreateSolidBrush(RGB(225, 225, 225));
            FillRect(item.hDC, &badge, badge_brush);
            DeleteObject(badge_brush);
            FrameRect(item.hDC, &badge,
                static_cast<HBRUSH>(GetStockObject(GRAY_BRUSH)));
            SetTextColor(item.hDC, RGB(55, 55, 55));
            DrawTextW(item.hDC, more.c_str(), -1, &badge,
                DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_NOPREFIX);
        } else if (headers.empty() && left + 145 <= right_limit) {
            const bool indexing = refresh_running.load();
            RECT badge{left, top, left + 145, bottom};
            HBRUSH badge_brush = CreateSolidBrush(
                indexing ? RGB(255, 244, 190) : RGB(255, 230, 230));
            FillRect(item.hDC, &badge, badge_brush);
            DeleteObject(badge_brush);
            FrameRect(item.hDC, &badge,
                static_cast<HBRUSH>(GetStockObject(GRAY_BRUSH)));
            SetTextColor(item.hDC,
                indexing ? RGB(110, 80, 0) : RGB(150, 25, 25));
            DrawTextW(item.hDC,
                indexing ? L"Indexing workbook..." : L"Headers unavailable",
                -1, &badge,
                DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_NOPREFIX);
        }
    }
    if (item.itemState & ODS_FOCUS)
        DrawFocusRect(item.hDC, &item.rcItem);
}

void import_input_files(const std::vector<std::filesystem::path>& files) {
    if (files.empty()) return;
    if (refresh_running.load()) {
        MessageBoxW(main_window,
            L"Wait for the current refresh to finish before adding files.",
            L"Point Input Files", MB_ICONINFORMATION);
        return;
    }
    const auto inbox = app_root / "Inbox";
    std::filesystem::create_directories(inbox);
    std::size_t imported = 0, rejected = 0;
    for (const auto& source : files) {
        if (!supported_input_file(source) || !std::filesystem::is_regular_file(source)) {
            ++rejected;
            continue;
        }
        const auto destination = inbox / source.filename();
        std::error_code equivalent_error;
        if (std::filesystem::exists(destination) &&
            std::filesystem::equivalent(source, destination, equivalent_error) &&
            !equivalent_error) continue;
        if (std::filesystem::exists(destination)) {
            const auto prompt = L"Replace the existing input file '" +
                destination.filename().wstring() + L"'?";
            if (MessageBoxW(main_window, prompt.c_str(), L"Point Input Files",
                    MB_ICONQUESTION | MB_YESNO | MB_DEFBUTTON2) != IDYES) continue;
        }
        std::error_code copy_error;
        std::filesystem::copy_file(source, destination,
            std::filesystem::copy_options::overwrite_existing, copy_error);
        copy_error ? ++rejected : ++imported;
    }
    refresh_input_file_list();
    if (imported) {
        point::append_audit(app_root, "INPUT_FILES_ADDED",
            std::to_string(imported) + " file(s); " +
            std::to_string(rejected) + " rejected");
        refresh_engine();
    } else if (rejected) {
        MessageBoxW(main_window,
            L"No files were added. Point accepts CSV, XLSX, XLS, and XLSM.",
            L"Point Input Files", MB_ICONWARNING);
    }
}

void choose_input_files() {
    std::vector<wchar_t> buffer(65'536, L'\0');
    OPENFILENAMEW dialog{};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = main_window;
    dialog.lpstrFilter = L"Supported data files (*.csv;*.xlsx;*.xls;*.xlsm)\0"
        L"*.csv;*.xlsx;*.xls;*.xlsm\0All files (*.*)\0*.*\0\0";
    dialog.lpstrFile = buffer.data();
    dialog.nMaxFile = static_cast<DWORD>(buffer.size());
    dialog.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST |
        OFN_PATHMUSTEXIST | OFN_ALLOWMULTISELECT;
    dialog.lpstrTitle = L"Add files to Point";
    if (!GetOpenFileNameW(&dialog)) return;
    std::vector<std::filesystem::path> files;
    const std::filesystem::path first(buffer.data());
    const wchar_t* cursor = buffer.data() + first.native().size() + 1;
    if (*cursor == L'\0') files.push_back(first);
    else while (*cursor != L'\0') {
        const std::filesystem::path name(cursor);
        files.push_back(first / name);
        cursor += name.native().size() + 1;
    }
    import_input_files(files);
}

void remove_selected_input_file() {
    if (refresh_running.load()) {
        MessageBoxW(main_window,
            L"Wait for the current refresh to finish before removing a file.",
            L"Point Input Files", MB_ICONINFORMATION);
        return;
    }
    const LRESULT selected_count =
        SendMessageW(input_file_list, LB_GETSELCOUNT, 0, 0);
    if (selected_count <= 0 || selected_count == LB_ERR) {
        MessageBoxW(main_window,
            L"Select one or more files to remove first.\r\n"
            L"Use Ctrl+click or Shift+click to select multiple files.",
            L"Point Input Files", MB_ICONINFORMATION);
        return;
    }
    std::vector<int> selected_indices(
        static_cast<std::size_t>(selected_count));
    const LRESULT returned = SendMessageW(
        input_file_list, LB_GETSELITEMS,
        static_cast<WPARAM>(selected_indices.size()),
        reinterpret_cast<LPARAM>(selected_indices.data()));
    if (returned != selected_count) {
        MessageBoxW(main_window, L"Point could not read the selected files.",
            L"Point Input Files", MB_ICONERROR);
        return;
    }
    std::vector<std::filesystem::path> filenames;
    for (const int index : selected_indices) {
        const LRESULT length = SendMessageW(
            input_file_list, LB_GETTEXTLEN,
            static_cast<WPARAM>(index), 0);
        if (length <= 0 || length > 32'767) continue;
        std::wstring name(static_cast<std::size_t>(length) + 1, L'\0');
        SendMessageW(input_file_list, LB_GETTEXT,
            static_cast<WPARAM>(index),
            reinterpret_cast<LPARAM>(name.data()));
        name.resize(static_cast<std::size_t>(length));
        const std::filesystem::path filename(name);
        if (filename == filename.filename() && supported_input_file(filename))
            filenames.push_back(filename);
    }
    if (filenames.empty()) {
        MessageBoxW(main_window, L"The selected entry is not a removable file.",
            L"Point Input Files", MB_ICONWARNING);
        return;
    }
    std::wstringstream prompt;
    prompt << L"Remove " << filenames.size()
           << L" selected file(s) from Point?\r\n\r\n";
    for (std::size_t index = 0;
         index < filenames.size() && index < 8; ++index)
        prompt << L"• " << filenames[index].wstring() << L"\r\n";
    if (filenames.size() > 8)
        prompt << L"• ...and " << filenames.size() - 8 << L" more\r\n";
    prompt << L"\r\nThe original files outside Point will not be deleted.";
    if (MessageBoxW(main_window, prompt.str().c_str(), L"Remove Input Files",
            MB_ICONWARNING | MB_YESNO | MB_DEFBUTTON2) != IDYES) return;

    std::size_t removed_count = 0;
    std::vector<std::filesystem::path> failed_files;
    for (const auto& filename : filenames) {
        std::error_code error;
        if (std::filesystem::remove(
                app_root / "Inbox" / filename, error) && !error)
            ++removed_count;
        else
            failed_files.push_back(filename);
    }
    if (!removed_count) {
        MessageBoxW(main_window,
            L"Point could not remove the selected input files.",
            L"Remove Input Files", MB_ICONERROR);
        return;
    }
    point::append_audit(app_root, "INPUT_FILES_REMOVED",
        std::to_string(removed_count) + " file(s) removed; " +
        std::to_string(failed_files.size()) + " failed");
    refresh_input_file_list();
    refresh_engine();
    if (!failed_files.empty()) {
        std::wstringstream failure;
        failure << failed_files.size()
                << L" selected file(s) could not be removed:\r\n\r\n";
        for (std::size_t index = 0;
             index < failed_files.size() && index < 8; ++index)
            failure << L"• " << failed_files[index].wstring() << L"\r\n";
        failure << L"\r\nClose the file in Excel and try again.";
        MessageBoxW(main_window, failure.str().c_str(),
            L"Remove Input Files", MB_ICONWARNING);
    }
}

void show_relationship_diagnostics() {
    if (!engine || engine->datasets().empty()) {
        MessageBoxW(main_window, L"Refresh data first.",
                    L"Relationship Diagnostics", MB_ICONINFORMATION);
        return;
    }
    std::wstringstream report;
    report << L"Reports: " << engine->datasets().size()
           << L"\nRelationships: " << engine->relationships().size()
           << L"\n\nDetected joins:\n";
    std::size_t shown = 0;
    for (const auto& relationship : engine->relationships()) {
        if (shown++ >= 20) {
            report << L"...additional relationships omitted\n";
            break;
        }
        const auto& left = engine->datasets()[relationship.left_dataset];
        const auto& right = engine->datasets()[relationship.right_dataset];
        report << widen(left.name) << L" ["
               << widen(left.headers[relationship.left_column]) << L"] -> "
               << widen(right.name) << L" ["
               << widen(right.headers[relationship.right_column]) << L"]\n  "
               << (relationship.left_unique && relationship.right_unique
                       ? L"one-to-one" : L"one-to-many")
               << L", confidence " << std::fixed << std::setprecision(0)
               << relationship.confidence * 100.0 << L"%\n";
    }

    const std::set<std::string> checked_fields{
        "username", "email", "userprincipalname", "displayname",
        "accountstatus", "department", "manager", "managername"};
    std::map<std::string,
             std::map<std::string, std::set<std::string>>> values;
    for (const auto& dataset : engine->datasets()) {
        const auto identity = dataset.normalized_header_index.find("employeeid");
        if (identity == dataset.normalized_header_index.end()) continue;
        for (const auto& row : dataset.rows) {
            if (identity->second >= row.size()) continue;
            const auto employee = point::trim(row[identity->second]);
            if (employee.empty()) continue;
            for (const auto& field : checked_fields) {
                const auto column = dataset.normalized_header_index.find(field);
                if (column == dataset.normalized_header_index.end() ||
                    column->second >= row.size()) continue;
                const auto value = point::trim(row[column->second]);
                if (!value.empty()) values[employee][field].insert(value);
            }
        }
    }
    std::size_t conflict_count = 0;
    std::wstringstream examples;
    for (const auto& [employee, fields] : values) {
        for (const auto& [field, distinct] : fields) {
            if (distinct.size() < 2) continue;
            ++conflict_count;
            if (conflict_count <= 10)
                examples << widen(employee) << L": " << widen(field)
                         << L" has " << distinct.size()
                         << L" different values\n";
        }
    }
    report << L"\nCross-report conflicts: " << conflict_count << L"\n";
    if (conflict_count) report << examples.str();
    else report << L"No conflicting identity profile values detected.";
    MessageBoxW(main_window, report.str().c_str(),
                L"Relationship and Conflict Diagnostics", MB_OK);
}

void explain_selected_result() {
    if (last_result.rows.empty()) {
        MessageBoxW(main_window, L"Run a search and select a result row first.",
                    L"Search Explanation", MB_ICONINFORMATION);
        return;
    }
    GridPosition position{};
    if (!grid_position_from_control(GetFocus(), position) || position.row < 0)
        position.row = 0;
    const int result_row = std::clamp(position.row, 0,
        static_cast<int>(last_result.rows.size()) - 1);
    std::wstringstream report;
    report << L"Result row " << result_row + 1 << L"\n\nWhy it matched:\n"
           << widen(last_result.explanation) << L"\n\nSource lineage:\n";
    if (last_result.sources.empty()) report << L"No source was recorded.\n";
    for (const auto& source : last_result.sources)
        report << L"• " << widen(source) << L"\n";
    report << L"\nDisplayed values:\n";
    const auto& row = last_result.rows[static_cast<std::size_t>(result_row)];
    for (std::size_t column = 0;
         column < row.size() && column < last_result.headers.size(); ++column)
        report << widen(last_result.headers[column]) << L": "
               << widen(row[column]) << L"\n";
    MessageBoxW(main_window, report.str().c_str(),
                L"Search Explanation and Lineage", MB_OK);
}

LRESULT CALLBACK synonym_window_proc(
        HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
    (void)lparam;
    switch (message) {
    case WM_CREATE: {
        CreateWindowW(
            L"STATIC",
            L"Excel-style 64 x 64 grid: canonical fields go in the yellow "
            L"heading row and synonyms underneath. Use the scrollbars or "
            L"mouse wheel. Suggestions come from current Inbox sheets. "
            L"Red = not found; blue = duplicate. Enter moves down.",
            WS_CHILD | WS_VISIBLE, 16, 14, 1180, 38,
            window, nullptr, nullptr, nullptr);
        synonym_headers.assign(SYNONYM_VISIBLE_COLUMNS, nullptr);
        synonym_cells.assign(
            SYNONYM_VISIBLE_ROWS,
            std::vector<HWND>(SYNONYM_VISIBLE_COLUMNS, nullptr));
        synonym_column_labels.assign(SYNONYM_VISIBLE_COLUMNS, nullptr);
        synonym_row_labels.assign(SYNONYM_VISIBLE_ROWS, nullptr);
        synonym_model_headers.assign(SYNONYM_FIELD_COLUMNS, {});
        synonym_model_cells.assign(
            SYNONYM_ROWS,
            std::vector<std::string>(SYNONYM_FIELD_COLUMNS));
        synonym_first_column = 0;
        synonym_first_row = 0;
        for (std::size_t column = 0;
             column < field_synonym_groups.size() &&
             column < static_cast<std::size_t>(SYNONYM_FIELD_COLUMNS);
             ++column) {
            synonym_model_headers[column] =
                field_synonym_groups[column].canonical_field;
            for (std::size_t row = 0;
                 row < field_synonym_groups[column].synonyms.size() &&
                 row < static_cast<std::size_t>(SYNONYM_ROWS); ++row) {
                synonym_model_cells[row][column] =
                    field_synonym_groups[column].synonyms[row];
            }
        }
        constexpr int left = 48;
        constexpr int top = 78;
        constexpr int cell_width = 143;
        constexpr int cell_height = 28;
        for (int column = 0;
             column < SYNONYM_VISIBLE_COLUMNS; ++column) {
            synonym_column_labels[static_cast<std::size_t>(column)] =
                CreateWindowW(
                L"STATIC", L"",
                WS_CHILD | WS_VISIBLE | SS_CENTER,
                left + column * cell_width, 56,
                cell_width - 3, 20, window, nullptr, nullptr, nullptr);
            synonym_headers[static_cast<std::size_t>(column)] =
                CreateWindowExW(
                    WS_EX_CLIENTEDGE, L"EDIT", nullptr,
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
                    left + column * cell_width, top,
                    cell_width - 3, cell_height, window,
                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(
                        ID_SYNONYM_HEADER_BASE + column)),
                    nullptr, nullptr);
            SendMessageW(
                synonym_headers[static_cast<std::size_t>(column)],
                EM_SETLIMITTEXT, 128, 0);
            SetWindowSubclass(
                synonym_headers[static_cast<std::size_t>(column)],
                synonym_edit_subclass,
                static_cast<UINT_PTR>(ID_SYNONYM_HEADER_BASE + column), 0);
        }
        for (int row = 0; row < SYNONYM_VISIBLE_ROWS; ++row) {
            synonym_row_labels[static_cast<std::size_t>(row)] =
                CreateWindowW(
                L"STATIC", L"",
                WS_CHILD | WS_VISIBLE | SS_CENTER,
                16, top + (row + 1) * cell_height + 4,
                28, 20, window, nullptr, nullptr, nullptr);
            for (int column = 0;
                 column < SYNONYM_VISIBLE_COLUMNS; ++column) {
                auto& cell = synonym_cells[static_cast<std::size_t>(row)]
                                           [static_cast<std::size_t>(column)];
                cell = CreateWindowExW(
                    WS_EX_CLIENTEDGE, L"EDIT", nullptr,
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
                    left + column * cell_width,
                    top + (row + 1) * cell_height,
                    cell_width - 3, cell_height, window,
                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(
                        ID_SYNONYM_CELL_BASE +
                        row * SYNONYM_VISIBLE_COLUMNS + column)),
                    nullptr, nullptr);
                SendMessageW(cell, EM_SETLIMITTEXT, 128, 0);
                SetWindowSubclass(
                    cell, synonym_edit_subclass,
                    static_cast<UINT_PTR>(ID_SYNONYM_CELL_BASE +
                        row * SYNONYM_VISIBLE_COLUMNS + column), 0);
            }
        }
        synonym_hscrollbar = CreateWindowW(
            L"SCROLLBAR", nullptr,
            WS_CHILD | WS_VISIBLE | SBS_HORZ,
            left, top + (SYNONYM_VISIBLE_ROWS + 1) * cell_height + 4,
            SYNONYM_VISIBLE_COLUMNS * cell_width - 3, 20, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_SYNONYM_HSCROLL)),
            nullptr, nullptr);
        synonym_vscrollbar = CreateWindowW(
            L"SCROLLBAR", nullptr,
            WS_CHILD | WS_VISIBLE | SBS_VERT,
            left + SYNONYM_VISIBLE_COLUMNS * cell_width,
            top, 20, (SYNONYM_VISIBLE_ROWS + 1) * cell_height,
            window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_SYNONYM_VSCROLL)),
            nullptr, nullptr);
        load_synonym_view();
        CreateWindowW(
            L"BUTTON", L"Save and Refresh",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
            914, 535, 150, 30, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_SYNONYM_SAVE)),
            nullptr, nullptr);
        CreateWindowW(
            L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            1074, 535, 106, 30, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_SYNONYM_CANCEL)),
            nullptr, nullptr);
        return 0;
    }
    case WM_CTLCOLOREDIT: {
        const auto control = reinterpret_cast<HWND>(lparam);
        const auto controls = synonym_grid_controls();
        if (std::find(controls.begin(), controls.end(), control) ==
            controls.end()) break;
        const auto names = split_synonym_cell(
            narrow(control_text(control)));
        const bool is_header = std::find(
            synonym_headers.begin(), synonym_headers.end(), control) !=
            synonym_headers.end();
        HBRUSH brush = is_header
            ? criteria_cell_brush : GetSysColorBrush(COLOR_WINDOW);
        COLORREF color = is_header
            ? RGB(255, 248, 214) : GetSysColor(COLOR_WINDOW);
        const bool duplicate = std::any_of(
            names.begin(), names.end(),
            [](const auto& name) {
                return synonym_grid_occurrences(
                    point::normalize_name(name)) > 1;
            });
        const bool missing = std::any_of(
            names.begin(), names.end(),
            [](const auto& name) {
                return !imported_field_exists(
                    point::normalize_name(name));
            });
        if (duplicate) {
            brush = synonym_duplicate_brush;
            color = RGB(205, 225, 255);
        } else if (missing) {
            brush = synonym_missing_brush;
            color = RGB(255, 205, 205);
        }
        auto context = reinterpret_cast<HDC>(wparam);
        SetTextColor(context, RGB(0, 0, 0));
        SetBkColor(context, color);
        return reinterpret_cast<LRESULT>(brush);
    }
    case WM_HSCROLL:
        if (reinterpret_cast<HWND>(lparam) == synonym_hscrollbar) {
            commit_synonym_view();
            SCROLLINFO info{sizeof(info)};
            info.fMask = SIF_ALL;
            GetScrollInfo(synonym_hscrollbar, SB_CTL, &info);
            int target = synonym_first_column;
            switch (LOWORD(wparam)) {
            case SB_LINELEFT: target -= 1; break;
            case SB_LINERIGHT: target += 1; break;
            case SB_PAGELEFT: target -= SYNONYM_VISIBLE_COLUMNS; break;
            case SB_PAGERIGHT: target += SYNONYM_VISIBLE_COLUMNS; break;
            case SB_THUMBTRACK:
            case SB_THUMBPOSITION: target = info.nTrackPos; break;
            default: return 0;
            }
            synonym_first_column = std::clamp(
                target, 0,
                SYNONYM_FIELD_COLUMNS - SYNONYM_VISIBLE_COLUMNS);
            load_synonym_view();
            return 0;
        }
        break;
    case WM_VSCROLL:
        if (reinterpret_cast<HWND>(lparam) == synonym_vscrollbar) {
            commit_synonym_view();
            SCROLLINFO info{sizeof(info)};
            info.fMask = SIF_ALL;
            GetScrollInfo(synonym_vscrollbar, SB_CTL, &info);
            int target = synonym_first_row;
            switch (LOWORD(wparam)) {
            case SB_LINEUP: target -= 1; break;
            case SB_LINEDOWN: target += 1; break;
            case SB_PAGEUP: target -= SYNONYM_VISIBLE_ROWS; break;
            case SB_PAGEDOWN: target += SYNONYM_VISIBLE_ROWS; break;
            case SB_THUMBTRACK:
            case SB_THUMBPOSITION: target = info.nTrackPos; break;
            default: return 0;
            }
            synonym_first_row = std::clamp(
                target, 0, SYNONYM_ROWS - SYNONYM_VISIBLE_ROWS);
            load_synonym_view();
            return 0;
        }
        break;
    case WM_MOUSEWHEEL: {
        commit_synonym_view();
        const int steps = GET_WHEEL_DELTA_WPARAM(wparam) / WHEEL_DELTA;
        if (steps == 0) return 0;
        if ((GET_KEYSTATE_WPARAM(wparam) & MK_SHIFT) != 0) {
            synonym_first_column = std::clamp(
                synonym_first_column - steps, 0,
                SYNONYM_FIELD_COLUMNS - SYNONYM_VISIBLE_COLUMNS);
        } else {
            synonym_first_row = std::clamp(
                synonym_first_row - steps * 3, 0,
                SYNONYM_ROWS - SYNONYM_VISIBLE_ROWS);
        }
        load_synonym_view();
        return 0;
    }
    case WM_COMMAND:
        if (HIWORD(wparam) == EN_CHANGE &&
            ((LOWORD(wparam) >= ID_SYNONYM_HEADER_BASE &&
              LOWORD(wparam) <
                  ID_SYNONYM_HEADER_BASE + SYNONYM_VISIBLE_COLUMNS) ||
             (LOWORD(wparam) >= ID_SYNONYM_CELL_BASE &&
              LOWORD(wparam) < ID_SYNONYM_CELL_BASE +
                  SYNONYM_VISIBLE_COLUMNS * SYNONYM_VISIBLE_ROWS))) {
            if (!synonym_view_loading) {
                commit_synonym_view();
                schedule_suggestions(
                    reinterpret_cast<HWND>(lparam), LOWORD(wparam));
            }
            refresh_synonym_grid_colors();
            return 0;
        }
        if (LOWORD(wparam) == ID_SYNONYM_CANCEL) {
            DestroyWindow(window);
            return 0;
        }
        if (LOWORD(wparam) == ID_SYNONYM_SAVE) {
            try {
                const auto groups = field_synonyms_from_grid();
                point::Engine validator = engine
                    ? point::Engine(*engine) : point::Engine{};
                validator.set_field_synonyms(groups);
                save_field_synonyms(groups);
                field_synonym_groups = groups;
                point::append_audit(
                    app_root, "FIELD_SYNONYMS_UPDATED",
                    std::to_string(groups.size()) +
                    " canonical field group(s)");
                DestroyWindow(window);
                refresh_engine(true);
            } catch (const std::exception& ex) {
                point::append_audit(
                    app_root, "FIELD_SYNONYMS_BLOCKED", ex.what());
                MessageBoxW(
                    window, widen(ex.what()).c_str(),
                    L"Field Synonym Manager", MB_ICONERROR);
            }
            return 0;
        }
        break;
    case WM_CLOSE:
        DestroyWindow(window);
        return 0;
    case WM_DESTROY:
        hide_suggestions();
        synonym_headers.clear();
        synonym_cells.clear();
        synonym_column_labels.clear();
        synonym_row_labels.clear();
        synonym_hscrollbar = nullptr;
        synonym_vscrollbar = nullptr;
        synonym_window = nullptr;
        return 0;
    }
    return DefWindowProcW(window, message, wparam, lparam);
}

void set_change_baseline() {
    if (!engine || engine->datasets().empty()) {
        MessageBoxW(
            main_window,
            L"No imported reports are available for a baseline.",
            L"Change baseline", MB_ICONWARNING);
        return;
    }
    previous_engine = std::make_unique<point::Engine>(*engine);
    set_status(
        L"Current imported reports are now the retained Change baseline.");
    point::append_audit(
        app_root, "CHANGE_BASELINE_SET",
        std::to_string(engine->datasets().size()) + " report(s)");
}

void add_count_related_data(
        point::QueryResult& count_result,
        const std::vector<std::string>& count_fields,
        const std::string& source_name) {
    const std::size_t count_column = count_fields.size();
    std::string normalized_count_fields;
    for (const auto& field : count_fields) {
        normalized_count_fields += point::normalize_name(field);
        normalized_count_fields.push_back(';');
    }
    auto contains_any = [&](const std::vector<std::string>& terms) {
        return std::any_of(
            terms.begin(), terms.end(),
            [&](const auto& term) {
                return normalized_count_fields.find(term) !=
                    std::string::npos;
            });
    };
    const bool device_security_count = contains_any({
        "device", "computer", "encryption", "patch", "crowdstrike",
        "vulnerability", "antivirus", "endpoint", "operatingsystem",
        "asset", "serialnumber"
    });
    const bool ticket_count = contains_any({
        "ticket", "incident", "priority", "resolution"
    });
    const bool user_access_count = contains_any({
        "employee", "user", "account", "mfa", "password", "locked",
        "disabled", "group", "role", "access"
    });

    std::string related_header = "Related Objects";
    std::vector<std::string> preferred_identity_fields;
    if (device_security_count) {
        related_header = "Affected Devices";
        preferred_identity_fields = {
            "Device ID", "Computer ID", "Computer Name", "Asset ID",
            "Asset Tag", "Serial Number", "Employee ID", "Username"
        };
    } else if (ticket_count) {
        related_header = "Affected Tickets";
        preferred_identity_fields = {
            "Ticket ID", "Incident ID", "Case ID", "Employee ID",
            "Username"
        };
    } else if (user_access_count) {
        related_header = "Affected Users";
        preferred_identity_fields = {
            "Employee ID", "Username", "User Principal Name", "Email",
            "Account Name", "Full Name", "Display Name"
        };
    } else {
        preferred_identity_fields = {
            "Employee ID", "Username", "Device ID", "Computer ID",
            "Computer Name", "Email", "Asset Tag", "Account Name",
            "Ticket ID", "Full Name", "Display Name"
        };
    }

    const point::DataSet* selected_dataset = nullptr;
    std::string identity_field;
    std::size_t identity_column = 0;
    std::size_t best_identity_rank =
        preferred_identity_fields.size();
    std::vector<std::size_t> group_columns;
    for (const auto& dataset : engine->datasets()) {
        if (!point::trim(source_name).empty() &&
            point::normalize_name(dataset.name) !=
                point::normalize_name(source_name)) {
            continue;
        }
        std::vector<std::size_t> candidate_columns;
        bool contains_groups = true;
        for (const auto& field : count_fields) {
            const auto found = dataset.normalized_header_index.find(
                point::normalize_name(field));
            if (found == dataset.normalized_header_index.end()) {
                contains_groups = false;
                break;
            }
            candidate_columns.push_back(found->second);
        }
        if (!contains_groups) continue;

        for (std::size_t rank = 0;
             rank < preferred_identity_fields.size(); ++rank) {
            const auto& preferred =
                preferred_identity_fields[rank];
            const auto found = dataset.normalized_header_index.find(
                point::normalize_name(preferred));
            if (found == dataset.normalized_header_index.end()) continue;
            if (rank < best_identity_rank) {
                selected_dataset = &dataset;
                identity_field = dataset.headers[found->second];
                identity_column = found->second;
                group_columns = candidate_columns;
                best_identity_rank = rank;
            }
            break;
        }
    }

    auto group_signature = [](const std::vector<std::string>& values) {
        std::string signature;
        for (const auto& value : values) {
            const auto normalized =
                point::normalize_name(point::trim(value));
            signature += std::to_string(normalized.size());
            signature.push_back(':');
            signature += normalized;
            signature.push_back(';');
        }
        return signature;
    };

    count_result.headers.push_back(related_header);
    const std::string output_identity_field =
        !count_result.related_identity_field.empty()
        ? count_result.related_identity_field : identity_field;
    std::vector<std::set<std::string>> related_values(
        count_result.rows.size());
    std::unordered_map<std::string, std::vector<std::size_t>>
        target_groups;

    for (std::size_t row_index = 0;
         row_index < count_result.rows.size(); ++row_index) {
        auto& row = count_result.rows[row_index];
        if (row.size() <= count_column) {
            row.push_back("Related data unavailable");
            continue;
        }

        char* count_end = nullptr;
        const auto count_value = std::strtoull(
            row[count_column].c_str(), &count_end, 10);
        if (!count_end || *count_end != '\0') {
            row.push_back("Related data unavailable");
            continue;
        }
        if (count_value > COUNT_DETAIL_LIMIT) {
            row.push_back("Count above 50 - narrow filters");
            continue;
        }
        if (count_value == 0) {
            row.push_back("No related records");
            continue;
        }
        if (row_index <
                count_result.related_identity_values.size() &&
            !count_result
                 .related_identity_values[row_index].empty()) {
            related_values[row_index].insert(
                count_result.related_identity_values[row_index].begin(),
                count_result.related_identity_values[row_index].end());
            row.emplace_back();
            continue;
        }
        if (!selected_dataset) {
            row.push_back(
                "No affected object ID available");
            continue;
        }

        std::vector<std::string> group;
        group.reserve(count_fields.size());
        for (std::size_t column = 0;
             column < count_fields.size(); ++column) {
            group.push_back(
                column < row.size() ? row[column] : std::string{});
        }
        target_groups[group_signature(group)].push_back(row_index);
        row.emplace_back();
    }

    if (selected_dataset && !target_groups.empty()) {
        for (const auto& source_row : selected_dataset->rows) {
            std::vector<std::string> group;
            group.reserve(group_columns.size());
            for (const auto column : group_columns) {
                group.push_back(
                    column < source_row.size()
                    ? source_row[column] : std::string{});
            }
            const auto target =
                target_groups.find(group_signature(group));
            if (target == target_groups.end() ||
                identity_column >= source_row.size()) {
                continue;
            }
            const auto identity =
                point::trim(source_row[identity_column]);
            if (identity.empty()) continue;
            for (const auto row_index : target->second) {
                related_values[row_index].insert(identity);
            }
        }
    }

    for (std::size_t row_index = 0;
         row_index < count_result.rows.size(); ++row_index) {
        auto& row = count_result.rows[row_index];
        if (row.size() != count_column + 2 ||
            !row.back().empty()) {
            continue;
        }
        const auto& values = related_values[row_index];
        if (values.empty()) {
            row.back() = "No related records";
            continue;
        }
        row.back() = output_identity_field + ": ";
        bool first = true;
        for (const auto& value : values) {
            if (!first) row.back() += "; ";
            row.back() += value;
            first = false;
        }
    }
}

void run_search() {
    static bool query_in_progress = false;
    if (query_in_progress) return;
    if (!engine) return;
    query_in_progress = true;
    data_tools_unfiltered_result.reset();
    hide_suggestions();
    bool query_progress_visible = false;
    try {
        if (engine->datasets().empty())
            throw std::runtime_error(
                last_import_issue.empty()
                ? "No supported reports were imported. Add CSV, XLSX, XLS, "
                  "or XLSM files to build\\Inbox and click Refresh."
                : "Last import error: " + last_import_issue);
        show_refresh_progress(L"Preparing search...", 3);
        query_progress_visible = true;
        commit_visible_cells();
        const auto previous_universal_missing = universal_missing_rows;
        const auto previous_universal_duplicates =
            universal_duplicate_rows;
        universal_missing_rows.clear();
        universal_duplicate_rows.clear();
        universal_duplicate_dark_rows.clear();
        clear_count_detail_columns();
        if (count_mode && generated_count_column >= 0) {
            internal_cell_update = true;
            HWND generated_header =
                header_cells[
                    static_cast<std::size_t>(generated_count_column)];
            if (point::normalize_name(
                    narrow(control_text(generated_header))) == "count") {
                SetWindowTextW(generated_header, L"");
            }
            generated_count_column = -1;
            internal_cell_update = false;
        }
        if (analyze_mode && generated_analysis_headers) {
            internal_cell_update = true;
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t i = 0;
                 i < analysis_key_fields_cache.size() &&
                 i < header_cells.size(); ++i) {
                set_control_text(
                    header_cells[i], analysis_key_fields_cache[i]);
            }
            clear_cell_store();
            for (std::size_t i = 0;
                 i < analysis_key_values_cache.size() &&
                 i < static_cast<std::size_t>(GRID_COLUMNS); ++i) {
                if (!analysis_key_values_cache[i].empty()) {
                    store_cell(
                        0, static_cast<int>(i),
                        analysis_key_values_cache[i]);
                }
            }
            generated_analysis_headers = false;
            internal_cell_update = false;
        }
        if (insight_mode && generated_insight_headers) {
            internal_cell_update = true;
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t i = 0;
                 i < insight_fields_cache.size() &&
                 i < header_cells.size(); ++i) {
                set_control_text(
                    header_cells[i], insight_fields_cache[i]);
            }
            clear_cell_store();
            generated_insight_headers = false;
            internal_cell_update = false;
        }
        if (chart_mode && generated_chart_headers) {
            internal_cell_update = true;
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t i = 0;
                 i < chart_fields_cache.size() &&
                 i < header_cells.size(); ++i) {
                set_control_text(
                    header_cells[i], chart_fields_cache[i]);
            }
            clear_cell_store();
            for (std::size_t i = 0;
                 i < chart_filter_values_cache.size() &&
                 i < static_cast<std::size_t>(GRID_COLUMNS); ++i) {
                if (!chart_filter_values_cache[i].empty()) {
                    store_cell(
                        0, static_cast<int>(i),
                        chart_filter_values_cache[i]);
                }
            }
            generated_chart_headers = false;
            internal_cell_update = false;
        }
        if (change_mode && generated_change_headers) {
            internal_cell_update = true;
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t i = 0;
                 i < change_key_fields_cache.size() &&
                 i < header_cells.size(); ++i) {
                set_control_text(
                    header_cells[i], change_key_fields_cache[i]);
            }
            clear_cell_store();
            generated_change_headers = false;
            internal_cell_update = false;
        }
        if (compare_mode && generated_compare_header &&
            !compare_identity_field_cache.empty()) {
            internal_cell_update = true;
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t i = 0;
                 i < compare_fields_cache.size() &&
                 i < header_cells.size(); ++i) {
                set_control_text(
                    header_cells[i], compare_fields_cache[i]);
            }
            clear_cell_store();
            for (std::size_t i = 0;
                 i < compare_inputs_cache.size() && i < 2; ++i) {
                store_cell(
                    static_cast<int>(i), 0,
                    compare_inputs_cache[i]);
            }
            first_visible_row = 0;
            load_visible_cells();
            generated_compare_header = false;
            internal_cell_update = false;
        }
        const auto headers =
            insight_mode && selected_headers().empty()
            ? std::vector<std::string>{}
            : resolved_headers(analyze_mode ? 3 : GRID_COLUMNS);
        commit_visible_cells();
        resolve_all_identity_name_inputs();
        const auto source_name = selected_source_name();

        point::QueryResult combined;
        combined.headers = headers;
        std::vector<std::string> preserved_criteria;
        std::vector<std::string> compare_inputs;
        std::vector<std::string> next_universal_lookup_history;
        bool update_universal_lookup_history = false;
        std::size_t condition_count = 0;
        std::string condition_fields;
        std::size_t first_output_row = 0;

        if (change_mode) {
            if (!previous_engine)
                throw std::runtime_error(
                    "Change mode needs a baseline. Import the original "
                    "reports and choose Workspace > Set Change Baseline");
            if (headers.empty() || headers.size() > 3)
                throw std::runtime_error(
                    "Change mode requires one to three key headings");
            change_key_fields_cache = headers;
            combined = engine->changes_since(
                *previous_engine, headers, source_name);
            condition_count = headers.size();
            for (const auto& header : headers) {
                if (!condition_fields.empty()) condition_fields += ",";
                condition_fields += point::normalize_name(header);
            }
        } else if (chart_mode) {
            if (headers.empty())
                throw std::runtime_error(
                    "Chart mode requires at least one category heading");
            chart_fields_cache = headers;
            chart_filter_values_cache.assign(headers.size(), {});
            std::vector<std::string> categories;
            std::string series_field;
            std::vector<point::QueryCondition> filters;
            for (std::size_t column = 0;
                 column < headers.size(); ++column) {
                chart_filter_values_cache[column] =
                    point::trim(stored_cell(
                        0, static_cast<int>(column)));
                if (chart_filter_values_cache[column].empty()) {
                    categories.push_back(headers[column]);
                } else if (is_chart_series_marker(
                               chart_filter_values_cache[column])) {
                    if (!series_field.empty())
                        throw std::runtime_error(
                            "Chart mode supports one shared @series field");
                    series_field = headers[column];
                } else {
                    filters.push_back({
                        headers[column],
                        chart_filter_values_cache[column]});
                }
            }
            if (categories.empty() || categories.size() > 4)
                throw std::runtime_error(
                    "Chart mode requires one to four headings with a blank "
                    "value underneath. Other headings may be exact filters.");
            if (categories.size() == 1 && series_field.empty()) {
                combined = engine->distribution(
                    categories.front(), filters, source_name);
            } else {
                combined.headers = series_field.empty()
                    ? std::vector<std::string>{
                        "Chart Field", "Category", "Count"}
                    : std::vector<std::string>{
                        "Chart Field", "Category", "Series", "Count"};
                for (const auto& category : categories) {
                    const auto chart = series_field.empty()
                        ? engine->distribution(
                            category, filters, source_name)
                        : engine->distribution(
                            std::vector<std::string>{
                                category, series_field},
                            filters, source_name);
                    for (const auto& row : chart.rows) {
                        if (row.size() < 2) continue;
                        if (series_field.empty())
                            combined.rows.push_back(
                                {category, row.front(), row.back()});
                        else if (row.size() >= 3)
                            combined.rows.push_back({
                                category, row[0], row[1], row.back()});
                    }
                    for (const auto& source : chart.sources)
                        if (std::find(
                                combined.sources.begin(),
                                combined.sources.end(), source) ==
                            combined.sources.end())
                            combined.sources.push_back(source);
                }
                combined.explanation =
                    "Independent exact-filtered distributions for a "
                    "multi-chart, multi-series dashboard.";
            }
            condition_count = filters.size();
            for (const auto& category : categories) {
                if (!condition_fields.empty()) condition_fields += ",";
                condition_fields += point::normalize_name(category);
            }
            for (const auto& filter : filters) {
                condition_fields += "," +
                    point::normalize_name(filter.field);
            }
            if (!series_field.empty()) {
                condition_fields += ",series:" +
                    point::normalize_name(series_field);
            }
        } else if (insight_mode) {
            insight_fields_cache = headers;
            combined = engine->deep_insights(
                headers, source_name);
            condition_count = headers.size();
            for (const auto& header : headers) {
                if (!condition_fields.empty()) condition_fields += ",";
                condition_fields += point::normalize_name(header);
            }
        } else if (analyze_mode) {
            if (headers.empty() || headers.size() > 3)
                throw std::runtime_error(
                    "Analyze mode requires one to three key headings");
            analysis_key_fields_cache = headers;
            analysis_key_values_cache.clear();
            analysis_key_values_cache.reserve(headers.size());
            for (std::size_t column = 0;
                 column < headers.size(); ++column) {
                analysis_key_values_cache.push_back(
                    point::trim(stored_cell(
                        0, static_cast<int>(column))));
            }
            combined = engine->analyze_keys(
                headers, source_name, analysis_key_values_cache);
            condition_count = headers.size();
            for (const auto& header : headers) {
                if (!condition_fields.empty()) condition_fields += ",";
                condition_fields += point::normalize_name(header);
            }
        } else if (compare_mode) {
            if (headers.size() < 2 ||
                headers.size() >
                    static_cast<std::size_t>(GRID_COLUMNS))
                throw std::runtime_error(
                    "Compare mode requires an identity heading followed "
                    "by one or more comparison headings");
            compare_inputs = {
                point::trim(stored_cell(0, 0)),
                point::trim(stored_cell(1, 0))};
            if (compare_inputs[0].empty() ||
                compare_inputs[1].empty()) {
                throw std::runtime_error(
                    "Enter User A and User B in the first two rows under "
                    "the identity field");
            }
            const std::vector<std::string> comparison_fields(
                headers.begin() + 1, headers.end());
            combined = engine->compare_profiles(
                headers[0], compare_inputs[0],
                compare_inputs[1], comparison_fields, source_name);
            compare_identity_field_cache = headers[0];
            compare_fields_cache = headers;
            compare_inputs_cache = compare_inputs;
            condition_count = headers.size();
            for (const auto& header : headers) {
                if (!condition_fields.empty()) condition_fields += ",";
                condition_fields += point::normalize_name(header);
            }
            first_output_row = 0;
        } else if (count_mode) {
            if (headers.size() >
                static_cast<std::size_t>(GRID_COLUMNS - 2)) {
                throw std::runtime_error(
                    "Count mode supports up to 254 selected fields because "
                    "Count and Affected Objects reserve two of the 256 columns");
            }
            preserved_criteria.resize(headers.size());
            std::vector<point::QueryCondition> count_conditions;
            std::size_t blank_criteria = 0;
            for (std::size_t column = 0;
                 column < headers.size(); ++column) {
                preserved_criteria[column] = point::trim(
                    stored_cell(0, static_cast<int>(column)));
                if (preserved_criteria[column].empty()) {
                    ++blank_criteria;
                    continue;
                }
                count_conditions.push_back(
                    {headers[column], preserved_criteria[column]});
            }
            if (blank_criteria == headers.size()) {
                combined =
                    engine->count_groups(headers, source_name);
                condition_count = headers.size();
            } else if (blank_criteria == 0) {
                combined =
                    engine->count_exact(
                        count_conditions, source_name);
                condition_count = count_conditions.size();
            } else {
                combined =
                    engine->count_mixed(
                        headers, preserved_criteria, source_name);
                condition_count = count_conditions.size();
            }
            add_count_related_data(
                combined, headers, source_name);
            for (const auto& header : headers) {
                if (!condition_fields.empty()) condition_fields += ",";
                condition_fields +=
                    point::normalize_name(header);
            }
            first_output_row = 1;
        } else if (narrow_mode) {
            preserved_criteria.resize(headers.size());
            point::QueryRequest request;
            request.output_fields = headers;
            for (std::size_t column = 0;
                 column < headers.size(); ++column) {
                preserved_criteria[column] = point::trim(
                    stored_cell(0, static_cast<int>(column)));
                if (!preserved_criteria[column].empty()) {
                    request.conditions.push_back(
                        {headers[column], preserved_criteria[column]});
                }
            }
            if (request.conditions.empty())
                throw std::runtime_error(
                    "Enter at least one exact-match value in the yellow "
                    "criteria row");
            combined = engine->query(request);
            condition_count = request.conditions.size();
            for (const auto& condition : request.conditions) {
                if (!condition_fields.empty()) condition_fields += ",";
                condition_fields +=
                    point::normalize_name(condition.field);
            }
            first_output_row = 1;
        } else {
            std::vector<std::string> lookups;
            std::set<int> detected_pending_rows =
                universal_pending_lookup_rows;
            if (universal_results_displayed) {
                const int last_stored_row = highest_stored_row();
                for (int row = 0; row <= last_stored_row; ++row) {
                    const auto lookup = point::trim(stored_cell(row, 0));
                    if (lookup.empty()) continue;
                    if (row >= static_cast<int>(last_result.rows.size()) ||
                        last_result.rows[
                            static_cast<std::size_t>(row)].empty() ||
                        point::normalize_name(lookup) !=
                            point::normalize_name(last_result.rows[
                                static_cast<std::size_t>(row)].front())) {
                        detected_pending_rows.insert(row);
                    }
                }
            }
            const bool has_pending_universal_inputs =
                universal_results_displayed &&
                !detected_pending_rows.empty();
            const bool append_to_universal_results =
                has_pending_universal_inputs &&
                same_normalized_headers(
                    headers, universal_result_headers);
            combined.headers = headers;
            auto result_signature = [](const auto& row) {
                std::string signature;
                for (const auto& value : row) {
                    signature += point::normalize_name(value);
                    signature.push_back('\x1e');
                }
                return signature;
            };
            if (append_to_universal_results) {
                for (std::size_t row = 0;
                     row < last_result.rows.size(); ++row) {
                    if (detected_pending_rows.contains(
                            static_cast<int>(row))) {
                        continue;
                    }
                    const int new_row =
                        static_cast<int>(combined.rows.size());
                    combined.rows.push_back(last_result.rows[row]);
                    if (previous_universal_missing.contains(
                            static_cast<int>(row)))
                        universal_missing_rows.insert(new_row);
                    if (previous_universal_duplicates.contains(
                            static_cast<int>(row)))
                        universal_duplicate_rows.insert(new_row);
                }
            }
            std::vector<int> lookup_rows;
            const bool rerun_prior_universal_inputs =
                universal_results_displayed &&
                (!append_to_universal_results ||
                 !same_normalized_headers(
                     headers, universal_result_headers));
            if (has_pending_universal_inputs) {
                lookup_rows.assign(
                    detected_pending_rows.begin(),
                    detected_pending_rows.end());
            } else if (!universal_results_displayed) {
                const int last_stored_row = highest_stored_row();
                lookup_rows.reserve(
                    static_cast<std::size_t>(last_stored_row + 1));
                for (int row = 0; row <= last_stored_row; ++row)
                    lookup_rows.push_back(row);
            }
            if (rerun_prior_universal_inputs)
                lookups = universal_lookup_history;
            std::vector<std::string> pending_lookups;
            for (const int row : lookup_rows) {
                const auto lookup =
                    point::trim(stored_cell(row, 0));
                if (!lookup.empty()) {
                    lookups.push_back(lookup);
                    pending_lookups.push_back(lookup);
                }
            }
            if (lookups.empty())
                throw std::runtime_error(
                    "Enter at least one lookup value under the first field");
            if (!universal_results_displayed) {
                next_universal_lookup_history = lookups;
            } else {
                next_universal_lookup_history = universal_lookup_history;
                next_universal_lookup_history.insert(
                    next_universal_lookup_history.end(),
                    pending_lookups.begin(), pending_lookups.end());
            }
            update_universal_lookup_history = true;

            std::set<std::string> sources;
            if (append_to_universal_results) {
                sources.insert(
                    last_result.sources.begin(),
                    last_result.sources.end());
            }
            int last_query_percent = -1;
            auto update_query_progress = [&](std::size_t completed) {
                const int percent = 15 + static_cast<int>(
                    70 * completed / std::max<std::size_t>(1, lookups.size()));
                if (percent != last_query_percent || completed == lookups.size()) {
                    last_query_percent = percent;
                    std::wstringstream progress;
                    progress << L"Resolving " << completed << L" of "
                             << lookups.size() << L" lookup value(s)...";
                    show_refresh_progress(progress.str(), percent);
                }
                pump_search_messages();
            };
            update_query_progress(0);
            for (std::size_t lookup_index = 0;
                 lookup_index < lookups.size(); ++lookup_index) {
                const auto& lookup = lookups[lookup_index];
                // Prefer the heading under which the value was entered. This
                // prevents a date/ID that already matches Order Date, for
                // example, from broadening when Ship Date or another same-
                // shaped output field is later added. If the value is not in
                // that field, retain Point's cross-field universal behavior.
                point::QueryRequest preferred_request;
                preferred_request.output_fields = headers;
                preferred_request.conditions.push_back(
                    {headers.front(), lookup});
                auto result = engine->query(preferred_request);
                if (result.rows.empty())
                    result = engine->universal_lookup(headers, lookup);
                if (result.rows.empty()) {
                    // Keep the exact value that the operator entered in its
                    // lookup field.  Empty output cells made a failed lookup
                    // look like an unfinished row in large production grids,
                    // so label every related field explicitly.  The existing
                    // universal_missing_rows marker paints the complete row
                    // red, including the retained lookup value and labels.
                    std::vector<std::string> missing_row(
                        std::max<std::size_t>(headers.size(), 2),
                        "NOT FOUND");
                    missing_row.front() = lookup;
                    universal_missing_rows.insert(
                        static_cast<int>(combined.rows.size()));
                    combined.rows.push_back(std::move(missing_row));
                    update_query_progress(lookup_index + 1);
                    continue;
                }
                if (combined.rows.size() + result.rows.size() >
                    static_cast<std::size_t>(LOGICAL_ROWS)) {
                    throw std::runtime_error(
                        "The matching records exceed the 2,000,000-row "
                        "grid limit");
                }
                const bool duplicate = result.rows.size() > 1;
                for (const auto& row : result.rows) {
                    if (duplicate) {
                        universal_duplicate_rows.insert(
                            static_cast<int>(combined.rows.size()));
                    }
                    combined.rows.push_back(row);
                }
                sources.insert(
                    result.sources.begin(), result.sources.end());
                update_query_progress(lookup_index + 1);
            }
            // Preserve one result per entered lookup. If different inputs
            // resolve to the same displayed profile (for example an Employee
            // ID and that employee's Email), keep both rows and mark every
            // repeated resolved row blue instead of silently discarding it.
            std::map<std::string, std::vector<int>> resolved_row_indexes;
            for (std::size_t row = 0; row < combined.rows.size(); ++row) {
                if (universal_missing_rows.contains(
                        static_cast<int>(row))) {
                    continue;
                }
                resolved_row_indexes[result_signature(combined.rows[row])]
                    .push_back(static_cast<int>(row));
            }
            for (const auto& [signature, rows] : resolved_row_indexes) {
                (void)signature;
                if (rows.size() < 2) continue;
                universal_duplicate_rows.insert(rows.begin(), rows.end());
            }
            std::size_t duplicate_identity_column = 0;
            const std::vector<std::string> identity_priority = {
                "employeeid", "userid", "username", "email",
                "userprincipalname", "displayname", "fullname"
            };
            for (const auto& identity : identity_priority) {
                const auto found = std::find_if(
                    headers.begin(), headers.end(),
                    [&](const auto& header) {
                        return point::normalize_name(header) == identity;
                    });
                if (found != headers.end()) {
                    duplicate_identity_column =
                        static_cast<std::size_t>(
                            std::distance(headers.begin(), found));
                    break;
                }
            }
            std::map<std::string, bool> duplicate_group_shades;
            bool next_duplicate_group_is_dark = false;
            for (const int row : universal_duplicate_rows) {
                const auto& values = combined.rows[
                    static_cast<std::size_t>(row)];
                std::string group_key;
                if (duplicate_identity_column < values.size())
                    group_key = point::normalize_name(
                        values[duplicate_identity_column]);
                if (group_key.empty())
                    group_key = result_signature(values);
                auto [shade, inserted] =
                    duplicate_group_shades.try_emplace(
                        group_key, next_duplicate_group_is_dark);
                if (inserted)
                    next_duplicate_group_is_dark =
                        !next_duplicate_group_is_dark;
                if (shade->second)
                    universal_duplicate_dark_rows.insert(row);
            }
            combined.sources.assign(sources.begin(), sources.end());
            condition_count = lookups.size();
            condition_fields =
                "universal_to_" + point::normalize_name(headers.front());
        }

        const std::size_t result_capacity =
            static_cast<std::size_t>(LOGICAL_ROWS) - first_output_row;
        if (combined.rows.size() > result_capacity)
            throw std::runtime_error(
                "The matching records exceed the available grid rows");

        show_refresh_progress(L"Populating workspace results...", 90);

        std::size_t output_row = first_output_row;
        internal_cell_update = true;
        clear_cell_store();
        if (narrow_mode || count_mode) {
            for (std::size_t column = 0;
                 column < preserved_criteria.size(); ++column) {
                store_cell(
                    0, static_cast<int>(column),
                    preserved_criteria[column]);
            }
        }
        if (compare_mode) {
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t column = 0;
                 column < combined.headers.size() &&
                 column < header_cells.size(); ++column) {
                set_control_text(
                    header_cells[column], combined.headers[column]);
            }
            generated_compare_header = true;
        }
        if (count_mode) {
            generated_count_column =
                static_cast<int>(headers.size());
            set_control_text(
                header_cells[
                    static_cast<std::size_t>(generated_count_column)],
                "Count");
            const int related_column = generated_count_column + 1;
            set_control_text(
                header_cells[
                    static_cast<std::size_t>(related_column)],
                combined.headers[
                    static_cast<std::size_t>(related_column)]);
            generated_count_details = true;
            count_detail_first_column = related_column;
            count_detail_column_count = 1;
        }
        if (analyze_mode) {
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t column = 0;
                 column < combined.headers.size() &&
                 column < header_cells.size(); ++column) {
                set_control_text(
                    header_cells[column], combined.headers[column]);
            }
            generated_analysis_headers = true;
        }
        if (insight_mode) {
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t column = 0;
                 column < combined.headers.size() &&
                 column < header_cells.size(); ++column) {
                set_control_text(
                    header_cells[column], combined.headers[column]);
            }
            generated_insight_headers = true;
        }
        if (chart_mode) {
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t column = 0;
                 column < combined.headers.size() &&
                 column < header_cells.size(); ++column) {
                set_control_text(
                    header_cells[column], combined.headers[column]);
            }
            generated_chart_headers = true;
        }
        if (change_mode) {
            for (HWND header : header_cells) SetWindowTextW(header, L"");
            for (std::size_t column = 0;
                 column < combined.headers.size() &&
                 column < header_cells.size(); ++column) {
                set_control_text(
                    header_cells[column], combined.headers[column]);
            }
            generated_change_headers = true;
        }
        const std::size_t rows_to_populate = combined.rows.size();
        std::size_t values_to_store = 0;
        for (const auto& result_row : combined.rows)
            values_to_store += std::min<std::size_t>(
                result_row.size(), static_cast<std::size_t>(GRID_COLUMNS));
        cell_store.reserve(cell_store.size() + values_to_store);
        std::size_t populated = 0;
        int last_population_percent = -1;
        for (const auto& result_row : combined.rows) {
            for (std::size_t column = 0;
                 column < result_row.size() &&
                 column < static_cast<std::size_t>(GRID_COLUMNS);
                 ++column) {
                store_cell(
                    static_cast<int>(output_row),
                    static_cast<int>(column),
                    result_row[column]);
            }
            ++output_row;
            ++populated;
            if ((populated % 250U) == 0U || populated == rows_to_populate) {
                const int percent = 90 + static_cast<int>(
                    9 * populated /
                    std::max<std::size_t>(1, rows_to_populate));
                if (percent != last_population_percent) {
                    last_population_percent = percent;
                    std::wstringstream progress;
                    progress << L"Populating " << populated << L" of "
                             << rows_to_populate << L" result row(s)...";
                    show_refresh_progress(progress.str(), percent);
                }
                pump_search_messages();
            }
        }
        internal_cell_update = false;
        first_visible_row = 0;
        load_visible_cells();
        update_scrollbar();

        combined.explanation =
            change_mode
            ? "Record changes since the retained Change baseline."
            : chart_mode
            ? "Dynamic exact-filtered category distribution."
            : insight_mode
            ? "Point Insight Agent local profiling and risk ranking."
            : analyze_mode
            ? "Detailed repeated-key and missing-key record analysis."
            : compare_mode
            ? "Cyber access comparison with similarity, source lineage, "
              "risk ranking, and remediation guidance."
            : count_mode
            ? (preserved_criteria.empty() ||
               std::all_of(
                   preserved_criteria.begin(),
                   preserved_criteria.end(),
                   [](const std::string& value) {
                       return value.empty();
                   })
               ? "Identity-safe grouped occurrence distribution."
               : "Identity-safe exact selected-value occurrence count.")
            : narrow_mode
            ? "Exact multi-field AND filtering with bounded relationship "
              "joins."
            : "Universal exact-value lookup to requested output fields with "
              "bounded relationship joins.";
        const std::size_t result_count = combined.rows.size();
        last_result = std::move(combined);
        universal_results_displayed = universal_mode_active();
        universal_pending_lookup_rows.clear();
        universal_result_headers = universal_mode_active()
            ? headers : std::vector<std::string>{};
        if (update_universal_lookup_history)
            universal_lookup_history =
                std::move(next_universal_lookup_history);

        SendMessageW(refresh_progress, PBM_SETPOS, 100, 0);
        UpdateWindow(refresh_progress);
        ShowWindow(refresh_progress, SW_HIDE);
        query_progress_visible = false;

        std::wstringstream message;
        message << result_count << L" matching row(s) populated in "
                << (change_mode ? L"Change" :
                    chart_mode ? L"Chart" :
                    insight_mode ? L"Insights" :
                    analyze_mode ? L"Analyze" :
                    compare_mode ? L"Compare" :
                    count_mode ? L"Count" :
                    narrow_mode ? L"Narrow" : L"Universal")
                << L" mode.";
        if (count_mode) {
            message << L" Affected object IDs were filled automatically for "
                       L"counts of 50 or less.";
        }
        set_status(message.str());
        point::append_audit(
            app_root, "QUERY",
            "mode=" + std::string(
                change_mode ? "change" :
                chart_mode ? "chart" :
                insight_mode ? "insights" :
                analyze_mode ? "analyze" :
                compare_mode ? "compare" :
                count_mode ? "count" :
                narrow_mode ? "narrow" : "normal") +
            "; condition_fields=" + condition_fields +
            "; condition_count=" +
            std::to_string(condition_count) +
            "; results=" + std::to_string(result_count));
        if (chart_mode) open_dynamic_chart();
        query_in_progress = false;
    } catch (const std::exception& ex) {
        internal_cell_update = false;
        if (query_progress_visible) {
            ShowWindow(refresh_progress, SW_HIDE);
            query_progress_visible = false;
        }
        MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Query blocked", MB_ICONWARNING);
        point::append_audit(app_root, "QUERY_BLOCKED", ex.what());
        query_in_progress = false;
    }
}

void open_count_drilldown(int logical_row) {
    if (!engine || !count_mode || generated_count_column <= 0 ||
        logical_row < 1) {
        return;
    }
    try {
        commit_visible_cells();
        clear_count_detail_columns();
        const auto headers = selected_headers();
        const std::size_t count_column =
            static_cast<std::size_t>(generated_count_column);
        if (headers.size() <= count_column ||
            point::normalize_name(headers[count_column]) != "count") {
            throw std::runtime_error(
                "Run Count again before opening its evidence");
        }

        const auto count_text = point::trim(stored_cell(
            logical_row, generated_count_column));
        if (count_text.empty()) return;
        char* count_end = nullptr;
        const auto count_value = std::strtoull(
            count_text.c_str(), &count_end, 10);
        if (!count_end || *count_end != '\0') {
            throw std::runtime_error(
                "The selected Count value is not a valid whole number");
        }
        if (count_value > COUNT_DETAIL_LIMIT) {
            set_status(
                L"Related data is available only when Count is 50 or "
                L"less. Narrow the Count with more exact filters.");
            point::append_audit(
                app_root, "COUNT_DRILLDOWN_SKIPPED",
                "count=" + count_text + "; limit=50");
            return;
        }
        if (count_value == 0) return;

        std::vector<point::QueryCondition> conditions;
        std::vector<std::string> output_fields;
        std::set<std::string> selected_outputs;
        std::vector<std::string> grouped_fields;
        bool contains_blank_group = false;
        for (int column = 0;
             column < generated_count_column; ++column) {
            const auto field =
                headers[static_cast<std::size_t>(column)];
            const auto value =
                point::trim(stored_cell(logical_row, column));
            if (value.empty()) {
                contains_blank_group = true;
                continue;
            }
            conditions.push_back({field, value});
            grouped_fields.push_back(field);
        }
        if (conditions.empty()) {
            throw std::runtime_error(
                "This Count row contains only blank grouped values");
        }
        if (contains_blank_group) {
            throw std::runtime_error(
                "This Count row includes a blank grouped value. Add an "
                "explicit value before opening object evidence.");
        }

        const auto all_fields = engine->all_fields();
        auto add_if_available =
            [&](const std::string& requested) {
                const auto normalized =
                    point::normalize_name(requested);
                if (selected_outputs.contains(normalized)) return;
                for (const auto& available : all_fields) {
                    if (point::normalize_name(available) == normalized) {
                        selected_outputs.insert(normalized);
                        output_fields.push_back(available);
                        return;
                    }
                }
            };

        // Identity and object fields are first so the user can immediately
        // understand which people, devices, accounts, or tickets were counted.
        const std::vector<std::string> preferred_fields = {
            "Employee ID", "Username", "Full Name", "Display Name",
            "Email", "Computer ID", "Computer Name", "Asset Tag",
            "Account Name", "Group Name", "Department", "Department Name",
            "Location", "Location Name", "Device Type", "Operating System",
            "MFA Status", "Encryption Status", "Patch Status",
            "CrowdStrike Status", "Vulnerability Count",
            "Critical Vulnerabilities", "Ticket ID", "Ticket Status",
            "Risk Level", "Manager Employee ID", "Manager Name"
        };
        for (const auto& field : preferred_fields) {
            add_if_available(field);
            if (output_fields.size() >= 40) break;
        }
        for (const auto& field : grouped_fields) {
            add_if_available(field);
            if (output_fields.size() >= 40) break;
        }
        if (output_fields.empty())
            throw std::runtime_error(
                "No displayable fields were found for this Count row");

        point::QueryRequest request;
        request.conditions = conditions;
        request.output_fields = output_fields;
        auto detail = engine->query(request);
        if (detail.rows.empty()) {
            throw std::runtime_error(
                "No source objects matched this Count row. Refresh the "
                "reports and run Count again.");
        }

        const auto export_path =
            app_root / "Exports" / "point-count-drilldown.csv";
        if (export_authorized)
            engine->export_csv(detail, export_path);

        const int detail_first_column = generated_count_column + 1;
        const int available_columns =
            GRID_COLUMNS - detail_first_column;
        if (available_columns <= 0) {
            throw std::runtime_error(
                "No columns are available to the right of Count for "
                "the matching object data");
        }
        const int displayed_columns = std::min(
            available_columns,
            static_cast<int>(detail.headers.size()));
        const std::size_t displayed_rows = std::min(
            detail.rows.size(),
            static_cast<std::size_t>(LOGICAL_ROWS - 1));

        internal_cell_update = true;
        for (int column = 0;
             column < displayed_columns; ++column) {
            set_control_text(
                header_cells[static_cast<std::size_t>(
                    detail_first_column + column)],
                detail.headers[static_cast<std::size_t>(column)]);
        }
        for (std::size_t row = 0;
             row < displayed_rows; ++row) {
            for (int column = 0;
                 column < displayed_columns; ++column) {
                const auto& values = detail.rows[row];
                if (static_cast<std::size_t>(column) < values.size()) {
                    store_cell(
                        static_cast<int>(row) + 1,
                        detail_first_column + column,
                        values[static_cast<std::size_t>(column)]);
                }
            }
        }
        generated_count_details = true;
        count_detail_first_column = detail_first_column;
        count_detail_column_count = displayed_columns;
        first_visible_row = 0;
        first_visible_column = generated_count_column;
        internal_cell_update = false;
        load_visible_cells();
        update_scrollbar();
        layout(main_window);
        load_visible_cells();
        InvalidateRect(main_window, nullptr, TRUE);

        std::wostringstream status;
        status << L"Count " << widen(count_text)
               << L": showing " << displayed_rows
               << L" related object row(s) in the columns after Count";
        if (displayed_columns <
            static_cast<int>(detail.headers.size())) {
            status << L"; " << displayed_columns
                   << L" of " << detail.headers.size()
                   << L" fields fit";
        }
        if (export_authorized) {
            status << L". Full details exported to "
                      L"Exports\\point-count-drilldown.csv.";
        } else {
            status << L". Disk export requires Point Exporters membership.";
        }
        set_status(status.str());
        point::append_audit(
            app_root, "COUNT_DRILLDOWN",
            "count=" + count_text +
            "; object_rows=" + std::to_string(detail.rows.size()) +
            "; inline_columns=" +
            std::to_string(displayed_columns));
    } catch (const std::exception& ex) {
        MessageBoxW(
            main_window, widen(ex.what()).c_str(),
            L"Count evidence", MB_ICONWARNING);
        point::append_audit(
            app_root, "COUNT_DRILLDOWN_BLOCKED", ex.what());
    }
}

void export_grid() {
    if (!engine) return;
    hide_suggestions();
    try {
        if (!export_authorized)
            throw std::runtime_error(
                "Export denied: Point Exporters membership is required");
        const auto snapshot = snapshot_grid();
        if (snapshot.headers.empty())
            throw std::runtime_error("Type at least one heading before export");
        const auto path = app_root / "Exports" / "point-result.csv";
        engine->export_csv(snapshot, path);
        set_status(L"Exported securely to Exports\\point-result.csv");
        point::append_audit(
            app_root, "EXPORT",
            std::to_string(snapshot.rows.size()) + " rows");
    } catch (const std::exception& ex) {
        MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Export failed", MB_ICONERROR);
    }
}

void clear_results() {
    commit_visible_cells();
    clear_count_detail_columns();
    const int first_result_row =
        first_result_row_for_mode();
    erase_rows_from(first_result_row);
    universal_missing_rows.clear();
    universal_duplicate_rows.clear();
    universal_duplicate_dark_rows.clear();
    universal_results_displayed = false;
    universal_pending_lookup_rows.clear();
    universal_result_headers.clear();
    universal_lookup_history.clear();
    data_tools_unfiltered_result.reset();
    last_result = {};
    last_find_row = -1;
    last_find_column = -1;
    first_visible_row = 0;
    load_visible_cells();
    update_scrollbar();
    set_status(L"Results cleared; mode inputs and headings preserved.");
}

void sort_results_by_column(int column) {
    try {
        const auto snapshot = snapshot_grid();
        if (snapshot.rows.empty())
            throw std::runtime_error("There are no result rows to sort");
        if (column < 0 ||
            column >= static_cast<int>(snapshot.headers.size()))
            throw std::runtime_error("Select a populated result heading");

        if (last_sort_column == column)
            sort_ascending = !sort_ascending;
        else {
            last_sort_column = column;
            sort_ascending = true;
        }
        auto rows = snapshot.rows;
        auto numeric_value = [](const std::string& value,
                                double& output) {
            const auto cleaned = point::trim(value);
            if (cleaned.empty()) return false;
            char* end = nullptr;
            output = std::strtod(cleaned.c_str(), &end);
            return end && *end == '\0';
        };
        std::stable_sort(
            rows.begin(), rows.end(),
            [&](const auto& left, const auto& right) {
                double left_number = 0.0;
                double right_number = 0.0;
                const bool both_numeric =
                    numeric_value(
                        left[static_cast<std::size_t>(column)],
                        left_number) &&
                    numeric_value(
                        right[static_cast<std::size_t>(column)],
                        right_number);
                if (both_numeric) {
                    if (left_number == right_number) return false;
                    return sort_ascending
                        ? left_number < right_number
                        : left_number > right_number;
                }
                const auto left_text = point::normalize_name(
                    left[static_cast<std::size_t>(column)]);
                const auto right_text = point::normalize_name(
                    right[static_cast<std::size_t>(column)]);
                if (left_text == right_text) return false;
                return sort_ascending
                    ? left_text < right_text
                    : left_text > right_text;
            });

        const int first_result_row =
            first_result_row_for_mode();
        erase_rows_from(first_result_row);
        int target_row = first_result_row;
        for (const auto& row : rows) {
            for (std::size_t current_column = 0;
                 current_column < row.size() &&
                 current_column <
                     static_cast<std::size_t>(GRID_COLUMNS);
                 ++current_column) {
                store_cell(
                    target_row,
                    static_cast<int>(current_column),
                    row[current_column]);
            }
            ++target_row;
        }
        last_result = snapshot;
        last_result.rows = std::move(rows);
        first_visible_row = 0;
        load_visible_cells();
        update_scrollbar();
        set_status(
            sort_ascending
            ? L"Results sorted ascending."
            : L"Results sorted descending.");
    } catch (const std::exception& ex) {
        MessageBoxW(
            main_window, widen(ex.what()).c_str(),
            L"Sort", MB_ICONINFORMATION);
    }
}

void find_next_value() {
    commit_visible_cells();
    const auto term =
        point::normalize_name(narrow(control_text(find_text)));
    if (term.empty()) {
        MessageBoxW(
            main_window, L"Type a value in the Find box.",
            L"Find", MB_ICONINFORMATION);
        return;
    }
    const auto headers = selected_headers();
    if (headers.empty()) return;
    const int first_row = first_result_row_for_mode();
    const int last_row = highest_stored_row();
    if (last_row < first_row) {
        MessageBoxW(
            main_window, L"There are no result rows to search.",
            L"Find", MB_ICONINFORMATION);
        return;
    }
    const std::int64_t total_positions =
        static_cast<std::int64_t>(last_row - first_row + 1) *
        static_cast<std::int64_t>(headers.size());
    std::int64_t start_position = 0;
    if (last_find_row >= first_row &&
        last_find_column >= 0) {
        start_position =
            (last_find_row - first_row) *
                static_cast<int>(headers.size()) +
            last_find_column + 1;
    }
    for (std::int64_t offset = 0;
         offset < total_positions; ++offset) {
        const std::int64_t position =
            (start_position + offset) % total_positions;
        const int row =
            first_row +
            static_cast<int>(
                position /
                static_cast<std::int64_t>(headers.size()));
        const int column = static_cast<int>(
            position %
            static_cast<std::int64_t>(headers.size()));
        const auto value =
            point::normalize_name(stored_cell(row, column));
        if (value.find(term) == std::string::npos) continue;
        int requested_first_column = first_visible_column;
        if (column < first_visible_column)
            requested_first_column = column;
        else if (column >=
                 first_visible_column + visible_grid_columns)
            requested_first_column =
                column - visible_grid_columns + 1;
        if (requested_first_column != first_visible_column) {
            commit_visible_cells();
            first_visible_column = requested_first_column;
            layout(main_window);
            load_visible_cells();
        }
        scroll_to(row);
        const int visible_row = row - first_visible_row;
        if (visible_row >= 0 && visible_row < visible_grid_rows) {
            HWND control =
                data_cells[
                    static_cast<std::size_t>(visible_row)]
                          [static_cast<std::size_t>(column)];
            SetFocus(control);
            SendMessageW(control, EM_SETSEL, 0, -1);
        }
        last_find_row = row;
        last_find_column = column;
        std::wstringstream message;
        message << L"Found at result row "
                << row - first_row + 1 << L", column "
                << column + 1 << L".";
        set_status(message.str());
        return;
    }
    MessageBoxW(
        main_window, L"No matching result was found.",
        L"Find", MB_ICONINFORMATION);
}

void drill_down_chart_bar(
        const std::wstring& group, const std::wstring& label,
        const std::wstring& series) {
    if (!engine || !chart_mode || chart_fields_cache.empty() ||
        last_result.sources.size() != 1) {
        set_status(
            L"Chart selection highlighted. Drill-down requires Chart mode "
            L"with one source.");
        return;
    }
    try {
        std::vector<point::QueryCondition> conditions;
        for (const auto& bar : chart_all_bars) {
            if (bar.group == group && bar.label == label &&
                bar.series == series) {
                conditions = bar.conditions;
                break;
            }
        }
        if (conditions.empty())
            throw std::runtime_error(
                "The selected chart group is no longer available");
        for (std::size_t i = 0;
             i < chart_fields_cache.size(); ++i) {
            if (i < chart_filter_values_cache.size() &&
                !chart_filter_values_cache[i].empty() &&
                !is_chart_series_marker(
                    chart_filter_values_cache[i])) {
                conditions.push_back({
                    chart_fields_cache[i],
                    chart_filter_values_cache[i]});
            }
        }
        const auto& source = last_result.sources.front();
        const point::DataSet* dataset = nullptr;
        for (const auto& candidate : engine->datasets()) {
            if (point::normalize_name(candidate.name) ==
                point::normalize_name(source)) {
                dataset = &candidate;
                break;
            }
        }
        if (!dataset)
            throw std::runtime_error(
                "The chart source is no longer available");
        std::vector<std::string> outputs;
        for (const auto& condition : conditions) {
            bool already_selected = false;
            for (const auto& output : outputs)
                if (point::normalize_name(output) ==
                    point::normalize_name(condition.field))
                    already_selected = true;
            if (!already_selected) outputs.push_back(condition.field);
        }
        for (const auto& header : dataset->headers) {
            if (outputs.size() >=
                static_cast<std::size_t>(GRID_COLUMNS)) break;
            bool already_selected = false;
            for (const auto& output : outputs) {
                if (point::normalize_name(output) ==
                    point::normalize_name(header)) {
                    already_selected = true;
                    break;
                }
            }
            if (!already_selected) outputs.push_back(header);
        }
        auto detail = engine->exact_rows(
            conditions, outputs, source);
        internal_cell_update = true;
        clear_cell_store();
        for (HWND header : header_cells) SetWindowTextW(header, L"");
        for (std::size_t column = 0;
             column < detail.headers.size() &&
             column < header_cells.size(); ++column) {
            set_control_text(header_cells[column], detail.headers[column]);
        }
        int row = 0;
        for (const auto& values : detail.rows) {
            for (std::size_t column = 0;
                 column < values.size() &&
                 column < static_cast<std::size_t>(GRID_COLUMNS);
                 ++column) {
                store_cell(
                    row, static_cast<int>(column), values[column]);
            }
            ++row;
        }
        internal_cell_update = false;
        generated_chart_headers = true;
        first_visible_row = 0;
        load_visible_cells();
        update_scrollbar();
        last_result = std::move(detail);
        set_status(
            std::to_wstring(row) + L" underlying record(s) for " + label +
            L". Search again to restore the Chart setup.");
    } catch (const std::exception& ex) {
        internal_cell_update = false;
        MessageBoxW(
            main_window, widen(ex.what()).c_str(),
            L"Chart drill-down", MB_ICONWARNING);
    }
}

// Interactive chart helpers.
COLORREF chart_color(std::size_t index) {
    constexpr COLORREF colors[] = {
        RGB(47, 111, 237), RGB(21, 153, 112),
        RGB(130, 86, 201), RGB(238, 145, 32),
        RGB(218, 72, 92), RGB(31, 157, 181)};
    return colors[index % (sizeof(colors) / sizeof(colors[0]))];
}

std::wstring chart_view_name() {
    switch (chart_view) {
    case ChartView::Bar: return L"Bar";
    case ChartView::Column: return L"Column";
    case ChartView::Pie: return L"Pie";
    case ChartView::Stacked: return L"Stacked %";
    default: return L"Auto";
    }
}

std::wstring chart_value_text(double number) {
    std::wostringstream value;
    value << std::fixed << std::setprecision(
        std::floor(number) == number ? 0 : 2) << number;
    return value.str();
}

void rebuild_chart_view() {
    const auto filter = chart_filter
        ? point::normalize_name(narrow(control_text(chart_filter)))
        : std::string{};
    chart_bars.clear();
    for (const auto& bar : chart_all_bars) {
        if (!filter.empty() &&
            point::normalize_name(
                narrow(bar.group + L" " + bar.label + L" " +
                       bar.series)).find(filter) ==
                std::string::npos) continue;
        chart_bars.push_back(bar);
    }
    std::stable_sort(
        chart_bars.begin(), chart_bars.end(),
        [](const ChartBar& left, const ChartBar& right) {
            const auto left_group =
                point::normalize_name(narrow(left.group));
            const auto right_group =
                point::normalize_name(narrow(right.group));
            if (left_group != right_group)
                return left_group < right_group;
            if (left.value == right.value)
                return point::normalize_name(
                           narrow(left.label + L" " + left.series)) <
                       point::normalize_name(
                           narrow(right.label + L" " + right.series));
            return chart_descending
                ? left.value > right.value : left.value < right.value;
        });
    if (chart_top_limit > 0) {
        const bool has_series = std::any_of(
            chart_bars.begin(), chart_bars.end(),
            [](const ChartBar& bar) { return !bar.series.empty(); });
        if (has_series) {
            using CategoryKey = std::pair<std::wstring, std::wstring>;
            std::map<CategoryKey, double> totals;
            for (const auto& bar : chart_bars)
                totals[{bar.group, bar.label}] += bar.value;
            std::map<std::wstring,
                     std::vector<std::pair<std::wstring, double>>> ranked;
            for (const auto& [key, total] : totals)
                ranked[key.first].push_back({key.second, total});
            std::set<CategoryKey> retained_categories;
            for (auto& [group, values] : ranked) {
                std::stable_sort(
                    values.begin(), values.end(),
                    [](const auto& left, const auto& right) {
                        if (left.second != right.second)
                            return chart_descending
                                ? left.second > right.second
                                : left.second < right.second;
                        return point::normalize_name(narrow(left.first)) <
                               point::normalize_name(narrow(right.first));
                    });
                const auto keep = std::min<std::size_t>(
                    values.size(),
                    static_cast<std::size_t>(chart_top_limit));
                for (std::size_t i = 0; i < keep; ++i)
                    retained_categories.insert({group, values[i].first});
            }
            chart_bars.erase(
                std::remove_if(
                    chart_bars.begin(), chart_bars.end(),
                    [&retained_categories](const ChartBar& bar) {
                        return !retained_categories.contains(
                            CategoryKey{bar.group, bar.label});
                    }),
                chart_bars.end());
        } else {
            std::map<std::wstring, int> retained;
            chart_bars.erase(
                std::remove_if(
                    chart_bars.begin(), chart_bars.end(),
                    [&retained](const ChartBar& bar) {
                        return retained[bar.group]++ >= chart_top_limit;
                    }),
                chart_bars.end());
        }
    }
    for (auto& bar : chart_bars) bar.bounds = {};
    chart_selected_index = -1;
    if (!chart_window) return;
    const auto type_label = L"View: " + chart_view_name();
    SetWindowTextW(
        GetDlgItem(chart_window, ID_CHART_TYPE), type_label.c_str());
    const auto top_label = chart_top_limit == 0
        ? std::wstring(L"Top: All")
        : L"Top: " + std::to_wstring(chart_top_limit);
    SetWindowTextW(
        GetDlgItem(chart_window, ID_CHART_TOP), top_label.c_str());
    SetWindowTextW(
        GetDlgItem(chart_window, ID_CHART_SORT),
        chart_descending ? L"Sort: High-Low" : L"Sort: Low-High");
    InvalidateRect(chart_window, nullptr, TRUE);
}

ChartView effective_chart_view() {
    std::set<std::wstring> groups;
    bool has_series = false;
    for (const auto& bar : chart_bars) {
        if (!bar.group.empty()) groups.insert(bar.group);
        if (!bar.series.empty()) has_series = true;
    }
    if (chart_view != ChartView::Auto) return chart_view;
    if (has_series) return ChartView::Stacked;
    if (groups.size() > 1) return ChartView::Bar;
    if (chart_bars.size() <= 6) return ChartView::Pie;
    if (chart_bars.size() <= 12) return ChartView::Column;
    return ChartView::Bar;
}

LRESULT CALLBACK chart_window_proc(
        HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
    switch (message) {
    case WM_CREATE:
        chart_filter = CreateWindowW(
            L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            20, 52, 260, 28, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_CHART_FILTER)),
            nullptr, nullptr);
        SendMessageW(
            chart_filter, EM_SETCUEBANNER, TRUE,
            reinterpret_cast<LPARAM>(L"Live category search"));
        CreateWindowW(
            L"BUTTON", L"View: Auto", WS_CHILD | WS_VISIBLE,
            292, 52, 120, 28, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_CHART_TYPE)),
            nullptr, nullptr);
        CreateWindowW(
            L"BUTTON", L"Top: 20", WS_CHILD | WS_VISIBLE,
            424, 52, 100, 28, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_CHART_TOP)),
            nullptr, nullptr);
        CreateWindowW(
            L"BUTTON", L"Sort: High-Low", WS_CHILD | WS_VISIBLE,
            536, 52, 140, 28, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_CHART_SORT)),
            nullptr, nullptr);
        rebuild_chart_view();
        return 0;
    case WM_SIZE: {
        const int width = static_cast<int>(LOWORD(lparam));
        const int controls_left = std::max(300, width - 410);
        MoveWindow(
            chart_filter, 20, 52,
            std::max(140, controls_left - 32), 28, TRUE);
        MoveWindow(
            GetDlgItem(window, ID_CHART_TYPE),
            controls_left, 52, 120, 28, TRUE);
        MoveWindow(
            GetDlgItem(window, ID_CHART_TOP),
            controls_left + 126, 52, 92, 28, TRUE);
        MoveWindow(
            GetDlgItem(window, ID_CHART_SORT),
            controls_left + 224, 52, 140, 28, TRUE);
        InvalidateRect(window, nullptr, TRUE);
        return 0;
    }
    case WM_COMMAND: {
        const int id = LOWORD(wparam);
        if (id == ID_CHART_FILTER && HIWORD(wparam) == EN_CHANGE) {
            rebuild_chart_view();
            return 0;
        }
        if (id == ID_CHART_TYPE) {
            chart_view = chart_view == ChartView::Auto
                ? ChartView::Bar
                : chart_view == ChartView::Bar
                ? ChartView::Column
                : chart_view == ChartView::Column
                ? ChartView::Pie
                : chart_view == ChartView::Pie
                ? ChartView::Stacked : ChartView::Auto;
            rebuild_chart_view();
            return 0;
        }
        if (id == ID_CHART_TOP) {
            chart_top_limit = chart_top_limit == 5 ? 10 :
                chart_top_limit == 10 ? 20 :
                chart_top_limit == 20 ? 0 : 5;
            rebuild_chart_view();
            return 0;
        }
        if (id == ID_CHART_SORT) {
            chart_descending = !chart_descending;
            rebuild_chart_view();
            return 0;
        }
        return DefWindowProcW(window, message, wparam, lparam);
    }
    case WM_MOUSEWHEEL:
        chart_top_limit =
            static_cast<short>(HIWORD(wparam)) > 0
            ? (chart_top_limit == 0 ? 20 :
               chart_top_limit == 20 ? 10 : 5)
            : (chart_top_limit == 5 ? 10 :
               chart_top_limit == 10 ? 20 : 0);
        rebuild_chart_view();
        return 0;
    case WM_PAINT: {
        PAINTSTRUCT paint{};
        HDC context = BeginPaint(window, &paint);
        RECT area{};
        GetClientRect(window, &area);
        const int client_width =
            static_cast<int>(area.right - area.left);
        const int client_height =
            static_cast<int>(area.bottom - area.top);
        FillRect(context, &area,
                 reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1));
        SetBkMode(context, TRANSPARENT);
        HFONT title_font = CreateFontW(
            24, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        HFONT old_font = reinterpret_cast<HFONT>(
            SelectObject(context, title_font));
        TextOutW(context, 24, 18, chart_title.c_str(),
                 static_cast<int>(chart_title.size()));
        SelectObject(context, old_font);
        DeleteObject(title_font);

        if (chart_bars.empty()) {
            const std::wstring empty =
                L"No numeric result data is available.";
            TextOutW(context, 24, 110, empty.c_str(),
                     static_cast<int>(empty.size()));
            EndPaint(window, &paint);
            return 0;
        }
        const double maximum = std::max_element(
            chart_bars.begin(), chart_bars.end(),
            [](const ChartBar& left, const ChartBar& right) {
                return left.value < right.value;
            })->value;
        const int top = 105;
        HFONT body_font = CreateFontW(
            17, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        old_font = reinterpret_cast<HFONT>(
            SelectObject(context, body_font));
        const ChartView view = effective_chart_view();
        std::vector<std::wstring> dashboard_groups;
        bool dashboard_has_series = false;
        for (const auto& bar : chart_bars) {
            if (!bar.series.empty()) dashboard_has_series = true;
            if (!bar.group.empty() &&
                std::find(dashboard_groups.begin(),
                          dashboard_groups.end(), bar.group) ==
                    dashboard_groups.end())
                dashboard_groups.push_back(bar.group);
        }
        if (dashboard_groups.size() > 1 || dashboard_has_series) {
            const int panel_columns =
                dashboard_groups.size() <= 2 ? 1 : 2;
            const int panel_rows = static_cast<int>(
                (dashboard_groups.size() + panel_columns - 1) /
                panel_columns);
            const int available_height =
                std::max(160, client_height - top - 20);
            const int panel_width =
                std::max(260, (client_width - 30) / panel_columns);
            const int panel_height =
                std::max(150, available_height / panel_rows);
            for (std::size_t group_index = 0;
                 group_index < dashboard_groups.size(); ++group_index) {
                const int panel_column =
                    static_cast<int>(group_index) % panel_columns;
                const int panel_row =
                    static_cast<int>(group_index) / panel_columns;
                const int panel_left = 15 + panel_column * panel_width;
                const int panel_top = top + panel_row * panel_height;
                RECT panel_rect{
                    panel_left, panel_top,
                    panel_left + panel_width - 10,
                    panel_top + panel_height - 10};
                FrameRect(
                    context, &panel_rect,
                    reinterpret_cast<HBRUSH>(
                        GetStockObject(LTGRAY_BRUSH)));
                RECT group_title{
                    panel_left + 10, panel_top + 6,
                    panel_rect.right - 8, panel_top + 30};
                std::wstring panel_title =
                    dashboard_groups[group_index];
                std::vector<std::wstring> panel_series;
                for (const auto& candidate : chart_bars) {
                    if (candidate.group != dashboard_groups[group_index] ||
                        candidate.series.empty() ||
                        std::find(panel_series.begin(), panel_series.end(),
                                  candidate.series) != panel_series.end())
                        continue;
                    panel_series.push_back(candidate.series);
                }
                if (!panel_series.empty()) {
                    panel_title += L"  |  Series: ";
                    for (std::size_t i = 0; i < panel_series.size(); ++i) {
                        if (i) panel_title += L", ";
                        panel_title += panel_series[i];
                        if (panel_title.size() > 80) {
                            panel_title += L"…";
                            break;
                        }
                    }
                }
                DrawTextW(
                    context, panel_title.c_str(), -1,
                    &group_title,
                    DT_LEFT | DT_VCENTER | DT_SINGLELINE |
                    DT_END_ELLIPSIS);
                std::vector<std::size_t> group_bars;
                double group_maximum = 0.0;
                for (std::size_t i = 0; i < chart_bars.size(); ++i) {
                    if (chart_bars[i].group ==
                        dashboard_groups[group_index]) {
                        group_bars.push_back(i);
                        group_maximum = std::max(
                            group_maximum, chart_bars[i].value);
                    }
                }
                if (group_bars.empty()) continue;
                if (view == ChartView::Stacked && dashboard_has_series) {
                    std::vector<std::wstring> categories;
                    for (const auto index : group_bars)
                        if (std::find(categories.begin(), categories.end(),
                                      chart_bars[index].label) ==
                            categories.end())
                            categories.push_back(chart_bars[index].label);
                    std::stable_sort(
                        categories.begin(), categories.end(),
                        [&group_bars](const auto& left, const auto& right) {
                            const auto total_for = [&group_bars](
                                    const std::wstring& category) {
                                double total = 0.0;
                                for (const auto index : group_bars)
                                    if (chart_bars[index].label == category)
                                        total += chart_bars[index].value;
                                return total;
                            };
                            const double left_total = total_for(left);
                            const double right_total = total_for(right);
                            if (left_total != right_total)
                                return chart_descending
                                    ? left_total > right_total
                                    : left_total < right_total;
                            return point::normalize_name(narrow(left)) <
                                   point::normalize_name(narrow(right));
                        });
                    const int label_width =
                        std::min(135, panel_width / 3);
                    const int graph_left =
                        panel_left + label_width + 18;
                    const int graph_width = std::max(
                        70, static_cast<int>(panel_rect.right) -
                            graph_left - 48);
                    const int row_height = std::max(
                        20, std::min(31,
                            (panel_height - 44) /
                            static_cast<int>(categories.size())));
                    for (std::size_t category_index = 0;
                         category_index < categories.size();
                         ++category_index) {
                        const int y = panel_top + 35 +
                            static_cast<int>(category_index) * row_height;
                        RECT label_rect{
                            panel_left + 8, y,
                            panel_left + label_width + 8, y + row_height};
                        DrawTextW(
                            context, categories[category_index].c_str(), -1,
                            &label_rect,
                            DT_LEFT | DT_VCENTER | DT_SINGLELINE |
                            DT_END_ELLIPSIS);
                        double total = 0.0;
                        for (const auto index : group_bars)
                            if (chart_bars[index].label ==
                                categories[category_index])
                                total += chart_bars[index].value;
                        int segment_left = graph_left;
                        for (const auto index : group_bars) {
                            auto& bar = chart_bars[index];
                            if (bar.label != categories[category_index])
                                continue;
                            const int segment_width = total > 0.0
                                ? static_cast<int>(std::round(
                                    graph_width * bar.value / total)) : 0;
                            const int segment_right = std::min(
                                graph_left + graph_width,
                                segment_left + std::max(2, segment_width));
                            bar.bounds = {
                                segment_left, y + 4,
                                segment_right, y + row_height - 4};
                            HBRUSH brush = CreateSolidBrush(chart_color(
                                std::hash<std::wstring>{}(bar.series)));
                            FillRect(context, &bar.bounds, brush);
                            DeleteObject(brush);
                            if (static_cast<int>(index) ==
                                chart_selected_index)
                                FrameRect(
                                    context, &bar.bounds,
                                    reinterpret_cast<HBRUSH>(
                                        GetStockObject(BLACK_BRUSH)));
                            segment_left = segment_right;
                        }
                        const auto total_text = chart_value_text(total);
                        RECT total_rect{
                            graph_left + graph_width + 5, y,
                            panel_rect.right - 3, y + row_height};
                        DrawTextW(
                            context, total_text.c_str(), -1, &total_rect,
                            DT_LEFT | DT_VCENTER | DT_SINGLELINE);
                    }
                } else {
                const int label_width =
                    std::min(150, panel_width / 3);
                const int graph_left =
                    panel_left + label_width + 18;
                const int graph_width = std::max(
                    55, static_cast<int>(panel_rect.right) -
                        graph_left - 45);
                const int bar_height = std::max(
                    17, std::min(28,
                        (panel_height - 42) /
                        static_cast<int>(group_bars.size())));
                for (std::size_t local = 0;
                     local < group_bars.size(); ++local) {
                    const auto index = group_bars[local];
                    auto& bar = chart_bars[index];
                    const int y = panel_top + 34 +
                        static_cast<int>(local) * bar_height;
                    RECT label_rect{
                        panel_left + 8, y,
                        panel_left + label_width + 8,
                        y + bar_height};
                    const auto dashboard_label = bar.series.empty()
                        ? bar.label
                        : bar.label + L" — " + bar.series;
                    DrawTextW(
                        context, dashboard_label.c_str(), -1, &label_rect,
                        DT_LEFT | DT_VCENTER | DT_SINGLELINE |
                        DT_END_ELLIPSIS);
                    const int width = group_maximum > 0.0
                        ? static_cast<int>(
                            graph_width * bar.value / group_maximum) : 0;
                    bar.bounds = {
                        graph_left, y + 4,
                        graph_left + std::max(2, width),
                        y + bar_height - 4};
                    const std::size_t color_index = bar.series.empty()
                        ? local
                        : std::hash<std::wstring>{}(bar.series);
                    HBRUSH brush = CreateSolidBrush(
                        chart_color(color_index));
                    FillRect(context, &bar.bounds, brush);
                    DeleteObject(brush);
                    if (static_cast<int>(index) == chart_selected_index)
                        FrameRect(
                            context, &bar.bounds,
                            reinterpret_cast<HBRUSH>(
                                GetStockObject(BLACK_BRUSH)));
                    const auto value = chart_value_text(bar.value);
                    RECT value_rect{
                        bar.bounds.right + 5, y,
                        panel_rect.right - 4, y + bar_height};
                    DrawTextW(
                        context, value.c_str(), -1, &value_rect,
                        DT_LEFT | DT_VCENTER | DT_SINGLELINE);
                }
                }
            }
        } else if (view == ChartView::Bar) {
            const int label_width = std::min(280, client_width / 3);
            const int graph_left = label_width + 40;
            const int graph_width =
                std::max(100, client_width - graph_left - 90);
            const int available_height =
                std::max(100, client_height - top - 30);
            const int bar_height = std::max(
                18, std::min(34, available_height /
                    static_cast<int>(chart_bars.size())));
            for (std::size_t i = 0; i < chart_bars.size(); ++i) {
                auto& bar = chart_bars[i];
                const int y =
                    top + static_cast<int>(i) * bar_height;
                RECT label_rect{
                    20, y, label_width + 20, y + bar_height};
                DrawTextW(
                    context, bar.label.c_str(), -1, &label_rect,
                    DT_LEFT | DT_VCENTER | DT_SINGLELINE |
                    DT_END_ELLIPSIS);
                const int width = maximum > 0.0
                    ? static_cast<int>(
                        graph_width * bar.value / maximum) : 0;
                bar.bounds = {
                    graph_left, y + 4,
                    graph_left + std::max(2, width),
                    y + bar_height - 4};
                HBRUSH brush = CreateSolidBrush(chart_color(i));
                FillRect(context, &bar.bounds, brush);
                DeleteObject(brush);
                if (static_cast<int>(i) == chart_selected_index)
                    FrameRect(
                        context, &bar.bounds,
                        reinterpret_cast<HBRUSH>(
                            GetStockObject(BLACK_BRUSH)));
                const auto value = chart_value_text(bar.value);
                RECT value_rect{
                    bar.bounds.right + 8, y,
                    client_width - 10, y + bar_height};
                DrawTextW(
                    context, value.c_str(), -1, &value_rect,
                    DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            }
        } else if (view == ChartView::Column) {
            const int left = 55;
            const int bottom = client_height - 62;
            const int graph_height = std::max(80, bottom - top);
            const int slot = std::max(
                20, (client_width - left - 25) /
                static_cast<int>(chart_bars.size()));
            MoveToEx(context, left, top, nullptr);
            LineTo(context, left, bottom);
            LineTo(context, client_width - 20, bottom);
            for (std::size_t i = 0; i < chart_bars.size(); ++i) {
                auto& bar = chart_bars[i];
                const int height = maximum > 0.0
                    ? static_cast<int>(
                        graph_height * bar.value / maximum) : 0;
                const int x =
                    left + static_cast<int>(i) * slot + 4;
                bar.bounds = {
                    x, bottom - std::max(2, height),
                    x + std::max(8, slot - 8), bottom};
                HBRUSH brush = CreateSolidBrush(chart_color(i));
                FillRect(context, &bar.bounds, brush);
                DeleteObject(brush);
                if (static_cast<int>(i) == chart_selected_index)
                    FrameRect(
                        context, &bar.bounds,
                        reinterpret_cast<HBRUSH>(
                            GetStockObject(BLACK_BRUSH)));
                const auto value = chart_value_text(bar.value);
                RECT value_rect{
                    x, bar.bounds.top - 20,
                    x + slot, bar.bounds.top};
                DrawTextW(
                    context, value.c_str(), -1, &value_rect,
                    DT_CENTER | DT_SINGLELINE);
                RECT label_rect{
                    x, bottom + 3, x + slot, client_height - 4};
                DrawTextW(
                    context, bar.label.c_str(), -1, &label_rect,
                    DT_CENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            }
        } else {
            const int diameter = std::max(
                120, std::min(client_height - top - 25,
                              client_width / 2 - 30));
            RECT pie_rect{
                25, top, 25 + diameter, top + diameter};
            const double total = std::accumulate(
                chart_bars.begin(), chart_bars.end(), 0.0,
                [](double sum, const ChartBar& bar) {
                    return sum + bar.value;
                });
            double angle = 0.0;
            for (std::size_t i = 0; i < chart_bars.size(); ++i) {
                auto& bar = chart_bars[i];
                const double sweep = total > 0.0
                    ? 6.283185307179586 * bar.value / total : 0.0;
                const int center_x =
                    static_cast<int>(
                        (pie_rect.left + pie_rect.right) / 2);
                const int center_y =
                    static_cast<int>(
                        (pie_rect.top + pie_rect.bottom) / 2);
                const int radius = diameter / 2;
                POINT start{
                    center_x + static_cast<int>(
                        radius * std::cos(angle)),
                    center_y - static_cast<int>(
                        radius * std::sin(angle))};
                POINT end{
                    center_x + static_cast<int>(
                        radius * std::cos(angle + sweep)),
                    center_y - static_cast<int>(
                        radius * std::sin(angle + sweep))};
                HBRUSH brush = CreateSolidBrush(chart_color(i));
                HBRUSH old_brush = reinterpret_cast<HBRUSH>(
                    SelectObject(context, brush));
                if (chart_bars.size() == 1) {
                    Ellipse(
                        context, pie_rect.left, pie_rect.top,
                        pie_rect.right, pie_rect.bottom);
                } else {
                    Pie(
                        context, pie_rect.left, pie_rect.top,
                        pie_rect.right, pie_rect.bottom,
                        start.x, start.y, end.x, end.y);
                }
                SelectObject(context, old_brush);
                DeleteObject(brush);
                const int legend_left =
                    static_cast<int>(pie_rect.right) + 28;
                const int legend_top =
                    top + static_cast<int>(i) * 28;
                bar.bounds = {
                    legend_left, legend_top,
                    client_width - 18, legend_top + 24};
                RECT swatch{
                    legend_left, legend_top + 3,
                    legend_left + 18, legend_top + 21};
                brush = CreateSolidBrush(chart_color(i));
                FillRect(context, &swatch, brush);
                DeleteObject(brush);
                std::wostringstream legend;
                legend << bar.label << L" — "
                       << chart_value_text(bar.value);
                if (total > 0.0)
                    legend << L" (" << std::fixed
                           << std::setprecision(1)
                           << 100.0 * bar.value / total << L"%)";
                RECT legend_text{
                    legend_left + 25, legend_top,
                    client_width - 18, legend_top + 24};
                DrawTextW(
                    context, legend.str().c_str(), -1, &legend_text,
                    DT_LEFT | DT_VCENTER | DT_SINGLELINE |
                    DT_END_ELLIPSIS);
                if (static_cast<int>(i) == chart_selected_index)
                    FrameRect(
                        context, &bar.bounds,
                        reinterpret_cast<HBRUSH>(
                            GetStockObject(BLACK_BRUSH)));
                angle += sweep;
            }
        }
        SelectObject(context, old_font);
        DeleteObject(body_font);
        EndPaint(window, &paint);
        return 0;
    }
    case WM_MOUSEMOVE: {
        POINT point{
            static_cast<short>(LOWORD(lparam)),
            static_cast<short>(HIWORD(lparam))};
        for (const auto& bar : chart_bars) {
            if (PtInRect(&bar.bounds, point)) {
                double category_total = 0.0;
                for (const auto& candidate : chart_bars)
                    if (candidate.group == bar.group &&
                        candidate.label == bar.label)
                        category_total += candidate.value;
                std::wostringstream title;
                title << L"Point Dynamic Chart — "
                      << (bar.group.empty() ? L"" : bar.group + L" / ")
                      << bar.label;
                if (!bar.series.empty())
                    title << L" / " << bar.series;
                title << L": " << chart_value_text(bar.value);
                if (!bar.series.empty() && category_total > 0.0)
                    title << L" (" << std::fixed << std::setprecision(1)
                          << 100.0 * bar.value / category_total << L"%)";
                SetWindowTextW(window, title.str().c_str());
                return 0;
            }
        }
        SetWindowTextW(window, L"Point Dynamic Chart");
        return 0;
    }
    case WM_LBUTTONDOWN: {
        POINT point{
            static_cast<short>(LOWORD(lparam)),
            static_cast<short>(HIWORD(lparam))};
        for (std::size_t i = 0; i < chart_bars.size(); ++i) {
            const auto& bar = chart_bars[i];
            if (PtInRect(&bar.bounds, point)) {
                chart_selected_index = static_cast<int>(i);
                InvalidateRect(window, nullptr, TRUE);
                drill_down_chart_bar(
                    bar.group, bar.label, bar.series);
                return 0;
            }
        }
        return 0;
    }
    case WM_DESTROY:
        chart_window = nullptr;
        chart_filter = nullptr;
        chart_bars.clear();
        chart_all_bars.clear();
        return 0;
    default:
        return DefWindowProcW(window, message, wparam, lparam);
    }
    return DefWindowProcW(window, message, wparam, lparam);
}

void open_dynamic_chart() {
    try {
        if (chart_window) DestroyWindow(chart_window);
        const auto snapshot = snapshot_grid();
        if (snapshot.rows.empty())
            throw std::runtime_error(
                "Run a query or Insight analysis before opening a chart");
        int numeric_column = -1;
        for (int column =
                 static_cast<int>(snapshot.headers.size()) - 1;
             column >= 0; --column) {
            bool any = false;
            bool numeric = true;
            for (const auto& row : snapshot.rows) {
                if (column >= static_cast<int>(row.size())) continue;
                const auto value = point::trim(
                    row[static_cast<std::size_t>(column)]);
                if (value.empty()) continue;
                char* end = nullptr;
                (void)std::strtod(value.c_str(), &end);
                if (!end || *end != '\0') {
                    numeric = false;
                    break;
                }
                any = true;
            }
            if (numeric && any) {
                numeric_column = column;
                break;
            }
        }
        if (numeric_column < 0)
            throw std::runtime_error(
                "The current result has no numeric column to chart");

        std::vector<ChartBar> bars;
        const bool dashboard_layout =
            snapshot.headers.size() >= 3 &&
            point::normalize_name(snapshot.headers[0]) == "chartfield" &&
            point::normalize_name(snapshot.headers[1]) == "category";
        const bool dashboard_series_layout =
            dashboard_layout && snapshot.headers.size() >= 4 &&
            point::normalize_name(snapshot.headers[2]) == "series";
        for (const auto& row : snapshot.rows) {
            if (numeric_column >= static_cast<int>(row.size())) continue;
            char* end = nullptr;
            const double value = std::strtod(
                row[static_cast<std::size_t>(numeric_column)].c_str(),
                &end);
            if (!end || *end != '\0' || !std::isfinite(value) ||
                value < 0.0) {
                continue;
            }
            std::string group;
            std::string label;
            std::string series;
            if (dashboard_layout && row.size() >= 3) {
                group = row[0];
                label = row[1];
                if (dashboard_series_layout && row.size() >= 4)
                    series = row[2];
            } else if (insight_mode && row.size() > 2)
                label = row[2] + " — " + row[1];
            else {
                for (int column = 0;
                     column < numeric_column; ++column) {
                    if (column >= static_cast<int>(row.size()) ||
                        point::trim(row[
                            static_cast<std::size_t>(column)]).empty()) {
                        continue;
                    }
                    if (!label.empty()) label += " / ";
                    label += row[static_cast<std::size_t>(column)];
                    if (label.size() > 120) break;
                }
            }
            if (label.empty()) label = "(unlabeled)";
            std::vector<point::QueryCondition> conditions;
            if (chart_mode) {
                if (dashboard_layout) {
                    conditions.push_back({group, label});
                    if (dashboard_series_layout && row.size() >= 4) {
                        for (std::size_t i = 0;
                             i < chart_fields_cache.size(); ++i) {
                            if (i < chart_filter_values_cache.size() &&
                                is_chart_series_marker(
                                    chart_filter_values_cache[i])) {
                                conditions.push_back({
                                    chart_fields_cache[i], series});
                                break;
                            }
                        }
                    }
                } else for (int column = 0;
                            column < numeric_column; ++column) {
                    if (column >= static_cast<int>(snapshot.headers.size()) ||
                        column >= static_cast<int>(row.size())) continue;
                    conditions.push_back({
                        snapshot.headers[static_cast<std::size_t>(column)],
                        row[static_cast<std::size_t>(column)]});
                }
            }
            bars.push_back({
                widen(group), widen(label), widen(series), value,
                std::move(conditions), {}});
        }
        chart_all_bars = std::move(bars);
        chart_bars.clear();
        chart_view = ChartView::Auto;
        chart_top_limit = dashboard_layout ? 5 : 20;
        chart_descending = true;
        chart_selected_index = -1;
        chart_title = dashboard_layout
            ? L"Point Multi-Field Chart Dashboard"
            : L"Top results by " +
              widen(snapshot.headers[
                  static_cast<std::size_t>(numeric_column)]);
        chart_window = CreateWindowExW(
            WS_EX_APPWINDOW, L"PointChartWindow",
            L"Point Dynamic Chart",
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            CW_USEDEFAULT, CW_USEDEFAULT, 1000, 720,
            main_window, nullptr, GetModuleHandleW(nullptr), nullptr);
        if (!chart_window)
            throw std::runtime_error(
                "Point could not create the chart window");
        ShowWindow(chart_window, SW_SHOW);
        UpdateWindow(chart_window);
    } catch (const std::exception& ex) {
        MessageBoxW(
            main_window, widen(ex.what()).c_str(),
            L"Dynamic Chart", MB_ICONINFORMATION);
    }
}

void show_import_summary() {
    if (!engine) return;
    std::wstringstream summary;
    summary << L"Imported reports: "
            << engine->datasets().size() << L"\r\n\r\n";
    for (const auto& dataset : engine->datasets()) {
        summary << widen(dataset.name) << L"\r\n"
                << L"  Rows: " << dataset.rows.size()
                << L"  Fields: " << dataset.headers.size()
                << L"\r\n";
    }
    if (!engine->issues().empty()) {
        summary << L"\r\nRejected reports: "
                << engine->issues().size() << L"\r\n";
        for (const auto& issue : engine->issues()) {
            summary << widen(issue.path.filename().string())
                    << L": " << widen(issue.message) << L"\r\n";
        }
    }
    auto text = summary.str();
    if (text.size() > 30000)
        text = text.substr(0, 30000) +
               L"\r\nSummary truncated.";
    MessageBoxW(
        main_window, text.c_str(),
        L"Point Import Summary", MB_ICONINFORMATION);
}

void select_next_source() {
    if (!engine || engine->datasets().empty()) return;
    ++selected_source_index;
    if (selected_source_index >=
        static_cast<int>(engine->datasets().size())) {
        selected_source_index = -1;
    }
    if (selected_source_index < 0)
        set_status(L"Source selection: Automatic.");
    else
        set_status(
            L"Selected source: " +
            widen(selected_source_name()));
}

bool is_grid_edit(HWND control) {
    if (!control) return false;
    const int id = GetDlgCtrlID(control);
    return is_header_id(id) ||
           (id >= ID_CELL_BASE &&
            id < ID_CELL_BASE + GRID_CONTROL_ROWS * GRID_COLUMNS);
}

bool write_clipboard_text(const std::wstring& text) {
    if (!OpenClipboard(main_window)) return false;
    EmptyClipboard();
    const std::size_t bytes =
        (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!memory) {
        CloseClipboard();
        return false;
    }
    void* destination = GlobalLock(memory);
    if (!destination) {
        GlobalFree(memory);
        CloseClipboard();
        return false;
    }
    std::memcpy(destination, text.c_str(), bytes);
    GlobalUnlock(memory);
    if (!SetClipboardData(CF_UNICODETEXT, memory)) {
        GlobalFree(memory);
        CloseClipboard();
        return false;
    }
    CloseClipboard();
    return true;
}

std::wstring clipboard_safe_cell(std::wstring value) {
    for (wchar_t& ch : value) {
        if (ch == L'\t' || ch == L'\r' || ch == L'\n') ch = L' ';
    }
    return value;
}

void copy_selected_range() {
    if (!selection_active) return;
    commit_visible_cells();
    const int first_row =
        std::min(selection_anchor.row, selection_end.row);
    const int last_row =
        std::max(selection_anchor.row, selection_end.row);
    const int first_column =
        std::min(selection_anchor.column, selection_end.column);
    const int last_column =
        std::max(selection_anchor.column, selection_end.column);
    const std::size_t maximum_characters =
        50ull * 1024ull * 1024ull / sizeof(wchar_t);
    std::wstring clipboard;

    try {
        for (int row = first_row; row <= last_row; ++row) {
            for (int column = first_column;
                 column <= last_column; ++column) {
                if (column != first_column) clipboard.push_back(L'\t');
                std::wstring value;
                if (row == -1) {
                    value = control_text(
                        header_cells[static_cast<std::size_t>(column)]);
                } else {
                    value = widen(stored_cell(row, column));
                }
                clipboard += clipboard_safe_cell(std::move(value));
            }
            clipboard += L"\r\n";
            if (clipboard.size() > maximum_characters)
                throw std::runtime_error(
                    "Selection exceeds the 50 MiB clipboard safety limit");
        }
        if (!write_clipboard_text(clipboard))
            throw std::runtime_error("Windows clipboard is unavailable");
        std::wstringstream message;
        message << L"Copied selected range: "
                << last_row - first_row + 1 << L" row(s) × "
                << last_column - first_column + 1 << L" column(s).";
        set_status(message.str());
    } catch (const std::exception& ex) {
        MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Copy blocked", MB_ICONWARNING);
    }
}

void clear_selected_range() {
    if (!selection_active) return;
    commit_visible_cells();
    const int first_row =
        std::min(selection_anchor.row, selection_end.row);
    const int last_row =
        std::max(selection_anchor.row, selection_end.row);
    const int first_column =
        std::min(selection_anchor.column, selection_end.column);
    const int last_column =
        std::max(selection_anchor.column, selection_end.column);

    internal_cell_update = true;
    if (first_row <= -1) {
        for (int column = first_column;
             column <= last_column; ++column) {
            SetWindowTextW(
                header_cells[static_cast<std::size_t>(column)], L"");
        }
    }
    for (auto iterator = cell_store.begin();
         iterator != cell_store.end();) {
        const int row = static_cast<int>(
            iterator->first /
            static_cast<std::uint64_t>(GRID_COLUMNS));
        const int column = static_cast<int>(
            iterator->first %
            static_cast<std::uint64_t>(GRID_COLUMNS));
        if (row >= std::max(0, first_row) &&
            row <= last_row &&
            column >= first_column &&
            column <= last_column) {
            iterator = cell_store.erase(iterator);
        } else {
            ++iterator;
        }
    }
    invalidate_cell_row_indexes();
    internal_cell_update = false;
    load_visible_cells();
    repaint_grid_cells();
    set_status(L"Selected cells cleared.");
}

void select_used_grid(HWND focus) {
    GridPosition current{};
    if (!grid_position_from_control(focus, current)) return;
    const int last_column = std::max(
        0, static_cast<int>(selected_headers().size()) - 1);
    const int last_row = std::max(0, highest_stored_row());
    selection_anchor = {-1, 0};
    selection_end = {last_row, last_column};
    selection_active = true;
    repaint_grid_cells();
    set_status(
        L"Used grid selected. Ctrl+C copies; Delete clears.");
}

std::string active_mode_name() {
    if (change_mode) return "change";
    if (chart_mode) return "chart";
    if (insight_mode) return "insights";
    if (analyze_mode) return "analyze";
    if (compare_mode) return "compare";
    if (count_mode) return "count";
    if (narrow_mode) return "narrow";
    return "normal";
}

void save_workspace_view() {
    try {
        commit_visible_cells();
        const auto path =
            app_root / "Workspace" / "point-saved-view.txt";
        const auto temporary_path =
            app_root / "Workspace" / "point-saved-view.tmp";
        std::ofstream output(
            temporary_path,
            std::ios::binary | std::ios::trunc);
        if (!output)
            throw std::runtime_error(
                "Unable to create the saved workspace view");
        output << "POINT_VIEW_V3\n"
               << GRID_COLUMNS << '\n'
               << grid_row_height << '\n';
        for (const auto width : column_widths)
            output << width << '\n';
        output << active_mode_name() << '\n'
               << std::quoted(selected_source_name()) << '\n';
        for (HWND header : header_cells)
            output << std::quoted(
                narrow(control_text(header))) << '\n';
        for (int row = 0; row < 2; ++row)
            for (int column = 0; column < GRID_COLUMNS; ++column)
                output << std::quoted(
                    stored_cell(row, column)) << '\n';
        if (!output)
            throw std::runtime_error(
                "Saved workspace view write failed");
        output.close();
        if (!output)
            throw std::runtime_error(
                "Saved workspace view could not be finalized");
        point::compliance::protect_file_for_current_user(
            temporary_path);
        if (!MoveFileExW(
                temporary_path.c_str(), path.c_str(),
                MOVEFILE_REPLACE_EXISTING |
                MOVEFILE_WRITE_THROUGH)) {
            std::error_code cleanup_error;
            std::filesystem::remove(
                temporary_path, cleanup_error);
            throw std::runtime_error(
                "Saved workspace view could not be published atomically");
        }
        set_status(
            L"Encrypted workspace view saved for this Windows user.");
        point::append_audit(
            app_root, "VIEW_SAVED",
            "mode=" + active_mode_name());
    } catch (const std::exception& ex) {
        MessageBoxW(
            main_window, widen(ex.what()).c_str(),
            L"Save View", MB_ICONERROR);
    }
}

void load_workspace_view() {
    try {
        const auto path =
            app_root / "Workspace" / "point-saved-view.txt";
        const auto protected_text =
            point::compliance::read_user_protected_file(path);
        std::istringstream input(protected_text);
        std::string signature;
        std::string mode;
        std::string source;
        int saved_columns = 0;
        int saved_row_height = 28;
        std::vector<int> saved_column_widths(
            static_cast<std::size_t>(GRID_COLUMNS),
            MIN_COLUMN_WIDTH);
        auto read_bounded_integer =
            [&](const char* label, int minimum, int maximum) {
                std::string text;
                if (!std::getline(input, text))
                    throw std::runtime_error(
                        std::string("Saved workspace ") +
                        label + " is missing");
                std::size_t consumed = 0;
                int value = 0;
                try {
                    value = std::stoi(text, &consumed);
                } catch (...) {
                    throw std::runtime_error(
                        std::string("Saved workspace ") +
                        label + " is invalid");
                }
                if (consumed != text.size() ||
                    value < minimum || value > maximum) {
                    throw std::runtime_error(
                        std::string("Saved workspace ") +
                        label + " is out of range");
                }
                return value;
            };
        if (!std::getline(input, signature)) {
            throw std::runtime_error(
                "Saved workspace view is missing or invalid");
        }
        if (signature == "POINT_VIEW_V1") {
            saved_columns = 6;
        } else if (signature == "POINT_VIEW_V2") {
            saved_columns = read_bounded_integer(
                "column count", 1, GRID_COLUMNS);
        } else if (signature == "POINT_VIEW_V3") {
            saved_columns = read_bounded_integer(
                "column count", 1, GRID_COLUMNS);
            saved_row_height = read_bounded_integer(
                "row height", 20, 80);
            for (int column = 0;
                 column < GRID_COLUMNS; ++column) {
                saved_column_widths[
                    static_cast<std::size_t>(column)] =
                    read_bounded_integer(
                        "column width", 60, 600);
            }
        } else {
            throw std::runtime_error(
                "Saved workspace view is missing or invalid");
        }
        if (saved_columns < 1 || saved_columns > GRID_COLUMNS ||
            !std::getline(input, mode) ||
            !(input >> std::quoted(source))) {
            throw std::runtime_error(
                "Saved workspace view is missing or invalid");
        }
        std::vector<std::string> headers(GRID_COLUMNS);
        std::vector<std::string> values(
            2 * static_cast<std::size_t>(GRID_COLUMNS));
        for (int column = 0; column < saved_columns; ++column) {
            auto& header =
                headers[static_cast<std::size_t>(column)];
            if (!(input >> std::quoted(header)) ||
                header.size() > 4096)
                throw std::runtime_error(
                    "Saved view heading is invalid");
        }
        for (int row = 0; row < 2; ++row) {
            for (int column = 0; column < saved_columns; ++column) {
                auto& value = values[
                    static_cast<std::size_t>(
                        row * GRID_COLUMNS + column)];
                if (!(input >> std::quoted(value)) ||
                    value.size() > 1024ull * 1024ull)
                    throw std::runtime_error(
                        "Saved view cell is invalid");
            }
        }

        clear_grid();
        column_widths = std::move(saved_column_widths);
        grid_row_height = saved_row_height;
        narrow_mode = mode == "narrow";
        count_mode = mode == "count";
        compare_mode = mode == "compare";
        analyze_mode = mode == "analyze";
        insight_mode = mode == "insights";
        chart_mode = mode == "chart";
        change_mode = mode == "change";
        if (mode != "normal" && !narrow_mode && !count_mode &&
            !compare_mode && !analyze_mode && !insight_mode &&
            !chart_mode && !change_mode)
            throw std::runtime_error(
                "Saved workspace mode is invalid");
        selected_source_index = -1;
        if (engine && !source.empty()) {
            for (std::size_t i = 0;
                 i < engine->datasets().size(); ++i) {
                if (point::normalize_name(
                        engine->datasets()[i].name) ==
                    point::normalize_name(source)) {
                    selected_source_index = static_cast<int>(i);
                    break;
                }
            }
        }
        internal_cell_update = true;
        for (int column = 0; column < GRID_COLUMNS; ++column) {
            set_control_text(
                header_cells[static_cast<std::size_t>(column)],
                headers[static_cast<std::size_t>(column)]);
        }
        for (int row = 0; row < 2; ++row)
            for (int column = 0; column < GRID_COLUMNS; ++column) {
                const auto& value = values[
                    static_cast<std::size_t>(
                        row * GRID_COLUMNS + column)];
                if (!value.empty()) store_cell(row, column, value);
            }
        internal_cell_update = false;
        first_visible_row = 0;
        first_visible_column = 0;
        load_visible_cells();
        update_scrollbar();
        update_mode_ui();
        layout(main_window);
        set_status(L"Saved workspace view loaded.");
        point::append_audit(
            app_root, "VIEW_LOADED", "mode=" + mode);
    } catch (const std::exception& ex) {
        internal_cell_update = false;
        MessageBoxW(
            main_window, widen(ex.what()).c_str(),
            L"Load View", MB_ICONWARNING);
    }
}

void run_risk_watchlist() {
    if (!engine) return;
    try {
        auto findings = engine->deep_insights(
            {}, selected_source_name());
        findings.rows.erase(
            std::remove_if(
                findings.rows.begin(), findings.rows.end(),
                [](const auto& row) {
                    return row.empty() ||
                        (row[0] != "High" && row[0] != "Medium");
                }),
            findings.rows.end());
        if (findings.rows.empty())
            findings.rows.push_back({
                "Info", selected_source_name().empty()
                    ? "All reports" : selected_source_name(),
                "Watchlist clear",
                "No High or Medium built-in rule was triggered",
                "0", "No immediate action is recommended."});
        narrow_mode = count_mode = compare_mode = analyze_mode = false;
        chart_mode = change_mode = false;
        insight_mode = true;
        internal_cell_update = true;
        clear_cell_store();
        for (HWND header : header_cells) SetWindowTextW(header, L"");
        for (std::size_t column = 0;
             column < findings.headers.size() &&
             column < header_cells.size(); ++column)
            set_control_text(
                header_cells[column], findings.headers[column]);
        int target_row = 0;
        for (const auto& row : findings.rows) {
            for (std::size_t column = 0;
                 column < row.size() &&
                 column < static_cast<std::size_t>(GRID_COLUMNS);
                 ++column)
                store_cell(
                    target_row, static_cast<int>(column), row[column]);
            ++target_row;
        }
        internal_cell_update = false;
        generated_insight_headers = true;
        insight_fields_cache.clear();
        first_visible_row = 0;
        load_visible_cells();
        update_scrollbar();
        update_mode_ui();
        last_result = std::move(findings);
        set_status(
            std::to_wstring(target_row) +
            L" High/Medium watchlist finding(s).");
        point::append_audit(
            app_root, "WATCHLIST",
            "results=" + std::to_string(target_row));
    } catch (const std::exception& ex) {
        internal_cell_update = false;
        MessageBoxW(
            main_window, widen(ex.what()).c_str(),
            L"Risk Watchlist", MB_ICONWARNING);
    }
}

void copy_all_grid() {
    commit_visible_cells();
    const auto headers = selected_headers();
    if (headers.empty()) {
        MessageBoxW(main_window, L"There are no headings to copy.",
                    L"Copy All", MB_ICONINFORMATION);
        return;
    }

    int last_used_row = highest_stored_row();
    const int first_result_row =
        first_result_row_for_mode();

    std::wstring clipboard;
    const std::size_t maximum_characters =
        50ull * 1024ull * 1024ull / sizeof(wchar_t);
    auto append_row = [&](const std::vector<std::wstring>& values) {
        for (std::size_t column = 0; column < values.size(); ++column) {
            if (column) clipboard.push_back(L'\t');
            clipboard += clipboard_safe_cell(values[column]);
        }
        clipboard += L"\r\n";
        if (clipboard.size() > maximum_characters)
            throw std::runtime_error(
                "Copy exceeds the 50 MiB clipboard safety limit");
    };

    try {
        std::vector<std::wstring> header_values;
        for (const auto& header : headers)
            header_values.push_back(widen(header));
        append_row(header_values);

        for (int row = first_result_row;
             row <= last_used_row; ++row) {
            std::vector<std::wstring> values;
            for (std::size_t column = 0;
                 column < headers.size(); ++column) {
                values.push_back(widen(stored_cell(
                    row, static_cast<int>(column))));
            }
            append_row(values);
        }
        if (!write_clipboard_text(clipboard))
            throw std::runtime_error("Windows clipboard is unavailable");
        std::wstringstream message;
        const int copied_rows = last_used_row < first_result_row
            ? 0 : last_used_row - first_result_row + 1;
        message << L"Copied " << copied_rows
                << L" row(s) and " << headers.size()
                << L" column(s), including headings.";
        set_status(message.str());
    } catch (const std::exception& ex) {
        MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Copy All blocked", MB_ICONWARNING);
    }
}

std::wstring read_clipboard_text() {
    if (!OpenClipboard(main_window)) return {};
    HANDLE handle = GetClipboardData(CF_UNICODETEXT);
    if (!handle) {
        CloseClipboard();
        return {};
    }
    const wchar_t* text =
        static_cast<const wchar_t*>(GlobalLock(handle));
    if (!text) {
        CloseClipboard();
        return {};
    }
    const SIZE_T bytes = GlobalSize(handle);
    const std::size_t maximum =
        static_cast<std::size_t>(bytes / sizeof(wchar_t));
    std::size_t length = 0;
    while (length < maximum && text[length] != L'\0') ++length;
    std::wstring result(text, length);
    GlobalUnlock(handle);
    CloseClipboard();
    return result;
}

std::vector<std::vector<std::wstring>> parse_clipboard_grid(
        const std::wstring& text) {
    std::vector<std::vector<std::wstring>> rows(1);
    std::wstring cell;
    for (std::size_t i = 0; i < text.size(); ++i) {
        const wchar_t ch = text[i];
        if (ch == L'\t') {
            rows.back().push_back(std::move(cell));
            cell.clear();
        } else if (ch == L'\r' || ch == L'\n') {
            if (ch == L'\r' && i + 1 < text.size() &&
                text[i + 1] == L'\n') ++i;
            rows.back().push_back(std::move(cell));
            cell.clear();
            rows.emplace_back();
        } else {
            cell.push_back(ch);
        }
    }
    rows.back().push_back(std::move(cell));
    if (rows.size() > 1 && rows.back().size() == 1 &&
        rows.back().front().empty()) rows.pop_back();
    return rows;
}

void copy_active_cell(HWND focus) {
    DWORD start = 0, end = 0;
    SendMessageW(focus, EM_GETSEL,
                 reinterpret_cast<WPARAM>(&start),
                 reinterpret_cast<LPARAM>(&end));
    if (start != end) {
        SendMessageW(focus, WM_COPY, 0, 0);
        return;
    }
    if (write_clipboard_text(control_text(focus)))
        set_status(L"Cell copied.");
}

void paste_at_cell(HWND focus) {
    const auto clipboard = read_clipboard_text();
    if (clipboard.empty()) return;
    if (clipboard.size() > 50ull * 1024ull * 1024ull) {
        MessageBoxW(main_window, L"Clipboard data exceeds 50 MiB.",
                    L"Paste blocked", MB_ICONWARNING);
        return;
    }
    const auto pasted = parse_clipboard_grid(clipboard);
    const int id = GetDlgCtrlID(focus);
    const int start_column = cell_column_from_id(id);
    if (start_column < 0) return;

    if (!is_header_id(id)) {
        if (universal_mode_active() && universal_results_displayed &&
            start_column == 0) {
            // Preserve earlier results and mark only the newly pasted first-
            // column rows as the next Universal lookup batch.
            for (std::size_t row = 0; row < pasted.size() &&
                 first_visible_row +
                     (id - ID_CELL_BASE) / GRID_COLUMNS +
                     static_cast<int>(row) < LOGICAL_ROWS; ++row) {
                universal_pending_lookup_rows.insert(
                    first_visible_row +
                    (id - ID_CELL_BASE) / GRID_COLUMNS +
                    static_cast<int>(row));
            }
        } else {
            commit_visible_cells();
        }
    }
    internal_cell_update = true;
    std::size_t converted = 0;
    if (is_header_id(id)) {
        if (!pasted.empty()) {
            for (std::size_t column = 0;
                 column < pasted.front().size() &&
                 start_column + static_cast<int>(column) < GRID_COLUMNS;
                 ++column) {
                SetWindowTextW(
                    header_cells[static_cast<std::size_t>(
                        start_column + static_cast<int>(column))],
                    pasted.front()[column].c_str());
            }
        }
    } else {
        const int visible_row =
            (id - ID_CELL_BASE) / GRID_COLUMNS;
        const int start_row = first_visible_row + visible_row;
        std::vector<std::pair<int, int>> pasted_cells;
        for (std::size_t row = 0;
             row < pasted.size() &&
             start_row + static_cast<int>(row) < LOGICAL_ROWS;
             ++row) {
            for (std::size_t column = 0;
                 column < pasted[row].size() &&
                 start_column + static_cast<int>(column) < GRID_COLUMNS;
                 ++column) {
                store_cell(
                    start_row + static_cast<int>(row),
                    start_column + static_cast<int>(column),
                    narrow(pasted[row][column]));
                pasted_cells.push_back({
                    start_row + static_cast<int>(row),
                    start_column + static_cast<int>(column)});
            }
        }
        try {
            for (const auto& [row, column] : pasted_cells) {
                if (resolve_identity_name_cell(row, column, true)) ++converted;
            }
        } catch (const std::exception& ex) {
            MessageBoxW(
                main_window, widen(ex.what()).c_str(),
                L"Ambiguous full name", MB_ICONWARNING);
        }
        load_visible_cells();
        if (converted) {
            point::append_audit(
                app_root, "PASTE_NAME_IDENTITY_RESOLVED",
                std::to_string(converted) +
                " unique full-name input(s) converted");
        }
    }
    internal_cell_update = false;
    hide_suggestions();
    set_status(converted
        ? L"Clipboard pasted; unique full names converted to identity values."
        : L"Clipboard data pasted.");
}

bool handle_keyboard_shortcut(const MSG& message) {
    if (message.message != WM_KEYDOWN ||
        (GetKeyState(VK_CONTROL) & 0x8000) == 0)
        return false;
    if (message.wParam == 'F') {
        SetFocus(find_text);
        SendMessageW(find_text, EM_SETSEL, 0, -1);
        return true;
    }
    if (message.wParam == 'L') {
        clear_results();
        return true;
    }
    HWND focus = GetFocus();
    if (message.wParam == 'U') {
        try {
            undo_learned_transformation();
        } catch (const std::exception& ex) {
            MessageBoxW(main_window, widen(ex.what()).c_str(),
                L"Text Transformation", MB_ICONINFORMATION);
        }
        return true;
    }
    if (message.wParam == VK_OEM_COMMA) {
        if (is_grid_edit(focus)) transform_target_cell = focus;
        apply_quick_text_transform(
            LearnedTextTransformation::Operation::SkipLeftKeepRight);
        return true;
    }
    if (message.wParam == VK_OEM_PERIOD) {
        if (is_grid_edit(focus)) transform_target_cell = focus;
        apply_quick_text_transform(
            LearnedTextTransformation::Operation::SkipRightKeepLeft);
        return true;
    }
    if (!is_grid_edit(focus)) return false;
    if (message.wParam == 'Z') {
        SendMessageW(focus, EM_UNDO, 0, 0);
        return true;
    }
    if (message.wParam == 'C' &&
        (GetKeyState(VK_SHIFT) & 0x8000) != 0) {
        copy_all_grid();
        return true;
    }
    if (message.wParam == 'C') {
        if (selection_active) {
            copy_selected_range();
            return true;
        }
        copy_active_cell(focus);
        return true;
    }
    if (message.wParam == 'V') {
        paste_at_cell(focus);
        return true;
    }
    if (message.wParam == 'A') {
        select_used_grid(focus);
        return true;
    }
    if (message.wParam == 'X') {
        if (selection_active) {
            copy_selected_range();
            clear_selected_range();
        } else {
            SendMessageW(focus, WM_CUT, 0, 0);
        }
        return true;
    }
    return false;
}

void set_input_tab(bool active) {
    hide_input_columns_popup();
    schedule_tab_active = false;
    input_tab_active = active;
    CheckDlgButton(main_window, ID_TAB_WORKSPACE,
        active ? BST_UNCHECKED : BST_CHECKED);
    CheckDlgButton(main_window, ID_TAB_INPUT,
        active ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(main_window, ID_TAB_SCHEDULE, BST_UNCHECKED);
    const int visibility = active ? SW_HIDE : SW_SHOW;
    const int ids[] = {ID_HELP, ID_MODE, ID_REFRESH, ID_SEARCH, ID_COPY_ALL,
        ID_EXPORT, ID_FIND_TEXT, ID_FIND_NEXT, ID_TRANSFORM_PATTERN,
        ID_QUICK_KEEP_LEFT,
        ID_QUICK_TEXT_TRANSFORM, ID_QUICK_KEEP_RIGHT, ID_GRID_SCROLL,
        ID_GRID_HSCROLL};
    for (const int id : ids) ShowWindow(GetDlgItem(main_window, id), visibility);
    ShowWindow(refresh_progress,
        !active && refresh_running.load() ? SW_SHOW : SW_HIDE);
    if (active) {
        for (int column = 0; column < GRID_COLUMNS; ++column) {
            HWND header = header_cells[static_cast<std::size_t>(column)];
            if (!IsWindowVisible(header)) continue;
            ShowWindow(header, SW_HIDE);
            for (int row = 0; row < GRID_CONTROL_ROWS; ++row)
                ShowWindow(data_cells[static_cast<std::size_t>(row)]
                    [static_cast<std::size_t>(column)], SW_HIDE);
        }
    }
    ShowWindow(input_drop_zone, active ? SW_SHOW : SW_HIDE);
    ShowWindow(input_upload_button, active ? SW_SHOW : SW_HIDE);
    ShowWindow(input_remove_button, active ? SW_SHOW : SW_HIDE);
    ShowWindow(input_field_search, active ? SW_SHOW : SW_HIDE);
    ShowWindow(input_file_list, active ? SW_SHOW : SW_HIDE);
    ShowWindow(input_imported_label, active ? SW_SHOW : SW_HIDE);
    ShowWindow(input_archive_label, active ? SW_SHOW : SW_HIDE);
    ShowWindow(input_archive_select_button, active ? SW_SHOW : SW_HIDE);
    ShowWindow(input_archive_list, active ? SW_SHOW : SW_HIDE);
    ShowWindow(input_archive_toggle_button, active ? SW_SHOW : SW_HIDE);
    ShowWindow(schedule_help_text, SW_HIDE);
    ShowWindow(schedule_add_button, SW_HIDE);
    ShowWindow(schedule_interval_combo, SW_HIDE);
    ShowWindow(schedule_list, SW_HIDE);
    ShowWindow(schedule_remove_button, SW_HIDE);
    ShowWindow(schedule_run_now_button, SW_HIDE);
    ShowWindow(GetDlgItem(main_window, ID_SCHEDULE_FETCHER), SW_HIDE);
    if (active) refresh_input_file_list();
    layout(main_window);
}

void set_schedule_tab(bool active) {
    if (!active) {
        set_input_tab(false);
        return;
    }
    set_input_tab(false);
    schedule_tab_active = true;
    input_tab_active = false;
    CheckDlgButton(main_window, ID_TAB_WORKSPACE, BST_UNCHECKED);
    CheckDlgButton(main_window, ID_TAB_INPUT, BST_UNCHECKED);
    CheckDlgButton(main_window, ID_TAB_SCHEDULE, BST_CHECKED);
    const int workspace_ids[] = {
        ID_HELP, ID_MODE, ID_REFRESH, ID_SEARCH, ID_COPY_ALL, ID_EXPORT,
        ID_FIND_TEXT, ID_FIND_NEXT, ID_TRANSFORM_PATTERN,
        ID_QUICK_KEEP_LEFT,
        ID_QUICK_TEXT_TRANSFORM, ID_QUICK_KEEP_RIGHT,
        ID_GRID_SCROLL, ID_GRID_HSCROLL};
    for (const int id : workspace_ids)
        ShowWindow(GetDlgItem(main_window, id), SW_HIDE);
    for (int column = 0; column < GRID_COLUMNS; ++column) {
        HWND header = header_cells[static_cast<std::size_t>(column)];
        if (!IsWindowVisible(header)) continue;
        ShowWindow(header, SW_HIDE);
        for (int row = 0; row < GRID_CONTROL_ROWS; ++row)
            ShowWindow(data_cells[static_cast<std::size_t>(row)]
                [static_cast<std::size_t>(column)], SW_HIDE);
    }
    ShowWindow(schedule_help_text, SW_SHOW);
    ShowWindow(schedule_add_button, SW_SHOW);
    ShowWindow(schedule_interval_combo, SW_SHOW);
    ShowWindow(schedule_list, SW_SHOW);
    ShowWindow(schedule_remove_button, SW_SHOW);
    ShowWindow(schedule_run_now_button, SW_SHOW);
    ShowWindow(GetDlgItem(main_window, ID_SCHEDULE_FETCHER), SW_SHOW);
    refresh_schedule_list();
    layout(main_window);
}

void layout(HWND window) {
    if (!help_text || !find_text || !grid_scrollbar ||
        !grid_hscrollbar || !status_text || !refresh_progress ||
        header_cells.size() !=
            static_cast<std::size_t>(GRID_COLUMNS) ||
        data_cells.size() !=
            static_cast<std::size_t>(GRID_CONTROL_ROWS)) {
        return;
    }
    RECT area{};
    GetClientRect(window, &area);
    const int width = static_cast<int>(area.right - area.left);
    const int height = static_cast<int>(area.bottom - area.top);
    MoveWindow(GetDlgItem(window, ID_TAB_WORKSPACE), 12, 4, 120, 28, TRUE);
    MoveWindow(GetDlgItem(window, ID_TAB_INPUT), 136, 4, 120, 28, TRUE);
    MoveWindow(GetDlgItem(window, ID_TAB_SCHEDULE), 260, 4, 140, 28, TRUE);
    if (schedule_tab_active) {
        const bool compact_schedule = width < 1020;
        const int schedule_list_top = compact_schedule ? 214 : 170;
        MoveWindow(schedule_help_text, 24, 52, std::max(300, width - 48), 54, TRUE);
        MoveWindow(schedule_interval_combo, 24, 122, 190, 180, TRUE);
        MoveWindow(schedule_add_button, 226, 120, 210, 34, TRUE);
        MoveWindow(schedule_remove_button, 448, 120, 180, 34, TRUE);
        MoveWindow(schedule_run_now_button, 640, 120, 160, 34, TRUE);
        MoveWindow(GetDlgItem(window, ID_SCHEDULE_FETCHER),
            compact_schedule ? 24 : width - 200,
            compact_schedule ? 164 : 120, 176, 34, TRUE);
        MoveWindow(schedule_list, 24, schedule_list_top,
            std::max(300, width - 48),
            std::max(160, height - schedule_list_top - 60), TRUE);
        MoveWindow(status_text, 12, height - 36, width - 24, 24, TRUE);
        return;
    }
    if (input_tab_active) {
        MoveWindow(input_drop_zone, 24, 48, std::max(300, width - 48), 100, TRUE);
        MoveWindow(input_upload_button, 24, 160, 150, 32, TRUE);
        MoveWindow(input_remove_button, 186, 160, 170, 32, TRUE);
        MoveWindow(input_field_search, 368, 160,
            std::max(80, width - 392), 32, TRUE);
        const int content_bottom = height - 48;
        const int available = std::max(260, content_bottom - 226);
        const int imported_height = input_archive_collapsed
            ? std::max(120, content_bottom - 226 - 34)
            : std::max(100, available * 45 / 100);
        const int archive_label_top = 222 + imported_height + 10;
        MoveWindow(input_imported_label, 24, 202, 180, 22, TRUE);
        MoveWindow(input_file_list, 24, 226, std::max(300, width - 48),
            imported_height, TRUE);
        MoveWindow(input_archive_label, 24, archive_label_top, 150, 28, TRUE);
        MoveWindow(input_archive_select_button, 176, archive_label_top - 4,
            190, 30, TRUE);
        MoveWindow(input_archive_toggle_button, 374, archive_label_top - 4,
            34, 30, TRUE);
        ShowWindow(input_archive_select_button,
            input_archive_collapsed ? SW_HIDE : SW_SHOW);
        ShowWindow(input_archive_list,
            input_archive_collapsed ? SW_HIDE : SW_SHOW);
        MoveWindow(input_archive_list, 24, archive_label_top + 30,
            std::max(300, width - 48),
            std::max(90, content_bottom - archive_label_top - 30), TRUE);
        MoveWindow(status_text, 12, height - 36, width - 24, 24, TRUE);
        return;
    }
    const int usable_width = std::max(300, width - 44);
    first_visible_column = std::clamp(
        first_visible_column, 0, GRID_COLUMNS - 1);
    const int header_top = 106;
    const int available_grid_height =
        std::max(40, height - header_top - 60);
    visible_grid_rows = std::clamp(
        available_grid_height / std::max(20, grid_row_height) - 1,
        1, GRID_CONTROL_ROWS);
    const int maximum_first_row =
        std::max(0, LOGICAL_ROWS - visible_grid_rows);
    if (first_visible_row > maximum_first_row) {
        commit_visible_cells();
        first_visible_row = maximum_first_row;
        load_visible_cells();
    }
    const int row_height = std::clamp(
        available_grid_height / (visible_grid_rows + 1), 20, 80);

    std::vector<int> display_left(
        static_cast<std::size_t>(GRID_COLUMNS), -1);
    std::vector<int> display_width(
        static_cast<std::size_t>(GRID_COLUMNS), 0);
    int next_left = 12;
    const int grid_right = 12 + usable_width;
    visible_grid_columns = 0;
    for (int column = first_visible_column;
         column < GRID_COLUMNS && next_left < grid_right;
         ++column) {
        const int width_for_column = std::min(
            column_widths[static_cast<std::size_t>(column)],
            grid_right - next_left);
        display_left[static_cast<std::size_t>(column)] =
            next_left;
        display_width[static_cast<std::size_t>(column)] =
            width_for_column;
        next_left += width_for_column;
        ++visible_grid_columns;
    }

    MoveWindow(help_text, 12, 38, std::max(100, width - 544), 22, TRUE);
    MoveWindow(GetDlgItem(window, ID_MODE),
               width - 520, 32, 110, 28, TRUE);
    MoveWindow(GetDlgItem(window, ID_REFRESH),
               width - 400, 32, 90, 28, TRUE);
    MoveWindow(GetDlgItem(window, ID_SEARCH),
               width - 300, 32, 90, 28, TRUE);
    MoveWindow(GetDlgItem(window, ID_COPY_ALL),
               width - 200, 32, 90, 28, TRUE);
    MoveWindow(GetDlgItem(window, ID_EXPORT),
               width - 100, 32, 90, 28, TRUE);
    MoveWindow(find_text, 12, 70, 260, 26, TRUE);
    MoveWindow(GetDlgItem(window, ID_FIND_NEXT),
               280, 70, 90, 26, TRUE);
    MoveWindow(transform_pattern_text,
               378, 70, 100, 26, TRUE);
    MoveWindow(GetDlgItem(window, ID_QUICK_KEEP_LEFT),
               486, 70, 34, 26, TRUE);
    MoveWindow(GetDlgItem(window, ID_QUICK_TEXT_TRANSFORM),
               524, 70, 44, 26, TRUE);
    MoveWindow(GetDlgItem(window, ID_QUICK_KEEP_RIGHT),
               572, 70, 34, 26, TRUE);

    for (int column = 0; column < GRID_COLUMNS; ++column) {
        const bool visible =
            display_left[static_cast<std::size_t>(column)] >= 0;
        HWND header = header_cells[static_cast<std::size_t>(column)];
        if (!visible) {
            // Hidden columns stay untouched during ordinary resize events.
            // Only a column that was visible in the previous layout needs
            // its controls hidden. This avoids thousands of Win32 calls.
            if (IsWindowVisible(header)) {
                ShowWindow(header, SW_HIDE);
                for (int row = 0; row < GRID_CONTROL_ROWS; ++row) {
                    ShowWindow(
                        data_cells[static_cast<std::size_t>(row)]
                                  [static_cast<std::size_t>(column)],
                        SW_HIDE);
                }
            }
            continue;
        }
        const int left = display_left[static_cast<std::size_t>(column)];
        const int cell_width =
            display_width[static_cast<std::size_t>(column)];
        ShowWindow(header, SW_SHOW);
        MoveWindow(header, left, header_top,
                   cell_width, row_height, FALSE);
        for (int row = 0; row < GRID_CONTROL_ROWS; ++row) {
            const bool row_visible = row < visible_grid_rows;
            ShowWindow(
                data_cells[static_cast<std::size_t>(row)]
                          [static_cast<std::size_t>(column)],
                row_visible ? SW_SHOW : SW_HIDE);
            if (row_visible) {
                MoveWindow(
                    data_cells[static_cast<std::size_t>(row)]
                              [static_cast<std::size_t>(column)],
                    left, header_top + row_height * (row + 1),
                    cell_width, row_height, FALSE);
            }
        }
    }
    MoveWindow(grid_scrollbar, 12 + usable_width, header_top,
               20, row_height * (visible_grid_rows + 1), TRUE);
    MoveWindow(
        grid_hscrollbar, 12,
        header_top + row_height * (visible_grid_rows + 1),
        usable_width, 20, TRUE);
    update_scrollbar();
    update_column_scrollbar();
    MoveWindow(status_text, 12, height - 36, width - 24, 24, TRUE);
    MoveWindow(refresh_progress, 12, height - 58, width - 24, 16, TRUE);
    RECT grid_area{
        12, header_top, 12 + usable_width + 20,
        header_top + row_height * (visible_grid_rows + 1) + 20};
    RedrawWindow(
        window, &grid_area, nullptr,
        RDW_INVALIDATE | RDW_ALLCHILDREN);
}

LRESULT CALLBACK window_proc(HWND window, UINT message,
                             WPARAM wparam, LPARAM lparam) {
    if (inbox_changed_message != 0 && message == inbox_changed_message) {
        if (refresh_running.load()) {
            SetTimer(window, INBOX_NOTIFICATION_TIMER_ID, 1000, nullptr);
            set_status(L"A downloaded file is waiting; Imported Files will refresh automatically.");
        } else {
            refresh_engine(true);
        }
        return 0;
    }
    switch (message) {
    case WM_CREATE: {
        main_window = window;
        HMENU menu_bar = CreateMenu();
        HMENU workspace_menu = CreatePopupMenu();
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_IMPORT_SUMMARY, L"Import Summary");
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_NEXT_SOURCE, L"Next Source");
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_DYNAMIC_CHART, L"Open Dynamic Chart");
        AppendMenuW(workspace_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_SAVE_VIEW, L"Save Workspace View");
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_LOAD_VIEW, L"Load Workspace View");
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_RISK_WATCHLIST, L"Run Risk Watchlist");
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_SET_CHANGE_BASELINE, L"Set Change Baseline");
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_FIELD_SYNONYMS, L"Field Synonym Manager");
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_RELATIONSHIP_DIAGNOSTICS,
            L"Relationship and Conflict Diagnostics");
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_EXPLAIN_RESULT,
            L"Explain Selected Result and Lineage");
        AppendMenuW(workspace_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(
            workspace_menu, MF_STRING,
            ID_MENU_CLEAR_RESULTS, L"Clear Results\tCtrl+L");
        AppendMenuW(
            menu_bar, MF_POPUP,
            reinterpret_cast<UINT_PTR>(workspace_menu),
            L"Workspace");
        HMENU data_tools_menu = CreatePopupMenu();
        HMENU filter_menu = CreatePopupMenu();
        AppendMenuW(filter_menu, MF_STRING,
                    ID_DATA_FILTER_EQUALS, L"Equals Selected Value");
        AppendMenuW(filter_menu, MF_STRING,
                    ID_DATA_FILTER_CONTAINS, L"Contains Selected Value");
        AppendMenuW(filter_menu, MF_STRING,
                    ID_DATA_FILTER_STARTS, L"Starts With Selected Value");
        AppendMenuW(filter_menu, MF_STRING,
                    ID_DATA_FILTER_ENDS, L"Ends With Selected Value");
        AppendMenuW(filter_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(filter_menu, MF_STRING,
                    ID_DATA_FILTER_BLANK, L"Blank Cells");
        AppendMenuW(filter_menu, MF_STRING,
                    ID_DATA_FILTER_NOT_BLANK, L"Non-Blank Cells");
        AppendMenuW(filter_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(filter_menu, MF_STRING,
                    ID_DATA_FILTER_CLEAR, L"Clear Filter");
        AppendMenuW(data_tools_menu, MF_POPUP,
                    reinterpret_cast<UINT_PTR>(filter_menu),
                    L"Filter Selected Column");
        HMENU split_menu = CreatePopupMenu();
        AppendMenuW(split_menu, MF_STRING,
                    ID_DATA_SPLIT_SLASH, L"Slash  /");
        AppendMenuW(split_menu, MF_STRING,
                    ID_DATA_SPLIT_BACKSLASH, L"Backslash  \\");
        AppendMenuW(split_menu, MF_STRING,
                    ID_DATA_SPLIT_PIPE, L"Pipe  |");
        AppendMenuW(split_menu, MF_STRING,
                    ID_DATA_SPLIT_COMMA, L"Comma  ,");
        AppendMenuW(split_menu, MF_STRING,
                    ID_DATA_SPLIT_SEMICOLON, L"Semicolon  ;");
        AppendMenuW(split_menu, MF_STRING,
                    ID_DATA_SPLIT_COLON, L"Colon  :");
        AppendMenuW(split_menu, MF_STRING,
                    ID_DATA_SPLIT_SPACE, L"Space");
        AppendMenuW(split_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(split_menu, MF_STRING,
                    ID_DATA_SPLIT_CUSTOM,
                    L"Custom Delimiter from Find Box...");
        AppendMenuW(data_tools_menu, MF_POPUP,
                    reinterpret_cast<UINT_PTR>(split_menu),
                    L"Split Selected Column");
        AppendMenuW(menu_bar, MF_POPUP,
                    reinterpret_cast<UINT_PTR>(data_tools_menu),
                    L"Data Tools");
        SetMenu(window, menu_bar);
        CreateWindowW(L"BUTTON", L"Workspace",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_GROUP |
                BS_AUTORADIOBUTTON | BS_PUSHLIKE,
            12, 4, 120, 28, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_TAB_WORKSPACE)),
            nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Input Files",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP |
                BS_AUTORADIOBUTTON | BS_PUSHLIKE,
            136, 4, 120, 28, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_TAB_INPUT)),
            nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Auto Schedule",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP |
                BS_AUTORADIOBUTTON | BS_PUSHLIKE,
            260, 4, 140, 28, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_TAB_SCHEDULE)),
            nullptr, nullptr);
        CheckDlgButton(window, ID_TAB_WORKSPACE, BST_CHECKED);
        help_text = CreateWindowW(
            L"STATIC",
            L"Universal: heading=desired output; paste any known value below.",
            WS_CHILD | WS_VISIBLE, 12, 24, 650, 22,
            window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_HELP)),
            nullptr, nullptr);
        CreateWindowW(
            L"BUTTON", L"Mode: Universal", WS_CHILD | WS_VISIBLE,
            380, 18, 110, 28, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_MODE)),
            nullptr, nullptr);
        CreateWindowW(
            L"BUTTON", L"Refresh", WS_CHILD | WS_VISIBLE,
            500, 18, 90, 28, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_REFRESH)),
            nullptr, nullptr);
        CreateWindowW(
            L"BUTTON", L"Search", WS_CHILD | WS_VISIBLE,
            600, 18, 90, 28, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_SEARCH)),
            nullptr, nullptr);
        CreateWindowW(
            L"BUTTON", L"Export", WS_CHILD | WS_VISIBLE,
            700, 18, 90, 28, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_EXPORT)),
            nullptr, nullptr);
        CreateWindowW(
            L"BUTTON", L"Copy All", WS_CHILD | WS_VISIBLE,
            700, 18, 90, 28, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_COPY_ALL)),
            nullptr, nullptr);
        find_text = CreateWindowW(
            L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            12, 56, 260, 26, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_FIND_TEXT)),
            nullptr, nullptr);
        SendMessageW(
            find_text, EM_SETCUEBANNER, TRUE,
            reinterpret_cast<LPARAM>(L"Find in results"));
        CreateWindowW(
            L"BUTTON", L"Find Next", WS_CHILD | WS_VISIBLE,
            280, 56, 90, 26, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_FIND_NEXT)),
            nullptr, nullptr);

        transform_pattern_text = CreateWindowW(
            L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            378, 56, 100, 26, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_TRANSFORM_PATTERN)),
            nullptr, nullptr);
        SendMessageW(
            transform_pattern_text, EM_SETCUEBANNER, TRUE,
            reinterpret_cast<LPARAM>(L"text / delimiter"));

        quick_keep_left_button = CreateWindowW(
            L"BUTTON", L"<<",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
            486, 56, 34, 26, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_QUICK_KEEP_LEFT)),
            nullptr, nullptr);
        quick_text_transform_button = CreateWindowW(
            L"BUTTON", L"✂",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
            524, 56, 44, 26, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_QUICK_TEXT_TRANSFORM)),
            nullptr, nullptr);
        quick_keep_right_button = CreateWindowW(
            L"BUTTON", L">>",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
            572, 56, 34, 26, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_QUICK_KEEP_RIGHT)),
            nullptr, nullptr);
        SendMessageW(
            quick_text_transform_button, WM_SETFONT,
            reinterpret_cast<WPARAM>(GetStockObject(DEFAULT_GUI_FONT)), TRUE);

        for (int column = 0; column < GRID_COLUMNS; ++column) {
            HWND header = CreateWindowW(
                L"EDIT", nullptr,
                WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
                12 + column * 125, 62, 125, 28, window,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(
                    ID_HEADER_BASE + column)), nullptr, nullptr);
            SetWindowSubclass(header, grid_edit_subclass, 1, 0);
            header_cells.push_back(header);
        }

        for (int row = 0; row < GRID_CONTROL_ROWS; ++row) {
            std::vector<HWND> row_cells;
            for (int column = 0; column < GRID_COLUMNS; ++column) {
                const int id =
                    ID_CELL_BASE + row * GRID_COLUMNS + column;
                HWND cell = CreateWindowW(
                    L"EDIT", nullptr,
                    WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
                    12 + column * 125, 90 + row * 28, 125, 28,
                    window,
                    reinterpret_cast<HMENU>(
                        static_cast<INT_PTR>(id)),
                    nullptr, nullptr);
                SetWindowSubclass(cell, grid_edit_subclass, 1, 0);
                row_cells.push_back(cell);
            }
            data_cells.push_back(std::move(row_cells));
        }

        grid_scrollbar = CreateWindowW(
            L"SCROLLBAR", nullptr,
            WS_CHILD | WS_VISIBLE | SBS_VERT,
            770, 62, 20, 448, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_GRID_SCROLL)),
            nullptr, nullptr);
        update_scrollbar();

        grid_hscrollbar = CreateWindowW(
            L"SCROLLBAR", nullptr,
            WS_CHILD | WS_VISIBLE | SBS_HORZ,
            12, 540, 758, 20, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_GRID_HSCROLL)),
            nullptr, nullptr);
        update_column_scrollbar();

        suggestion_list = CreateWindowExW(
            WS_EX_CLIENTEDGE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
            L"LISTBOX", nullptr,
            WS_POPUP | WS_BORDER | WS_VSCROLL |
            LBS_NOTIFY | LBS_NOINTEGRALHEIGHT,
            0, 0, 250, 160, window,
            nullptr, GetModuleHandleW(nullptr), nullptr);
        if (!suggestion_list) {
            MessageBoxW(
                window,
                L"Point could not create the suggestion panel.",
                L"Point startup error", MB_ICONERROR);
            return -1;
        }
        SetWindowSubclass(
            suggestion_list, suggestion_list_subclass, 1, 0);
        status_text = CreateWindowW(
            L"STATIC", L"Starting...", WS_CHILD | WS_VISIBLE,
            12, 540, 760, 24, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_STATUS)),
            nullptr, nullptr);
        input_drop_zone = CreateWindowW(L"STATIC",
            L"Drag and drop CSV or Excel files here\r\n"
            L"Supported: .csv, .xlsx, .xls, .xlsm",
            WS_CHILD | WS_BORDER | SS_CENTER | SS_CENTERIMAGE,
            24, 52, 600, 150, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_INPUT_DROP_ZONE)),
            nullptr, nullptr);
        input_upload_button = CreateWindowW(L"BUTTON", L"Upload Files...",
            WS_CHILD | WS_TABSTOP | BS_DEFPUSHBUTTON,
            24, 216, 150, 32, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_INPUT_UPLOAD)),
            nullptr, nullptr);
        input_remove_button = CreateWindowW(
            L"BUTTON", L"Remove Selected",
            WS_CHILD | WS_TABSTOP,
            186, 216, 170, 32, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_INPUT_REMOVE)),
            nullptr, nullptr);
        input_field_search = CreateWindowExW(
            WS_EX_CLIENTEDGE, L"EDIT", nullptr,
            WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL,
            368, 216, 300, 32, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_INPUT_FIELD_SEARCH)),
            nullptr, nullptr);
        SendMessageW(input_field_search, EM_SETCUEBANNER, TRUE,
            reinterpret_cast<LPARAM>(L"Search files or fields..."));
        input_file_list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", nullptr,
            WS_CHILD | WS_VSCROLL | LBS_NOINTEGRALHEIGHT | LBS_EXTENDEDSEL |
                LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
            24, 264, 600, 300, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_INPUT_FILE_LIST)),
            nullptr, nullptr);
        SendMessageW(input_file_list, LB_SETITEMHEIGHT, 0, 48);
        SetWindowSubclass(
            input_file_list, input_file_list_subclass, 1, 0);
        input_imported_label = CreateWindowW(
            L"STATIC", L"Imported Files", WS_CHILD,
            24, 252, 180, 22, window, nullptr, nullptr, nullptr);
        input_archive_label = CreateWindowW(
            L"STATIC", L"Available Files", WS_CHILD,
            24, 500, 180, 22, window, nullptr, nullptr, nullptr);
        input_archive_select_button = CreateWindowW(
            L"BUTTON", L"Select Folder / Drive...",
            WS_CHILD | WS_TABSTOP,
            190, 494, 190, 30, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_INPUT_ARCHIVE_SELECT)),
            nullptr, nullptr);
        input_archive_list = CreateWindowExW(
            WS_EX_CLIENTEDGE, L"LISTBOX", nullptr,
            WS_CHILD | WS_VSCROLL | LBS_NOINTEGRALHEIGHT |
                LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
            24, 530, 600, 180, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_INPUT_ARCHIVE_LIST)),
            nullptr, nullptr);
        SendMessageW(input_archive_list, LB_SETITEMHEIGHT, 0, 36);
        SetWindowSubclass(
            input_archive_list, input_archive_list_subclass, 1, 0);
        input_archive_toggle_button = CreateWindowW(
            L"BUTTON", L"▲", WS_CHILD | WS_TABSTOP,
            374, 494, 34, 30, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_INPUT_ARCHIVE_TOGGLE)),
            nullptr, nullptr);
        schedule_help_text = CreateWindowW(
            L"STATIC",
            L"Select one or more source files and an interval. While Point is "
            L"open, due files are copied into Inbox and refreshed as one batch.",
            WS_CHILD, 24, 52, 900, 54, window, nullptr, nullptr, nullptr);
        schedule_interval_combo = CreateWindowW(
            L"COMBOBOX", nullptr,
            WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
            24, 122, 190, 180, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_SCHEDULE_INTERVAL)),
            nullptr, nullptr);
        const int intervals[] = {1, 2, 3, 4, 6, 8, 12, 24};
        for (const int hours : intervals) {
            const auto label = L"Every " + std::to_wstring(hours) + L" hour(s)";
            const LRESULT index = SendMessageW(schedule_interval_combo,
                CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str()));
            SendMessageW(schedule_interval_combo, CB_SETITEMDATA, index, hours);
        }
        SendMessageW(schedule_interval_combo, CB_SETCURSEL, 0, 0);
        schedule_add_button = CreateWindowW(
            L"BUTTON", L"Add Scheduled File(s)...", WS_CHILD | WS_TABSTOP,
            226, 120, 210, 34, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SCHEDULE_ADD)),
            nullptr, nullptr);
        schedule_remove_button = CreateWindowW(
            L"BUTTON", L"Remove Selected", WS_CHILD | WS_TABSTOP,
            448, 120, 180, 34, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SCHEDULE_REMOVE)),
            nullptr, nullptr);
        schedule_run_now_button = CreateWindowW(
            L"BUTTON", L"Run Selected Now", WS_CHILD | WS_TABSTOP,
            640, 120, 160, 34, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SCHEDULE_RUN_NOW)),
            nullptr, nullptr);
        CreateWindowW(
            L"BUTTON", L"Open Point Fetcher", WS_CHILD | WS_TABSTOP,
            812, 120, 176, 34, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_SCHEDULE_FETCHER)),
            nullptr, nullptr);
        schedule_list = CreateWindowExW(
            WS_EX_CLIENTEDGE, L"LISTBOX", nullptr,
            WS_CHILD | WS_VSCROLL | WS_HSCROLL | LBS_NOINTEGRALHEIGHT |
                LBS_EXTENDEDSEL,
            24, 170, 900, 300, window,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SCHEDULE_LIST)),
            nullptr, nullptr);
        SendMessageW(schedule_list, LB_SETHORIZONTALEXTENT, 1800, 0);
        refresh_schedule_list();
        SetTimer(window, AUTO_SCHEDULE_TIMER_ID, 30'000, nullptr);
        DragAcceptFiles(window, TRUE);
        refresh_progress = CreateWindowExW(
            0, PROGRESS_CLASSW, nullptr,
            WS_CHILD | PBS_SMOOTH,
            12, 518, 760, 16, window,
            reinterpret_cast<HMENU>(
                static_cast<INT_PTR>(ID_REFRESH_PROGRESS)),
            GetModuleHandleW(nullptr), nullptr);
        if (!refresh_progress) {
            MessageBoxW(window,
                        L"Point could not create the refresh progress bar.",
                        L"Point startup error", MB_ICONERROR);
            return -1;
        }
        SendMessageW(refresh_progress, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
        selected_cell_brush =
            CreateSolidBrush(RGB(196, 222, 255));
        if (!selected_cell_brush) {
            MessageBoxW(window,
                        L"Point could not create the selection highlight.",
                        L"Point startup error", MB_ICONERROR);
            return -1;
        }
        duplicate_light_brush =
            CreateSolidBrush(RGB(218, 234, 255));
        duplicate_dark_brush =
            CreateSolidBrush(RGB(152, 195, 242));
        if (!duplicate_light_brush || !duplicate_dark_brush) {
            MessageBoxW(window,
                        L"Point could not create duplicate row colors.",
                        L"Point startup error", MB_ICONERROR);
            return -1;
        }
        criteria_cell_brush =
            CreateSolidBrush(RGB(255, 248, 214));
        if (!criteria_cell_brush) {
            MessageBoxW(window,
                        L"Point could not create the criteria highlight.",
                        L"Point startup error", MB_ICONERROR);
            return -1;
        }
        changed_row_brush =
            CreateSolidBrush(RGB(255, 224, 224));
        if (!changed_row_brush) {
            MessageBoxW(window,
                        L"Point could not create the change highlight.",
                        L"Point startup error", MB_ICONERROR);
            return -1;
        }
        synonym_missing_brush =
            CreateSolidBrush(RGB(255, 205, 205));
        synonym_duplicate_brush =
            CreateSolidBrush(RGB(205, 225, 255));
        if (!synonym_missing_brush || !synonym_duplicate_brush) {
            MessageBoxW(window,
                        L"Point could not create synonym validation colors.",
                        L"Point startup error", MB_ICONERROR);
            return -1;
        }
        update_mode_ui();
        layout(window);
        PostMessageW(window, WM_COMMAND, ID_REFRESH, 0);
        return 0;
    }
    case WM_SIZE: {
        const bool grid_ready =
            header_cells.size() == static_cast<std::size_t>(GRID_COLUMNS) &&
            data_cells.size() ==
                static_cast<std::size_t>(GRID_CONTROL_ROWS);
        if (grid_ready && !input_tab_active && !schedule_tab_active)
            commit_visible_cells();
        layout(window);
        if (grid_ready && !input_tab_active && !schedule_tab_active)
            load_visible_cells();
        return 0;
    }
    case WM_VSCROLL:
        if (reinterpret_cast<HWND>(lparam) == grid_scrollbar) {
            SCROLLINFO info{sizeof(info)};
            info.fMask = SIF_ALL;
            GetScrollInfo(grid_scrollbar, SB_CTL, &info);
            int target = first_visible_row;
            switch (LOWORD(wparam)) {
            case SB_LINEUP: target -= 1; break;
            case SB_LINEDOWN: target += 1; break;
            case SB_PAGEUP: target -= visible_grid_rows; break;
            case SB_PAGEDOWN: target += visible_grid_rows; break;
            case SB_THUMBTRACK:
            case SB_THUMBPOSITION: target = info.nTrackPos; break;
            case SB_TOP: target = 0; break;
            case SB_BOTTOM:
                target = LOGICAL_ROWS - visible_grid_rows;
                break;
            default: return 0;
            }
            scroll_to(target);
            return 0;
        }
        break;
    case WM_HSCROLL:
        if (reinterpret_cast<HWND>(lparam) == grid_hscrollbar) {
            SCROLLINFO info{sizeof(info)};
            info.fMask = SIF_ALL;
            GetScrollInfo(grid_hscrollbar, SB_CTL, &info);
            int target = first_visible_column;
            switch (LOWORD(wparam)) {
            case SB_LINELEFT: target -= 1; break;
            case SB_LINERIGHT: target += 1; break;
            case SB_PAGELEFT:
                target -= std::max(1, visible_grid_columns - 1);
                break;
            case SB_PAGERIGHT:
                target += std::max(1, visible_grid_columns - 1);
                break;
            case SB_THUMBTRACK:
            case SB_THUMBPOSITION: target = info.nTrackPos; break;
            case SB_LEFT: target = 0; break;
            case SB_RIGHT:
                target = GRID_COLUMNS - visible_grid_columns;
                break;
            default: return 0;
            }
            commit_visible_cells();
            first_visible_column = std::clamp(
                target, 0,
                std::max(0, GRID_COLUMNS - visible_grid_columns));
            hide_suggestions();
            layout(window);
            load_visible_cells();
            return 0;
        }
        break;
    case WM_MOUSEWHEEL:
        {
        const int delta =
            static_cast<int>(GET_WHEEL_DELTA_WPARAM(wparam));
        if ((GET_KEYSTATE_WPARAM(wparam) & MK_SHIFT) != 0) {
            horizontal_wheel_remainder += delta;
            const int steps =
                horizontal_wheel_remainder / WHEEL_DELTA;
            horizontal_wheel_remainder %= WHEEL_DELTA;
            if (steps != 0) {
                commit_visible_cells();
                first_visible_column = std::clamp(
                    first_visible_column - steps,
                    0, GRID_COLUMNS - 1);
                hide_suggestions();
                layout(window);
                load_visible_cells();
            }
        } else {
            vertical_wheel_remainder += delta;
            const int steps =
                vertical_wheel_remainder / WHEEL_DELTA;
            vertical_wheel_remainder %= WHEEL_DELTA;
            if (steps != 0)
                scroll_to(first_visible_row - steps * 3);
        }
        return 0;
        }
    case WM_MOUSEHWHEEL:
        {
        horizontal_wheel_remainder +=
            static_cast<int>(GET_WHEEL_DELTA_WPARAM(wparam));
        const int steps =
            horizontal_wheel_remainder / WHEEL_DELTA;
        horizontal_wheel_remainder %= WHEEL_DELTA;
        if (steps != 0) {
            commit_visible_cells();
            first_visible_column = std::clamp(
                first_visible_column + steps,
                0, GRID_COLUMNS - 1);
            hide_suggestions();
            layout(window);
            load_visible_cells();
        }
        return 0;
        }
    case WM_DROPFILES: {
        HDROP drop = reinterpret_cast<HDROP>(wparam);
        const UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
        std::vector<std::filesystem::path> files;
        for (UINT index = 0; index < count; ++index) {
            const UINT length = DragQueryFileW(drop, index, nullptr, 0);
            std::wstring path(static_cast<std::size_t>(length) + 1, L'\0');
            DragQueryFileW(drop, index, path.data(), length + 1);
            path.resize(length);
            files.emplace_back(std::move(path));
        }
        DragFinish(drop);
        set_input_tab(true);
        import_input_files(files);
        return 0;
    }
    case WM_CTLCOLOREDIT: {
        HWND control = reinterpret_cast<HWND>(lparam);
        GridPosition position{};
        if (grid_position_from_control(control, position) &&
            position.row >= 0) {
            const int logical_row = position.row;
            if (universal_missing_rows.contains(logical_row)) {
                HDC context = reinterpret_cast<HDC>(wparam);
                SetTextColor(context, RGB(150, 0, 0));
                SetBkColor(context, RGB(255, 224, 224));
                return reinterpret_cast<LRESULT>(changed_row_brush);
            }
            if (universal_duplicate_rows.contains(logical_row)) {
                HDC context = reinterpret_cast<HDC>(wparam);
                const bool dark =
                    universal_duplicate_dark_rows.contains(logical_row);
                SetTextColor(
                    context, dark ? RGB(0, 30, 80) : RGB(0, 45, 105));
                SetBkColor(
                    context,
                    dark ? RGB(152, 195, 242) : RGB(218, 234, 255));
                return reinterpret_cast<LRESULT>(
                    dark ? duplicate_dark_brush : duplicate_light_brush);
            }
        }
        if (grid_position_from_control(control, position) &&
            position_in_selection(position)) {
            HDC context = reinterpret_cast<HDC>(wparam);
            SetTextColor(context, RGB(0, 0, 0));
            SetBkColor(context, RGB(196, 222, 255));
            return reinterpret_cast<LRESULT>(selected_cell_brush);
        }
        if (change_mode && generated_change_headers &&
            grid_position_from_control(control, position) &&
            position.row >= 0) {
            const int logical_row =
                first_visible_row + position.row;
            const auto change_type = point::normalize_name(
                stored_cell(logical_row, 0));
            if (change_type == "added" ||
                change_type == "removed" ||
                change_type == "modified" ||
                change_type == "ambiguouskey") {
                HDC context = reinterpret_cast<HDC>(wparam);
                SetTextColor(context, RGB(150, 0, 0));
                SetBkColor(context, RGB(255, 224, 224));
                return reinterpret_cast<LRESULT>(
                    changed_row_brush);
            }
        }
        if ((narrow_mode || count_mode || compare_mode ||
             (analyze_mode && !generated_analysis_headers) ||
             (chart_mode && !generated_chart_headers)) &&
            grid_position_from_control(control, position) &&
            (position.row == 0 ||
             (compare_mode && position.row == 1))) {
            HDC context = reinterpret_cast<HDC>(wparam);
            SetTextColor(context, RGB(0, 0, 0));
            SetBkColor(context, RGB(255, 248, 214));
            return reinterpret_cast<LRESULT>(criteria_cell_brush);
        }
        break;
    }
    case WM_COMMAND: {
        const int id = LOWORD(wparam);
        const int notification = HIWORD(wparam);
        if (id == ID_TRANSFORM_PATTERN && notification == EN_CHANGE) {
            update_transform_lens();
            return 0;
        }
        if (id == ID_INPUT_FIELD_SEARCH && notification == EN_CHANGE) {
            refresh_input_file_list();
            return 0;
        }
        if (id == ID_INPUT_FILE_LIST && notification == LBN_DBLCLK) {
            show_input_file_columns();
            return 0;
        }
        if (notification == EN_CHANGE && !internal_cell_update &&
            id >= ID_CELL_BASE &&
            id < ID_CELL_BASE + GRID_CONTROL_ROWS * GRID_COLUMNS) {
            const int visible_index = id - ID_CELL_BASE;
            const int logical_row =
                first_visible_row + visible_index / GRID_COLUMNS;
            const int column = visible_index % GRID_COLUMNS;
            pending_identity_resolution_cells.insert(
                cell_key(logical_row, column));
        }
        if (notification == EN_CHANGE && !internal_cell_update &&
            universal_mode_active() && universal_results_displayed &&
            id >= ID_CELL_BASE &&
            id < ID_CELL_BASE + GRID_CONTROL_ROWS * GRID_COLUMNS) {
            const int visible_index = id - ID_CELL_BASE;
            const int visible_row = visible_index / GRID_COLUMNS;
            const int column = visible_index % GRID_COLUMNS;
            if (column == 0) {
                const int logical_row = first_visible_row + visible_row;
                universal_pending_lookup_rows.insert(logical_row);
                set_status(
                    L"New Universal lookup ready; Search will fill this row "
                    L"and preserve previous rows.");
            }
        }
        if (id >= ID_CELL_BASE &&
            id < ID_CELL_BASE + GRID_CONTROL_ROWS * GRID_COLUMNS &&
            notification == EN_KILLFOCUS && !internal_cell_update) {
            const int visible_row =
                (id - ID_CELL_BASE) / GRID_COLUMNS;
            const int column =
                (id - ID_CELL_BASE) % GRID_COLUMNS;
            commit_visible_cells();
            try {
                if (resolve_identity_name_cell(
                        first_visible_row + visible_row,
                        column, true)) {
                    load_visible_cells();
                    set_status(
                        L"Full name converted to its unique identity value.");
                    point::append_audit(
                        app_root, "NAME_IDENTITY_RESOLVED",
                        "1 unique full-name input converted on edit");
                }
                pending_identity_resolution_cells.erase(cell_key(
                    first_visible_row + visible_row, column));
            } catch (const std::exception& ex) {
                MessageBoxW(
                    main_window, widen(ex.what()).c_str(),
                    L"Ambiguous full name", MB_ICONWARNING);
            }
        }
        if ((is_header_id(id) ||
             (id >= ID_CELL_BASE &&
              id < ID_CELL_BASE + GRID_CONTROL_ROWS * GRID_COLUMNS)) &&
            notification == EN_CHANGE) {
            if (is_header_id(id) && chart_mode &&
                generated_chart_headers && !internal_cell_update) {
                begin_chart_heading_edit(
                    id - ID_HEADER_BASE,
                    narrow(control_text(reinterpret_cast<HWND>(lparam))));
            }
            schedule_suggestions(
                reinterpret_cast<HWND>(lparam), id);
            return 0;
        }
        switch (id) {
        case ID_TAB_WORKSPACE: set_input_tab(false); return 0;
        case ID_TAB_INPUT: set_input_tab(true); return 0;
        case ID_TAB_SCHEDULE: set_schedule_tab(true); return 0;
        case ID_INPUT_UPLOAD: choose_input_files(); return 0;
        case ID_INPUT_REMOVE: remove_selected_input_file(); return 0;
        case ID_INPUT_ARCHIVE_SELECT:
            choose_archive_folder_or_drive(); return 0;
        case ID_INPUT_ARCHIVE_TOGGLE:
            input_archive_collapsed = !input_archive_collapsed;
            SetWindowTextW(input_archive_toggle_button,
                input_archive_collapsed ? L"▼" : L"▲");
            layout(main_window);
            return 0;
        case ID_SCHEDULE_ADD: add_auto_import_schedules(); return 0;
        case ID_SCHEDULE_REMOVE:
            try {
                remove_selected_auto_schedules();
            } catch (const std::exception& ex) {
                MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Auto Schedule", MB_ICONERROR);
            }
            return 0;
        case ID_SCHEDULE_RUN_NOW:
            try {
                run_auto_import_schedules(true);
            } catch (const std::exception& ex) {
                MessageBoxW(main_window, widen(ex.what()).c_str(),
                    L"Auto Schedule", MB_ICONERROR);
            }
            return 0;
        case ID_SCHEDULE_FETCHER:
            ShellExecuteW(main_window, L"open",
                (app_root / "PointFetcher.exe").c_str(),
                nullptr, app_root.c_str(), SW_SHOWNORMAL);
            return 0;
        case ID_MODE: toggle_query_mode(); return 0;
        case ID_REFRESH: refresh_engine(); return 0;
        case ID_SEARCH: run_search(); return 0;
        case ID_COPY_ALL: copy_all_grid(); return 0;
        case ID_EXPORT: export_grid(); return 0;
        case ID_FIND_NEXT: find_next_value(); return 0;
        case ID_QUICK_KEEP_LEFT:
            apply_quick_text_transform(
                LearnedTextTransformation::Operation::SkipRightKeepLeft);
            return 0;
        case ID_QUICK_TEXT_TRANSFORM:
            apply_quick_text_transform(
                LearnedTextTransformation::Operation::RemoveHighlighted);
            return 0;
        case ID_QUICK_KEEP_RIGHT:
            apply_quick_text_transform(
                LearnedTextTransformation::Operation::SkipLeftKeepRight);
            return 0;
        case ID_MENU_IMPORT_SUMMARY:
            show_import_summary(); return 0;
        case ID_MENU_NEXT_SOURCE:
            select_next_source(); return 0;
        case ID_MENU_DYNAMIC_CHART:
            open_dynamic_chart(); return 0;
        case ID_MENU_SAVE_VIEW:
            save_workspace_view(); return 0;
        case ID_MENU_LOAD_VIEW:
            load_workspace_view(); return 0;
        case ID_MENU_RISK_WATCHLIST:
            run_risk_watchlist(); return 0;
        case ID_MENU_SET_CHANGE_BASELINE:
            set_change_baseline(); return 0;
        case ID_MENU_FIELD_SYNONYMS:
            open_field_synonyms(); return 0;
        case ID_MENU_RELATIONSHIP_DIAGNOSTICS:
            show_relationship_diagnostics(); return 0;
        case ID_MENU_EXPLAIN_RESULT:
            explain_selected_result(); return 0;
        case ID_MENU_CLEAR_RESULTS:
            clear_results(); return 0;
        case ID_DATA_FILTER_EQUALS:
            apply_data_filter(point::RowFilterOperator::Equals); return 0;
        case ID_DATA_FILTER_CONTAINS:
            apply_data_filter(point::RowFilterOperator::Contains); return 0;
        case ID_DATA_FILTER_STARTS:
            apply_data_filter(point::RowFilterOperator::StartsWith); return 0;
        case ID_DATA_FILTER_ENDS:
            apply_data_filter(point::RowFilterOperator::EndsWith); return 0;
        case ID_DATA_FILTER_BLANK:
            apply_data_filter(point::RowFilterOperator::IsBlank); return 0;
        case ID_DATA_FILTER_NOT_BLANK:
            apply_data_filter(point::RowFilterOperator::IsNotBlank); return 0;
        case ID_DATA_FILTER_CLEAR:
            clear_data_filter(); return 0;
        case ID_DATA_SPLIT_SLASH: split_selected_column("/"); return 0;
        case ID_DATA_SPLIT_BACKSLASH:
            split_selected_column("\\"); return 0;
        case ID_DATA_SPLIT_PIPE: split_selected_column("|"); return 0;
        case ID_DATA_SPLIT_COMMA: split_selected_column(","); return 0;
        case ID_DATA_SPLIT_SEMICOLON:
            split_selected_column(";"); return 0;
        case ID_DATA_SPLIT_COLON: split_selected_column(":"); return 0;
        case ID_DATA_SPLIT_SPACE: split_selected_column(" "); return 0;
        case ID_DATA_SPLIT_CUSTOM: {
            const auto delimiter = narrow(control_text(find_text));
            if (point::trim(delimiter).empty()) {
                MessageBoxW(
                    main_window,
                    L"Type the custom delimiter in the Find box first.",
                    L"Data Tools — Custom Delimiter", MB_ICONINFORMATION);
            } else {
                split_selected_column(delimiter);
            }
            return 0;
        }
        }
        break;
    }
    case WM_POINT_REFRESH_PROGRESS: {
        std::unique_ptr<RefreshProgressUpdate> update(
            reinterpret_cast<RefreshProgressUpdate*>(lparam));
        if (update && refresh_running.load())
            show_refresh_progress(update->text, update->percent);
        return 0;
    }
    case WM_POINT_REFRESH_COMPLETE: {
        std::unique_ptr<RefreshCompletion> completion(
            reinterpret_cast<RefreshCompletion*>(lparam));
        if (completion) complete_refresh(std::move(completion));
        return 0;
    }
    case WM_POINT_ARCHIVE_SCAN_COMPLETE: {
        std::unique_ptr<ArchiveScanCompletion> completion(
            reinterpret_cast<ArchiveScanCompletion*>(lparam));
        input_archive_scan_running.store(false);
        EnableWindow(input_archive_select_button, TRUE);
        if (!completion) return 0;
        input_archive_root = completion->root;
        input_archive_files = std::move(completion->files);
        SendMessageW(input_archive_list, LB_RESETCONTENT, 0, 0);
        if (!completion->error.empty()) {
            SendMessageW(input_archive_list, LB_ADDSTRING, 0,
                reinterpret_cast<LPARAM>(L"Scan failed"));
            set_status(L"Available-file scan failed.");
            MessageBoxW(main_window, completion->error.c_str(),
                L"Available Files", MB_ICONWARNING);
            return 0;
        }
        if (input_archive_files.empty()) {
            SendMessageW(input_archive_list, LB_ADDSTRING, 0,
                reinterpret_cast<LPARAM>(
                    L"No supported CSV or Excel files found"));
        } else {
            for (const auto& path : input_archive_files) {
                const auto filename = path.filename().wstring();
                SendMessageW(input_archive_list, LB_ADDSTRING, 0,
                    reinterpret_cast<LPARAM>(filename.c_str()));
            }
        }
        InvalidateRect(input_archive_list, nullptr, TRUE);
        std::wstringstream status;
        status << input_archive_files.size()
               << L" available CSV/Excel file(s) found in "
               << input_archive_root.wstring();
        if (completion->limited)
            status << L" (first 10,000 shown)";
        set_status(status.str());
        return 0;
    }
    case WM_DRAWITEM: {
        const auto* item = reinterpret_cast<const DRAWITEMSTRUCT*>(lparam);
        if (item && item->CtlID == ID_INPUT_FILE_LIST) {
            draw_input_file_item(*item);
            return TRUE;
        }
        if (item && item->CtlID == ID_INPUT_ARCHIVE_LIST) {
            draw_archive_file_item(*item);
            return TRUE;
        }
        break;
    }
    case WM_TIMER:
        if (wparam == INBOX_NOTIFICATION_TIMER_ID) {
            if (!refresh_running.load()) {
                KillTimer(window, INBOX_NOTIFICATION_TIMER_ID);
                refresh_engine(true);
            }
            return 0;
        }
        if (wparam == AUTO_SCHEDULE_TIMER_ID) {
            try {
                run_auto_import_schedules(false);
            } catch (const std::exception& ex) {
                set_status(L"Scheduled import failed: " + widen(ex.what()));
                point::append_audit(app_root, "AUTO_IMPORT_FAILED", ex.what());
            }
            return 0;
        }
        if (wparam == SUGGESTION_TIMER_ID) {
            KillTimer(window, SUGGESTION_TIMER_ID);
            HWND target = pending_suggestion_target;
            const int id = pending_suggestion_id;
            pending_suggestion_target = nullptr;
            pending_suggestion_id = 0;
            if (IsWindow(target))
                show_suggestions_for(target, id);
            return 0;
        }
        break;
    case WM_DESTROY:
        refresh_cancel_requested.store(true);
        input_archive_scan_cancel.store(true);
        if (refresh_thread.joinable()) refresh_thread.join();
        if (input_archive_scan_thread.joinable())
            input_archive_scan_thread.join();
        if (selected_cell_brush) DeleteObject(selected_cell_brush);
        if (duplicate_light_brush) DeleteObject(duplicate_light_brush);
        if (duplicate_dark_brush) DeleteObject(duplicate_dark_brush);
        if (criteria_cell_brush) DeleteObject(criteria_cell_brush);
        if (changed_row_brush) DeleteObject(changed_row_brush);
        if (synonym_missing_brush) DeleteObject(synonym_missing_brush);
        if (synonym_duplicate_brush) DeleteObject(synonym_duplicate_brush);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(window, message, wparam, lparam);
}

}  // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int show) {
    const HRESULT com_result = CoInitializeEx(
        nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    wchar_t executable[MAX_PATH]{};
    GetModuleFileNameW(nullptr, executable, MAX_PATH);
    app_root = std::filesystem::path(executable).parent_path();
    try {
        point::ensure_directories(app_root);
        compliance_policy =
            point::compliance::load_policy(app_root);
        point::compliance::authorize_current_user(
            compliance_policy);
        export_authorized =
            point::compliance::current_user_can_export(
                compliance_policy);
        point::compliance::harden_data_directories(app_root);
        point::compliance::enforce_retention(
            app_root, compliance_policy);
        load_field_synonyms();
        load_auto_import_schedules();
        point::append_audit(
            app_root, "STARTUP",
            "authorized=true; policy=point-security.conf");
    } catch (const std::exception& ex) {
        try {
            point::append_audit(
                app_root, "STARTUP_BLOCKED", ex.what());
        } catch (...) {
        }
        MessageBoxW(nullptr, widen(ex.what()).c_str(),
                    L"Point startup failed", MB_ICONERROR);
        return 1;
    }

    INITCOMMONCONTROLSEX controls{
        sizeof(controls), ICC_STANDARD_CLASSES | ICC_PROGRESS_CLASS};
    InitCommonControlsEx(&controls);
    inbox_changed_message = RegisterWindowMessageW(L"Point.InboxChanged.v1");

    const HICON point_icon = LoadIconW(
        instance, MAKEINTRESOURCEW(IDI_POINT));
    const HICON point_small_icon = reinterpret_cast<HICON>(LoadImageW(
        instance, MAKEINTRESOURCEW(IDI_POINT), IMAGE_ICON,
        GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON),
        LR_DEFAULTCOLOR));
    if (!point_icon || !point_small_icon) {
        MessageBoxW(nullptr,
            L"Point's embedded application icon could not be loaded.",
            L"Point startup failed", MB_OK | MB_ICONERROR);
        return 1;
    }

    WNDCLASSEXW window_class{sizeof(window_class)};
    window_class.lpfnWndProc = window_proc;
    window_class.hInstance = instance;
    window_class.lpszClassName = L"PointMainWindow";
    window_class.hCursor = LoadCursor(nullptr, IDC_ARROW);
    window_class.hIcon = point_icon;
    window_class.hIconSm = point_small_icon;
    window_class.hbrBackground =
        reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    if (!RegisterClassExW(&window_class)) return 1;

    WNDCLASSEXW chart_class{sizeof(chart_class)};
    chart_class.lpfnWndProc = chart_window_proc;
    chart_class.hInstance = instance;
    chart_class.lpszClassName = L"PointChartWindow";
    chart_class.hCursor = LoadCursor(nullptr, IDC_HAND);
    chart_class.hIcon = point_icon;
    chart_class.hIconSm = point_small_icon;
    chart_class.hbrBackground =
        reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    if (!RegisterClassExW(&chart_class)) return 1;

    WNDCLASSEXW synonym_class{sizeof(synonym_class)};
    synonym_class.lpfnWndProc = synonym_window_proc;
    synonym_class.hInstance = instance;
    synonym_class.lpszClassName = L"PointSynonymWindow";
    synonym_class.hCursor = LoadCursor(nullptr, IDC_ARROW);
    synonym_class.hIcon = point_icon;
    synonym_class.hIconSm = point_small_icon;
    synonym_class.hbrBackground =
        reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    if (!RegisterClassExW(&synonym_class)) return 1;

    HWND window = CreateWindowExW(
        0, window_class.lpszClassName,
        L"Point — Secure Local Data Workspace",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT,
        1100, 680, nullptr, nullptr, instance, nullptr);
    if (!window) return 1;
    SendMessageW(window, WM_SETICON, ICON_BIG,
        reinterpret_cast<LPARAM>(point_icon));
    SendMessageW(window, WM_SETICON, ICON_SMALL,
        reinterpret_cast<LPARAM>(point_small_icon));
    ShowWindow(window, show == SW_HIDE ? SW_SHOWNORMAL : show);
    SetForegroundWindow(window);
    UpdateWindow(window);

    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        if (handle_keyboard_shortcut(message)) continue;
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
    if (SUCCEEDED(com_result)) CoUninitialize();
    return static_cast<int>(message.wParam);
}
