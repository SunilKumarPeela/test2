#pragma once

#include <cstdint>
#include <filesystem>
#include <functional>
#include <optional>
#include <string>
#include <unordered_map>
#include <vector>

namespace point {

struct Limits {
    std::uintmax_t max_file_bytes = 500ull * 1024ull * 1024ull;
    std::size_t max_rows = 2'000'000;
    std::size_t max_columns = 2'000;
    std::size_t max_header_bytes = 256;
    std::size_t max_cell_bytes = 1024ull * 1024ull;
    std::size_t max_result_rows = 100'000;
};

struct DataSet {
    std::filesystem::path path;
    std::string name;
    std::vector<std::string> headers;
    std::vector<std::vector<std::string>> rows;
    std::unordered_map<std::string, std::size_t> normalized_header_index;
    // Direct column-to-canonical-field lookup. Exact indexes touch every row,
    // so avoiding a reverse hash-map scan per cell is a major speedup.
    std::vector<std::string> canonical_headers;
    std::uintmax_t source_size = 0;
    std::int64_t source_modified = 0;
};

struct Relationship {
    std::size_t left_dataset = 0;
    std::size_t right_dataset = 0;
    std::size_t left_column = 0;
    std::size_t right_column = 0;
    double confidence = 0.0;
    bool left_unique = false;
    bool right_unique = false;
};

struct QueryCondition {
    std::string field;
    std::string value;
};

struct QueryRequest {
    std::vector<QueryCondition> conditions;
    // Retained for compatibility with the original single-field API.
    std::string lookup_field;
    std::string lookup_value;
    std::vector<std::string> output_fields;
};

struct QueryResult {
    std::vector<std::string> headers;
    std::vector<std::vector<std::string>> rows;
    std::vector<std::string> sources;
    // Count-only metadata used by the UI to display the exact affected
    // objects without rerunning cross-report relationship queries.
    std::string related_identity_field;
    std::vector<std::vector<std::string>> related_identity_values;
    std::string explanation;
};

enum class RowFilterOperator {
    Equals,
    Contains,
    StartsWith,
    EndsWith,
    IsBlank,
    IsNotBlank
};

QueryResult filter_query_result(
    const QueryResult& source,
    std::size_t column,
    RowFilterOperator operation,
    const std::string& value = {});

QueryResult split_query_result_column(
    const QueryResult& source,
    std::size_t column,
    const std::string& delimiter,
    std::size_t maximum_parts = 8);

enum class TextMatchPosition {
    Prefix,
    Suffix,
    Anywhere
};

// Applies the exact transformation learned from a user's in-cell text
// selection. Matching is case-insensitive, while all unselected characters
// retain their original spelling. A null result means the rule did not match.
std::optional<std::string> remove_text_pattern(
    const std::string& value,
    const std::string& selected_text,
    TextMatchPosition position);

enum class TextSide {
    Left,
    Right
};

// Splits on the first exact delimiter and retains only the requested side.
// A null result means the row did not contain the delimiter.
std::optional<std::string> keep_text_side(
    const std::string& value,
    const std::string& delimiter,
    TextSide side);

struct ImportIssue {
    std::filesystem::path path;
    std::string message;
};

struct ExportPolicy {
    bool block_payment_card_numbers = true;
    bool mask_highly_sensitive_fields = true;
};

struct FieldSynonymGroup {
    std::string canonical_field;
    std::vector<std::string> synonyms;
};

enum class IdentityResolutionStatus {
    NotApplicable,
    NoMatch,
    Unique,
    Ambiguous
};

struct IdentityResolution {
    IdentityResolutionStatus status =
        IdentityResolutionStatus::NotApplicable;
    std::string value;
    std::size_t distinct_matches = 0;
};

std::string normalize_name(const std::string& value);
std::string trim(const std::string& value);
std::string escape_csv_for_spreadsheet(const std::string& value);
bool looks_like_payment_card_number(const std::string& value);
bool is_highly_sensitive_field(const std::string& header);
std::string mask_sensitive_value(const std::string& value);

class Engine {
public:
    explicit Engine(Limits limits = {});
    Engine(const Engine& other);

    void set_field_synonyms(
        const std::vector<FieldSynonymGroup>& groups);
    const std::vector<FieldSynonymGroup>& field_synonyms() const {
        return field_synonyms_;
    }

    void load_inbox(const std::filesystem::path& inbox);
    void load_files(const std::vector<std::filesystem::path>& files);
    void load_files_incremental(
        const std::vector<std::filesystem::path>& files,
        const Engine* previous,
        const std::function<void(
            std::size_t, std::size_t,
            const std::filesystem::path&, bool)>& progress = {},
        const std::function<bool()>& cancelled = {});
    QueryResult query(const QueryRequest& request) const;
    QueryResult count_groups(
        const std::vector<std::string>& group_fields,
        const std::string& source_name = {}) const;
    QueryResult count_exact(
        const std::vector<QueryCondition>& conditions,
        const std::string& source_name = {}) const;
    QueryResult count_mixed(
        const std::vector<std::string>& fields,
        const std::vector<std::string>& criteria,
        const std::string& source_name = {}) const;
    QueryResult common_values(
        const std::string& identity_field,
        const std::string& first_identity,
        const std::string& second_identity,
        const std::string& value_field,
        const std::string& source_name = {}) const;
    QueryResult compare_values(
        const std::string& identity_field,
        const std::string& first_identity,
        const std::string& second_identity,
        const std::string& value_field,
        const std::string& source_name = {}) const;
    QueryResult compare_profiles(
        const std::string& identity_field,
        const std::string& first_identity,
        const std::string& second_identity,
        const std::vector<std::string>& value_fields,
        const std::string& source_name = {}) const;
    QueryResult analyze_field(
        const std::string& field,
        const std::string& source_name = {}) const;
    QueryResult analyze_keys(
        const std::vector<std::string>& key_fields,
        const std::string& source_name = {},
        const std::vector<std::string>& exact_key_values = {}) const;
    QueryResult deep_insights(
        const std::vector<std::string>& fields = {},
        const std::string& source_name = {}) const;
    QueryResult distribution(
        const std::string& category_field,
        const std::vector<QueryCondition>& filters = {},
        const std::string& source_name = {}) const;
    QueryResult distribution(
        const std::vector<std::string>& category_fields,
        const std::vector<QueryCondition>& filters = {},
        const std::string& source_name = {}) const;
    QueryResult changes_since(
        const Engine& previous,
        const std::vector<std::string>& key_fields,
        const std::string& source_name = {}) const;
    QueryResult exact_rows(
        const std::vector<QueryCondition>& conditions,
        const std::vector<std::string>& output_fields,
        const std::string& source_name) const;
    void export_csv(const QueryResult& result,
                    const std::filesystem::path& destination,
                    ExportPolicy policy = {}) const;

    const std::vector<DataSet>& datasets() const { return datasets_; }
    const std::vector<Relationship>& relationships() const {
        return relationships_;
    }
    const std::vector<ImportIssue>& issues() const { return issues_; }
    std::vector<std::string> all_fields() const;
    std::vector<std::string> values_for_field(
        const std::string& field,
        const std::string& prefix,
        std::size_t limit = 50) const;
    IdentityResolution resolve_identity_from_name(
        const std::string& target_field,
        const std::string& entered_name) const;
    QueryResult universal_lookup(
        const std::vector<std::string>& output_fields,
        const std::string& entered_value,
        std::size_t max_anchor_rows = 500) const;

private:
    using ExactValueIndex =
        std::unordered_map<std::string, std::vector<std::size_t>>;

    Limits limits_;
    std::vector<DataSet> datasets_;
    std::vector<Relationship> relationships_;
    std::vector<ImportIssue> issues_;
    // Built only for columns actually queried, then reused. This avoids
    // scanning 100,000+ rows for every exact lookup and relationship hop.
    mutable std::unordered_map<std::uint64_t, ExactValueIndex>
        exact_value_indexes_;
    std::vector<FieldSynonymGroup> field_synonyms_;
    std::unordered_map<std::string, std::string> field_aliases_;
    std::unordered_map<std::string, std::string> canonical_display_names_;

    DataSet parse_csv(const std::filesystem::path& path) const;
    std::string canonical_field_name(const std::string& field) const;
    void rebuild_dataset_field_indexes();
    void discover_relationships();
    const ExactValueIndex& exact_value_index(
        std::size_t dataset_id, std::size_t column) const;
    std::string indexed_value_key(
        std::size_t dataset_id, std::size_t column,
        const std::string& value) const;
    bool indexed_unique_column(
        std::size_t dataset_id, std::size_t column) const;
    QueryResult count_related(
        const std::vector<std::string>& fields,
        const std::vector<std::string>& exact_values,
        const std::string& source_name) const;
};

void ensure_directories(const std::filesystem::path& root);
void append_audit(const std::filesystem::path& root,
                  const std::string& event,
                  const std::string& detail);

}  // namespace point
