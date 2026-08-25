#include "point_core.h"

#include <cassert>
#include <filesystem>
#include <fstream>
#include <iostream>

int main() {
    const auto root = std::filesystem::temp_directory_path() /
        "point-schema-mapping-test";
    std::filesystem::create_directories(root);
    const auto users = root / "Users.csv";
    const auto devices = root / "Devices.csv";

    point::QueryResult text_tools;
    text_tools.headers = {"Username", "Status"};
    text_tools.rows = {
        {"global/speela", "Active"},
        {"global/jsmith", "Disabled"},
        {"local/adoe", "Active"},
        {"", "Active"}
    };
    const auto active_filter = point::filter_query_result(
        text_tools, 1, point::RowFilterOperator::Equals, "active");
    assert(active_filter.rows.size() == 3);
    const auto global_filter = point::filter_query_result(
        text_tools, 0, point::RowFilterOperator::StartsWith, "GLOBAL/");
    assert(global_filter.rows.size() == 2);
    const auto blank_filter = point::filter_query_result(
        text_tools, 0, point::RowFilterOperator::IsBlank);
    assert(blank_filter.rows.size() == 1);
    const auto split_usernames = point::split_query_result_column(
        text_tools, 0, "/");
    assert(split_usernames.headers.size() == 4);
    assert(split_usernames.headers[2] == "Username Part 1");
    assert(split_usernames.headers[3] == "Username Part 2");
    assert(split_usernames.rows[0][0] == "global/speela");
    assert(split_usernames.rows[0][2] == "global");
    assert(split_usernames.rows[0][3] == "speela");
    const auto remove_domain = point::remove_text_pattern(
        "global/speela", "global/", point::TextMatchPosition::Prefix);
    assert(remove_domain && *remove_domain == "speela");
    const auto remove_domain_case = point::remove_text_pattern(
        "GLOBAL/jsmith", "global/", point::TextMatchPosition::Prefix);
    assert(remove_domain_case && *remove_domain_case == "jsmith");
    const auto wrong_position = point::remove_text_pattern(
        "local/global/user", "global/", point::TextMatchPosition::Prefix);
    assert(!wrong_position);
    const auto remove_suffix = point::remove_text_pattern(
        "sunil@company.com", "@company.com", point::TextMatchPosition::Suffix);
    assert(remove_suffix && *remove_suffix == "sunil");
    const auto skip_left = point::keep_text_side(
        "different-domain/sunil", "/", point::TextSide::Right);
    assert(skip_left && *skip_left == "sunil");
    const auto skip_right = point::keep_text_side(
        "Engineering, Grace", ",", point::TextSide::Left);
    assert(skip_right && *skip_right == "Engineering");
    assert(!point::keep_text_side(
        "no delimiter", "/", point::TextSide::Right));
    {
        std::ofstream out(users);
        out << "Employee ID,Username,Email,Display Name,Manager Employee ID\n"
               "E1001,speela,sunil@example.com,\"Peela, Sunil Kumar\",\n"
               "E1002,jsmith,john@example.com,Smith John,E1001\n";
    }
    {
        std::ofstream out(devices);
        out << "Associate Number,Computer Name\n"
               "E1001,PC-001\nE1002,PC-002\n";
    }

    point::Engine unmapped;
    unmapped.load_files({users, devices});
    assert(unmapped.relationships().empty());

    point::Engine mapped;
    mapped.set_field_synonyms({
        {"Employee ID", {"EID", "Associate Number", "Worker ID"}}
    });
    mapped.load_files({users, devices});
    assert(mapped.relationships().size() == 1);
    point::QueryRequest request;
    request.lookup_field = "Employee ID";
    request.lookup_value = "E1001";
    request.output_fields = {"Display Name", "Computer Name"};
    const auto result = mapped.query(request);
    assert(result.rows.size() == 1);
    assert(result.rows[0][0] == "Peela, Sunil Kumar");
    assert(result.rows[0][1] == "PC-001");

    point::Engine incremental;
    incremental.set_field_synonyms({
        {"Employee ID", {"EID", "Associate Number", "Worker ID"}}
    });
    std::size_t reused_files = 0;
    incremental.load_files_incremental(
        {users, devices}, &mapped,
        [&](std::size_t, std::size_t,
            const std::filesystem::path&, bool reused) {
            if (reused) ++reused_files;
        });
    assert(reused_files == 2);
    assert(incremental.relationships().size() == 1);
    const auto incremental_result = incremental.query(request);
    assert(incremental_result.rows == result.rows);

    const auto username = mapped.resolve_identity_from_name(
        "Username", "Sunil Kumar Peela");
    assert(username.status == point::IdentityResolutionStatus::Unique);
    assert(username.value == "speela");
    const auto employee = mapped.resolve_identity_from_name(
        "Employee ID", "Peela, Sunil Kumar");
    assert(employee.status == point::IdentityResolutionStatus::Unique);
    assert(employee.value == "E1001");
    const auto email = mapped.resolve_identity_from_name(
        "Email", "Sunil Kumar Peela");
    assert(email.status == point::IdentityResolutionStatus::Unique);
    assert(email.value == "sunil@example.com");

    const auto computer_to_username = mapped.universal_lookup(
        {"Username"}, "PC-001");
    assert(computer_to_username.rows.size() == 1);
    assert(computer_to_username.rows[0][0] == "speela");
    const auto email_to_computer = mapped.universal_lookup(
        {"Computer Name"}, "sunil@example.com");
    assert(email_to_computer.rows.size() == 1);
    assert(email_to_computer.rows[0][0] == "PC-001");
    const auto email_to_employee = mapped.universal_lookup(
        {"Employee ID"}, "john@example.com");
    assert(email_to_employee.rows.size() == 1);
    assert(email_to_employee.rows[0][0] == "E1002");
    const auto computer_to_name = mapped.universal_lookup(
        {"Display Name"}, "PC-001");
    assert(computer_to_name.rows.size() == 1);
    assert(computer_to_name.rows[0][0] == "Peela, Sunil Kumar");
    const auto reordered_name = mapped.universal_lookup(
        {"Display Name", "Username", "Computer Name"},
        "Sunil Kumar Peela");
    assert(reordered_name.rows.size() == 1);
    assert(reordered_name.rows[0][0] == "Peela, Sunil Kumar");
    assert(reordered_name.rows[0][1] == "speela");
    assert(reordered_name.rows[0][2] == "PC-001");

    // The exact Username seed must remain the identity authority even when a
    // related group report contains a different Employee ID. Relationship
    // traversal may fill Department, but it must never replace the seed EID.
    const auto seed_identity = root / "SeedIdentity.csv";
    const auto related_group = root / "RelatedGroup.csv";
    {
        std::ofstream out(seed_identity);
        out << "Username,Employee ID,Group Name\n"
               "speela,E1001,Security Team\n";
    }
    {
        std::ofstream out(related_group);
        out << "Group Name,Employee ID,Department\n"
               "Security Team,E9999,Cybersecurity\n";
    }
    point::Engine seed_authority_engine;
    seed_authority_engine.load_files({seed_identity, related_group});
    point::QueryRequest seed_authority_request;
    seed_authority_request.conditions.push_back({"Username", "speela"});
    seed_authority_request.output_fields = {
        "Employee ID", "Department"};
    const auto seed_authority_result =
        seed_authority_engine.query(seed_authority_request);
    assert(seed_authority_result.rows.size() == 1);
    assert(seed_authority_result.rows[0][0] == "E1001");
    assert(seed_authority_result.rows[0][1] == "Cybersecurity");
    const auto one_employee = mapped.universal_lookup(
        {"Employee ID", "Username", "Email"}, "E1001");
    assert(one_employee.rows.size() == 1);
    assert(one_employee.rows[0][0] == "E1001");
    assert(one_employee.rows[0][1] == "speela");
    assert(one_employee.rows[0][2] == "sunil@example.com");
    const auto missing_lookup = mapped.universal_lookup(
        {"Username"}, "COMPUTER-DOES-NOT-EXIST");
    assert(missing_lookup.rows.empty());

    const auto zero_users = root / "ZeroPaddedUsers.csv";
    const auto zero_devices = root / "UnpaddedDevices.csv";
    {
        std::ofstream out(zero_users);
        out << "Employee ID,Username\n00039929,zero.user\n";
    }
    {
        std::ofstream out(zero_devices);
        out << "Associate Number,Computer Name\n39929,PC-ZERO\n";
    }
    point::Engine zero_engine;
    zero_engine.set_field_synonyms({
        {"Employee ID", {"Associate Number"}}
    });
    zero_engine.load_files({zero_users, zero_devices});
    assert(zero_engine.relationships().size() == 1);
    const auto short_eid = zero_engine.universal_lookup(
        {"Username"}, "39929");
    assert(short_eid.rows.size() == 1);
    assert(short_eid.rows[0][0] == "zero.user");
    point::QueryRequest short_eid_request;
    short_eid_request.lookup_field = "Employee ID";
    short_eid_request.lookup_value = "39929";
    short_eid_request.output_fields = {"Username"};
    const auto direct_short_eid = zero_engine.query(short_eid_request);
    assert(direct_short_eid.rows.size() == 1);
    assert(direct_short_eid.rows[0][0] == "zero.user");
    const auto padded_eid = zero_engine.universal_lookup(
        {"Computer Name"}, "00039929");
    assert(padded_eid.rows.size() == 1);
    assert(padded_eid.rows[0][0] == "PC-ZERO");
    const auto preserved_eid = zero_engine.universal_lookup(
        {"Employee ID"}, "zero.user");
    assert(preserved_eid.rows.size() == 1);
    assert(preserved_eid.rows[0][0] == "00039929");

    const auto group_memberships = root / "GroupMemberships.csv";
    const auto group_catalog = root / "GroupCatalog.csv";
    {
        std::ofstream out(group_memberships);
        out << "Employee ID,Group ID,Group Name\n"
               "E000004,G0106,Legal Standard Access\n";
    }
    {
        std::ofstream out(group_catalog);
        out << "Group ID,Group Name,Group Owner\n"
               "G0106,Legal Standard Access,Legal\n";
    }
    point::Engine group_engine;
    group_engine.load_files({group_memberships, group_catalog});
    assert(!group_engine.relationships().empty());
    point::QueryRequest owner_request;
    owner_request.lookup_field = "Employee ID";
    owner_request.lookup_value = "E000004";
    owner_request.output_fields = {
        "Group ID", "Group Name", "Group Owner"};
    const auto owner_result = group_engine.query(owner_request);
    assert(owner_result.rows.size() == 1);
    assert(owner_result.rows[0][0] == "G0106");
    assert(owner_result.rows[0][2] == "Legal");

    const auto unsafe_groups_left = root / "UnsafeGroupsLeft.csv";
    const auto unsafe_groups_right = root / "UnsafeGroupsRight.csv";
    {
        std::ofstream out(unsafe_groups_left);
        out << "Group ID,Left Value\nG1,A\nG1,B\n";
    }
    {
        std::ofstream out(unsafe_groups_right);
        out << "Group ID,Right Value\nG1,C\nG1,D\n";
    }
    point::Engine unsafe_group_engine;
    unsafe_group_engine.load_files({unsafe_groups_left, unsafe_groups_right});
    assert(unsafe_group_engine.relationships().empty());

    const auto request_assignments = root / "RequestAssignments.csv";
    const auto request_catalog = root / "RequestCatalog.csv";
    {
        std::ofstream out(request_assignments);
        out << "Employee ID,Request Code\nE000004,R-42\n";
    }
    {
        std::ofstream out(request_catalog);
        out << "Request Code,Approver\nR-42,Security Manager\n";
    }
    point::Engine generic_catalog_engine;
    generic_catalog_engine.load_files({request_assignments, request_catalog});
    point::QueryRequest approver_request;
    approver_request.lookup_field = "Employee ID";
    approver_request.lookup_value = "E000004";
    approver_request.output_fields = {"Request Code", "Approver"};
    const auto approver_result = generic_catalog_engine.query(approver_request);
    assert(approver_result.rows.size() == 1);
    assert(approver_result.rows[0][1] == "Security Manager");

    const auto duplicates = root / "DuplicateNames.csv";
    {
        std::ofstream out(duplicates);
        out << "Employee ID,Username,Display Name\n"
               "E2001,alex.one,Alex Lee\n"
               "E2002,alex.two,Lee Alex\n";
    }
    point::Engine ambiguous_engine;
    ambiguous_engine.load_files({duplicates});
    const auto ambiguous = ambiguous_engine.resolve_identity_from_name(
        "Username", "Alex Lee");
    assert(ambiguous.status ==
        point::IdentityResolutionStatus::Ambiguous);
    assert(ambiguous.distinct_matches == 2);

    const auto blank_leading_header = root / "BlankLeadingHeader.csv";
    {
        std::ofstream out(blank_leading_header);
        out << ",First Name,Last Name,Gender,Country,Age,Date,Id\n"
               "1,Dulce,Abril,Female,United States,32,15/10/2017,1562\n";
    }
    point::Engine blank_header_engine;
    blank_header_engine.load_files({blank_leading_header});
    assert(blank_header_engine.issues().empty());
    assert(blank_header_engine.all_fields().size() == 8);
    point::QueryRequest blank_header_request;
    blank_header_request.lookup_field = "Id";
    blank_header_request.lookup_value = "1562";
    blank_header_request.output_fields = {
        "Unnamed Column 1", "First Name", "Last Name"};
    const auto blank_header_result =
        blank_header_engine.query(blank_header_request);
    assert(blank_header_result.rows.size() == 1);
    assert(blank_header_result.rows[0][0] == "1");
    assert(blank_header_result.rows[0][1] == "Dulce");
    assert(blank_header_result.rows[0][2] == "Abril");

    const auto unique_first_names = root / "UniqueFirstNames.csv";
    const auto repeated_first_names = root / "RepeatedFirstNames.csv";
    {
        std::ofstream out(unique_first_names);
        out << "First Name,Last Name\nAlice,One\nBob,Two\n";
    }
    {
        std::ofstream out(repeated_first_names);
        out << "First Name,Last Name\n"
               "Shanice,Mccrystal\nShanice,Mccrystal\n";
    }
    point::Engine mixed_uniqueness_engine;
    mixed_uniqueness_engine.load_files(
        {unique_first_names, repeated_first_names});
    const auto shanice_result = mixed_uniqueness_engine.universal_lookup(
        {"First Name", "Last Name"}, "Shanice");
    assert(!shanice_result.rows.empty());
    assert(shanice_result.rows[0][0] == "Shanice");
    assert(shanice_result.rows[0][1] == "Mccrystal");

    const auto date_fields = root / "DateFieldPriority.csv";
    {
        std::ofstream out(date_fields);
        out << "Order Date,Ship Date,Order ID,City\n"
               "44933,44938,US-2023-105417,Huntsville\n"
               "44929,44933,US-2023-103800,Houston\n"
               "44932,44933,US-2023-106054,Athens\n";
    }
    point::Engine date_priority_engine;
    date_priority_engine.load_files({date_fields});
    const auto broad_date_result = date_priority_engine.universal_lookup(
        {"Order Date", "City", "Order ID", "Ship Date"}, "44933");
    assert(broad_date_result.rows.size() == 3);
    point::QueryRequest preferred_date_request;
    preferred_date_request.output_fields =
        {"Order Date", "City", "Order ID", "Ship Date"};
    preferred_date_request.conditions.push_back({"Order Date", "44933"});
    const auto preferred_date_result =
        date_priority_engine.query(preferred_date_request);
    assert(preferred_date_result.rows.size() == 1);
    assert(preferred_date_result.rows[0][2] == "US-2023-105417");

    point::Engine contextual_synonym_engine;
    contextual_synonym_engine.set_field_synonyms({
        {"Department", {"Business Unit"}}
    });

    bool conflict_rejected = false;
    try {
        mapped.set_field_synonyms({
            {"Employee ID", {"Worker Number"}},
            {"User ID", {"Worker Number"}}
        });
    } catch (...) {
        conflict_rejected = true;
    }
    assert(conflict_rejected);

    std::filesystem::remove_all(root);
    std::cout << "schema mapping tests passed\n";
}
