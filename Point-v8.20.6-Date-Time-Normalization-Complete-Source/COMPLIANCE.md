# Point v8.0 Compliance Control Matrix

## Scope and accurate status

Point v8.0 is a compliance-ready local analytics application. It implements
technical safeguards but is not, by itself, a NIST, PCI DSS, or GDPR
certification. Compliance depends on the deployed environment, data flows,
policies, training, contracts, incident response, evidence, and assessment.

The default PCI posture is **cardholder-data avoidance**: Point blocks exports
containing values that pass payment-card length and Luhn validation. Point
must not be placed in the cardholder data environment unless a Qualified
Security Assessor approves the complete deployment scope.

## Implemented technical controls

| Control | Implementation | Primary alignment |
|---|---|---|
| Named-user access | Windows token must belong to `Point Users` or `Point Administrators` | NIST PR.AA; PCI 7; GDPR 32 |
| Export authorization | CSV export requires `Point Exporters` or `Point Administrators` | NIST PR.AA; PCI 7; GDPR 25/32 |
| Fail-closed configuration | Startup fails when the policy is missing, invalid, or access is unauthorized | NIST PR.AA; PCI 7 |
| Data-directory permissions | Protected DACL for Inbox, Workspace, Exports, and Logs | NIST PR.AA; PCI 7; GDPR 25/32 |
| Workspace encryption | Windows DPAPI, current-user scope, UI-free operation | NIST PR.DS; PCI 3; GDPR 25/32 |
| Payment-card avoidance | 13–19 digit normalization plus Luhn detection blocks export | PCI 3 |
| Sensitive export masking | Secrets, credentials, IDs, PAN, and verification codes show only the last four characters | NIST PR.DS; PCI 3; GDPR 5/25 |
| Spreadsheet injection defense | Formula-leading CSV values are escaped | NIST PR.DS; PCI 6 |
| Audit trail | UTC event log with control-character neutralization and chained SHA-256 hashes | NIST DE.CM; PCI 10; GDPR 5(2)/32 |
| Retention | Configurable automatic expiry for exports, workspaces, and logs | NIST GV.PO; PCI 10; GDPR 5(1)(e) |
| Secure build | Warning-as-error, SDL checks, CFG, ASLR, and DEP | NIST PR.PS; PCI 6; GDPR 32 |
| Bounded parsing | Rejects links, null bytes, malformed CSV, oversized files/cells/rows/columns | NIST PR.DS; PCI 6; GDPR 32 |
| Local-only processing | No network, telemetry, macros, scripting, or shell execution | NIST PR.DS; GDPR 25 |

## Deployment controls still owned by the organization

1. Approve the lawful purpose and legal basis for every personal-data report.
2. Maintain a record of processing activities and complete a DPIA where needed.
3. Limit Inbox data to necessary fields and approved populations.
4. Provision and review Point Windows-group membership.
5. Send Point and Windows security events to a protected central SIEM.
6. Back up required evidence and test restoration.
7. Maintain incident-response and breach-notification procedures.
8. Patch Windows, build tools, and Point under a documented vulnerability
   management program.
9. Perform code-signing, malware scanning, penetration testing, and release
   approval before production deployment.
10. Obtain the required PCI SAQ/ROC/AOC assessment if Point enters PCI scope.
11. Operate GDPR access, correction, restriction, portability, objection, and
   erasure processes outside Point where the controller determines they apply.
12. Validate retention values with Legal, Privacy, Security, and Records teams.

## Default retention

| Data | Default | Configuration |
|---|---:|---|
| Exports | 30 days | `export_retention_days` |
| Saved workspaces | 30 days | `workspace_retention_days` |
| Local audit logs | 365 days | `log_retention_days` |

Deletion is logical filesystem deletion. On SSDs and managed endpoints, use
full-disk encryption and the organization's media-sanitization process; the
application does not claim physical overwrite guarantees.

## Evidence package

Retain these items for review:

- approved `point-security.conf`;
- Windows group membership export and access-review approval;
- directory ACL evidence;
- Point release hash and code-signing evidence;
- successful test output;
- audit-chain verification output;
- retention-job evidence;
- vulnerability and penetration-test reports;
- data-flow diagram and inventory;
- privacy/legal approval and PCI scope determination.
# Point Fetcher network boundary

`Point.exe` does not perform website or API downloads. Optional outbound HTTPS
is isolated in `PointFetcher.exe`. Fetcher rejects non-HTTPS URLs, accepts only
supported CSV/Excel output names, limits responses to 2 GiB, validates the
download before atomic publication, and stores authentication secrets in the
current user's Windows Credential Manager. Passwords, bearer tokens, and API
keys are not written to Point or Fetcher configuration files or audit logs.

Interactive browser authentication, CAPTCHA, MFA, and form-login automation
are outside the generic Fetcher boundary and require an approved API or
site-specific connector.
