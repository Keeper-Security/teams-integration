# Changelog

All notable changes to the Keeper Security app for Microsoft Teams are documented in this file.

## [1.2.0] - 2026-06-29

### Added
- **Nested Share Folder (NSF) support.** Record creation and access grants now detect whether the target shared folder is a Nested Share Folder and route through dedicated NSF, role-based paths (`createNsfRecord`, `grantNsfRecordAccess`, `grantNsfFolderAccess`) instead of the classic shared-folder flow. Subfolders inherit NSF status from their parent. Applies to `keeper-create-secret`, record requests, and folder requests.
- **Rotate-credentials-on-expiry for time-bound access.** Record and folder request approval cards now include a "Rotate credentials when access expires" toggle. When enabled — and rotation is configured on the underlying records — credentials auto-rotate via PAM when the granted access expires (`--rotate-on-expiration`). The result card surfaces a "Credential Rotation: Enabled on expiry" status, and rotation-not-configured errors are detected and reported.
- **Asynchronous post-create record resolution.** A new flow (`handlers/approval/postCreate.js`, `lookupRecordUidAfterCreate`) resolves the record UID after a record-add / nsf-record-add completes and updates the approval card in place.

### Changed
- CI: upgraded `actions/checkout` to v6 and `docker/login-action` to v4 for Node 24 compatibility.

### Security / Maintenance
- Dependency bumps (transitive): `qs` 6.15.1 → 6.15.3, `form-data` 4.0.5 → 4.0.6, `ws` 7.5.10 → 7.5.11, `side-channel` 1.1.0 → 1.1.1, `hasown` 2.0.3 → 2.0.4.

## [1.1.0] - 2026-05-11

### Added
- Enhanced the `keeper-create-secret` command with shared-folder selection, subfolder support, and improved UX.

### Fixed
- Addressed findings from the penetration-test report.

### Security / Maintenance
- Upgraded `dotenv` to v17.
- Dependency group bumps via Dependabot.

## [1.0.0] - 2026-03-13

### Added
- Initial release: Keeper Security bot for Microsoft Teams — request access to records and folders, share credentials securely with one-time links, SSO Cloud device approvals, and manage approval workflows directly from Teams.

[1.2.0]: https://github.com/Keeper-Security/teams-integration/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/Keeper-Security/teams-integration/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/Keeper-Security/teams-integration/releases/tag/v1.0.0
