# Change Log

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to [Semantic Versioning](https://semver.org/).

## [2.0.0] - 2026-07-29

This is a major release that migrates to certificate-based authentication and implements comprehensive improvements to error handling, retry logic, and code quality.

### Added

- Added GitHub Actions workflow for automated release creation (`Create Release` workflow)
- Added GitHub Actions workflow to verify CHANGELOG.md updates on pull requests (`Verify CHANGELOG Updated` workflow)
- Added certificate-based authentication functions (`Get-MSEntraCertificate` and `Get-MSEntraAccessToken`)
  - Implements JWT token generation with X.509 certificate signing
  - Uses SHA-256 certificate thumbprint (`x5t#S256`) for enhanced security
- Added comprehensive error handling with `Resolve-MicrosoftGraphAPIError` function
  - Extracts detailed error messages from Microsoft Graph API responses
  - Provides line numbers and friendly error messages for better debugging
- Added do-until loop retry mechanism for authentication method removal
  - Implements automatic retry for default authentication methods
  - Tracks retry attempts per authentication method using object properties
  - Maximum retry count configurable (default: 5 attempts)
- Added configurable authentication methods via `$authenticationMethodsConfig` hashtable
  - Currently supports: Microsoft Authenticator, Phone Authentication, Email Authentication
  - Easy to extend with additional authentication method types
- Added extensive `Write-Verbose` logging throughout all scripts for enhanced debugging
- Added `$actionMessage` variables before major operations for better audit logging
- Added early filtering of authenticators based on configured types for improved performance
- Added pagination support with `@odata.nextLink` for complete result sets
- Added parameter validation attributes (`[ValidateNotNullOrEmpty()]`) to functions

### Changed

- BREAKING: Updated global variable names to use Entra ID naming convention:
  - Old: `EntraAppId` → New: `EntraIdAppId`
  - Old: `EntraTenantId` → New: `EntraIdTenantId`
  - Old: `EntraAppSecret` → New: `EntraIdCertificateBase64String` and `EntraIdCertificatePassword`
- BREAKING: Replaced client secret authentication with certificate-based authentication for Microsoft Graph API access
  - Now uses JWT (JSON Web Tokens) generated from X.509 certificates
  - Requires certificate to be uploaded to Entra ID App Registration
- Updated API permissions to minimal required set:
  - `UserAuthenticationMethod.ReadWrite.All` - To manage user authentication methods
  - `User.Read.All` - To read user information
- Refactored authentication method removal logic from separate retry blocks to unified do-until loop
- Improved success messages from verbose "Successfully deleted" to concise "Deleted" (IsError=false indicates success)
- Enhanced delegated form with improved user search functionality and grid display options
- Updated data source model to prioritize DisplayName and UserPrincipalName columns
- Improved error handling function renamed from `Get-ErrorMessage` to `Resolve-MicrosoftGraphAPIError`
- Updated README.md with comprehensive documentation including:
  - Certificate-based authentication setup instructions
  - API permissions requirements
  - Connection settings with detailed variable descriptions
  - Remarks on supported authentication methods, retry mechanism, and error handling
  - Development resources with full v1.0 API endpoints
  - Getting help section with HelloID forum link
- Improved code formatting and consistency across all PowerShell scripts
- Enhanced audit logging for all authentication method removal operations
- Updated form categories order to "User Management", "Entra ID"

### Deprecated

- Deprecated support for client secret-based authentication (use certificate-based authentication instead)

### Removed

- Removed support for client secret-based authentication in favor of certificate-based authentication
- Removed excessive API permissions that were not required for MFA reset operations
- Removed old datasource file with incorrect naming convention
- Removed hardcoded authentication method types from switch statements

### Fixed

- Fixed pagination handling to ensure all users and authentication methods are retrieved when result sets exceed page limits
- Fixed error handling to provide more detailed error messages with line numbers and friendly messages
- Fixed logging to properly distinguish between configured types and found types before operation
- Fixed variable reference issues where variables were used before being defined
- Fixed emoji rendering in README.md by using `:bulb:` markdown notation instead of actual emoji

## [1.1.0] - 2025-06-06

### Added

- Added do-until loop retry mechanism for handling default authentication methods

### Changed

- Updated task script to improve retry logic for authentication method removal

### Deprecated

### Removed

### Fixed

- Fixed handling of default authentication methods that cannot be removed on first attempt

## [1.0.0] - 2024-10-29

### Added

- Initial release of HelloID-Conn-SA-Full-EntraID-MFA-Reset
- Basic Entra ID MFA reset functionality
- Form-based user selection with wildcard search across displayName, userPrincipalName, and mail
- Support for resetting Microsoft Authenticator and Phone authentication methods
- Client secret-based authentication for Microsoft Graph API
- Data source for searching Entra ID users with wildcard support
- Task for authentication method removal operations
- All-in-one setup script for HelloID form deployment
- Basic error handling and audit logging

### Changed

### Deprecated

### Removed

### Fixed
