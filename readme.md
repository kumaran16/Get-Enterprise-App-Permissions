# GetEnterpriseAppPermissions.ps1

A PowerShell script to retrieve and analyze permissions for Enterprise Applications in Microsoft Entra ID (formerly Azure AD).

## Description

This script connects to Microsoft Graph and retrieves detailed permission information for Enterprise Applications, including delegated permissions, application permissions, and role assignments.

## Prerequisites

- PowerShell 5.1 or later
- Microsoft.Graph PowerShell module
- Appropriate permissions to read Enterprise Applications in Entra ID

## Installation

1. Install the Microsoft Graph PowerShell module:
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

2. Download the script to your local machine

## Usage

```powershell
.\GetEnterpriseAppPermissions.ps1
```

## Parameters

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| TBD | TBD | TBD | TBD |

## Output

The script generates a report containing:
- Enterprise Application details
- Granted permissions
- Permission scopes
- Role assignments

## Examples

```powershell
# Basic usage
.\GetEnterpriseAppPermissions.ps1

# With specific parameters (example)
.\GetEnterpriseAppPermissions.ps1 -OutputPath "C:\Reports\"
```

## Notes

- Requires Global Reader or Application Administrator permissions
- Script may take time to complete for tenants with many Enterprise Applications

## License

MIT License