<#
.SYNOPSIS
    Retrieves Enterprise App permissions and their consent status from Microsoft Entra ID.

.DESCRIPTION
    This script connects to Microsoft Graph and retrieves detailed information about Enterprise Applications,
    including their permissions, consent types, and various filtering options.

.PARAMETER CreateSession
    Forces a new Microsoft Graph connection session.

.PARAMETER TenantId
    The Tenant ID for certificate-based authentication.

.PARAMETER ClientId
    The Client ID for certificate-based authentication.

.PARAMETER CertificateThumbPrint
    The certificate thumbprint for certificate-based authentication.

.PARAMETER ApplicationId
    Filter by specific Application ID.

.PARAMETER ApplicationName
    Filter by specific Application Name.

.PARAMETER ObjectId
    Filter by specific Object ID.

.PARAMETER APIName
    Filter by specific API Name.

.PARAMETER AppVisibility
    Filter by app visibility (VisibleApps or HiddenApps).

.PARAMETER AppOrigin
    Filter by app origin (HomeTenant or ExternalTenant).

.PARAMETER UsersSignIn
    Filter by user sign-in status (Enabled or Disabled).

.PARAMETER ConsentType
    Filter by consent type (AdminConsent or UserConsent).

.PARAMETER AdminConsentApplicationPermissions
    Filter by specific admin consented application permissions.

.PARAMETER AdminConsentDelegatedPermissions
    Filter by specific admin consented delegated permissions.

.PARAMETER UserConsents
    Filter by specific user consented permissions.

.PARAMETER AccessScopeToAllUsers
    Include only apps accessible to all users.

.PARAMETER RoleAssignmentRequiredApps
    Include only apps that require role assignment.

.PARAMETER OwnerlessApps
    Include only apps without owners.

.PARAMETER IncludeAppsWithNoPermissions
    Include apps that have no permissions assigned.

.EXAMPLE
    .\GetEnterpriseAppPermissions.ps1 -IncludeAppsWithNoPermissions

.EXAMPLE
    .\GetEnterpriseAppPermissions.ps1 -AppVisibility "HiddenApps" -OwnerlessApps

.NOTES
    Requires Microsoft.Graph PowerShell module.
    Requires appropriate permissions to read Enterprise Applications.
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false, HelpMessage = "Forces a new Microsoft Graph connection session")]
    [switch]$CreateSession,
    
    [Parameter(Mandatory = $false, HelpMessage = "Tenant ID for certificate authentication")]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false, HelpMessage = "Client ID for certificate authentication")]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false, HelpMessage = "Certificate thumbprint for authentication")]
    [ValidatePattern('^[0-9a-fA-F]{40}$')]
    [string]$CertificateThumbPrint,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by Application ID")]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$ApplicationId,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by Application Name")]
    [ValidateNotNullOrEmpty()]
    [string]$ApplicationName,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by Object ID")]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$ObjectId,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by API Name")]
    [ValidateNotNullOrEmpty()]
    [string]$APIName,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by app visibility")]
    [ValidateSet("VisibleApps", "HiddenApps")]
    [string]$AppVisibility,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by app origin")]
    [ValidateSet("HomeTenant", "ExternalTenant")]
    [string]$AppOrigin,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by user sign-in status")]
    [ValidateSet("Enabled", "Disabled")]
    [string]$UsersSignIn,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by consent type")]
    [ValidateSet("AdminConsent", "UserConsent")]
    [string]$ConsentType,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by admin consented application permissions")]
    [string[]]$AdminConsentApplicationPermissions,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by admin consented delegated permissions")]
    [string[]]$AdminConsentDelegatedPermissions,
    
    [Parameter(Mandatory = $false, HelpMessage = "Filter by user consented permissions")]
    [string[]]$UserConsents,
    
    [Parameter(Mandatory = $false, HelpMessage = "Include only apps accessible to all users")]
    [switch]$AccessScopeToAllUsers,
    
    [Parameter(Mandatory = $false, HelpMessage = "Include only apps that require role assignment")]
    [switch]$RoleAssignmentRequiredApps,
    
    [Parameter(Mandatory = $false, HelpMessage = "Include only apps without owners")]
    [switch]$OwnerlessApps,
    
    [Parameter(Mandatory = $false, HelpMessage = "Include apps with no permissions")]
    [switch]$IncludeAppsWithNoPermissions
)

# Global variables
$ErrorActionPreference = 'Stop'
$script:ServicePrincipalCache = @{}

function Connect-MgGraphSession {
    <#
    .SYNOPSIS
    Connects to Microsoft Graph with proper error handling.
    #>
    [CmdletBinding()]
    param()
    
    try {
        $MsGraphModule = Get-Module Microsoft.Graph -ListAvailable
        if (-not $MsGraphModule) {
            Write-Host "Microsoft Graph module is not installed. It is required to run this script." -ForegroundColor Red
            $confirm = Read-Host "Microsoft Graph module is required but not installed. Install it now? [Y/N]"
            if ($confirm -match "^[yY]") {
                Write-Host "Installing Microsoft Graph module. This may take a few minutes..." -ForegroundColor Yellow
                Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber -Force
            }
            else {
                Write-Error "Microsoft Graph PowerShell module is required but not installed. Please install it using 'Install-Module Microsoft.Graph -Scope CurrentUser'"
                exit 1
            }
        }

        if ($CreateSession.IsPresent) {
            try {
                Disconnect-MgGraph -ErrorAction SilentlyContinue
            }
            catch {
                # Ignore disconnection errors
            }
        }

        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
        
        if ($TenantId -and $ClientId -and $CertificateThumbPrint) {
            Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbPrint -NoWelcome
        }
        else {
            Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All" -NoWelcome
        }
        
        # Verify connection
        $context = Get-MgContext
        if (-not $context) {
            throw "Failed to establish Microsoft Graph connection"
        }
        
        Write-Host "Successfully connected to tenant: $($context.TenantId)" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        exit 1
    }
}

function Get-CachedServicePrincipal {
    <#
    .SYNOPSIS
    Gets a service principal with caching to improve performance.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServicePrincipalId
    )
    
    if (-not $script:ServicePrincipalCache.ContainsKey($ServicePrincipalId)) {
        try {
            $sp = Get-MgServicePrincipal -ServicePrincipalId $ServicePrincipalId -ErrorAction Stop
            $script:ServicePrincipalCache[$ServicePrincipalId] = $sp
        }
        catch {
            Write-Warning "Could not retrieve service principal $ServicePrincipalId`: $($_.Exception.Message)"
            return $null
        }
    }
    
    return $script:ServicePrincipalCache[$ServicePrincipalId]
}

function Test-FilterCriteria {
    <#
    .SYNOPSIS
    Tests if an application meets the filtering criteria.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$ServicePrincipal,
        [Parameter(Mandatory = $true)]
        [string]$UserVisibility,
        [Parameter(Mandatory = $true)]
        [string]$AccessScope,
        [Parameter(Mandatory = $true)]
        [string]$AppOrg,
        [Parameter(Mandatory = $true)]
        [string]$Owners
    )
    
    # Application filters
    if ($ApplicationId -and ($ApplicationId -ne $ServicePrincipal.AppId)) { return $false }
    if ($ApplicationName -and ($ApplicationName -ne $ServicePrincipal.DisplayName)) { return $false }
    if ($ObjectId -and ($ObjectId -ne $ServicePrincipal.Id)) { return $false }
    
    # Status filters
    if ($UsersSignIn -eq "Enabled" -and (-not $ServicePrincipal.AccountEnabled)) { return $false }
    if ($UsersSignIn -eq "Disabled" -and $ServicePrincipal.AccountEnabled) { return $false }
    
    # Visibility filters
    if ($AppVisibility -eq "VisibleApps" -and ($UserVisibility -ne "Visible")) { return $false }
    if ($AppVisibility -eq "HiddenApps" -and ($UserVisibility -ne "Hidden")) { return $false }
    
    # Access scope filters
    if ($AccessScopeToAllUsers.IsPresent -and ($AccessScope -eq "Only assigned users can access")) { return $false }
    if ($RoleAssignmentRequiredApps.IsPresent -and ($AccessScope -eq "All users can access")) { return $false }
    
    # Owner filters
    if ($OwnerlessApps.IsPresent -and ($Owners -ne "-")) { return $false }
    
    # Origin filters
    if ($AppOrigin -eq "HomeTenant" -and ($AppOrg -eq "External tenant")) { return $false }
    if ($AppOrigin -eq "ExternalTenant" -and ($AppOrg -eq "Home tenant")) { return $false }
    
    return $true
}

function Test-PermissionFilters {
    <#
    .SYNOPSIS
    Tests if permissions meet the filtering criteria.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AdminApps,
        [Parameter(Mandatory = $true)]
        [string]$AdminDelegated,
        [Parameter(Mandatory = $true)]
        [string]$UserDelegated,
        [Parameter(Mandatory = $true)]
        [string]$ResourceName
    )
    
    # Include apps with no permissions filter
    if ((-not $IncludeAppsWithNoPermissions.IsPresent) -and 
        ($AdminDelegated -eq "-" -and $AdminApps -eq "-" -and $UserDelegated -eq "-")) {
        return $false
    }
    
    # Permission-specific filters
    if ($AdminConsentApplicationPermissions -and 
        ((($AdminApps -split ", ") | Where-Object { $_ -in $AdminConsentApplicationPermissions }).Count -eq 0)) {
        return $false
    }
    
    if ($AdminConsentDelegatedPermissions -and 
        ((($AdminDelegated -split ", ") | Where-Object { $_ -in $AdminConsentDelegatedPermissions }).Count -eq 0)) {
        return $false
    }
    
    if ($UserConsents -and 
        ((($UserDelegated -split ", ") | Where-Object { $_ -in $UserConsents }).Count -eq 0)) {
        return $false
    }
    
    # API Name filter
    if ($APIName -and ($APIName -ne $ResourceName)) { return $false }
    
    # Consent type filters
    if ($ConsentType -eq "AdminConsent" -and $UserDelegated -ne '-') { return $false }
    if ($ConsentType -eq "UserConsent" -and ($AdminApps -ne '-' -or $AdminDelegated -ne '-')) { return $false }
    
    return $true
}

# Main script execution
try {
    Connect-MgGraphSession
    
    $ExportCSV = Join-Path (Get-Location) "EntraAppPermissions_Report_$((Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')).csv"
    $TenantGUID = (Get-MgOrganization).Id
    Write-Host "Retrieving Enterprise applications and analyzing permissions..." -ForegroundColor Yellow
    
    $AppCount = 0 
    $PrintCount = 0

    $ServicePrincipals = Get-MgServicePrincipal -All
    $TotalApps = $ServicePrincipals.Count
    
    Write-Host "Discovered $TotalApps Enterprise applications for analysis..." -ForegroundColor Yellow
    
    foreach ($ServicePrincipal in $ServicePrincipals) {
        $AppCount++
        $AppName = $ServicePrincipal.DisplayName
        $AppId = $ServicePrincipal.AppId
        $ObjId = $ServicePrincipal.Id
        
        Write-Progress -Activity "Analyzing Enterprise Applications" -Status "Processing $AppCount of $TotalApps`: $AppName" -PercentComplete (($AppCount / $TotalApps) * 100) -Id 1
        
        try {
            # Get application properties
            $ServicePrincipalType = $ServicePrincipal.ServicePrincipalType
            $CreatedDateTime = $ServicePrincipal.AdditionalProperties.createdDateTime
            if ($CreatedDateTime) {
                $CreatedDateTime = [datetime]$CreatedDateTime
            }
            else {
                $CreatedDateTime = "Unknown"
            }
            
            $AccountEnabled = if ($ServicePrincipal.AccountEnabled) { "Enabled" } else { "Disabled" }
            
            # Get owners with error handling
            try {
                $OwnersList = Get-MgServicePrincipalOwner -ServicePrincipalId $ObjId -ErrorAction Stop
                $Owners = ($OwnersList | ForEach-Object { $_.AdditionalProperties["displayName"] } | Where-Object { $_ }) -join ", "
                if (-not $Owners) { $Owners = "-" }
            }
            catch {
                Write-Warning "Could not retrieve owners for app: $AppName"
                $Owners = "-"
            }
            
            $Tags = $ServicePrincipal.Tags
            $IsRoleAssignmentRequired = $ServicePrincipal.AppRoleAssignmentRequired
            
            # Determine visibility and access scope
            $UserVisibility = if ($Tags -contains "HideApp") { "Hidden" } else { "Visible" }
            $AccessScope = if ($IsRoleAssignmentRequired) { "Only assigned users can access" } else { "All users can access" }
            
            # Determine app origin
            $AppOwnerOrgId = $ServicePrincipal.AppOwnerOrganizationId
            $AppOrg = if ($AppOwnerOrgId -eq $TenantGUID) { "Home tenant" } else { "External tenant" }
            
            # Test basic filters first to avoid unnecessary API calls
            if (-not (Test-FilterCriteria -ServicePrincipal $ServicePrincipal -UserVisibility $UserVisibility -AccessScope $AccessScope -AppOrg $AppOrg -Owners $Owners)) {
                continue
            }
            
            # Get permissions
            try {
                $DelegatedGrants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $ObjId -All -ErrorAction Stop
                $AppAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ObjId -All -ErrorAction Stop
            }
            catch {
                Write-Warning "Could not retrieve permissions for app: $AppName - $($_.Exception.Message)"
                continue
            }
            
            $AllAPIids = @($DelegatedGrants.ResourceId; $AppAssignments.ResourceId) | Sort-Object -Unique
            if (-not $AllAPIids) { $AllAPIids = @('-') }
            
            foreach ($ResourceId in $AllAPIids) {
                if ($ResourceId -eq '-') {
                    $ResourceName = '-'
                    $AdminDelegated = '-'
                    $UserDelegated = '-'
                    $AdminApps = '-'
                }
                else {
                    $ResourceSp = Get-CachedServicePrincipal -ServicePrincipalId $ResourceId
                    if (-not $ResourceSp) {
                        Write-Warning "Failed to retrieve resource service principal '$ResourceId' for app '$AppName' - skipping this resource"
                        continue
                    }
                    
                    $ResourceName = $ResourceSp.DisplayName
                    
                    # Process admin delegated permissions
                    $AdminDelegatedList = $DelegatedGrants | Where-Object { $_.ResourceId -eq $ResourceId -and $_.ConsentType -eq "AllPrincipals" } | ForEach-Object { $_.Scope.Trim() }
                    $AdminDelegated = if (-not $AdminDelegatedList) { "-" } else { ($AdminDelegatedList -split "\s+" | Where-Object { $_ }) -join ", " }
                    
                    # Process user delegated permissions
                    $UserDelegatedList = $DelegatedGrants | Where-Object { $_.ResourceId -eq $ResourceId -and $_.ConsentType -eq "Principal" } | ForEach-Object { $_.Scope.Trim() }
                    $UserDelegated = if (-not $UserDelegatedList) { "-" } else { ($UserDelegatedList -split "\s+" | Where-Object { $_ }) -join ", " }
                    
                    # Process application permissions
                    $AdminAppsList = $AppAssignments | Where-Object { $_.ResourceId -eq $ResourceId } | ForEach-Object {
                        $role = $ResourceSp.AppRoles | Where-Object Id -eq $_.AppRoleId
                        if ($role) { $role.Value }
                    }
                    $AdminApps = if (-not $AdminAppsList) { "-" } else { $AdminAppsList -join ", " }
                }
                
                # Test permission filters
                if (-not (Test-PermissionFilters -AdminApps $AdminApps -AdminDelegated $AdminDelegated -UserDelegated $UserDelegated -ResourceName $ResourceName)) {
                    continue
                }
                
                # Create output object
                $PrintCount++
                $OutputObject = [PSCustomObject]@{
                    'App Name'                               = $AppName
                    'Object Id'                              = $ObjId
                    'API Name'                               = $ResourceName
                    'Admin Consented App Permissions'        = $AdminApps
                    'Admin Consented Delegated Permissions'  = $AdminDelegated
                    'User Consented Permissions'             = $UserDelegated
                    'Owners'                                 = $Owners
                    'Users Sign In'                          = $AccountEnabled
                    'User Visibility'                        = $UserVisibility
                    'Role Assignment Required'               = $AccessScope
                    'Service Principal Type'                 = $ServicePrincipalType
                    'App Id'                                 = $AppId
                    'App Origin'                             = $AppOrg
                    'App Org Id'                             = $AppOwnerOrgId
                    'API Id'                                 = $ResourceId
                    'Created Date'                           = $CreatedDateTime
                }
                
                $OutputObject | Export-Csv -Path $ExportCSV -Append -NoTypeInformation
            }
        }
        catch {
            Write-Warning "Error processing app '$AppName': $($_.Exception.Message)"
            continue
        }
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}
finally {
    try {
        Disconnect-MgGraph | Out-Null
    }
    catch {
        # Ignore disconnection errors
    }
}

Write-Host "Enterprise App permissions analysis completed successfully!" -ForegroundColor Green

if (Test-Path -Path $ExportCSV) {
    Write-Host ""
    Write-Host "Processing Summary:" -ForegroundColor Green
    Write-Host ("=" * 75) -ForegroundColor Cyan
    Write-Host "Total Enterprise Applications Analyzed: " -NoNewline -ForegroundColor Yellow
    Write-Host "$AppCount" -ForegroundColor White
    Write-Host "Records Generated in Report:            " -NoNewline -ForegroundColor Yellow  
    Write-Host "$PrintCount" -ForegroundColor White
    Write-Host ("=" * 75) -ForegroundColor Cyan
    Write-Host ""
    
    # Improved file opening logic without COM object
    $OpenFile = Read-Host -Prompt "Would you like to open the CSV report file? [Y/N]"
    if ($OpenFile -match "^[yY]") {
        try {
            Invoke-Item $ExportCSV
        }
        catch {
            Write-Warning "Failed to open the CSV file automatically. Please open it manually from: $ExportCSV"
        }
    }
    
    Write-Host "CSV report file location: " -NoNewline -ForegroundColor Cyan
    Write-Host $ExportCSV -ForegroundColor White
}
else {
    Write-Host "No report file was created - no Enterprise Applications matched the specified filter criteria." -ForegroundColor Red
    Write-Host "Consider adjusting your filter parameters or use -IncludeAppsWithNoPermissions to include applications without permissions." -ForegroundColor Yellow
}