<#
.SYNOPSIS
    Office 365 E1 + Microsoft 365 E3 / F3 / E5 license usage audit, including no Teams variants.

.DESCRIPTION
    This script identifies users assigned with Office 365 E1 and Microsoft 365 E3, F3 or E5 licenses,
    including EEA / no Teams / HUB variants, then consolidates Microsoft 365 usage
    reports for Exchange, OneDrive, SharePoint, Teams and Microsoft 365 Apps.

    The script itself is written in English.
    Console output and exported CSV columns are in French.

    Important:
    - Teams usage is reported.
    - For "no Teams" bundles, Teams usage is NOT counted as core Microsoft 365 bundle usage.
      This avoids considering a no Teams bundle as used only because the user may have Teams
      through another license.

.PREREQUISITES
    Install-Module Microsoft.Graph -Scope CurrentUser

.REQUIRED GRAPH SCOPES
    Reports.Read.All
    User.Read.All
    Organization.Read.All
    Directory.Read.All

.NOTES
    Official Microsoft reference:
    https://learn.microsoft.com/fr-fr/entra/identity/users/licensing-service-plan-reference
#>

param(
    [ValidateSet("D7", "D30", "D90", "D180")]
    [string]$Period = "D90",

    [string]$OutputFolder = ".\Audit-Usage-Licences-M365",

    [string]$CsvDelimiter = ";",

    [string]$MicrosoftLicensingReferenceCsvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv",

    [switch]$SkipSkuReferenceDownload
)

$ErrorActionPreference = "Stop"

# =========================
# Target Office 365 / Microsoft 365 SKUs
# =========================
# The list includes standard SKUs and known no Teams / EEA / HUB variants.
# Matching is performed by both skuPartNumber and skuId to handle SKU names with spaces or special characters.

$TargetProducts = @(
    [pscustomobject]@{
        ProductName           = "Office 365 E1"
        ExpectedSkuPartNumber = "STANDARDPACK"
        ExpectedSkuId         = "18181a46-0d4e-45cd-891e-60aabd171b4e"
        LicenseFamily         = "E1"
        IncludesTeams         = $true
    },
    [pscustomobject]@{
        ProductName           = "Office 365 E1 (no Teams)"
        ExpectedSkuPartNumber = "Office_365_E1_(no_Teams)"
        ExpectedSkuId         = "f8ced641-8e17-4dc5-b014-f5a2d53f6ac8"
        LicenseFamily         = "E1"
        IncludesTeams         = $false
    },
    [pscustomobject]@{
        ProductName           = "Office 365 E1 EEA (no Teams)"
        ExpectedSkuPartNumber = "Office_365_w/o_Teams_Bundle_E1"
        ExpectedSkuId         = "b57282e3-65bd-4252-9502-c0eae1e5ab7f"
        LicenseFamily         = "E1"
        IncludesTeams         = $false
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E3"
        ExpectedSkuPartNumber = "SPE_E3"
        ExpectedSkuId         = "05e9a617-0261-4cee-bb44-138d3ef5d965"
        LicenseFamily         = "E3"
        IncludesTeams         = $true
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E3 EEA (no Teams)"
        ExpectedSkuPartNumber = "O365_w/o Teams Bundle_M3"
        ExpectedSkuId         = "c2fe850d-fbbb-4858-b67d-bd0c6e746da3"
        LicenseFamily         = "E3"
        IncludesTeams         = $false
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E3 EEA (no Teams) - HUB 500 seats minimum"
        ExpectedSkuPartNumber = "O365_w/o Teams Bundle_M3_(500_seats_min)_HUB"
        ExpectedSkuId         = "602e6573-55a3-46b1-a1a0-cc267991501a"
        LicenseFamily         = "E3"
        IncludesTeams         = $false
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E3 - HUB 500 seats minimum"
        ExpectedSkuPartNumber = "Microsoft_365_E3"
        ExpectedSkuId         = "0c21030a-7e60-4ec7-9a0f-0042e0e0211a"
        LicenseFamily         = "E3"
        IncludesTeams         = $true
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 F3"
        ExpectedSkuPartNumber = "SPE_F1"
        ExpectedSkuId         = "66b55226-6b4f-492c-910c-a3b7a3c9d993"
        LicenseFamily         = "F3"
        IncludesTeams         = $true
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 F3 EEA (no Teams)"
        ExpectedSkuPartNumber = "Microsoft_365_F3_EEA_(no_Teams)"
        ExpectedSkuId         = "f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87"
        LicenseFamily         = "F3"
        IncludesTeams         = $false
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E5"
        ExpectedSkuPartNumber = "SPE_E5"
        ExpectedSkuId         = "06ebc4ee-1bb5-47dd-8120-11324bc54e06"
        LicenseFamily         = "E5"
        IncludesTeams         = $true
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E5 - HUB 500 seats minimum"
        ExpectedSkuPartNumber = "Microsoft_365_E5"
        ExpectedSkuId         = "db684ac5-c0e7-4f92-8284-ef9ebde75d33"
        LicenseFamily         = "E5"
        IncludesTeams         = $true
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E5 EEA (no Teams)"
        ExpectedSkuPartNumber = "O365_w/o_Teams_Bundle_M5"
        ExpectedSkuId         = "3271cf8e-2be5-4a09-a549-70fd05baaa17"
        LicenseFamily         = "E5"
        IncludesTeams         = $false
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E5 EEA (no Teams) - HUB 500 seats minimum"
        ExpectedSkuPartNumber = "O365_w/o_Teams_Bundle_M5_(500_seats_min)_HUB"
        ExpectedSkuId         = "1e988bf3-8b7c-4731-bec0-4e2a2946600c"
        LicenseFamily         = "E5"
        IncludesTeams         = $false
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E5 EEA (no Teams) without Audio Conferencing"
        ExpectedSkuPartNumber = "Microsoft_365_E5_EEA_(no_Teams)_without_Audio_Conferencing"
        ExpectedSkuId         = "90277bc7-a6fe-4181-99d8-712b08b8d32b"
        LicenseFamily         = "E5"
        IncludesTeams         = $false
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E5 EEA (no Teams) without Audio Conferencing - HUB 500 seats minimum"
        ExpectedSkuPartNumber = "Microsoft_365_E5_EEA_(no_Teams)without_Audio_Conferencing(500_seats_min)_HUB"
        ExpectedSkuId         = "a640eead-25f6-4bec-97e3-23cfd382d7c2"
        LicenseFamily         = "E5"
        IncludesTeams         = $false
    },
    [pscustomobject]@{
        ProductName           = "Microsoft 365 E5 EEA (no Teams) with Calling Minutes"
        ExpectedSkuPartNumber = "Microsoft_365_E5_EEA_(no_Teams)_with_Calling_Minutes"
        ExpectedSkuId         = "6ee4114a-9b2d-4577-9e7a-49fa43d222d3"
        LicenseFamily         = "E5"
        IncludesTeams         = $false
    }
)

# =========================
# Modules
# =========================

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null

$RawFolder = Join-Path $OutputFolder "Rapports-Bruts"
New-Item -Path $RawFolder -ItemType Directory -Force | Out-Null

$ReferenceFolder = Join-Path $OutputFolder "Reference-Microsoft"
New-Item -Path $ReferenceFolder -ItemType Directory -Force | Out-Null

$DetailOutput = Join-Path $OutputFolder "Usage_Licences_O365_M365_E1_E3_F3_E5_Detail_$Period.csv"
$SummaryOutput = Join-Path $OutputFolder "Usage_Licences_O365_M365_E1_E3_F3_E5_Synthese_$Period.csv"
$SkuOutput = Join-Path $OutputFolder "Usage_Licences_O365_M365_E1_E3_F3_E5_SKU_Audites.csv"

# =========================
# Graph connection
# =========================

$Scopes = @(
    "Reports.Read.All",
    "User.Read.All",
    "Organization.Read.All",
    "Directory.Read.All"
)

Write-Host "Connexion à Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $Scopes -NoWelcome

# =========================
# Helper functions
# =========================

function Get-PropertyValue {
    param(
        [object]$Object,
        [string[]]$Names
    )

    if ($null -eq $Object) {
        return $null
    }

    foreach ($Name in $Names) {
        $Property = $Object.PSObject.Properties |
            Where-Object { $_.Name -eq $Name } |
            Select-Object -First 1

        if ($Property) {
            return $Property.Value
        }
    }

    return $null
}

function Get-MicrosoftLicensingReference {
    param(
        [string]$CsvUrl,
        [string]$Folder,
        [switch]$SkipDownload
    )

    if ($SkipDownload) {
        Write-Warning "Téléchargement de la référence Microsoft ignoré. Utilisation du mapping local."
        return @()
    }

    $ReferenceCsvPath = Join-Path $Folder "Microsoft-Licensing-Service-Plan-Reference.csv"

    try {
        Write-Host "Téléchargement de la référence Microsoft des licences..." -ForegroundColor Cyan

        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
        catch {}

        Invoke-WebRequest -Uri $CsvUrl -OutFile $ReferenceCsvPath -UseBasicParsing -ErrorAction Stop

        if (!(Test-Path $ReferenceCsvPath) -or ((Get-Item $ReferenceCsvPath).Length -eq 0)) {
            Write-Warning "Référence Microsoft téléchargée mais fichier vide. Utilisation du mapping local."
            return @()
        }

        $Rows = @(Import-Csv -Path $ReferenceCsvPath)

        if ($Rows.Count -eq 0) {
            Write-Warning "Référence Microsoft importée mais aucune ligne détectée. Utilisation du mapping local."
            return @()
        }

        Write-Host "Référence Microsoft importée : $($Rows.Count) lignes." -ForegroundColor Green
        return $Rows
    }
    catch {
        Write-Warning "Impossible de télécharger/importer la référence Microsoft : $($_.Exception.Message)"
        Write-Warning "Le script continue avec le mapping local E1/E3/F3/E5."
        return @()
    }
}

function Resolve-TargetSkuCatalog {
    param(
        [array]$ReferenceRows,
        [array]$TargetProducts
    )

    $Catalog = @()

    foreach ($Target in $TargetProducts) {
        $MatchedRow = $null

        if ($ReferenceRows -and $ReferenceRows.Count -gt 0) {
            $Candidates = foreach ($Row in $ReferenceRows) {
                $ProductName = Get-PropertyValue -Object $Row -Names @(
                    "Product_Display_Name",
                    "Product Display Name",
                    "Product name",
                    "Product Name",
                    "Nom du produit"
                )

                $StringId = Get-PropertyValue -Object $Row -Names @(
                    "String_Id",
                    "String ID",
                    "String Id",
                    "ID de chaîne",
                    "SkuPartNumber"
                )

                $Guid = Get-PropertyValue -Object $Row -Names @(
                    "GUID",
                    "SkuId",
                    "skuId"
                )

                if (
                    $StringId -eq $Target.ExpectedSkuPartNumber -or
                    $Guid -eq $Target.ExpectedSkuId
                ) {
                    $Row
                }
            }

            $MatchedRow = @($Candidates | Select-Object -First 1)
        }

        if ($MatchedRow) {
            $ResolvedProductName = Get-PropertyValue -Object $MatchedRow -Names @(
                "Product_Display_Name",
                "Product Display Name",
                "Product name",
                "Product Name",
                "Nom du produit"
            )

            $ResolvedSkuPartNumber = Get-PropertyValue -Object $MatchedRow -Names @(
                "String_Id",
                "String ID",
                "String Id",
                "ID de chaîne",
                "SkuPartNumber"
            )

            $ResolvedSkuId = Get-PropertyValue -Object $MatchedRow -Names @(
                "GUID",
                "SkuId",
                "skuId"
            )

            if ([string]::IsNullOrWhiteSpace($ResolvedProductName)) {
                $ResolvedProductName = $Target.ProductName
            }

            if ([string]::IsNullOrWhiteSpace($ResolvedSkuPartNumber)) {
                $ResolvedSkuPartNumber = $Target.ExpectedSkuPartNumber
            }

            if ([string]::IsNullOrWhiteSpace($ResolvedSkuId)) {
                $ResolvedSkuId = $Target.ExpectedSkuId
            }

            $Catalog += [pscustomobject]@{
                ProductName   = $ResolvedProductName
                SkuPartNumber = $ResolvedSkuPartNumber
                SkuId         = $ResolvedSkuId
                LicenseFamily = $Target.LicenseFamily
                IncludesTeams = $Target.IncludesTeams
                Source        = "Microsoft Learn CSV"
            }
        }
        else {
            $Catalog += [pscustomobject]@{
                ProductName   = $Target.ProductName
                SkuPartNumber = $Target.ExpectedSkuPartNumber
                SkuId         = $Target.ExpectedSkuId
                LicenseFamily = $Target.LicenseFamily
                IncludesTeams = $Target.IncludesTeams
                Source        = "Mapping local"
            }
        }
    }

    $Catalog |
        Sort-Object SkuId -Unique
}

function Get-GraphReportCsv {
    param(
        [Parameter(Mandatory)]
        [string]$ReportName,

        [Parameter(Mandatory)]
        [string]$Period,

        [Parameter(Mandatory)]
        [string]$Folder
    )

    $FilePath = Join-Path $Folder "$ReportName-$Period.csv"
    $Uri = "https://graph.microsoft.com/v1.0/reports/$ReportName(period='$Period')?`$format=text/csv"

    Write-Host "Téléchargement du rapport : $ReportName ($Period)" -ForegroundColor Cyan

    try {
        Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputFilePath $FilePath | Out-Null
    }
    catch {
        Write-Warning "Impossible de récupérer le rapport $ReportName : $($_.Exception.Message)"
        return @()
    }

    if (!(Test-Path $FilePath) -or ((Get-Item $FilePath).Length -eq 0)) {
        Write-Warning "Rapport vide ou non récupéré : $ReportName"
        return @()
    }

    return @(Import-Csv -Path $FilePath)
}

function Convert-ToIntSafe {
    param($Value)

    if ([string]::IsNullOrWhiteSpace([string]$Value)) {
        return 0
    }

    $CleanValue = ([string]$Value).Trim().Replace(",", "")
    $Parsed = 0

    if ([int]::TryParse($CleanValue, [ref]$Parsed)) {
        return $Parsed
    }

    return 0
}

function Convert-ToDateSafe {
    param($Value)

    if ([string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    try {
        return ([datetime]$Value).ToString("yyyy-MM-dd")
    }
    catch {
        return [string]$Value
    }
}

function Convert-ToFrenchYesNo {
    param($Value)

    if ($Value -is [bool]) {
        if ($Value) { return "Oui" } else { return "Non" }
    }

    if ([string]::IsNullOrWhiteSpace([string]$Value)) {
        return "Non"
    }

    if ([string]$Value -match "^(true|yes|1|oui)$") {
        return "Oui"
    }

    return "Non"
}

function New-ReportIndexByUpn {
    param([array]$Rows)

    $Index = @{}

    foreach ($Row in $Rows) {
        $Upn = Get-PropertyValue -Object $Row -Names @(
            "User Principal Name",
            "Principal Name",
            "UPN"
        )

        if (![string]::IsNullOrWhiteSpace($Upn)) {
            $Index[$Upn.ToLowerInvariant()] = $Row
        }
    }

    return $Index
}

function Get-MaxDateString {
    param([string[]]$Values)

    $Dates = foreach ($Value in $Values) {
        if (![string]::IsNullOrWhiteSpace($Value)) {
            try {
                [datetime]$Value
            }
            catch {}
        }
    }

    if ($Dates) {
        return ($Dates | Sort-Object -Descending | Select-Object -First 1).ToString("yyyy-MM-dd")
    }

    return $null
}

function Export-CsvFrench {
    param(
        [Parameter(Mandatory)]
        [object[]]$InputObject,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Delimiter
    )

    try {
        $InputObject | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8BOM -Delimiter $Delimiter
    }
    catch {
        $InputObject | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter
    }
}

# =========================
# Microsoft licensing reference
# =========================

$MicrosoftReferenceRows = Get-MicrosoftLicensingReference `
    -CsvUrl $MicrosoftLicensingReferenceCsvUrl `
    -Folder $ReferenceFolder `
    -SkipDownload:$SkipSkuReferenceDownload

$SkuCatalog = Resolve-TargetSkuCatalog `
    -ReferenceRows $MicrosoftReferenceRows `
    -TargetProducts $TargetProducts

$SkuCatalogByPartNumber = @{}
$SkuCatalogById = @{}

foreach ($SkuItem in $SkuCatalog) {
    if (![string]::IsNullOrWhiteSpace($SkuItem.SkuPartNumber)) {
        $SkuCatalogByPartNumber[$SkuItem.SkuPartNumber] = $SkuItem
    }

    if (![string]::IsNullOrWhiteSpace($SkuItem.SkuId)) {
        $SkuCatalogById[$SkuItem.SkuId.ToLowerInvariant()] = $SkuItem
    }
}

Write-Host ""
Write-Host "SKU ciblés après résolution de la référence Microsoft :" -ForegroundColor Green

$SkuCatalog |
    Select-Object `
        @{Name = "Produit"; Expression = { $_.ProductName }},
        @{Name = "Famille"; Expression = { $_.LicenseFamily }},
        @{Name = "SKU"; Expression = { $_.SkuPartNumber }},
        @{Name = "GUID référence"; Expression = { $_.SkuId }},
        @{Name = "Teams inclus"; Expression = { if ($_.IncludesTeams) { "Oui" } else { "Non" } }},
        @{Name = "Source"; Expression = { $_.Source }} |
    Format-Table -AutoSize

Export-CsvFrench -InputObject $SkuCatalog -Path $SkuOutput -Delimiter $CsvDelimiter

# =========================
# Tenant SKUs
# =========================

Write-Host "Récupération des licences disponibles dans le tenant..." -ForegroundColor Cyan

$AllSkus = @(Get-MgSubscribedSku -All)

$SkuById = @{}

foreach ($Sku in $AllSkus) {
    $SkuById[$Sku.SkuId.ToString()] = $Sku
}

$TargetSkuPartNumbers = @($SkuCatalog.SkuPartNumber | Where-Object { ![string]::IsNullOrWhiteSpace($_) })
$TargetSkuIdsFromCatalog = @(
    $SkuCatalog.SkuId |
        Where-Object { ![string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.ToLowerInvariant() }
)

$TargetSkus = @(
    $AllSkus | Where-Object {
        ($TargetSkuPartNumbers -contains $_.SkuPartNumber) -or
        ($TargetSkuIdsFromCatalog -contains $_.SkuId.ToString().ToLowerInvariant())
    }
)

$ExistingSkuPartNumbers = @($AllSkus.SkuPartNumber)
$ExistingSkuIds = @($AllSkus | ForEach-Object { $_.SkuId.ToString().ToLowerInvariant() })

$MissingCatalogSkus = @(
    $SkuCatalog | Where-Object {
        ($ExistingSkuPartNumbers -notcontains $_.SkuPartNumber) -and
        ($ExistingSkuIds -notcontains $_.SkuId.ToLowerInvariant())
    }
)

if ($MissingCatalogSkus.Count -gt 0) {
    Write-Host ""
    Write-Host "SKU ciblés non présents dans le tenant :" -ForegroundColor Yellow

    $MissingCatalogSkus |
        Select-Object ProductName, SkuPartNumber, SkuId, LicenseFamily |
        Format-Table -AutoSize
}

if ($TargetSkus.Count -eq 0) {
    throw "Aucune licence Office 365 E1 ou Microsoft 365 E3/F3/E5 (y compris variantes sans Teams) trouvée dans le tenant."
}

$TargetSkuIds = @($TargetSkus | ForEach-Object { $_.SkuId.ToString() })

Write-Host ""
Write-Host "Licences auditées dans le tenant :" -ForegroundColor Green

$TenantSkuDisplay = foreach ($TenantSku in $TargetSkus) {
    $SkuIdKey = $TenantSku.SkuId.ToString().ToLowerInvariant()
    $CatalogItem = $null

    if ($SkuCatalogById.ContainsKey($SkuIdKey)) {
        $CatalogItem = $SkuCatalogById[$SkuIdKey]
    }
    elseif ($SkuCatalogByPartNumber.ContainsKey($TenantSku.SkuPartNumber)) {
        $CatalogItem = $SkuCatalogByPartNumber[$TenantSku.SkuPartNumber]
    }

    [pscustomobject]@{
        "Licence"            = if ($CatalogItem) { $CatalogItem.ProductName } else { $TenantSku.SkuPartNumber }
        "Famille"            = if ($CatalogItem) { $CatalogItem.LicenseFamily } else { "" }
        "SKU"                = $TenantSku.SkuPartNumber
        "GUID tenant"        = $TenantSku.SkuId
        "Teams inclus"       = if ($CatalogItem -and $CatalogItem.IncludesTeams) { "Oui" } else { "Non" }
        "Unités consommées"  = $TenantSku.ConsumedUnits
    }
}

$TenantSkuDisplay | Format-Table -AutoSize

# =========================
# Licensed users
# =========================

Write-Host "Récupération des utilisateurs licenciés Office 365 E1 / Microsoft 365 E3/F3/E5..." -ForegroundColor Cyan

$Users = @(
    Get-MgUser -All -Property `
        Id,
        DisplayName,
        UserPrincipalName,
        AccountEnabled,
        UserType,
        AssignedLicenses
)

$LicensedUsers = foreach ($User in $Users) {
    $MatchedLicenses = @()

    foreach ($AssignedLicense in $User.AssignedLicenses) {
        $AssignedSkuId = $AssignedLicense.SkuId.ToString()

        if ($TargetSkuIds -contains $AssignedSkuId) {
            $MatchedLicenses += $SkuById[$AssignedSkuId]
        }
    }

    if ($MatchedLicenses.Count -gt 0) {
        $SkuPartNumbers = @($MatchedLicenses.SkuPartNumber)
        $SkuIds = @($MatchedLicenses | ForEach-Object { $_.SkuId.ToString() })

        $FriendlyNames = foreach ($MatchedLicense in $MatchedLicenses) {
            $SkuIdKey = $MatchedLicense.SkuId.ToString().ToLowerInvariant()

            if ($SkuCatalogById.ContainsKey($SkuIdKey)) {
                $SkuCatalogById[$SkuIdKey].ProductName
            }
            elseif ($SkuCatalogByPartNumber.ContainsKey($MatchedLicense.SkuPartNumber)) {
                $SkuCatalogByPartNumber[$MatchedLicense.SkuPartNumber].ProductName
            }
            else {
                $MatchedLicense.SkuPartNumber
            }
        }

        $LicenseFamilies = foreach ($MatchedLicense in $MatchedLicenses) {
            $SkuIdKey = $MatchedLicense.SkuId.ToString().ToLowerInvariant()

            if ($SkuCatalogById.ContainsKey($SkuIdKey)) {
                $SkuCatalogById[$SkuIdKey].LicenseFamily
            }
            elseif ($SkuCatalogByPartNumber.ContainsKey($MatchedLicense.SkuPartNumber)) {
                $SkuCatalogByPartNumber[$MatchedLicense.SkuPartNumber].LicenseFamily
            }
        }

        $IncludesTeamsInBundle = $false

        foreach ($MatchedLicense in $MatchedLicenses) {
            $SkuIdKey = $MatchedLicense.SkuId.ToString().ToLowerInvariant()

            if ($SkuCatalogById.ContainsKey($SkuIdKey)) {
                if ($SkuCatalogById[$SkuIdKey].IncludesTeams -eq $true) {
                    $IncludesTeamsInBundle = $true
                }
            }
            elseif ($SkuCatalogByPartNumber.ContainsKey($MatchedLicense.SkuPartNumber)) {
                if ($SkuCatalogByPartNumber[$MatchedLicense.SkuPartNumber].IncludesTeams -eq $true) {
                    $IncludesTeamsInBundle = $true
                }
            }
        }

        $ReferenceSources = foreach ($MatchedLicense in $MatchedLicenses) {
            $SkuIdKey = $MatchedLicense.SkuId.ToString().ToLowerInvariant()

            if ($SkuCatalogById.ContainsKey($SkuIdKey)) {
                $SkuCatalogById[$SkuIdKey].Source
            }
            elseif ($SkuCatalogByPartNumber.ContainsKey($MatchedLicense.SkuPartNumber)) {
                $SkuCatalogByPartNumber[$MatchedLicense.SkuPartNumber].Source
            }
            else {
                "Tenant Graph"
            }
        }

        [pscustomobject]@{
            Id                    = $User.Id
            DisplayName           = $User.DisplayName
            UserPrincipalName     = $User.UserPrincipalName
            AccountEnabled        = $User.AccountEnabled
            UserType              = $User.UserType
            LicenseNames          = (($FriendlyNames | Select-Object -Unique) -join ";")
            LicenseFamilies       = (($LicenseFamilies | Where-Object { $_ } | Select-Object -Unique) -join ";")
            LicenseSkuPartNumbers = ($SkuPartNumbers -join ";")
            LicenseSkuIds         = ($SkuIds -join ";")
            LicenseReference      = (($ReferenceSources | Select-Object -Unique) -join ";")
            IncludesTeamsInBundle = $IncludesTeamsInBundle
        }
    }
}

Write-Host "Utilisateurs avec licence Office 365 E1 / Microsoft 365 E3/F3/E5 ciblée : $($LicensedUsers.Count)" -ForegroundColor Green

# =========================
# Usage reports
# =========================

Write-Host ""
Write-Host "Récupération des rapports d'usage Microsoft 365 sur la période $Period..." -ForegroundColor Cyan

$Office365ActiveReport = Get-GraphReportCsv -ReportName "getOffice365ActiveUserDetail" -Period $Period -Folder $RawFolder
$EmailReport           = Get-GraphReportCsv -ReportName "getEmailActivityUserDetail" -Period $Period -Folder $RawFolder
$OneDriveReport        = Get-GraphReportCsv -ReportName "getOneDriveActivityUserDetail" -Period $Period -Folder $RawFolder
$SharePointReport      = Get-GraphReportCsv -ReportName "getSharePointActivityUserDetail" -Period $Period -Folder $RawFolder
$TeamsReport           = Get-GraphReportCsv -ReportName "getTeamsUserActivityUserDetail" -Period $Period -Folder $RawFolder
$M365AppsReport        = Get-GraphReportCsv -ReportName "getM365AppUserDetail" -Period $Period -Folder $RawFolder

$Office365ByUpn  = New-ReportIndexByUpn -Rows $Office365ActiveReport
$EmailByUpn      = New-ReportIndexByUpn -Rows $EmailReport
$OneDriveByUpn   = New-ReportIndexByUpn -Rows $OneDriveReport
$SharePointByUpn = New-ReportIndexByUpn -Rows $SharePointReport
$TeamsByUpn      = New-ReportIndexByUpn -Rows $TeamsReport
$AppsByUpn       = New-ReportIndexByUpn -Rows $M365AppsReport

# =========================
# Consolidation
# =========================

Write-Host ""
Write-Host "Consolidation des usages par utilisateur..." -ForegroundColor Cyan

$Results = foreach ($LicensedUser in $LicensedUsers) {
    $UpnKey = $LicensedUser.UserPrincipalName.ToLowerInvariant()

    $Active = $Office365ByUpn[$UpnKey]
    $Mail   = $EmailByUpn[$UpnKey]
    $OD     = $OneDriveByUpn[$UpnKey]
    $SP     = $SharePointByUpn[$UpnKey]
    $Teams  = $TeamsByUpn[$UpnKey]
    $Apps   = $AppsByUpn[$UpnKey]

    # Exchange usage
    $ExchangeLastActivity = Convert-ToDateSafe (
        Get-PropertyValue -Object $Active -Names @("Exchange Last Activity Date")
    )

    if (!$ExchangeLastActivity) {
        $ExchangeLastActivity = Convert-ToDateSafe (
            Get-PropertyValue -Object $Mail -Names @("Last Activity Date")
        )
    }

    $MailSendCount = Convert-ToIntSafe (
        Get-PropertyValue -Object $Mail -Names @("Send Count")
    )

    $MailReceiveCount = Convert-ToIntSafe (
        Get-PropertyValue -Object $Mail -Names @("Receive Count")
    )

    $MailReadCount = Convert-ToIntSafe (
        Get-PropertyValue -Object $Mail -Names @("Read Count")
    )

    # OneDrive usage
    $OneDriveLastActivity = Convert-ToDateSafe (
        Get-PropertyValue -Object $Active -Names @("OneDrive Last Activity Date")
    )

    if (!$OneDriveLastActivity) {
        $OneDriveLastActivity = Convert-ToDateSafe (
            Get-PropertyValue -Object $OD -Names @("Last Activity Date")
        )
    }

    $OneDriveViewedEdited = Convert-ToIntSafe (
        Get-PropertyValue -Object $OD -Names @("Viewed Or Edited File Count", "Viewed or Edited File Count")
    )

    $OneDriveSynced = Convert-ToIntSafe (
        Get-PropertyValue -Object $OD -Names @("Synced File Count")
    )

    $OneDriveSharedInternal = Convert-ToIntSafe (
        Get-PropertyValue -Object $OD -Names @("Shared Internally File Count")
    )

    $OneDriveSharedExternal = Convert-ToIntSafe (
        Get-PropertyValue -Object $OD -Names @("Shared Externally File Count")
    )

    # SharePoint usage
    $SharePointLastActivity = Convert-ToDateSafe (
        Get-PropertyValue -Object $Active -Names @("SharePoint Last Activity Date")
    )

    if (!$SharePointLastActivity) {
        $SharePointLastActivity = Convert-ToDateSafe (
            Get-PropertyValue -Object $SP -Names @("Last Activity Date")
        )
    }

    $SharePointViewedEdited = Convert-ToIntSafe (
        Get-PropertyValue -Object $SP -Names @("Viewed Or Edited File Count", "Viewed or Edited File Count")
    )

    $SharePointVisitedPages = Convert-ToIntSafe (
        Get-PropertyValue -Object $SP -Names @("Visited Page Count")
    )

    $SharePointSharedInternal = Convert-ToIntSafe (
        Get-PropertyValue -Object $SP -Names @("Shared Internally File Count")
    )

    $SharePointSharedExternal = Convert-ToIntSafe (
        Get-PropertyValue -Object $SP -Names @("Shared Externally File Count")
    )

    # Teams usage
    $TeamsLastActivity = Convert-ToDateSafe (
        Get-PropertyValue -Object $Active -Names @("Teams Last Activity Date")
    )

    if (!$TeamsLastActivity) {
        $TeamsLastActivity = Convert-ToDateSafe (
            Get-PropertyValue -Object $Teams -Names @("Last Activity Date")
        )
    }

    $TeamsTeamChat = Convert-ToIntSafe (
        Get-PropertyValue -Object $Teams -Names @(
            "Team Chat Message Count",
            "Channel Message Count"
        )
    )

    $TeamsPrivateChat = Convert-ToIntSafe (
        Get-PropertyValue -Object $Teams -Names @(
            "Private Chat Message Count",
            "Chat Message Count"
        )
    )

    $TeamsCalls = Convert-ToIntSafe (
        Get-PropertyValue -Object $Teams -Names @("Call Count")
    )

    $TeamsMeetings = Convert-ToIntSafe (
        Get-PropertyValue -Object $Teams -Names @(
            "Meeting Count",
            "Meetings Organized Count",
            "Meetings Attended Count"
        )
    )

    # Microsoft 365 Apps usage
    $M365AppsLastActivation = Convert-ToDateSafe (
        Get-PropertyValue -Object $Apps -Names @("Last Activation Date")
    )

    $M365AppsLastActivity = Convert-ToDateSafe (
        Get-PropertyValue -Object $Apps -Names @("Last Activity Date")
    )

    $AppsWindows = Convert-ToFrenchYesNo (
        Get-PropertyValue -Object $Apps -Names @("Windows")
    )

    $AppsMac = Convert-ToFrenchYesNo (
        Get-PropertyValue -Object $Apps -Names @("Mac")
    )

    $AppsMobile = Convert-ToFrenchYesNo (
        Get-PropertyValue -Object $Apps -Names @("Mobile")
    )

    $AppsWeb = Convert-ToFrenchYesNo (
        Get-PropertyValue -Object $Apps -Names @("Web")
    )

    $AppsOutlook = Convert-ToFrenchYesNo (
        Get-PropertyValue -Object $Apps -Names @("Outlook")
    )

    $AppsWord = Convert-ToFrenchYesNo (
        Get-PropertyValue -Object $Apps -Names @("Word")
    )

    $AppsExcel = Convert-ToFrenchYesNo (
        Get-PropertyValue -Object $Apps -Names @("Excel")
    )

    $AppsPowerPoint = Convert-ToFrenchYesNo (
        Get-PropertyValue -Object $Apps -Names @("PowerPoint")
    )

    $AppsTeams = Convert-ToFrenchYesNo (
        Get-PropertyValue -Object $Apps -Names @("Teams")
    )

    $LatestActivityDate = Get-MaxDateString -Values @(
        $ExchangeLastActivity,
        $OneDriveLastActivity,
        $SharePointLastActivity,
        $TeamsLastActivity,
        $M365AppsLastActivity,
        $M365AppsLastActivation
    )

    $CoreActivitySignals =
        $MailSendCount +
        $MailReceiveCount +
        $MailReadCount +
        $OneDriveViewedEdited +
        $OneDriveSynced +
        $OneDriveSharedInternal +
        $OneDriveSharedExternal +
        $SharePointViewedEdited +
        $SharePointVisitedPages +
        $SharePointSharedInternal +
        $SharePointSharedExternal

    $TeamsActivitySignals =
        $TeamsTeamChat +
        $TeamsPrivateChat +
        $TeamsCalls +
        $TeamsMeetings

    $AppsActivitySignals = 0
    if ($M365AppsLastActivity -or $M365AppsLastActivation) {
        $AppsActivitySignals = 1
    }

    $TotalActivitySignals =
        $CoreActivitySignals +
        $TeamsActivitySignals +
        $AppsActivitySignals

    $HasExchangeUsage = if (($MailSendCount + $MailReceiveCount + $MailReadCount) -gt 0 -or $ExchangeLastActivity) { "Oui" } else { "Non" }

    $HasOneDriveUsage = if (($OneDriveViewedEdited + $OneDriveSynced + $OneDriveSharedInternal + $OneDriveSharedExternal) -gt 0 -or $OneDriveLastActivity) { "Oui" } else { "Non" }

    $HasSharePointUsage = if (($SharePointViewedEdited + $SharePointVisitedPages + $SharePointSharedInternal + $SharePointSharedExternal) -gt 0 -or $SharePointLastActivity) { "Oui" } else { "Non" }

    $HasTeamsUsage = if (($TeamsActivitySignals) -gt 0 -or $TeamsLastActivity) { "Oui" } else { "Non" }

    $HasAppsUsage = if (
        $M365AppsLastActivity -or
        $M365AppsLastActivation -or
        @(
            $AppsWindows,
            $AppsMac,
            $AppsMobile,
            $AppsWeb,
            $AppsOutlook,
            $AppsWord,
            $AppsExcel,
            $AppsPowerPoint,
            $AppsTeams
        ) -contains "Oui"
    ) {
        "Oui"
    }
    else {
        "Non"
    }

    $HasCoreM365Usage = if (
        $HasExchangeUsage -eq "Oui" -or
        $HasOneDriveUsage -eq "Oui" -or
        $HasSharePointUsage -eq "Oui" -or
        $HasAppsUsage -eq "Oui"
    ) {
        "Oui"
    }
    else {
        "Non"
    }

    $HasAnyUsage = if (
        $HasCoreM365Usage -eq "Oui" -or
        ($LicensedUser.IncludesTeamsInBundle -eq $true -and $HasTeamsUsage -eq "Oui")
    ) {
        "Oui"
    }
    else {
        "Non"
    }

    $LicenseFamilyList = @(
        ($LicensedUser.LicenseFamilies -split ";") |
            ForEach-Object { $_.Trim() } |
            Where-Object { ![string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    $HasOnlyE1LicenseFamily = (
        $LicenseFamilyList.Count -gt 0 -and
        @($LicenseFamilyList | Where-Object { $_ -ne "E1" }).Count -eq 0
    )

    $UsesDesktopApps = @(
        $AppsWindows,
        $AppsMac,
        $AppsOutlook,
        $AppsWord,
        $AppsExcel,
        $AppsPowerPoint
    ) -contains "Oui"

    $HasFrontlineLikeUsagePattern = (
        (
            $AppsWeb -eq "Oui" -or
            $AppsMobile -eq "Oui" -or
            $HasExchangeUsage -eq "Oui" -or
            $HasOneDriveUsage -eq "Oui" -or
            $HasSharePointUsage -eq "Oui" -or
            $HasTeamsUsage -eq "Oui"
        ) -and
        !$UsesDesktopApps
    )

    $IsE1ToF3Candidate = (
        $HasOnlyE1LicenseFamily -and
        $HasAnyUsage -eq "Oui" -and
        $HasFrontlineLikeUsagePattern
    )

    $Recommendation = if ($HasAnyUsage -eq "Non") {
        "Aucun usage détecté sur $Period - candidat retrait ou réaffectation, à valider métier"
    }
    elseif ($HasCoreM365Usage -eq "Non" -and $HasTeamsUsage -eq "Oui" -and $LicensedUser.IncludesTeamsInBundle -eq $false) {
        "Usage Teams détecté uniquement, mais Teams n'est pas inclus dans le bundle - ne pas considérer comme usage du bundle M365"
    }
    elseif ($IsE1ToF3Candidate) {
        "Profil E1 orienté web/mobile sans usage Apps desktop - candidat optimisation E1 vers F3, à valider métier/prérequis"
    }
    elseif (($CoreActivitySignals + $AppsActivitySignals) -le 10) {
        "Usage faible sur $Period - candidat optimisation ou changement de licence, à valider métier"
    }
    else {
        "Usage détecté sur $Period - conserver ou analyser plus finement selon le profil"
    }

    [pscustomobject]@{
        "Nom complet"                              = $LicensedUser.DisplayName
        "UPN"                                      = $LicensedUser.UserPrincipalName
        "Compte actif"                             = if ($LicensedUser.AccountEnabled) { "Oui" } else { "Non" }
        "Type utilisateur"                         = $LicensedUser.UserType

        "Licence"                                  = $LicensedUser.LicenseNames
        "Famille de licence"                       = $LicensedUser.LicenseFamilies
        "SKU"                                      = $LicensedUser.LicenseSkuPartNumbers
        "GUID SKU"                                 = $LicensedUser.LicenseSkuIds
        "Référence SKU"                            = $LicensedUser.LicenseReference
        "Teams inclus dans le bundle"              = if ($LicensedUser.IncludesTeamsInBundle) { "Oui" } else { "Non" }

        "Usage détecté"                            = $HasAnyUsage
        "Usage cœur Microsoft 365"                 = $HasCoreM365Usage
        "Candidat optimisation E1 vers F3"         = if ($IsE1ToF3Candidate) { "Oui" } else { "Non" }
        "Dernière activité détectée"               = $LatestActivityDate
        "Score d'activité total"                   = $TotalActivitySignals
        "Score d'activité cœur M365"               = ($CoreActivitySignals + $AppsActivitySignals)
        "Score d'activité Teams"                   = $TeamsActivitySignals

        "Usage Exchange"                           = $HasExchangeUsage
        "Dernière activité Exchange"               = $ExchangeLastActivity
        "Mails envoyés"                            = $MailSendCount
        "Mails reçus"                              = $MailReceiveCount
        "Mails lus"                                = $MailReadCount

        "Usage OneDrive"                           = $HasOneDriveUsage
        "Dernière activité OneDrive"               = $OneDriveLastActivity
        "Fichiers OneDrive consultés/modifiés"     = $OneDriveViewedEdited
        "Fichiers OneDrive synchronisés"           = $OneDriveSynced
        "Partages OneDrive internes"               = $OneDriveSharedInternal
        "Partages OneDrive externes"               = $OneDriveSharedExternal

        "Usage SharePoint"                         = $HasSharePointUsage
        "Dernière activité SharePoint"             = $SharePointLastActivity
        "Fichiers SharePoint consultés/modifiés"   = $SharePointViewedEdited
        "Pages SharePoint visitées"                = $SharePointVisitedPages
        "Partages SharePoint internes"             = $SharePointSharedInternal
        "Partages SharePoint externes"             = $SharePointSharedExternal

        "Usage Teams"                              = $HasTeamsUsage
        "Dernière activité Teams"                  = $TeamsLastActivity
        "Messages Teams canal"                     = $TeamsTeamChat
        "Messages Teams privés"                    = $TeamsPrivateChat
        "Appels Teams"                             = $TeamsCalls
        "Réunions Teams"                           = $TeamsMeetings

        "Usage Microsoft 365 Apps"                 = $HasAppsUsage
        "Dernière activation Apps"                 = $M365AppsLastActivation
        "Dernière activité Apps"                   = $M365AppsLastActivity
        "Apps Windows"                             = $AppsWindows
        "Apps Mac"                                 = $AppsMac
        "Apps Mobile"                              = $AppsMobile
        "Apps Web"                                 = $AppsWeb
        "Outlook utilisé"                          = $AppsOutlook
        "Word utilisé"                             = $AppsWord
        "Excel utilisé"                            = $AppsExcel
        "PowerPoint utilisé"                       = $AppsPowerPoint
        "Teams App utilisé"                        = $AppsTeams

        "Recommandation"                           = $Recommendation
    }
}

# =========================
# Exports
# =========================

$SortedResults = $Results |
    Sort-Object "Famille de licence", "Licence", "Usage détecté", "Dernière activité détectée", "UPN"

Export-CsvFrench -InputObject $SortedResults -Path $DetailOutput -Delimiter $CsvDelimiter

$Summary = $Results |
    Group-Object "Licence" |
    ForEach-Object {
        $Group = $_.Group

        [pscustomobject]@{
            "Licence"                                  = $_.Name
            "Famille de licence"                       = (($Group."Famille de licence" | Select-Object -Unique) -join ";")
            "Teams inclus dans le bundle"              = (($Group."Teams inclus dans le bundle" | Select-Object -Unique) -join ";")
            "Nombre d'utilisateurs"                    = $Group.Count
            "Utilisateurs avec usage détecté"          = @($Group | Where-Object { $_."Usage détecté" -eq "Oui" }).Count
            "Utilisateurs sans usage détecté"          = @($Group | Where-Object { $_."Usage détecté" -eq "Non" }).Count
            "Utilisateurs avec usage cœur M365"        = @($Group | Where-Object { $_."Usage cœur Microsoft 365" -eq "Oui" }).Count
            "Utilisateurs avec Exchange"               = @($Group | Where-Object { $_."Usage Exchange" -eq "Oui" }).Count
            "Utilisateurs avec OneDrive"               = @($Group | Where-Object { $_."Usage OneDrive" -eq "Oui" }).Count
            "Utilisateurs avec SharePoint"             = @($Group | Where-Object { $_."Usage SharePoint" -eq "Oui" }).Count
            "Utilisateurs avec Teams"                  = @($Group | Where-Object { $_."Usage Teams" -eq "Oui" }).Count
            "Utilisateurs avec Apps M365"              = @($Group | Where-Object { $_."Usage Microsoft 365 Apps" -eq "Oui" }).Count
            "Candidats optimisation E1 vers F3"        = @($Group | Where-Object { $_."Candidat optimisation E1 vers F3" -eq "Oui" }).Count
            "Candidats usage faible"                   = @($Group | Where-Object { $_."Recommandation" -like "Usage faible*" }).Count
            "Candidats sans usage"                     = @($Group | Where-Object { $_."Recommandation" -like "Aucun usage*" }).Count
            "Teams seul non compté comme usage bundle" = @($Group | Where-Object { $_."Recommandation" -like "Usage Teams détecté uniquement*" }).Count
        }
    }

Export-CsvFrench -InputObject $Summary -Path $SummaryOutput -Delimiter $CsvDelimiter

Write-Host ""
Write-Host "Exports générés :" -ForegroundColor Green
Write-Host "- Détail utilisateur : $DetailOutput"
Write-Host "- Synthèse licences : $SummaryOutput"
Write-Host "- SKU audités : $SkuOutput"
Write-Host "- Rapports bruts : $RawFolder"
Write-Host "- Référence Microsoft : $ReferenceFolder"
Write-Host ""

Write-Host "Lecture rapide :" -ForegroundColor Yellow
Write-Host "- Le fichier de détail permet d'analyser utilisateur par utilisateur."
Write-Host "- Le fichier de synthèse permet de comparer les usages Office 365 E1 et Microsoft 365 E3, F3, E5, ainsi que leurs variantes sans Teams."
Write-Host "- Le script valide les SKU avec la référence Microsoft quand le CSV est disponible."
Write-Host "- Le matching est réalisé par skuPartNumber et par skuId."
Write-Host "- Pour les bundles sans Teams, l'usage Teams est visible mais non compté comme usage cœur du bundle M365."
Write-Host "- Les profils E1 orientés web/mobile sans usage Apps desktop sont marqués comme candidats E1 vers F3."
Write-Host "- Un utilisateur sans usage détecté doit être validé métier avant retrait de licence."
Write-Host "- Si les utilisateurs apparaissent anonymisés, vérifier le paramètre de confidentialité des rapports Microsoft 365."
Write-Host ""

Disconnect-MgGraph | Out-Null