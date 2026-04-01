# EntraMapper.ps1 -- Parses Entra ID CSV export and builds UPN -> Department map.
# Exports: Invoke-EntraMapper

function Invoke-EntraMapper {
    <#
    .SYNOPSIS
        Parses an Entra ID CSV export (as a string) and returns a UPN-to-Department hashtable.
    .PARAMETER CsvText
        The raw CSV text content (not a file path).
    .OUTPUTS
        Hashtable with keys:
          Success    = $true/$false
          Map        = hashtable of lowercase-UPN -> department-string (when Success=$true)
          UserCount  = int (when Success=$true)
          Error      = error message string (when Success=$false)
    .NOTES
        - Column name lookup is case-insensitive and matches any column whose name contains "department".
          Checked names include: "Department", "department", "Job Department", "JobDepartment".
        - UPN lookup is case-insensitive (all keys stored lowercase).
        - Users with blank/null Department are mapped to "Unassigned".
    #>
    param(
        [string]$CsvText
    )

    if ([string]::IsNullOrWhiteSpace($CsvText)) {
        return @{ Success = $false; Error = 'CSV text is empty' }
    }

    try {
        # Write to a temp file so Import-Csv can handle it (Import-Csv needs a file path in PS 5.1)
        $tempCsvPath = [System.IO.Path]::GetTempFileName() + '.csv'
        [System.IO.File]::WriteAllText($tempCsvPath, $CsvText, [System.Text.Encoding]::UTF8)

        # Wrap in @() to force array: Import-Csv returns PSCustomObject (not array) for single-row CSVs in PS 5.1
        $rows = @(Import-Csv -Path $tempCsvPath)
        Remove-Item $tempCsvPath -Force -ErrorAction SilentlyContinue

        if ($rows.Count -eq 0) {
            return @{ Success = $false; Error = 'CSV contained no data rows' }
        }

        # Find the UPN column (case-insensitive; must contain "userprincipalname" or "upn")
        $allProps = $rows[0].PSObject.Properties.Name
        $upnProp  = $allProps | Where-Object { $_ -match '^(userprincipalname|upn)$' } | Select-Object -First 1
        if (-not $upnProp) {
            # Fall back: any column containing "principal"
            $upnProp = $allProps | Where-Object { $_ -imatch 'principal' } | Select-Object -First 1
        }
        if (-not $upnProp) {
            return @{ Success = $false; Error = "No UPN column found. Expected 'UserPrincipalName' from an Intune export. Found: $($allProps -join ', ')" }
        }

        # Find the Department column (case-insensitive; any column whose name contains "department")
        $deptProp = $allProps | Where-Object { $_ -imatch 'department' } | Select-Object -First 1
        if (-not $deptProp) {
            return @{ Success = $false; Error = "No Department column found. Expected a column containing 'department'. Found: $($allProps -join ', ')" }
        }

        # Build map: lowercase UPN -> department name (trimmed, original casing preserved for display)
        $map = @{}
        foreach ($row in $rows) {
            $upn  = ($row.$upnProp  -replace '\s', '').ToLower()
            # Guard against $null: PowerShell returns $null for empty CSV fields when the trailing comma is absent
            $dept = if ($null -ne $row.$deptProp) { $row.$deptProp.Trim() } else { '' }
            if ([string]::IsNullOrEmpty($upn)) { continue }
            if ([string]::IsNullOrEmpty($dept)) { $dept = 'Unassigned' }
            $map[$upn] = $dept
        }

        return @{
            Success   = $true
            Map       = $map
            UserCount = $map.Count
        }
    }
    catch {
        return @{ Success = $false; Error = "Entra CSV parse error: $($_.Exception.Message)" }
    }
}
