# Aggregator.ps1 -- Aggregates per-user print data by department.
# Exports: Invoke-Aggregator

function Invoke-Aggregator {
    <#
    .SYNOPSIS
        Groups PdfParser user rows by department (using EntraMapper map) and computes totals.
    .PARAMETER Users
        Array of user hashtables from Invoke-PdfParser (Users field).
    .PARAMETER DepartmentMap
        Hashtable of lowercase-UPN -> department from Invoke-EntraMapper (Map field).
    .PARAMETER DateRange
        Hashtable @{ From = string; To = string } from Invoke-PdfParser.
    .OUTPUTS
        Hashtable with keys:
          Departments = array of department summary hashtables
          OrgTotals   = org-wide aggregate hashtable
          DateRange   = @{ From; To }

        Each department summary:
          Name        = string
          Users       = array of per-user row hashtables (same fields as input, plus Department)
          ActiveUsers = count of users with Pages > 0
          TotalPages  = int
          TotalPrint  = int
          TotalCopy   = int
          TotalBW     = int
          TotalColor  = int
          TotalJobs   = int
          TotalCost   = double
          PctBW       = double (0-100, rounded to 1 decimal)
          PctColor    = double (0-100, rounded to 1 decimal)

        OrgTotals has the same numeric fields (no Name or Users).
    #>
    param(
        [array]$Users,
        [hashtable]$DepartmentMap,
        [hashtable]$DateRange
    )

    # Group users by department
    $groups = @{}  # department name -> list of user rows

    foreach ($user in $Users) {
        $upn  = $user.UPN.ToLower()
        $dept = if ($DepartmentMap.ContainsKey($upn)) { $DepartmentMap[$upn] } else { 'Unassigned' }

        $userWithDept = $user.Clone()
        $userWithDept['Department'] = $dept

        if (-not $groups.ContainsKey($dept)) {
            $groups[$dept] = [System.Collections.Generic.List[hashtable]]::new()
        }
        $groups[$dept].Add($userWithDept)
    }

    # Build per-department summaries
    $departments = @()
    foreach ($deptName in ($groups.Keys | Sort-Object)) {
        $deptUsers = $groups[$deptName]
        $summary   = Compute-Summary -DeptName $deptName -DeptUsers $deptUsers
        $departments += $summary
    }

    # Build org-wide totals
    # Note: Measure-Object -Property does NOT work on hashtable arrays in PS 5.1 (only PSCustomObject).
    # Sum manually instead.
    $orgPages = 0; $orgPrint2 = 0; $orgCopy = 0; $orgBW2 = 0; $orgColor2 = 0
    $orgJobs  = 0; $orgCost2  = 0.0; $orgActive = 0
    foreach ($d in $departments) {
        $orgPages   += $d.TotalPages
        $orgPrint2  += $d.TotalPrint
        $orgCopy    += $d.TotalCopy
        $orgBW2     += $d.TotalBW
        $orgColor2  += $d.TotalColor
        $orgJobs    += $d.TotalJobs
        $orgCost2   += $d.TotalCost
        $orgActive  += $d.ActiveUsers
    }
    $orgTotals = @{
        TotalPages  = $orgPages
        TotalPrint  = $orgPrint2
        TotalCopy   = $orgCopy
        TotalBW     = $orgBW2
        TotalColor  = $orgColor2
        TotalJobs   = $orgJobs
        TotalCost   = [Math]::Round($orgCost2, 2)
        ActiveUsers = $orgActive
    }
    $orgBW    = $orgTotals.TotalBW
    $orgColor = $orgTotals.TotalColor
    $orgPrint = $orgBW + $orgColor
    $orgTotals['PctBW']    = if ($orgPrint -gt 0) { [Math]::Round(($orgBW    / $orgPrint) * 100, 1) } else { 0.0 }
    $orgTotals['PctColor'] = if ($orgPrint -gt 0) { [Math]::Round(($orgColor / $orgPrint) * 100, 1) } else { 0.0 }

    return @{
        Departments = $departments
        OrgTotals   = $orgTotals
        DateRange   = $DateRange
    }
}

function Compute-Summary {
    param(
        [string]$DeptName,
        [System.Collections.Generic.List[hashtable]]$DeptUsers
    )

    $totalPages  = 0; $totalPrint = 0; $totalCopy = 0
    $totalBW     = 0; $totalColor = 0; $totalJobs  = 0
    $totalCost   = 0.0; $activeUsers = 0
    $totalOneSided = 0; $totalTwoSided = 0; $totalScans = 0; $totalFax = 0

    foreach ($u in $DeptUsers) {
        $totalPages  += $u.Pages
        $totalPrint  += $u.Print
        $totalCopy   += $u.Copy
        $totalBW     += $u.BW
        $totalColor  += $u.Color
        $totalJobs   += $u.Jobs
        $totalCost   += $u.Cost
        if ($u.Pages -gt 0) { $activeUsers++ }
        $totalOneSided += $u.OneSided
        $totalTwoSided += $u.TwoSided
        $totalScans    += $u.Scans
        $totalFax      += $u.Fax
    }

    $printTotal = $totalBW + $totalColor
    $pctBW    = if ($printTotal -gt 0) { [Math]::Round(($totalBW    / $printTotal) * 100, 1) } else { 0.0 }
    $pctColor = if ($printTotal -gt 0) { [Math]::Round(($totalColor / $printTotal) * 100, 1) } else { 0.0 }

    return @{
        Name        = $DeptName
        Users       = @($DeptUsers)
        ActiveUsers = $activeUsers
        TotalPages  = $totalPages
        TotalPrint  = $totalPrint
        TotalCopy   = $totalCopy
        TotalBW     = $totalBW
        TotalColor  = $totalColor
        TotalJobs   = $totalJobs
        TotalCost   = [Math]::Round($totalCost, 2)
        PctBW       = $pctBW
        PctColor    = $pctColor
        TotalOneSided = $totalOneSided
        TotalTwoSided = $totalTwoSided
        TotalScans    = $totalScans
        TotalFax      = $totalFax
    }
}
