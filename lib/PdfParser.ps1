# PdfParser.ps1 -- Extracts structured data from a PaperCut Hive "User activity summary" PDF.
# Exports: Invoke-PdfParser

function Invoke-PdfParser {
    <#
    .SYNOPSIS
        Converts a PaperCut Hive PDF to structured data using pdftotext.
    .PARAMETER PdfPath
        Absolute path to the PDF file on disk.
    .PARAMETER ProjectRoot
        Root directory of the PapayaFlow project (used to locate bin\pdftotext.exe).
    .OUTPUTS
        Hashtable with keys:
          Success  = $true/$false
          Users    = array of user row hashtables (when Success=$true)
          DateRange = @{ From = string; To = string } (when Success=$true)
          Error    = error message string (when Success=$false)
    #>
    param(
        [string]$PdfPath,
        [string]$ProjectRoot
    )

    # Locate pdftotext.exe
    $pdftotextExe = Join-Path $ProjectRoot 'bin\pdftotext.exe'
    if (-not (Test-Path $pdftotextExe)) {
        return @{
            Success = $false
            Error   = "pdftotext.exe not found at $pdftotextExe. Download Poppler for Windows and place pdftotext.exe in the bin\ directory."
        }
    }

    if (-not (Test-Path $PdfPath)) {
        return @{ Success = $false; Error = "PDF file not found: $PdfPath" }
    }

    # Create a temp path for the text output
    $tempTxtPath = [System.IO.Path]::ChangeExtension($PdfPath, '.txt')

    try {
        # Invoke pdftotext with -layout to preserve column alignment
        # -layout is critical: without it, columns merge and regex fails
        $proc = Start-Process -FilePath $pdftotextExe `
                              -ArgumentList '-layout', "`"$PdfPath`"", "`"$tempTxtPath`"" `
                              -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -ne 0) {
            return @{ Success = $false; Error = "pdftotext exited with code $($proc.ExitCode)" }
        }

        if (-not (Test-Path $tempTxtPath)) {
            return @{ Success = $false; Error = "pdftotext produced no output file" }
        }

        $rawText = Get-Content -Path $tempTxtPath -Raw -Encoding UTF8

        # Extract date range
        # PaperCut header looks like: "January 1, 2026 to March 26, 2026"
        $dateRange = $null
        if ($rawText -match '(\w+ \d{1,2},\s*\d{4})\s+to\s+(\w+ \d{1,2},\s*\d{4})') {
            $dateRange = @{
                From = $Matches[1].Trim()
                To   = $Matches[2].Trim()
            }
        }

        if (-not $dateRange) {
            return @{
                Success = $false
                Error   = "No date range found. This does not appear to be a PaperCut Hive User Activity Summary report."
            }
        }

        # Extract per-user rows
        # PaperCut -layout output produces lines like:
        #   1  user1@example.org   1234  800  434  900  334  120  600  634  0  0  $12.34
        # Columns: Rank  UPN  Pages  Print  Copy  B&W  Color  Jobs  1-Sided  2-Sided  Scans  Fax  Cost
        #
        # Regex explanation:
        #   ^\s*          -- optional leading whitespace
        #   (\d+)         -- rank (integer)
        #   \s+           -- separator
        #   ([\w.%+-]+@[\w.-]+\.[a-zA-Z]{2,}) -- UPN (email format)
        #   \s+(\d[\d,]*) -- pages
        #   \s+(\d[\d,]*) -- print
        #   \s+(\d[\d,]*) -- copy
        #   \s+(\d[\d,]*) -- B&W
        #   \s+(\d[\d,]*) -- color
        #   \s+(\d[\d,]*) -- jobs
        #   \s+(\d[\d,]*) -- 1-sided
        #   \s+(\d[\d,]*) -- 2-sided
        #   \s+(\d[\d,]*) -- scans
        #   \s+(\d[\d,]*) -- fax
        #   \s+\$?([\d,.]+) -- cost (may have $ prefix)
        $rowPattern = '^\s*(\d+)\s+([\w.%+\-]+@[\w.\-]+\.[a-zA-Z]{2,})\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+\$?([\d,.]+)'

        $users = @()
        $seenUpns = @{}  # deduplicate: PaperCut sometimes repeats header rows mid-report

        foreach ($line in ($rawText -split "`n")) {
            if ($line -match $rowPattern) {
                $upn = $Matches[2].Trim().ToLower()
                if ($seenUpns.ContainsKey($upn)) { continue }  # skip duplicate rows
                $seenUpns[$upn] = $true

                $users += @{
                    Rank     = [int]$Matches[1]
                    UPN      = $upn
                    Pages    = Parse-PdfInt $Matches[3]
                    Print    = Parse-PdfInt $Matches[4]
                    Copy     = Parse-PdfInt $Matches[5]
                    BW       = Parse-PdfInt $Matches[6]
                    Color    = Parse-PdfInt $Matches[7]
                    Jobs     = Parse-PdfInt $Matches[8]
                    OneSided = Parse-PdfInt $Matches[9]
                    TwoSided = Parse-PdfInt $Matches[10]
                    Scans    = Parse-PdfInt $Matches[11]
                    Fax      = Parse-PdfInt $Matches[12]
                    Cost     = Parse-PdfDecimal $Matches[13]
                }
            }
        }

        if ($users.Count -eq 0) {
            return @{
                Success = $false
                Error   = "Date range found but no user rows extracted. PDF may be a different PaperCut report type, or pdftotext layout parsing failed."
            }
        }

        return @{
            Success   = $true
            Users     = $users
            DateRange = $dateRange
        }
    }
    catch {
        return @{ Success = $false; Error = "PDF parsing exception: $($_.Exception.Message)" }
    }
    finally {
        # Always clean up the temp text file
        if (Test-Path $tempTxtPath) { Remove-Item $tempTxtPath -Force -ErrorAction SilentlyContinue }
    }
}

function Invoke-PaperCutCsvParser {
    <#
    .SYNOPSIS
        Parses a PaperCut CSV export. Handles two formats automatically:
          - Transaction log (has 'Transaction type' column): aggregates rows per user
          - User activity summary: reads one row per user
    .PARAMETER CsvText
        Raw CSV text (not a file path).
    .OUTPUTS
        Hashtable: Success, Users, DateRange (may be null), Error
    #>
    param([string]$CsvText)

    if ([string]::IsNullOrWhiteSpace($CsvText)) {
        return @{ Success = $false; Error = 'PaperCut CSV is empty' }
    }

    try {
        # Strip UTF-8 BOM if present (causes first column name to appear as "﻿User" instead of "User")
        $CsvText = $CsvText.TrimStart([char]0xFEFF)

        $tempCsvPath = [System.IO.Path]::GetTempFileName() + '.csv'
        [System.IO.File]::WriteAllText($tempCsvPath, $CsvText, [System.Text.Encoding]::UTF8)
        $rows = @(Import-Csv -Path $tempCsvPath)
        Remove-Item $tempCsvPath -Force -ErrorAction SilentlyContinue

        if ($rows.Count -eq 0) {
            return @{ Success = $false; Error = 'PaperCut CSV contained no data rows' }
        }

        $props = $rows[0].PSObject.Properties.Name

        # Dispatch by format
        if ($props -contains 'Transaction type') {
            return Invoke-PaperCutTransactionLogParser -Rows $rows
        } elseif ($props -contains 'Print+Copy total pages' -or $props -contains 'Print job BW pages') {
            return Invoke-PaperCutHiveActivityParser -Rows $rows
        } else {
            return Invoke-PaperCutActivitySummaryParser -Rows $rows -Props $props
        }
    }
    catch {
        return @{ Success = $false; Error = "PaperCut CSV parse error: $($_.Exception.Message)" }
    }
}

function Invoke-PaperCutTransactionLogParser {
    <#
    .SYNOPSIS
        Aggregates a PaperCut transaction log (one row per job) into per-user totals.
        Columns used: 'User Email', 'Transaction type', 'Amount', 'Date'.
        Transaction types counted: 'Print job', 'Copy job' (and their refunds).
    #>
    param($Rows)

    $userMap = @{}
    $dates   = [System.Collections.Generic.List[datetime]]::new()

    foreach ($row in $Rows) {
        $txType = $row.'Transaction type'.Trim()
        # Only count print and copy activity
        if ($txType -notmatch '^(Print job|Copy job)') { continue }

        $email = $row.'User Email'.Trim().ToLower()
        if ([string]::IsNullOrEmpty($email)) { continue }

        if (-not $userMap.ContainsKey($email)) {
            $userMap[$email] = @{ UPN = $email; Print = 0; Copy = 0; Cost = 0.0 }
        }

        # Amount is negative for charges, positive for refunds — negate to get spend
        $amount = [double]($row.Amount -replace '[,$\s]', '')
        $userMap[$email].Cost += -$amount

        if ($txType -eq 'Print job') { $userMap[$email].Print++ }
        if ($txType -eq 'Copy job')  { $userMap[$email].Copy++  }

        # Collect dates for range calculation
        try {
            $d = [datetime]::Parse($row.Date.Trim())
            $dates.Add($d)
        } catch { }
    }

    if ($userMap.Count -eq 0) {
        return @{ Success = $false; Error = 'No print or copy transactions found in transaction log' }
    }

    # Build date range from min/max dates in the data
    $dateRange = $null
    if ($dates.Count -gt 0) {
        $sorted = $dates | Sort-Object
        $dateRange = @{
            From = $sorted[0].ToString('MMMM d, yyyy')
            To   = $sorted[$sorted.Count - 1].ToString('MMMM d, yyyy')
        }
    }

    $users = @()
    $rank  = 1
    foreach ($upn in ($userMap.Keys | Sort-Object)) {
        $u = $userMap[$upn]
        $users += @{
            Rank     = $rank++
            UPN      = $u.UPN
            Pages    = 0
            Print    = $u.Print
            Copy     = $u.Copy
            BW       = 0
            Color    = 0
            Jobs     = $u.Print + $u.Copy
            OneSided = 0
            TwoSided = 0
            Scans    = 0
            Fax      = 0
            Cost     = [Math]::Round($u.Cost, 2)
        }
    }

    return @{ Success = $true; Users = $users; DateRange = $dateRange }
}

function Invoke-PaperCutHiveActivityParser {
    <#
    .SYNOPSIS
        Parses the PaperCut Hive activity export (one row per user, split print/copy columns).
        Known columns: User, Print+Copy total pages, Print job BW pages, Print job color pages,
        Print job total pages, Copy job BW/color/total pages, Print/Copy jobs 1-sided/2-sided/total,
        Scan pages, Fax pages, Cost.
    #>
    param($Rows)

    $users = @()
    $rank  = 1
    foreach ($row in $Rows) {
        $upn = ($row.User -replace '\s', '').ToLower()
        if ([string]::IsNullOrEmpty($upn)) { continue }

        $users += @{
            Rank     = $rank++
            UPN      = $upn
            Pages    = Parse-PdfInt $row.'Print+Copy total pages'
            Print    = Parse-PdfInt $row.'Print job total pages'
            Copy     = Parse-PdfInt $row.'Copy job total pages'
            BW       = (Parse-PdfInt $row.'Print job BW pages') + (Parse-PdfInt $row.'Copy job BW pages')
            Color    = (Parse-PdfInt $row.'Print job color pages') + (Parse-PdfInt $row.'Copy job color pages')
            Jobs     = (Parse-PdfInt $row.'Print jobs total') + (Parse-PdfInt $row.'Copy jobs total')
            OneSided = (Parse-PdfInt $row.'Print jobs 1-sided') + (Parse-PdfInt $row.'Copy jobs 1-sided')
            TwoSided = (Parse-PdfInt $row.'Print jobs 2-sided') + (Parse-PdfInt $row.'Copy jobs 2-sided')
            Scans    = Parse-PdfInt $row.'Scan pages'
            Fax      = Parse-PdfInt $row.'Fax pages'
            Cost     = Parse-PdfDecimal ($row.Cost -replace '[$,]', '')
        }
    }

    if ($users.Count -eq 0) {
        return @{ Success = $false; Error = 'No user rows extracted from PaperCut Hive activity export' }
    }

    return @{ Success = $true; Users = $users; DateRange = $null }
}

function Invoke-PaperCutActivitySummaryParser {
    <#
    .SYNOPSIS
        Parses a PaperCut activity summary CSV (one row per user) using flexible column matching.
    #>
    param($Rows, [string[]]$Props)

    function Find-Col {
        param([string[]]$Patterns, [string[]]$P)
        foreach ($pat in $Patterns) {
            $col = $P | Where-Object { $_ -imatch $pat } | Select-Object -First 1
            if ($col) { return $col }
        }
        return $null
    }

    $upnProp = Find-Col @('^username$', '^user$', '^email') $Props
    if (-not $upnProp) {
        return @{ Success = $false; Error = "No user column found. Expected 'Username' or 'User'. Found: $($Props -join ', ')" }
    }

    $pagesProp    = Find-Col @('total.?pages', '^pages$')    $Props
    $printProp    = Find-Col @('^print.?pages', '^print$')   $Props
    $copyProp     = Find-Col @('^copy.?pages',  '^copy$')    $Props
    $bwProp       = Find-Col @('b.?w', 'grayscale', 'mono')  $Props
    $colorProp    = Find-Col @('^color.?pages', '^color$', 'colour') $Props
    $jobsProp     = Find-Col @('^total.?jobs$', '^jobs$')    $Props
    $oneSidedProp = Find-Col @('1.?sided', 'simplex')        $Props
    $twoSidedProp = Find-Col @('2.?sided', 'duplex')         $Props
    $scansProp    = Find-Col @('scan')                        $Props
    $faxProp      = Find-Col @('fax')                        $Props
    $costProp     = Find-Col @('^total.?cost$', '^cost$', '^amount$') $Props

    $users = @()
    $rank  = 1
    foreach ($row in $Rows) {
        $upn = ($row.$upnProp -replace '\s', '').ToLower()
        if ([string]::IsNullOrEmpty($upn)) { continue }

        $users += @{
            Rank     = $rank++
            UPN      = $upn
            Pages    = if ($pagesProp)    { Parse-PdfInt     $row.$pagesProp } else { 0 }
            Print    = if ($printProp)    { Parse-PdfInt     $row.$printProp } else { 0 }
            Copy     = if ($copyProp)     { Parse-PdfInt     $row.$copyProp  } else { 0 }
            BW       = if ($bwProp)       { Parse-PdfInt     $row.$bwProp    } else { 0 }
            Color    = if ($colorProp)    { Parse-PdfInt     $row.$colorProp } else { 0 }
            Jobs     = if ($jobsProp)     { Parse-PdfInt     $row.$jobsProp  } else { 0 }
            OneSided = if ($oneSidedProp) { Parse-PdfInt     $row.$oneSidedProp } else { 0 }
            TwoSided = if ($twoSidedProp) { Parse-PdfInt     $row.$twoSidedProp } else { 0 }
            Scans    = if ($scansProp)    { Parse-PdfInt     $row.$scansProp } else { 0 }
            Fax      = if ($faxProp)      { Parse-PdfInt     $row.$faxProp   } else { 0 }
            Cost     = if ($costProp)     { Parse-PdfDecimal ($row.$costProp -replace '[$,]', '') } else { 0.0 }
        }
    }

    if ($users.Count -eq 0) {
        return @{ Success = $false; Error = 'No user rows extracted from PaperCut CSV' }
    }

    return @{ Success = $true; Users = $users; DateRange = $null }
}

function Parse-PdfInt {
    param([string]$s)
    # Remove commas (e.g. "1,234" -> 1234) and parse
    [int]($s -replace ',', '')
}

function Parse-PdfDecimal {
    param([string]$s)
    # Remove commas and parse as double
    [double]($s -replace ',', '')
}
