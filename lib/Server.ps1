# Server.ps1 -- HttpListener server with route dispatch and multipart parsing.
# Exports: Send-JsonResponse, Send-StaticFile, Parse-MultipartBoundary,
#          Read-MultipartStream, Find-ByteSequence, Handle-Process, Start-Server

function Send-JsonResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [string]$Body,
        [int]$StatusCode = 200
    )
    $Response.ContentType = 'application/json; charset=utf-8'
    $Response.StatusCode  = $StatusCode
    $Response.Headers.Add('Access-Control-Allow-Origin', '*')
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Response.ContentLength64 = $buffer.Length
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $Response.OutputStream.Close()
}

function Send-StaticFile {
    param(
        $Context,
        [string]$ProjectRoot
    )
    $request  = $Context.Request
    $response = $Context.Response
    $urlPath  = $request.Url.LocalPath

    if ($urlPath -eq '/' -or [string]::IsNullOrEmpty($urlPath)) {
        $urlPath = '/index.html'
    }

    # Prevent path traversal
    $relativePath = $urlPath.TrimStart('/') -replace '/', '\'
    $filePath = Join-Path $ProjectRoot "web\$relativePath"
    $webRoot  = Join-Path $ProjectRoot 'web'
    if (-not $filePath.StartsWith($webRoot)) {
        $response.StatusCode = 403
        $buf = [System.Text.Encoding]::UTF8.GetBytes('403 Forbidden')
        $response.ContentLength64 = $buf.Length
        $response.OutputStream.Write($buf, 0, $buf.Length)
        $response.OutputStream.Close()
        return
    }

    if (-not (Test-Path $filePath)) {
        $response.StatusCode = 404
        $buf = [System.Text.Encoding]::UTF8.GetBytes('404 Not Found')
        $response.ContentLength64 = $buf.Length
        $response.OutputStream.Write($buf, 0, $buf.Length)
        $response.OutputStream.Close()
        return
    }

    $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
    $response.ContentType = switch ($ext) {
        '.html' { 'text/html; charset=utf-8' }
        '.js'   { 'application/javascript; charset=utf-8' }
        '.css'  { 'text/css; charset=utf-8' }
        '.json' { 'application/json; charset=utf-8' }
        '.png'  { 'image/png' }
        '.ico'  { 'image/x-icon' }
        '.svg'  { 'image/svg+xml' }
        default { 'application/octet-stream' }
    }
    $response.StatusCode = 200
    $bytes = [System.IO.File]::ReadAllBytes($filePath)
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
    $response.OutputStream.Close()
}

function Parse-MultipartBoundary {
    param([string]$ContentType)
    # Content-Type: multipart/form-data; boundary=----WebKitFormBoundary...
    if ($ContentType -match 'boundary=([^\s;]+)') {
        return $Matches[1].Trim('"')
    }
    return $null
}

function Find-ByteSequence {
    param(
        [byte[]]$Haystack,
        [byte[]]$Needle,
        [int]$StartAt = 0
    )
    $needleLen = $Needle.Length
    $limit     = $Haystack.Length - $needleLen
    for ($i = $StartAt; $i -le $limit; $i++) {
        $match = $true
        for ($j = 0; $j -lt $needleLen; $j++) {
            if ($Haystack[$i + $j] -ne $Needle[$j]) { $match = $false; break }
        }
        if ($match) { return $i }
    }
    return -1
}

function Read-MultipartStream {
    <#
    .SYNOPSIS
        Parses a multipart/form-data stream.
        Returns @{ PdfBytes = [byte[]]; CsvText = [string] }
        PDF is read as raw bytes to avoid StreamReader corruption of binary data.
        CSV is decoded as UTF-8 text.
    #>
    param(
        [System.Net.HttpListenerRequest]$Request
    )

    $boundary = Parse-MultipartBoundary -ContentType $Request.ContentType
    if (-not $boundary) {
        throw "No multipart boundary found in Content-Type header"
    }

    # Read entire request body as raw bytes (MUST NOT use StreamReader here)
    $ms = New-Object System.IO.MemoryStream
    $Request.InputStream.CopyTo($ms)
    $allBytes = $ms.ToArray()
    $ms.Dispose()

    # Boundary delimiters as bytes
    # Note: [System.Text.Encoding]::Latin1 is null in PowerShell 5.1; use GetEncoding(28591) instead.
    $enc           = [System.Text.Encoding]::GetEncoding(28591)  # ISO-8859-1 / Latin1 — byte-transparent (0x00-0xFF = same index)
    $boundaryBytes = $enc.GetBytes("--$boundary")
    $crlf          = [byte[]](13, 10)  # \r\n

    # Split allBytes on boundary markers
    # Returns list of byte[] segments (each is one part including its headers)
    $parts = [System.Collections.Generic.List[byte[]]]::new()
    $start = 0
    while ($true) {
        $idx = Find-ByteSequence -Haystack $allBytes -Needle $boundaryBytes -StartAt $start
        if ($idx -lt 0) { break }
        if ($start -gt 0) {
            # The segment before this boundary (trim leading \r\n)
            $segStart = $start
            $segEnd   = $idx - 2  # strip trailing \r\n before boundary
            if ($segEnd -gt $segStart) {
                $seg = $allBytes[$segStart..($segEnd - 1)]
                $parts.Add($seg)
            }
        }
        $start = $idx + $boundaryBytes.Length + 2  # skip boundary + \r\n
    }

    $result = @{ PdfBytes = $null; PaperCutCsvText = $null; CsvText = $null }

    foreach ($part in $parts) {
        # Find the double-CRLF that separates headers from body
        $doubleCrlf   = [byte[]](13, 10, 13, 10)
        $headerEndIdx = Find-ByteSequence -Haystack $part -Needle $doubleCrlf -StartAt 0
        if ($headerEndIdx -lt 0) { continue }

        $headerBytes = $part[0..($headerEndIdx - 1)]
        $headerText  = [System.Text.Encoding]::UTF8.GetString($headerBytes)
        $bodyStart   = $headerEndIdx + 4
        $bodyBytes   = if ($bodyStart -lt $part.Length) { $part[$bodyStart..($part.Length - 1)] } else { [byte[]]@() }

        if ($headerText -match 'name="pdf"') {
            # Detect by filename extension whether user uploaded a CSV instead of a PDF
            $isCsv = $headerText -match 'filename="[^"]*\.csv"'
            if ($isCsv) {
                $result.PaperCutCsvText = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
            } else {
                # Binary PDF — keep as bytes, never decode through string
                $result.PdfBytes = $bodyBytes
            }
        }
        elseif ($headerText -match 'name="csv"') {
            # CSV is text — decode as UTF-8
            $result.CsvText = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
        }
    }

    return $result
}

function Handle-Process {
    param(
        $Context,
        [string]$ProjectRoot
    )
    $request  = $Context.Request
    $response = $Context.Response

    if ($request.HttpMethod -ne 'POST') {
        Send-JsonResponse -Response $response -Body (@{ error = 'Method not allowed' } | ConvertTo-Json) -StatusCode 405
        return
    }

    try {
        $parts = Read-MultipartStream -Request $request

        $hasPdf = ($null -ne $parts.PdfBytes -and $parts.PdfBytes.Length -gt 0)
        $hasPaperCutCsv = -not [string]::IsNullOrWhiteSpace($parts.PaperCutCsvText)

        if (-not $hasPdf -and -not $hasPaperCutCsv) {
            Send-JsonResponse -Response $response -Body (@{ error = 'No PaperCut report received (upload a PDF or CSV)' } | ConvertTo-Json) -StatusCode 400
            return
        }
        if ([string]::IsNullOrWhiteSpace($parts.CsvText)) {
            Send-JsonResponse -Response $response -Body (@{ error = 'No user directory CSV received' } | ConvertTo-Json) -StatusCode 400
            return
        }

        # Step 1: Parse PaperCut report — PDF or CSV
        $pdfResult = $null
        if ($hasPaperCutCsv) {
            $pdfResult = Invoke-PaperCutCsvParser -CsvText $parts.PaperCutCsvText
        } else {
            # Write PDF bytes to temp file (binary-safe — NOT WriteAllText)
            $tempDir     = [System.IO.Path]::GetTempPath()
            $tempPdfPath = Join-Path $tempDir "papayaflow_$([System.IO.Path]::GetRandomFileName()).pdf"
            [System.IO.File]::WriteAllBytes($tempPdfPath, $parts.PdfBytes)
            try {
                $pdfResult = Invoke-PdfParser -PdfPath $tempPdfPath -ProjectRoot $ProjectRoot
            } finally {
                if (Test-Path $tempPdfPath) { Remove-Item $tempPdfPath -Force -ErrorAction SilentlyContinue }
            }
        }
        if (-not $pdfResult.Success) {
            Send-JsonResponse -Response $response -Body (@{ error = $pdfResult.Error } | ConvertTo-Json) -StatusCode 422
            return
        }

        # Step 2: Parse Intune CSV
        $entraResult = Invoke-EntraMapper -CsvText $parts.CsvText
        if (-not $entraResult.Success) {
            Send-JsonResponse -Response $response -Body (@{ error = $entraResult.Error } | ConvertTo-Json) -StatusCode 422
            return
        }

        # Step 3: Aggregate
        $aggregated = Invoke-Aggregator `
            -Users         $pdfResult.Users `
            -DepartmentMap $entraResult.Map `
            -DateRange     $pdfResult.DateRange

        # Step 4: Serialize to JSON and return
        # ConvertTo-Json with depth 5 handles nested user arrays
        $jsonBody = $aggregated | ConvertTo-Json -Depth 5
        Send-JsonResponse -Response $response -Body $jsonBody
    }
    catch {
        Send-JsonResponse -Response $response -Body (@{ error = $_.Exception.Message } | ConvertTo-Json) -StatusCode 500
    }
}

function Start-Server {
    param(
        [int]$Port = 8080,
        [string]$ProjectRoot
    )

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$Port/")
    $listener.Start()
    Write-Host "PapayaFlow running at http://localhost:$Port  (Ctrl+C to stop)"

    try {
        while ($listener.IsListening) {
            $context = $listener.GetContext()
            try {
                $path = $context.Request.Url.LocalPath
                switch -Exact ($path) {
                    '/process'  { Handle-Process  $context $ProjectRoot }
                    '/shutdown' {
                        $buf = [System.Text.Encoding]::UTF8.GetBytes('OK')
                        $context.Response.ContentLength64 = $buf.Length
                        $context.Response.OutputStream.Write($buf, 0, $buf.Length)
                        $context.Response.OutputStream.Close()
                        $listener.Stop()
                    }
                    default     { Send-StaticFile $context $ProjectRoot }
                }
            }
            catch {
                try {
                    Send-JsonResponse -Response $context.Response -Body (@{ error = $_.Exception.Message } | ConvertTo-Json) -StatusCode 500
                } catch { }
            }
        }
    }
    finally {
        $listener.Stop()
        $listener.Close()
    }
}
