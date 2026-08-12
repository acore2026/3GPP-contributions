---
name: 3gpp-review
description: Use when the user asks to read, summarize, or analyze 3GPP contribution documents (TDocs) for a Key Issue (KI). Handles downloading the TDoc list Excel from 3GPP by meeting number (e.g., SA2#175-AH-e), downloading individual .docx proposals, extracting text, reading each proposal, and generating a structured markdown summary organized by Solution Variant. Triggers on phrases like "read TDocs", "summarize KI proposals", "3GPP contribution summary", "download and analyze SA2 documents", "download meeting documents", "总结提案", "分析KI".
---

# 3GPP TDocs Review & Summarizer

自动化下载、解析、分析 3GPP 工作组贡献文档（TDocs），针对指定 Key Issue (KI) 生成结构化中文总结。支持流程图提取（Visio XML / VML / EMF→PNG 视觉识别）。

**IMPORTANT**: Follow each step EXACTLY as written. Do NOT skip steps. Do NOT guess URLs or folder names. Execute every PowerShell command as-is.

---

## WORKFLOW OVERVIEW (7 steps, execute in order)

```
Step 0: Initialize cache & environment
Step 1: Find the meeting folder on 3GPP website
Step 2: Download the TDoc Index Excel
Step 3: Parse Excel, filter KI proposals, download .zip files
Step 4: Extract text + diagrams from .docx (paragraphs, tables, Visio, VML, EMF→PNG)
Step 5: Read proposals in parallel batches (use task tool, with visual model for figures)
Step 6: Verify coverage & generate structured Markdown summary (with flowcharts)
```

---

## STEP 0: Initialize Cache & Environment

### 0.1 Set up cache directory structure

All downloads and intermediate files use a persistent cache so re-runs skip already-downloaded files.

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
$cacheRoot = "C:\Users\Administrator\Downloads\3gpp_cache"
New-Item -ItemType Directory -Path $cacheRoot -Force | Out-Null

# Sub-directories will be created dynamically:
#   $cacheRoot\{meetingFolder}\index\       - Index Excel
#   $cacheRoot\{meetingFolder}\docs\        - Downloaded .zip/.docx
#   $cacheRoot\{meetingFolder}\texts\       - Extracted .txt files
#   $cacheRoot\{meetingFolder}\figures\     - Converted PNG images from EMF
#   C:\Users\Administrator\Desktop\OpenCodeDiscussion\  - Final output
# === END OF BLOCK ===
```

### 0.2 Ensure ImportExcel module is available

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
    Install-Module ImportExcel -Force -Scope CurrentUser
    Write-Output "ImportExcel installed."
} else {
    Write-Output "ImportExcel already available."
}
# === END OF BLOCK ===
```

---

## STEP 1: Find the Meeting Folder

### What you need from the user

The user provides a meeting identifier and optionally a KI number. Parse them:

| User says | WG | Number | Type | KI |
|-----------|-----|--------|------|----|
| `SA2#175-AH-e KI#22` | SA2 | 175 | AH-e | 22 |
| `SA2#173 KI#19` | SA2 | 173 | (unknown) | 19 |
| `174` | SA2 (default) | 174 | (unknown) | (ask user) |

### Supported Working Groups

| WG | FTP Path |
|----|----------|
| SA2 | `tsg_sa/WG2_Arch` |
| SA1 | `tsg_sa/WG1_Serv` |
| SA3 | `tsg_sa/WG3_Security` |
| SA5 | `tsg_sa/WG5_OAM` |
| CT1 | `tsg_ct/WG1_NAS` |
| CT4 | `tsg_ct/WG4_MAP` |

**CRITICAL**: SA2 FTP path is `WG2_Arch` NOT `WG2_Architecture`.

### How to find the exact folder name

**You MUST use the `webfetch` tool to browse the directory listing.** Folder names include city and date and are NOT predictable.

```
Use webfetch tool with:
  url: "https://www.3gpp.org/ftp/{wgPath}/"
  format: "text"
```

Then search the response text for the meeting number.

**Known folder name examples** (for reference only, always verify):

```
SA2#173 → TSGS2_173_Goa_2026-02
SA2#174 → TSGS2_174_Malta_2026-04
SA2#175 → TSGS2_175_Dalian_2026-05
SA2#175-AH-e → TSGS2_175-AH-e_Electronic_2026-06
```

### What to record

After finding the folder, write down these values for use in later steps:

```
$wgPath = "tsg_sa/WG2_Arch"
$meetingFolder = "TSGS2_173_Goa_2026-02"   ← replace with actual
$meetingNum = "173"                          ← extract number
$meetingLocation = "Goa"                     ← extract city
$meetingDate = "2026-02"                     ← extract date
$wgPrefix = "S2"                             ← document number prefix (S1/S2/S3/...)
```

### If webfetch fails or listing is truncated

If the directory listing is very long and appears truncated, use `grep` on the output to search for the meeting number. If webfetch fails entirely:

> "3GPP FTP 目录无法访问。请手动访问 https://www.3gpp.org/ftp/{wgPath}/ 找到对应会议的文件夹名，然后告诉我。"

---

## STEP 2: Download the TDoc Index Excel

### 2.1 Browse the meeting folder to find the Index zip

```
Use webfetch tool with:
  url: "https://www.3gpp.org/ftp/{wgPath}/{meetingFolder}/"
  format: "text"
```

Look for a file matching pattern: `SA2-{number}_Index_{year}.zip`

### 2.2 Download and extract (with cache check)

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
$indexUrl = "https://www.3gpp.org/ftp/{wgPath}/{meetingFolder}/{indexZipName}"  # ← replace
$indexDir = Join-Path $cacheRoot "$meetingFolder\index"
$zipOut = Join-Path $indexDir $indexZipName

# Skip download if already cached
if (Test-Path -LiteralPath $zipOut) {
    $fileSize = (Get-Item -LiteralPath $zipOut).Length
    if ($fileSize -gt 1000) {
        Write-Output "Index zip already cached ($fileSize bytes). Skipping download."
    } else {
        Remove-Item -LiteralPath $zipOut -Force
    }
}

if (-not (Test-Path -LiteralPath $zipOut)) {
    New-Item -ItemType Directory -Path $indexDir -Force | Out-Null
    Invoke-WebRequest -Uri $indexUrl -OutFile $zipOut -TimeoutSec 120
    $fileSize = (Get-Item -LiteralPath $zipOut).Length
    if ($fileSize -lt 1000) {
        Write-Output "ERROR: Downloaded file is too small ($fileSize bytes). Likely a 404 page."
        Remove-Item -LiteralPath $zipOut -Force
    } else {
        Write-Output "Downloaded: $fileSize bytes"
    }
}

$extractPath = Join-Path $indexDir "extracted"
if (Test-Path -LiteralPath $zipOut) {
    Expand-Archive -LiteralPath $zipOut -DestinationPath $extractPath -Force
    Get-ChildItem -LiteralPath $extractPath -Recurse | Select-Object FullName, Length
}
# === END OF BLOCK ===
```

### 2.3 Verify the Excel file

After extraction, find the `.xlsx` file. Record its full path as `$excelPath`.

---

## STEP 3: Parse Excel, Filter KI Proposals, Download .zip Files

### 3.1 List Excel sheets and find the target

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
$sheets = Get-ExcelSheetInfo -Path $excelPath
Write-Output "=== Sheet Names ==="
$sheets | ForEach-Object { Write-Output "  $($_.Name) (Index: $($_.Index))" }
# === END OF BLOCK ===
```

Find the sheet matching the meeting (e.g., `SA2#173_Goa`). Record as `$sheetName`.

### 3.2 Read the sheet with dynamic column detection

**CRITICAL**: You MUST use `-NoHeader` parameter. The 3GPP Index Excel has duplicate column headers.

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
$allDocs = Import-Excel -Path $excelPath -WorksheetName $sheetName -NoHeader

Write-Output "Total rows: $($allDocs.Count)"

# Dynamically detect header row and column mapping
# Scan first 20 rows to find the one containing column labels
Write-Output "`n=== First 20 rows (for column detection) ==="
$allDocs | Select-Object -First 20 | ForEach-Object {
    $row = $_
    $vals = @()
    for ($i = 1; $i -le 15; $i++) {
        $propName = "P$i"
        $val = $row.$propName
        if ($val) { $vals += "P${i}=$val" }
    }
    Write-Output ($vals -join " | ")
}
# === END OF BLOCK ===
```

### 3.3 Identify column mapping from header row

From the output above, identify:

```
$colAgenda  = "P?"   # Column containing agenda item numbers (e.g., "20.6.22")
$colDocNum   = "P?"   # Column containing document numbers (e.g., "S2-2600098")
$colTitle    = "P?"   # Column containing proposal titles
$colSource   = "P?"   # Column containing source company names
$colStatus   = "P?"   # Column containing status (e.g., "Available")
$dataStartRow = ??     # Row number where actual data begins (after headers)
```

**Common mapping for SA2** (verify from output):

```
P1 = Status, P4 = Agenda Item, P5 = Doc Number, P8 = Title, P9 = Source
Data typically starts around row 15.
```

### 3.4 Find the agenda item for the target KI

Search the header rows for the KI description to find the agenda item number:

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
$kiNumber = "22"  # ← replace with target KI number

# Search header rows for KI-related text
Write-Output "=== Searching for KI#$kiNumber in headers ==="
$allDocs | Select-Object -First 15 | ForEach-Object {
    $row = $_
    for ($i = 1; $i -le 15; $i++) {
        $propName = "P$i"
        $val = $row.$propName
        if ($val -and ($val -match "KI" -or $val -match "Key Issue" -or $val -match $kiNumber)) {
            Write-Output "Found in P${i}: $val"
        }
    }
}

# Also list all unique agenda items from data rows
Write-Output "`n=== All unique agenda items ==="
$dataRows = $allDocs | Select-Object -Skip ($dataStartRow - 1)
$dataRows | Where-Object { $_.$colDocNum -match $wgPrefix } |
    Select-Object -ExpandProperty $colAgenda -Unique |
    Sort-Object |
    ForEach-Object { Write-Output "  $_" }
# === END OF BLOCK ===
```

### 3.5 Filter proposals by agenda item

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
$agendaItem = "20.6.22"  # ← replace with actual agenda item

$kiProposals = $dataRows | Where-Object {
    $_.$colAgenda -eq $agendaItem -and $_.$colDocNum -match $wgPrefix
}

Write-Output "Found $($kiProposals.Count) proposals for agenda item $agendaItem"
Write-Output "`n=== Proposal List ==="
$kiProposals | ForEach-Object {
    Write-Output "$($_.$colDocNum) | $($_.$colSource) | $($_.$colTitle)"
}

$docNumbers = @($kiProposals | ForEach-Object { $_.$colDocNum })
$docMeta = @{}
$kiProposals | ForEach-Object {
    $docMeta[$_.$colDocNum] = @{
        Source = $_.$colSource
        Title  = $_.$colTitle
    }
}
Write-Output "`nDocument numbers to download: $($docNumbers.Count)"
# === END OF BLOCK ===
```

### 3.6 Download proposal .zip files (parallel with retry)

**CRITICAL**: 3GPP stores proposals as `.zip` files, NOT `.docx`. Each `.zip` contains a `.docx` inside.

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
$docsBaseUrl = "https://www.3gpp.org/ftp/{wgPath}/{meetingFolder}/Docs"  # ← replace
$docsDir = Join-Path $cacheRoot "$meetingFolder\docs"
New-Item -ItemType Directory -Path $docsDir -Force | Out-Null

# Filter out already-downloaded docs (cache check)
$toDownload = @()
foreach ($doc in $docNumbers) {
    $docDir = Join-Path $docsDir $doc
    $docxFiles = @()
    if (Test-Path -LiteralPath $docDir) {
        $docxFiles = @(Get-ChildItem -LiteralPath $docDir -Filter "*.docx" -Recurse -ErrorAction SilentlyContinue)
    }
    if ($docxFiles.Count -eq 0) {
        $toDownload += $doc
    }
}

Write-Output "Already cached: $($docNumbers.Count - $toDownload.Count)"
Write-Output "To download: $($toDownload.Count)"

# Download with parallel jobs (batches of 5)
$batchSize = 5
$downloaded = 0; $failed = 0; $retries = 2

for ($i = 0; $i -lt $toDownload.Count; $i += $batchSize) {
    $batch = $toDownload[$i..([Math]::Min($i + $batchSize - 1, $toDownload.Count - 1))]
    $jobs = @()

    foreach ($doc in $batch) {
        $zipUrl = "$docsBaseUrl/$doc.zip"
        $docDir = Join-Path $docsDir $doc
        $outFile = Join-Path $docDir "$doc.zip"

        $jobs += Start-Job -ScriptBlock {
            param($url, $dir, $file, $maxRetries)
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
                try {
                    Invoke-WebRequest -Uri $url -OutFile $file -TimeoutSec 30 -ErrorAction Stop
                    $size = (Get-Item -LiteralPath $file).Length
                    if ($size -lt 500) {
                        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
                        throw "File too small ($size bytes)"
                    }
                    Expand-Archive -LiteralPath $file -DestinationPath $dir -Force
                    Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
                    return "OK"
                } catch {
                    if ($attempt -eq $maxRetries) { return "FAIL: $($_.Exception.Message)" }
                    Start-Sleep -Seconds 2
                }
            }
        } -ArgumentList $zipUrl, $docDir, $outFile, $retries
    }

    $jobs | Wait-Job -Timeout 60 | Out-Null
    foreach ($job in $jobs) {
        $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
        if ($result -eq "OK") { $downloaded++ } else { $failed++; Write-Output $result }
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
    Write-Output "Progress: $([Math]::Min($i + $batchSize, $toDownload.Count))/$($toDownload.Count)"
}

Write-Output "`nDone. Downloaded: $downloaded, Failed: $failed, Cached: $($docNumbers.Count - $toDownload.Count)"
# === END OF BLOCK ===
```

---

## STEP 4: Extract Text + Diagrams from .docx Files

### 4.1 Overview of extraction layers

This step extracts content from .docx files using a **3-layer diagram extraction strategy**:

| Layer | Source | Format | Output Marker |
|-------|--------|--------|---------------|
| Text | `word/document.xml` paragraphs + tables | Plain text | `[H1]`, `[TABLE START]` |
| Layer 1 | `word/embeddings/*.vsdx` (Visio) | XML shapes + arrows | `[VISIO FLOWCHART]` |
| Layer 2 | `document.xml` VML/WPS textboxes | Regex `<v:textbox>` | `[VML FLOWCHART]` |
| Layer 3 | `word/media/*.emf` → PNG | Image (visual model) | `[FIGURE: path]` |

**Diagram format statistics** (from analysis of 15 proposals across multiple companies/KIs):
- ~73% of proposals contain Visio `.vsdx` embeddings (Nokia, LG, OPPO, Ericsson, Qualcomm, InterDigital, Samsung)
- ~87% contain VML inline shapes (China Mobile, Huawei, Nokia)
- ~73% contain EMF preview images (always paired with Visio)
- ~7% use PNG screenshots (ZTE) — only Layer 3 can handle these
- ~7% are pure text with no diagrams

### 4.2 Full extraction script (text + diagrams)

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
Add-Type -AssemblyName System.Drawing

$docsDir = Join-Path $cacheRoot "$meetingFolder\docs"
$textDir = Join-Path $cacheRoot "$meetingFolder\texts"
$figDir = Join-Path $cacheRoot "$meetingFolder\figures"
New-Item -ItemType Directory -Path $textDir -Force | Out-Null
New-Item -ItemType Directory -Path $figDir -Force | Out-Null

$dirs = Get-ChildItem -LiteralPath $docsDir -Directory
$total = $dirs.Count; $current = 0

foreach ($d in $dirs) {
    $current++
    $textFile = Join-Path $textDir "$($d.Name).txt"

    # Skip if already extracted (check text file exists and is substantial)
    if ((Test-Path -LiteralPath $textFile) -and (Get-Item -LiteralPath $textFile).Length -gt 100) {
        Write-Output "SKIP (cached): $($d.Name)"
        continue
    }

    $docx = Get-ChildItem -LiteralPath $d.FullName -Filter "*.docx" -Recurse | Select-Object -First 1
    if (-not $docx) {
        Write-Output "NO DOCX: $($d.Name)"
        continue
    }

    $workDir = Join-Path $textDir "_work_$($d.Name)"
    $zipPath = Join-Path $textDir "$($d.Name).zip"
    $docFigDir = Join-Path $figDir $d.Name

    try {
        Copy-Item -LiteralPath $docx.FullName -Destination $zipPath -Force
        Expand-Archive -LiteralPath $zipPath -DestinationPath $workDir -Force

        $xmlPath = Join-Path $workDir "word\document.xml"
        if (-not (Test-Path -LiteralPath $xmlPath)) {
            Write-Output "NO XML: $($d.Name)"
            continue
        }

        [xml]$xml = Get-Content -LiteralPath $xmlPath -Encoding UTF8
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

        $text = ""
        $diagramStats = @{ Visio = 0; VML = 0; Figures = 0 }

        # ============================================================
        # PART A: Extract paragraphs and tables (same as before)
        # ============================================================
        $body = $xml.SelectSingleNode("//w:body", $ns)
        if (-not $body) { $body = $xml.SelectSingleNode("//w:document/w:body", $ns) }

        foreach ($child in $body.ChildNodes) {
            if ($child.LocalName -eq "p") {
                # === PARAGRAPH ===
                $pStyle = ""
                $pPr = $child.SelectSingleNode("w:pPr/w:pStyle/@w:val", $ns)
                if ($pPr) { $pStyle = $pPr.Value }

                $headingPrefix = ""
                if ($pStyle -match "^Heading(\d)$" -or $pStyle -match "^heading(\d)$" -or $pStyle -match "^(\d)$") {
                    $level = $Matches[1]
                    $headingPrefix = "[H$level] "
                } elseif ($pStyle -match "Title" -or $pStyle -match "title") {
                    $headingPrefix = "[H1] "
                }

                $line = ""
                $runs = $child.SelectNodes(".//w:r", $ns)
                foreach ($r in $runs) {
                    $parent = $r.ParentNode
                    if ($parent -and $parent.LocalName -eq "del") { continue }
                    $tNodes = $r.SelectNodes("w:t", $ns)
                    foreach ($t in $tNodes) { $line += $t.InnerText }
                }

                if ($line.Trim() -or $headingPrefix) {
                    $text += "$headingPrefix$line`n"
                }
            }
            elseif ($child.LocalName -eq "tbl") {
                # === TABLE ===
                $text += "[TABLE START]`n"
                $rows = $child.SelectNodes(".//w:tr", $ns)
                foreach ($row in $rows) {
                    $cells = $row.SelectNodes("w:tc", $ns)
                    $cellTexts = @()
                    foreach ($cell in $cells) {
                        $cellText = ""
                        $cellParas = $cell.SelectNodes(".//w:p", $ns)
                        foreach ($cp in $cellParas) {
                            $cRuns = $cp.SelectNodes(".//w:r", $ns)
                            foreach ($cr in $cRuns) {
                                $parent = $cr.ParentNode
                                if ($parent -and $parent.LocalName -eq "del") { continue }
                                $cTNodes = $cr.SelectNodes("w:t", $ns)
                                foreach ($ct in $cTNodes) { $cellText += $ct.InnerText }
                            }
                        }
                        $cellTexts += $cellText.Trim()
                    }
                    $text += ($cellTexts -join " | ") + "`n"
                }
                $text += "[TABLE END]`n"
            }
        }

        # ============================================================
        # PART B: Layer 1 - Visio .vsdx extraction (FIXED)
        # ============================================================
        # CRITICAL: In 3GPP Visio diagrams, arrows and text labels are SEPARATE shapes.
        # Arrows have BeginX/EndX/BeginY/EndY but NO text.
        # Text labels have text but NO arrow properties.
        # We use a 2-pass approach: collect arrows first, then classify text shapes
        # by distance to arrow midpoints (< 1.0 = label, >= 1.0 = entity).

        $embedDir = Join-Path $workDir "word\embeddings"
        if (Test-Path -LiteralPath $embedDir) {
            $vsdxFiles = @(Get-ChildItem -LiteralPath $embedDir -Filter "*.vsdx" -ErrorAction SilentlyContinue)
            foreach ($vsdx in $vsdxFiles) {
                $vsdxZip = Join-Path $workDir "_vsdx_temp.zip"
                $vsdxExtract = Join-Path $workDir "_vsdx_$($vsdx.BaseName)"

                try {
                    Copy-Item -LiteralPath $vsdx.FullName -Destination $vsdxZip -Force
                    Expand-Archive -LiteralPath $vsdxZip -DestinationPath $vsdxExtract -Force
                    Remove-Item -LiteralPath $vsdxZip -Force

                    $pageXml = Join-Path $vsdxExtract "visio\pages\page1.xml"
                    if (-not (Test-Path -LiteralPath $pageXml)) { continue }

                    [xml]$vxml = Get-Content -LiteralPath $pageXml -Encoding UTF8
                    $vns = New-Object System.Xml.XmlNamespaceManager($vxml.NameTable)
                    $vns.AddNamespace("v", "http://schemas.microsoft.com/office/visio/2012/main")

                    $shapes = $vxml.SelectNodes("//v:Shape", $vns)
                    if ($shapes.Count -eq 0) { continue }

                    $text += "`n[VISIO FLOWCHART: $($vsdx.Name)]`n"
                    $diagramStats.Visio++

                    # Use generic lists to avoid PowerShell array type issues
                    $entityList = New-Object System.Collections.Generic.List[PSObject]
                    $arrowList = New-Object System.Collections.Generic.List[PSObject]
                    $candidateList = New-Object System.Collections.Generic.List[PSObject]

                    # Pass 1: Collect arrows and candidate text shapes
                    foreach ($shape in $shapes) {
                        $sid = $shape.GetAttribute("ID")
                        $textNode = $shape.SelectSingleNode("v:Text", $vns)
                        $shapeText = if ($textNode) { $textNode.InnerText.Trim() } else { "" }

                        $endArrow = $shape.SelectSingleNode(".//v:Cell[@N='EndArrow']", $vns)
                        $beginArrow = $shape.SelectSingleNode(".//v:Cell[@N='BeginArrow']", $vns)
                        $eaVal = if ($endArrow) { $endArrow.GetAttribute("V") } else { "0" }
                        $baVal = if ($beginArrow) { $beginArrow.GetAttribute("V") } else { "0" }
                        $hasArrow = ($eaVal -ne "0") -or ($baVal -ne "0")

                        $beginXNode = $shape.SelectSingleNode(".//v:Cell[@N='BeginX']", $vns)
                        $beginYNode = $shape.SelectSingleNode(".//v:Cell[@N='BeginY']", $vns)
                        $endXNode = $shape.SelectSingleNode(".//v:Cell[@N='EndX']", $vns)
                        $endYNode = $shape.SelectSingleNode(".//v:Cell[@N='EndY']", $vns)
                        $hasLineCoords = $beginXNode -and $endXNode

                        $pinXNode = $shape.SelectSingleNode(".//v:Cell[@N='PinX']", $vns)
                        $pinYNode = $shape.SelectSingleNode(".//v:Cell[@N='PinY']", $vns)
                        $px = if ($pinXNode) { [double]$pinXNode.GetAttribute("V") } else { 0 }
                        $py = if ($pinYNode) { [double]$pinYNode.GetAttribute("V") } else { 0 }

                        if ($hasArrow -and $hasLineCoords) {
                            $bx = [math]::Round([double]$beginXNode.GetAttribute("V"), 2)
                            $by = [math]::Round([double]$beginYNode.GetAttribute("V"), 2)
                            $ex = [math]::Round([double]$endXNode.GetAttribute("V"), 2)
                            $ey = [math]::Round([double]$endYNode.GetAttribute("V"), 2)
                            if ($eaVal -ne "0" -and $baVal -ne "0") { $dir = "<->" }
                            elseif ($eaVal -ne "0") { $dir = "-->" }
                            elseif ($baVal -ne "0") { $dir = "<--" }
                            else { $dir = "---" }
                            $arrowList.Add([PSCustomObject]@{
                                ID = $sid; BeginX = $bx; BeginY = $by
                                EndX = $ex; EndY = $ey
                                MidX = [math]::Round(($bx + $ex) / 2, 2)
                                MidY = [math]::Round(($by + $ey) / 2, 2)
                                Direction = $dir; Label = $shapeText
                            }) | Out-Null
                        }
                        elseif ($shapeText -and -not $hasArrow) {
                            $candidateList.Add([PSCustomObject]@{
                                Text = $shapeText
                                X = [math]::Round($px, 2)
                                Y = [math]::Round($py, 2)
                                ID = $sid
                            }) | Out-Null
                        }
                    }

                    # Pass 2: Classify candidates as entities or labels by distance to arrows
                    $labelList = New-Object System.Collections.Generic.List[PSObject]
                    foreach ($cand in $candidateList) {
                        $bestDist = [double]::MaxValue
                        foreach ($arrow in $arrowList) {
                            $dist = [math]::Sqrt([math]::Pow($cand.X - $arrow.MidX, 2) + [math]::Pow($cand.Y - $arrow.MidY, 2))
                            if ($dist -lt $bestDist) { $bestDist = $dist }
                        }
                        if ($bestDist -lt 1.0 -and $arrowList.Count -gt 0) {
                            $labelList.Add($cand) | Out-Null
                        } else {
                            $entityList.Add($cand) | Out-Null
                        }
                    }

                    $entities = $entityList.ToArray()
                    $arrows = $arrowList.ToArray()
                    $textLabels = $labelList.ToArray()

                    # Match text labels to arrows by coordinate proximity
                    foreach ($label in $textLabels) {
                        $bestArrow = $null; $bestDist = [double]::MaxValue
                        foreach ($arrow in $arrows) {
                            if ($arrow.Label) { continue }
                            $dist = [math]::Sqrt([math]::Pow($label.X - $arrow.MidX, 2) + [math]::Pow($label.Y - $arrow.MidY, 2))
                            if ($dist -lt $bestDist) { $bestDist = $dist; $bestArrow = $arrow }
                        }
                        if ($bestArrow -and $bestDist -lt 1.0) {
                            $bestArrow.Label = $label.Text
                        }
                    }

                    # Output entities sorted by X (left to right)
                    if ($entities.Count -gt 0) {
                        $text += "Entities:`n"
                        $entities | Sort-Object X | ForEach-Object {
                            $text += "  [$($_.X), $($_.Y)] $($_.Text)`n"
                        }
                    }

                    # Output flow sorted by MidY (top to bottom = chronological)
                    $labeledArrows = $arrows | Where-Object { $_.Label }
                    $unlabeledArrows = $arrows | Where-Object { -not $_.Label }

                    if ($labeledArrows.Count -gt 0) {
                        $text += "Flow (top-to-bottom = chronological):`n"
                        $stepNum = 0
                        $labeledArrows | Sort-Object { -$_.MidY } | ForEach-Object {
                            $stepNum++
                            $arrow = $_
                            $srcX = $arrow.BeginX
                            $dstX = $arrow.EndX
                            $srcEntity = ($entities | Sort-Object { [math]::Abs($_.X - $srcX) } | Select-Object -First 1).Text
                            $dstEntity = ($entities | Sort-Object { [math]::Abs($_.X - $dstX) } | Select-Object -First 1).Text
                            $text += "  Step $stepNum`: $srcEntity $($arrow.Direction) $dstEntity | $($arrow.Label)`n"
                        }
                    }

                    if ($unlabeledArrows.Count -gt 0) {
                        $text += "Unlabeled arrows ($($unlabeledArrows.Count)):`n"
                        $unlabeledArrows | Sort-Object { -$_.MidY } | ForEach-Object {
                            $arrow = $_
                            $srcX = $arrow.BeginX
                            $dstX = $arrow.EndX
                            $srcEntity = ($entities | Sort-Object { [math]::Abs($_.X - $srcX) } | Select-Object -First 1).Text
                            $dstEntity = ($entities | Sort-Object { [math]::Abs($_.X - $dstX) } | Select-Object -First 1).Text
                            $text += "  $srcEntity $($arrow.Direction) $dstEntity (no label)`n"
                        }
                    }

                    $text += "[END VISIO FLOWCHART]`n"

                    Remove-Item -LiteralPath $vsdxExtract -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch {
                    $text += "[VISIO ERROR: $($vsdx.Name) - $($_.Exception.Message)]`n"
                }
            }
        }
                        }
                        elseif ($shapeText -and -not $hasArrow -and -not $hasLineCoords) {
                            # Entity/NF box (has text, no arrow, no line coords)
                            $entities += [PSCustomObject]@{
                                Text = $shapeText
                                X = [math]::Round($px, 2)
                                Y = [math]::Round($py, 2)
                                ID = $sid
                            }
                        }
                        elseif ($shapeText -and -not $hasArrow) {
                            # Text label (has text, no arrow, may have PinX/PinY)
                            $textLabels += [PSCustomObject]@{
                                Text = $shapeText
                                X = [math]::Round($px, 2)
                                Y = [math]::Round($py, 2)
                                ID = $sid
                            }
                        }
                        elseif ($shapeText -and $hasArrow -and $hasLineCoords) {
                            # Labeled arrow (rare: has both text and arrow)
                            $bx = [math]::Round([double]$beginXNode.GetAttribute("V"), 2)
                            $by = [math]::Round([double]$beginYNode.GetAttribute("V"), 2)
                            $ex = [math]::Round([double]$endXNode.GetAttribute("V"), 2)
                            $ey = [math]::Round([double]$endYNode.GetAttribute("V"), 2)
                            if ($eaVal -ne "0" -and $baVal -ne "0") { $dir = "<->" }
                            elseif ($eaVal -ne "0") { $dir = "-->" }
                            elseif ($baVal -ne "0") { $dir = "<--" }
                            else { $dir = "---" }
                            $arrows += [PSCustomObject]@{
                                ID = $sid
                                BeginX = $bx; BeginY = $by
                                EndX = $ex; EndY = $ey
                                MidX = [math]::Round(($bx + $ex) / 2, 2)
                                MidY = [math]::Round(($by + $ey) / 2, 2)
                                Direction = $dir
                                Label = $shapeText
                            }
                        }
                        # else: lifeline or other → ignore
                    }

                    # Match text labels to arrows by coordinate proximity
                    foreach ($label in $textLabels) {
                        $bestArrow = $null; $bestDist = [double]::MaxValue
                        foreach ($arrow in $arrows) {
                            if ($arrow.Label) { continue }
                            $dist = [math]::Sqrt([math]::Pow($label.X - $arrow.MidX, 2) + [math]::Pow($label.Y - $arrow.MidY, 2))
                            if ($dist -lt $bestDist) { $bestDist = $dist; $bestArrow = $arrow }
                        }
                        if ($bestArrow -and $bestDist -lt 2.5) {
                            $bestArrow.Label = $label.Text
                        }
                    }

                    # Output entities sorted by X (left to right)
                    if ($entities.Count -gt 0) {
                        $text += "Entities:`n"
                        $entities | Sort-Object X | ForEach-Object {
                            $text += "  [$($_.X), $($_.Y)] $($_.Text)`n"
                        }
                    }

                    # Output labeled arrows sorted by MidY (top to bottom = chronological)
                    $labeledArrows = $arrows | Where-Object { $_.Label }
                    $unlabeledArrows = $arrows | Where-Object { -not $_.Label }

                    if ($labeledArrows.Count -gt 0) {
                        $text += "Flow (top-to-bottom = chronological):`n"
                        $stepNum = 0
                        $labeledArrows | Sort-Object { -$_.MidY } | ForEach-Object {
                            $stepNum++
                            $srcEntity = ($entities | Sort-Object { [math]::Abs($_.X - $_.BeginX) } | Select-Object -First 1).Text
                            $dstEntity = ($entities | Sort-Object { [math]::Abs($_.X - $_.EndX) } | Select-Object -First 1).Text
                            $text += "  Step $stepNum`: $srcEntity $($_.Direction) $dstEntity | $($_.Label)`n"
                        }
                    }

                    if ($unlabeledArrows.Count -gt 0) {
                        $text += "Unlabeled arrows ($($unlabeledArrows.Count)):`n"
                        $unlabeledArrows | Sort-Object { -$_.MidY } | ForEach-Object {
                            $srcEntity = ($entities | Sort-Object { [math]::Abs($_.X - $_.BeginX) } | Select-Object -First 1).Text
                            $dstEntity = ($entities | Sort-Object { [math]::Abs($_.X - $_.EndX) } | Select-Object -First 1).Text
                            $text += "  $srcEntity $($_.Direction) $dstEntity (no label)`n"
                        }
                    }

                    $text += "[END VISIO FLOWCHART]`n"

                    Remove-Item -LiteralPath $vsdxExtract -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch {
                    $text += "[VISIO ERROR: $($vsdx.Name) - $($_.Exception.Message)]`n"
                }
            }
        }

        # ============================================================
        # PART C: Layer 2 - VML/WPS inline textbox extraction
        # ============================================================
        $xmlRaw = Get-Content -LiteralPath $xmlPath -Encoding UTF8 -Raw

        # Extract VML textboxes (used by China Mobile, Huawei, Nokia)
        $vmlMatches = [regex]::Matches($xmlRaw, '<v:textbox[^>]*>(.*?)</v:textbox>', 'Singleline')
        if ($vmlMatches.Count -gt 0) {
            $text += "`n[VML FLOWCHART]`n"
            $diagramStats.VML++
            foreach ($m in $vmlMatches) {
                $inner = $m.Groups[1].Value
                $textMatches = [regex]::Matches($inner, '<w:t[^>]*>([^<]+)</w:t>')
                $lineText = ($textMatches | ForEach-Object { $_.Groups[1].Value }) -join ""
                if ($lineText.Trim()) {
                    $text += "  $($lineText.Trim())`n"
                }
            }
            $text += "[END VML FLOWCHART]`n"
        }

        # If no VML found, try WPS textboxes (fallback, usually paired with VML)
        if ($vmlMatches.Count -eq 0) {
            $wpsMatches = [regex]::Matches($xmlRaw, '<wps:txbx>(.*?)</wps:txbx>', 'Singleline')
            if ($wpsMatches.Count -gt 0) {
                $text += "`n[WPS FLOWCHART]`n"
                $diagramStats.VML++
                foreach ($m in $wpsMatches) {
                    $inner = $m.Groups[1].Value
                    $textMatches = [regex]::Matches($inner, '<w:t[^>]*>([^<]+)</w:t>')
                    $lineText = ($textMatches | ForEach-Object { $_.Groups[1].Value }) -join ""
                    if ($lineText.Trim()) {
                        $text += "  $($lineText.Trim())`n"
                    }
                }
                $text += "[END WPS FLOWCHART]`n"
            }
        }

        # ============================================================
        # PART D: Layer 3 - EMF → PNG conversion (for visual model)
        # ============================================================
        $mediaDir = Join-Path $workDir "word\media"
        if (Test-Path -LiteralPath $mediaDir) {
            $emfFiles = @(Get-ChildItem -LiteralPath $mediaDir -Filter "*.emf" -ErrorAction SilentlyContinue)
            $pngFiles = @(Get-ChildItem -LiteralPath $mediaDir -Filter "*.png" -ErrorAction SilentlyContinue)
            $allImages = @()

            # Convert EMF to PNG
            foreach ($emf in $emfFiles) {
                try {
                    $metafile = New-Object System.Drawing.Imaging.Metafile($emf.FullName)
                    New-Item -ItemType Directory -Path $docFigDir -Force | Out-Null
                    $pngPath = Join-Path $docFigDir "$($emf.BaseName).png"
                    $bmp = New-Object System.Drawing.Bitmap($metafile.Width, $metafile.Height)
                    $g = [System.Drawing.Graphics]::FromImage($bmp)
                    $g.Clear([System.Drawing.Color]::White)
                    $g.DrawImage($metafile, 0, 0, $metafile.Width, $metafile.Height)
                    $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
                    $g.Dispose(); $bmp.Dispose(); $metafile.Dispose()
                    $allImages += $pngPath
                }
                catch {
                    Write-Output "  EMF convert error: $($emf.Name) - $($_.Exception.Message)"
                }
            }

            # Copy existing PNG files
            foreach ($png in $pngFiles) {
                New-Item -ItemType Directory -Path $docFigDir -Force | Out-Null
                $pngDest = Join-Path $docFigDir $png.Name
                Copy-Item -LiteralPath $png.FullName -Destination $pngDest -Force
                $allImages += $pngDest
            }

            # Add figure markers to text
            foreach ($imgPath in $allImages) {
                $text += "`n[FIGURE: $imgPath]`n"
                $diagramStats.Figures++
            }
        }

        # ============================================================
        # Save extracted text
        # ============================================================
        Set-Content -LiteralPath $textFile -Value $text -Encoding UTF8
        $lines = (Get-Content -LiteralPath $textFile).Count
        $tableCount = ([regex]::Matches($text, "\[TABLE START\]")).Count
        Write-Output "OK ($current/$total): $($d.Name) - $lines lines, $tableCount tables, Visio=$($diagramStats.Visio), VML=$($diagramStats.VML), Figs=$($diagramStats.Figures)"
    }
    catch {
        Write-Output "ERROR ($current/$total): $($d.Name) - $($_.Exception.Message)"
    }
    finally {
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Output "`nDone extracting text + diagrams."
# === END OF BLOCK ===
```

### 4.3 Verify extraction results

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
$textFiles = Get-ChildItem -LiteralPath $textDir -Filter "*.txt"
Write-Output "=== Extraction Summary ==="
$totalVisio = 0; $totalVML = 0; $totalFig = 0; $noDiagrams = 0
foreach ($tf in $textFiles) {
    $content = Get-Content -LiteralPath $tf.FullName -Encoding UTF8 -Raw
    $v = ([regex]::Matches($content, "\[VISIO FLOWCHART")).Count
    $m = ([regex]::Matches($content, "\[VML FLOWCHART\]|\[WPS FLOWCHART\]")).Count
    $f = ([regex]::Matches($content, "\[FIGURE:")).Count
    if ($v + $m + $f -eq 0) { $noDiagrams++ }
    $totalVisio += $v; $totalVML += $m; $totalFig += $f
}
Write-Output "Total files: $($textFiles.Count)"
Write-Output "Visio flowcharts: $totalVisio"
Write-Output "VML/WPS flowcharts: $totalVML"
Write-Output "Figure images (for visual model): $totalFig"
Write-Output "Proposals with no diagrams: $noDiagrams"
# === END OF BLOCK ===
```

---

## STEP 5: Read and Analyze Proposals

### 5.1 List extracted text files

```powershell
# === COPY AND EXECUTE THIS BLOCK ===
$textFiles = Get-ChildItem -LiteralPath $textDir -Filter "*.txt" |
    Sort-Object Name |
    ForEach-Object { $_.Name }

Write-Output "Total text files: $($textFiles.Count)"
$textFiles | ForEach-Object { Write-Output "  $_" }
# === END OF BLOCK ===
```

### 5.2 Split into batches

Split files into batches of **12-15 files** each. Use smaller batches if files are large (>500 lines each).

```
Batch 1: Files 1-15 (alphabetically)
Batch 2: Files 16-30
Batch 3: Files 31-45
Batch 4: Remaining files (if any)
```

### 5.3 Launch parallel task agents

**CRITICAL**: Send ALL batch tasks in a SINGLE message to maximize parallelism.

For each batch, use the `task` tool with `subagent_type: "general"` and this prompt structure:

```
You are analyzing 3GPP {WG} contribution documents for KI#{kiNumber}
({KI Title}) from meeting {WG}#{meetingNum} ({Location}, {Date}).

Read the following {N} text files and for EACH proposal, provide:
1. Document number (e.g., S2-XXXXXXX)
2. Source company
3. Title
4. Which Solution # it targets (e.g., #22.1, #22.2, etc.)
5. Which bullet(s) of KI#{kiNumber} it addresses
6. Core technical proposal summary (2-4 sentences in Chinese)
7. Key new network functions or architecture elements proposed
8. Key procedures or call flows defined
9. Any tables found and their content summary
10. Flowchart/diagram analysis (see instructions below)

The files are at: {textDir}\

Read these files:
1. {filename1}.txt (Source: {company1})
2. {filename2}.txt (Source: {company2})
...

=== FLOWCHART/DIAGRAM ANALYSIS INSTRUCTIONS (MANDATORY) ===

The extracted text files may contain the following diagram markers:

a) [VISIO FLOWCHART: name] ... [END VISIO FLOWCHART]
   - Contains structured data extracted from embedded Visio diagrams
   - "Entities" section: network functions with [X, Y] coordinates (sorted left-to-right)
   - "Flow" section: message exchanges with Step numbers, sorted top-to-bottom (chronological)
   - Format: "Step N: SourceEntity --> DestEntity | Message Label"
   - Arrows: "-->" = forward, "<--" = reverse, "<->" = bidirectional
   - Use this structured data to understand the complete call flow

b) [VML FLOWCHART] ... [END VML FLOWCHART] or [WPS FLOWCHART] ... [END WPS FLOWCHART]
   - Contains text labels extracted from inline VML/WPS shapes
   - These are typically entity names and message step descriptions
   - Reconstruct the call flow from the numbered steps

c) [FIGURE: /path/to/image.png]
   - These are rasterized images from the original document
   - **MANDATORY**: You MUST use the Read tool to open and view EACH PNG image file listed
   - The Read tool supports image viewing - use it to see the actual flowchart content
   - After viewing each image, determine the diagram type:
     * If it is a signaling/call flow diagram: extract entities and message sequence, output as a Markdown table (same format as Visio/VML flowcharts)
     * If it is an architecture diagram: describe entities, connections, and reference points as bullet points
     * If it is a protocol stack diagram: describe each layer from top to bottom
   - For signaling flowcharts from PNG images, the output format MUST be identical to Visio/VML flowcharts:

| 步骤 | 发送方 → 接收方 | 消息 | 说明 |
|------|----------------|------|------|
| ① | ... | ... | ... |

For EACH proposal that contains flowcharts/diagrams, you MUST output flowchart analysis in this exact format:

**流程图分析:**
- 图数量: N

**图 1: [Figure title/caption from document text or image content]**
- 类型: Visio / VML / 图片
- 涉及实体: UE, AMF, SMF, ...
- 信令流程:

| 步骤 | 发送方 → 接收方 | 消息 | 说明 |
|------|----------------|------|------|
| ① | UE AI agent → UE | Notify: service required | Agent 通知 UE 需要网络服务 |
| ② | UE → AIMF | UE AI agent Registration request | UE 通过 NAS 发送注册请求 |
|  | AIMF ↔ UDM | Subscription check | AIMF 从 UDM 获取订阅数据 |
| ... | ... | ... | ... |

**图 2: [Figure title]**
- 类型: Visio / VML / 图片
- 涉及实体: ...
- 信令流程: (same table format)

If a proposal has NO flowcharts, output: **流程图分析:** 无

=== END FLOWCHART INSTRUCTIONS ===

IMPORTANT:
- Text marked with [H1], [H2] etc. are headings - use them to understand document structure
- Text between [TABLE START] and [TABLE END] are table contents - analyze them carefully
- Return your analysis in a structured format
- Be thorough: capture ALL technical details including NF names, reference points, procedure steps
- Write summaries in Chinese (中文)
- For flowcharts, always reconstruct the complete message sequence with entity names and message labels
- **PNG images containing signaling flows MUST use the same Markdown table format as Visio/VML flowcharts** — do NOT describe them as prose text

Return a JSON-like structured output for each proposal.
```

### 5.4 Verify coverage after all batches complete

After all task agents return, verify:

```
Expected proposals: {docNumbers.Count}
Analyzed proposals: {count from task results}
Missing: {list any doc numbers not covered}
```

If any proposals are missing, launch an additional task agent for them.

---

## STEP 6: Verify Coverage & Generate Structured Markdown Summary

### 6.1 Output file location

```
C:\Users\Administrator\Desktop\OpenCodeDiscussion\{MeetingNumber}_KI{kiNumber}.md
```

### 6.2 REQUIRED document structure

```markdown
# {WG}#{MeetingNum} ({Location}, {Date}) KI#{kiNumber} ({KI Title}) 提案总结

## 一、总览

- **会议**: {WG}#{MeetingNum}, {Location}, {Date}
- **KI#{kiNumber}**: {KI Title}
- **议程项**: {agenda item number}
- **提案总数**: {N}篇
- **来源公司**: {list of companies}

### 与上次会议对比 (if applicable)

| 维度 | {WG}#{PrevNum} | {WG}#{MeetingNum} |
|------|---------------|------------------|
| 提案数 | X篇 | Y篇 |
| ... | ... | ... |

---

## 二、按 Solution 分类总览

### Solution #{kiNumber}.1: [Solution Title]
**对应 KI#{kiNumber} Bullet Y & Z**

| 提案编号 | 来源公司 | 标题/主题 |
|---------|---------|----------|
| S2-XXXXXXX | Company A | Brief description |

**Variants/Options:**

| Option/Variant | 描述 | 支持提案 |
|----------------|------|---------|
| Var A | Description | List of proposals |

**Solution 总结:**
[Technical summary paragraph]

---

(repeat for each Solution)

---

## 三、各公司技术方向对比表

| 公司 | 核心 NF/方案 | 关键特色 |
|------|------------|---------|
| Company A | NF name | Key differentiator |

---

## 四、各提案详细总结

### S2-XXXXXXX (Company) - Solution #XX.Y: Title

**技术总结:**
[1 paragraph detailed technical summary]

**流程图:**

(If the proposal contains flowcharts, include the following for EACH figure:)

**图1: [Figure title/caption from document]**

涉及实体: UE, AMF, SMF, UPF, ...

信令流程:

| 步骤 | 发送方 → 接收方 | 消息 | 说明 |
|------|----------------|------|------|
| ① | UE → AMF | Registration Request | UE 发起注册请求 |
| ② | AMF → SMF | Nsmf_PDUSession_CreateSMContext | AMF 向 SMF 创建会话上下文 |
| ③ | SMF → AMF | Response | SMF 返回响应 |
| ④ | AMF → UE | Registration Accept | AMF 返回注册接受 |

**图2: [Figure title/caption]**
(Repeat for each figure)

(If no flowcharts:)
**流程图:** 无

---

(repeat for each proposal)

---

## 五、关键分歧与共识

### 共识点
- [List of areas where most companies agree]

### 分歧点
- [List of areas where companies disagree, with supporting proposal references]

### 待进一步讨论
- [Open issues that need further discussion in next meeting]
```

### 6.3 Summary content guidelines

For each **Solution summary**, cover:
1. Architecture choices (new NFs vs reused)
2. Signaling paths (NAS vs UP vs AF-proxied)
3. Key procedures
4. What makes each company's approach unique
5. Where proposals agree (consensus)
6. Where proposals disagree (divergence)

For each **individual proposal summary**, cover:
1. Source company and title
2. Which Solution and Variant it targets
3. Core technical contribution
4. Key call flows
5. Unique aspects
6. **Flowchart analysis** (if diagrams exist):
   - Reconstruct the complete message sequence as a Markdown table
   - List all participating entities (NFs, UE, AF, etc.)
   - Include ALL message labels from the flowchart
   - If multiple figures exist, create a separate table for each
   - If the figure is an architecture diagram (not a call flow), describe the architecture elements and reference points as bullet points

### 6.4 Flowchart output format guidelines

**All signaling/call flow diagrams** (regardless of source: Visio XML, VML, or PNG image) MUST use the same Markdown table format:

```markdown
| 步骤 | 发送方 → 接收方 | 消息 | 说明 |
|------|----------------|------|------|
| ① | UE → RAN | RRC Connection Request | UE 发起 RRC 连接 |
|  | RAN → AMF | N2 Initial UE Message | RAN 转发初始消息 |
| ③ | AMF → SMF | Nsmf_PDUSession_CreateSMContext | AMF 请求创建会话 |
| ④ | SMF → UPF | N4 Session Establishment | SMF 配置 UPF |
| ⑤ | UPF → SMF | Response | UPF 返回响应 |
| ⑥ | SMF → AMF | Response | SMF 返回响应 |
| ⑦ | AMF → UE | Registration Accept | AMF 接受注册 |
```

**Architecture diagrams** (from any source) use descriptive bullet points:

```
架构图描述:
- 新增 NF: NCEF (Network Capability Exposure Function)
- 参考点: Ncef1 (NCEF ↔ AF), Ncef2 (NCEF ↔ NW Agent)
- 数据流: AF → NCEF → NW Agent → 6G UPF
```

**Protocol stack diagrams** (from any source) use layered description:

```
协议栈描述 (从上到下):
- 应用层: AI Agent Protocol
- 传输层: HTTP/2 over TLS
- 网络层: IPv6
- 接入层: 6G RAN
```

**IMPORTANT**:
- Always use Chinese for descriptions and annotations (说明列)
- Entity names should use standard 3GPP abbreviations (AMF, SMF, UPF, NEF, AF, etc.)
- Message labels should preserve the original English text from the flowchart (消息列)
- Number the steps using circled numbers (①...) to match the original figure
- For bidirectional arrows, use "<->" in the direction column
- The "说明" column should briefly explain the purpose of each step in Chinese
- **PNG images containing signaling flows MUST be converted to the same table format as Visio/VML flowcharts** — do NOT describe them as prose

### 6.5 Flowchart assembly logic (Step 5 → Step 6)

When assembling the final summary document, for EACH proposal:

1. **Extract flowchart data from task agent output**: The task agent returns structured analysis including a "流程图分析" section. Extract this section verbatim.

2. **Insert into proposal summary template**: Place the flowchart analysis immediately after the "技术总结" paragraph, under a "**流程图:**" heading.

3. **Template for each proposal in Section 四**:

```markdown
### S2-XXXXXXX (Company) - Solution #XX.Y: Title

**技术总结:**
[1 paragraph detailed technical summary from task agent]

**流程图:**

[Insert the complete "流程图分析" section from task agent output here, including all figures with their tables]

---
```

4. **If task agent output has NO flowchart analysis**: Output "**流程图:** 无"

5. **Quality check**: After assembly, verify that every proposal with Visio/VML/FIGURE markers in its .txt file has a corresponding flowchart section in the final document. If missing, re-run the task agent for that specific proposal.

### 6.6 CRITICAL: No truncation or omission rules

**ABSOLUTELY PROHIBITED** in the final summary document:

- **NEVER** use phrases like "由于篇幅限制"、"以下提案仅列出关键信息"、"完整的分析数据已保存在 task agent 返回结果中" or any similar truncation notices
- **NEVER** omit any proposal's flowchart analysis, even if the document becomes very long
- **NEVER** replace a proposal's full flowchart tables with a one-sentence summary
- **NEVER** write "流程图分析见上文" or "同前" or any reference that avoids including the actual content

**MANDATORY rules**:

- **EVERY** proposal MUST have its own complete `**技术总结:**` paragraph (2-4 sentences in Chinese)
- **EVERY** proposal MUST have its own complete `**流程图:**` section with ALL figure tables from the task agent output
- If a proposal has NO flowcharts, output `**流程图:** 无` — do NOT skip the section entirely
- If the document is too large for a single write operation, split the output into multiple files (e.g., `175-AH-e_KI22_part1.md`, `175-AH-e_KI22_part2.md`) rather than omitting content
- The final document MUST contain exactly N proposal sections where N = total number of proposals analyzed

**Verification step**: After generating the document, count the number of `### S2-` headings in Section 四 and verify it matches the expected proposal count. If it does not match, regenerate the missing proposals.

---

## TROUBLESHOOTING

### Problem: "Duplicate column headers" error when reading Excel
**Solution**: You forgot `-NoHeader`. Always use:
```powershell
Import-Excel -Path $excelPath -WorksheetName $sheetName -NoHeader
```

### Problem: Cannot find proposals when filtering
**Solution**: The agenda item number may not match. Print all unique agenda items:
```powershell
$dataRows | Where-Object { $_.$colDocNum -match $wgPrefix } |
    Select-Object -ExpandProperty $colAgenda -Unique | Sort-Object
```

### Problem: .docx download returns 404
**Solution**: 3GPP stores `.zip` files, not `.docx`. Use:
```
https://www.3gpp.org/ftp/.../Docs/S2-2600098.zip   ← CORRECT
https://www.3gpp.org/ftp/.../Docs/S2-2600098.docx   ← WRONG
```

### Problem: Cannot access 3GPP FTP
**Solution**: Tell the user to manually download from:
```
https://www.3gpp.org/ftp/{wgPath}/
```
Then provide the local file path.

### Problem: Context overflow when reading proposals
**Solution**: Use smaller batches (8-10 files per task).

### Problem: Downloaded zip is too small / corrupt
**Solution**: The script checks file size < 500 bytes and re-downloads. If persistent, the document may not exist on the server.

### Problem: Text extraction produces garbled characters
**Solution**: Ensure `-Encoding UTF8` is used in both `Get-Content` and `Set-Content`.

### Problem: EMF → PNG conversion fails
**Solution**: Ensure `Add-Type -AssemblyName System.Drawing` is executed at the start of Step 4. Some EMF files with unsupported record types may fail — these will be logged but won't block other extraction.

### Problem: Visio .vsdx has no page1.xml
**Solution**: Some Visio files use different page names. Check `visio/pages/pages.xml` for the actual page file name.

### Problem: VML textboxes produce garbled or split text
**Solution**: VML text is sometimes split across multiple `<w:t>` elements with formatting. The regex extraction joins them, but spacing may be off. This is acceptable for analysis purposes.

### Problem: Old binary .vsd files cannot be parsed
**Solution**: `.vsd` (Visio 2003-2010 binary format) cannot be parsed as XML. The corresponding EMF preview image will be converted to PNG and can be analyzed by the visual model (Layer 3).

---

## KEY FACTS CHEAT SHEET

```
3GPP base URL:         https://www.3gpp.org/ftp/
SA2 FTP path:          tsg_sa/WG2_Arch  (NOT WG2_Architecture)
SA1 FTP path:          tsg_sa/WG1_Serv
SA3 FTP path:          tsg_sa/WG3_Security
CT1 FTP path:          tsg_ct/WG1_NAS
Index file format:     SA2-{num}_Index_{year}.zip (contains .xlsx)
Excel read parameter:  -NoHeader (MANDATORY)
Column detection:      Dynamic - scan first 20 rows for headers
Common SA2 columns:    P4=Agenda, P5=DocNum, P8=Title, P9=Source
Proposal file format:  .zip (contains .docx inside)
Text extraction:       .docx → unzip → word/document.xml → parse paragraphs + tables
Track changes:         Skip w:del elements, only extract final text
Heading detection:     w:pStyle with Heading1/2/3 or Title
Table extraction:      w:tbl → w:tr → w:tc → pipe-delimited
Diagram Layer 1:       word/embeddings/*.vsdx → visio/pages/page1.xml → shapes + arrows
Diagram Layer 2:       document.xml → regex <v:textbox> or <wps:txbx> → text labels
Diagram Layer 3:       word/media/*.emf → .NET System.Drawing → PNG → visual model
Diagram coverage:      ~93% of proposals have extractable diagrams
EMF conversion:        Add-Type -AssemblyName System.Drawing; Metafile → Bitmap → PNG
Download strategy:     Parallel Jobs, batch of 5, with retry
Cache directory:       C:\Users\Administrator\Downloads\3gpp_cache\
Figures directory:     C:\Users\Administrator\Downloads\3gpp_cache\{meeting}\figures\
Reading strategy:      Split into batches of 12-15, use task tool in parallel
Visual model:          Task agents use Read tool to view [FIGURE: path] PNG images
Flowchart output:      Markdown table (步骤/发送方→接收方/消息/说明) for call flows, bullet points for architecture
Coverage check:        Verify all doc numbers analyzed before generating summary
Output language:       Chinese (中文)
Output location:       C:\Users\Administrator\Desktop\OpenCodeDiscussion\
```

---

## EXAMPLE: Complete execution for SA2#173 KI#22

```
User: "请帮我总结 SA2#173 的 KI#22 提案"

Step 0: Cache dir created, ImportExcel verified

Step 1: webfetch "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/"
        → Found: TSGS2_173_Goa_2026-02

Step 2: webfetch ".../TSGS2_173_Goa_2026-02/"
        → Found: SA2-173_Index_2026.zip
        → Downloaded to cache\TSGS2_173_Goa_2026-02\index\

Step 3: Import-Excel with -NoHeader, dynamic column detection
        → Detected: P4=Agenda, P5=DocNum, P8=Title, P9=Source
        → Filter P4 = "20.6.22" → 45 proposals
        → Parallel download (batches of 5) → 45 OK, 0 failed

Step 4: Extracted text + diagrams from 45 .docx → .txt files
        → Paragraphs + tables (all 45)
        → Visio flowcharts: 33 proposals
        → VML/WPS flowcharts: 6 proposals
        → EMF→PNG figures: 33 proposals (for visual model verification)
        → No diagrams: 3 proposals (pure text)

Step 5: 3 parallel task batches (15+15+15)
        → All 45 proposals analyzed
        → Task agents read [FIGURE: path] PNGs via Read tool (visual model)
        → Flowcharts reconstructed as Markdown tables (步骤/发送方→接收方/消息/说明)
        → Coverage verified: 45/45

Step 6: Generated 173_KI22.md with:
        - Overview (45 proposals, 24+ companies)
        - Solution classification (#22.1-#22.13)
        - Company comparison table
        - 45 individual proposal summaries with flowchart tables
        - Key consensus and divergence analysis
```
