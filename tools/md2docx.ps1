# md2docx.ps1 - convert a Markdown manual into a Word (.docx) document.
#
# No pandoc / python / Microsoft Word required: this builds the OOXML package
# directly with System.IO.Compression plus hand-written XML.
# Compatible with Windows PowerShell 5.1 and PowerShell 7+.
#
# NOTE: this file is intentionally ASCII-only. Non-ASCII comments were observed
# to be corrupted on write in some Windows toolchains, which silently broke the
# parser (a CJK comment can collapse onto the following line and comment out code).
#
# Usage:
#   & .\tools\md2docx.ps1 -Input .\CONTRIBUTING.md -Output .\GIT-manual.docx

param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [string]$Title = 'Team Git / GitHub Collaboration Manual',
    [string]$Author = ''
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Double-quote character, kept in a variable to avoid escaping problems.
$Q = [string][char]34

# ---------- XML escaping ----------
function Esc([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $s = $s.Replace('&', '&amp;')
    $s = $s.Replace('<', '&lt;')
    $s = $s.Replace('>', '&gt;')
    $s = $s.Replace('"', '&quot;')
    return $s
}

# ---------- inline Markdown -> Word runs ----------
function Runs([string]$text, [string]$baseProps) {
    if ([string]::IsNullOrEmpty($text)) { return '' }
    $sb = New-Object System.Text.StringBuilder
    $pattern = '(\*\*.+?\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\))'
    $parts = [regex]::Split($text, $pattern)
    foreach ($p in $parts) {
        if ([string]::IsNullOrEmpty($p)) { continue }
        $inner = ''
        if ($p -match '^\*\*(.+)\*\*$') {
            $inner = '<w:r><w:rPr>' + $baseProps + '<w:b/></w:rPr><w:t xml:space=' + $Q + 'preserve' + $Q + '>' + (Esc $Matches[1]) + '</w:t></w:r>'
        }
        elseif ($p -match '^`(.+)`$') {
            $inner = '<w:r><w:rPr>' + $baseProps + '<w:rFonts w:ascii=' + $Q + 'Consolas' + $Q + ' w:hAnsi=' + $Q + 'Consolas' + $Q + '/><w:shd w:fill=' + $Q + 'F2F2F2' + $Q + '/></w:rPr><w:t xml:space=' + $Q + 'preserve' + $Q + '>' + (Esc $Matches[1]) + '</w:t></w:r>'
        }
        elseif ($p -match '^\[([^\]]+)\]\(([^)]+)\)$') {
            $inner = '<w:r><w:rPr>' + $baseProps + '<w:color w:val=' + $Q + '0563C1' + $Q + '/><w:u w:val=' + $Q + 'single' + $Q + '/></w:rPr><w:t xml:space=' + $Q + 'preserve' + $Q + '>' + (Esc $Matches[1]) + '</w:t></w:r>'
        }
        else {
            $inner = '<w:r><w:rPr>' + $baseProps + '</w:rPr><w:t xml:space=' + $Q + 'preserve' + $Q + '>' + (Esc $p) + '</w:t></w:r>'
        }
        $sb.Append($inner) | Out-Null
    }
    return $sb.ToString()
}

# ---------- paragraph emitters ----------
$Script:Body = New-Object System.Text.StringBuilder

function Para([string]$runsXml, [string]$style, [string]$extraPPr) {
    $ppr = ''
    if (-not [string]::IsNullOrEmpty($style)) { $ppr += '<w:pStyle w:val=' + $Q + $style + $Q + '/>' }
    if (-not [string]::IsNullOrEmpty($extraPPr)) { $ppr += $extraPPr }
    if ($ppr.Length -gt 0) { $ppr = '<w:pPr>' + $ppr + '</w:pPr>' }
    $Script:Body.Append('<w:p>' + $ppr + $runsXml + '</w:p>') | Out-Null
}

# Compact paragraph used inside table cells.
function CellPara([string]$runsXml) {
    $p = '<w:p><w:pPr><w:spacing w:before=' + $Q + '0' + $Q + ' w:after=' + $Q + '0' + $Q +
         ' w:line=' + $Q + '240' + $Q + ' w:lineRule=' + $Q + 'auto' + $Q + '/></w:pPr>' + $runsXml + '</w:p>'
    return $p
}

# ---------- real Word table (not flattened text) ----------
function TableXml($rows) {
    if ($rows.Count -eq 0) { return }
    $sb = New-Object System.Text.StringBuilder
    $sb.Append('<w:tbl><w:tblPr><w:tblStyle w:val=' + $Q + 'TableGrid' + $Q + '/>') | Out-Null
    $sb.Append('<w:tblW w:w=' + $Q + '5000' + $Q + ' w:type=' + $Q + 'pct' + $Q + '/>') | Out-Null
    $sb.Append('<w:tblBorders>') | Out-Null
    foreach ($side in @('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) {
        $sb.Append('<w:' + $side + ' w:val=' + $Q + 'single' + $Q + ' w:sz=' + $Q + '4' + $Q +
                   ' w:space=' + $Q + '0' + $Q + ' w:color=' + $Q + 'BFBFBF' + $Q + '/>') | Out-Null
    }
    $sb.Append('</w:tblBorders></w:tblPr>') | Out-Null

    $rowIdx = 0
    foreach ($row in $rows) {
        $isHeader = ($rowIdx -eq 0)
        $sb.Append('<w:tr>') | Out-Null
        if ($isHeader) { $sb.Append('<w:trPr><w:tblHeader/></w:trPr>') | Out-Null }
        foreach ($cell in $row) {
            $shd = 'FFFFFF'
            if ($isHeader) { $shd = 'DEEAF6' }
            elseif ($rowIdx % 2 -eq 0) { $shd = 'F7F9FC' }
            $base = '<w:sz w:val=' + $Q + '19' + $Q + '/>'
            if ($isHeader) { $base += '<w:b/>' }
            $runs = Runs ([string]$cell) $base
            $sb.Append('<w:tc><w:tcPr><w:shd w:val=' + $Q + 'clear' + $Q + ' w:fill=' + $Q + $shd + $Q + '/></w:tcPr>') | Out-Null
            $sb.Append((CellPara $runs)) | Out-Null
            $sb.Append('</w:tc>') | Out-Null
        }
        $sb.Append('</w:tr>') | Out-Null
        $rowIdx++
    }
    $sb.Append('</w:tbl>') | Out-Null
    $sb.Append('<w:p><w:pPr><w:spacing w:before=' + $Q + '0' + $Q + ' w:after=' + $Q + '0' + $Q + '/></w:pPr></w:p>') | Out-Null
    $Script:Body.Append($sb.ToString()) | Out-Null
}

# ---------- line-by-line Markdown parsing ----------
$md = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8
$lines = $md -split "`r?`n"

$inCode = $false
$codeBuf = New-Object System.Collections.Generic.List[string]
$tableRows = New-Object System.Collections.Generic.List[object]

function FlushTable() {
    if ($tableRows.Count -gt 0) {
        TableXml $tableRows
        $tableRows.Clear()
    }
}

foreach ($line in $lines) {

    # fenced code block
    if ($line -match '^\s*```') {
        if (-not $inCode) {
            $inCode = $true
            FlushTable
            $codeBuf.Clear()
        }
        else {
            $inCode = $false
            foreach ($cl in $codeBuf) {
                $base = '<w:rFonts w:ascii=' + $Q + 'Consolas' + $Q + ' w:hAnsi=' + $Q + 'Consolas' + $Q +
                        '/><w:sz w:val=' + $Q + '18' + $Q + '/><w:color w:val=' + $Q + '1F1F1F' + $Q + '/>'
                $r = Runs $cl $base
                $extra = '<w:shd w:val=' + $Q + 'clear' + $Q + ' w:fill=' + $Q + 'F6F8FA' + $Q +
                         '/><w:spacing w:before=' + $Q + '0' + $Q + ' w:after=' + $Q + '0' + $Q +
                         ' w:line=' + $Q + '240' + $Q + ' w:lineRule=' + $Q + 'auto' + $Q +
                         '/><w:ind w:left=' + $Q + '240' + $Q + ' w:right=' + $Q + '120' + $Q + '/>'
                Para $r '' $extra
            }
            Para '' '' ''
        }
        continue
    }
    if ($inCode) { $codeBuf.Add($line); continue }

    # table row
    if ($line -match '^\s*\|.*\|\s*$') {
        if ($line -match '^\s*\|[\s:\-\|]+\|\s*$') { continue }
        $cells = New-Object System.Collections.Generic.List[string]
        foreach ($c in ($line.Trim().Trim('|') -split '\|')) { $cells.Add($c.Trim()) }
        $tableRows.Add($cells) | Out-Null
        continue
    }
    else { FlushTable }

    # blank line
    if ($line -match '^\s*$') { continue }

    # horizontal rule
    if ($line -match '^\s*(---+|\*\*\*+|___+)\s*$') {
        Para '' '' ('<w:pBdr><w:bottom w:val=' + $Q + 'single' + $Q + ' w:sz=' + $Q + '6' + $Q +
                    ' w:space=' + $Q + '1' + $Q + ' w:color=' + $Q + 'CCCCCC' + $Q + '/></w:pBdr>')
        continue
    }

    # heading
    if ($line -match '^(#{1,6})\s+(.*)$') {
        $level = $Matches[1].Length
        $text = $Matches[2].Trim()
        Para (Runs $text '') ('Heading' + $level) ''
        continue
    }

    # blockquote
    if ($line -match '^\s*>\s?(.*)$') {
        $text = $Matches[1]
        $extra = '<w:pBdr><w:left w:val=' + $Q + 'single' + $Q + ' w:sz=' + $Q + '18' + $Q +
                 ' w:space=' + $Q + '6' + $Q + ' w:color=' + $Q + 'D0D7DE' + $Q +
                 '/></w:pBdr><w:shd w:val=' + $Q + 'clear' + $Q + ' w:fill=' + $Q + 'F6F8FA' + $Q +
                 '/><w:ind w:left=' + $Q + '240' + $Q + '/><w:spacing w:before=' + $Q + '60' + $Q +
                 ' w:after=' + $Q + '60' + $Q + '/>'
        Para (Runs $text '') '' $extra
        continue
    }

    # checkbox item
    if ($line -match '^\s*[-*]\s+\[\s?\]\s+(.*)$') {
        $ind = '<w:ind w:left=' + $Q + '360' + $Q + ' w:hanging=' + $Q + '240' + $Q +
               '/><w:spacing w:before=' + $Q + '0' + $Q + ' w:after=' + $Q + '0' + $Q + '/>'
        Para (Runs ([string][char]0x2610 + '  ' + $Matches[1]) '') '' $ind
        continue
    }

    # unordered list
    if ($line -match '^\s*[-*]\s+(.*)$') {
        $ind = '<w:ind w:left=' + $Q + '360' + $Q + ' w:hanging=' + $Q + '240' + $Q +
               '/><w:spacing w:before=' + $Q + '0' + $Q + ' w:after=' + $Q + '0' + $Q + '/>'
        Para (Runs ([string][char]0x2022 + '  ' + $Matches[1]) '') '' $ind
        continue
    }

    # ordered list
    if ($line -match '^\s*(\d+)\.\s+(.*)$') {
        $ind = '<w:ind w:left=' + $Q + '360' + $Q + ' w:hanging=' + $Q + '240' + $Q +
               '/><w:spacing w:before=' + $Q + '0' + $Q + ' w:after=' + $Q + '0' + $Q + '/>'
        Para (Runs ($Matches[1] + '.  ' + $Matches[2]) '') '' $ind
        continue
    }

    # plain paragraph
    Para (Runs $line '') '' ''
}
FlushTable

# ---------- document.xml ----------
$documentXml = '<?xml version=' + $Q + '1.0' + $Q + ' encoding=' + $Q + 'UTF-8' + $Q + ' standalone=' + $Q + 'yes' + $Q + '?>' + "`n" +
    '<w:document xmlns:w=' + $Q + 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' + $Q + '><w:body>' + "`n" +
    $Script:Body.ToString() + "`n" +
    '<w:sectPr><w:pgSz w:w=' + $Q + '11906' + $Q + ' w:h=' + $Q + '16838' + $Q + '/><w:pgMar w:top=' + $Q + '1134' + $Q +
    ' w:right=' + $Q + '1134' + $Q + ' w:bottom=' + $Q + '1134' + $Q + ' w:left=' + $Q + '1134' + $Q +
    ' w:header=' + $Q + '851' + $Q + ' w:footer=' + $Q + '992' + $Q + ' w:gutter=' + $Q + '0' + $Q +
    '/></w:sectPr>' + "`n" + '</w:body></w:document>'

# ---------- styles.xml ----------
$stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/>
<w:sz w:val="21"/><w:szCs w:val="21"/>
</w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>
<w:pPr><w:spacing w:after="120" w:line="288" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
<w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="360" w:after="160"/><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="2E74B5"/></w:pBdr></w:pPr>
<w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="1F4E79"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
<w:pPr><w:outlineLvl w:val="1"/><w:spacing w:before="280" w:after="120"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="2E74B5"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>
<w:pPr><w:outlineLvl w:val="2"/><w:spacing w:before="220" w:after="100"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="1F4E79"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/>
<w:pPr><w:outlineLvl w:val="3"/><w:spacing w:before="180" w:after="80"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/><w:basedOn w:val="Normal"/>
<w:pPr><w:outlineLvl w:val="4"/><w:spacing w:before="150" w:after="60"/></w:pPr>
<w:rPr><w:b/><w:i/><w:sz w:val="22"/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="heading 6"/><w:basedOn w:val="Normal"/>
<w:pPr><w:outlineLvl w:val="5"/><w:spacing w:before="120" w:after="60"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="21"/><w:color w:val="595959"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>
<w:tblPr><w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
</w:tblBorders></w:tblPr></w:style>
</w:styles>
'@

$contentTypesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
'@

$relsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

$docRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@

$now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$coreXml = '<?xml version=' + $Q + '1.0' + $Q + ' encoding=' + $Q + 'UTF-8' + $Q + ' standalone=' + $Q + 'yes' + $Q + '?>' + "`n" +
    '<cp:coreProperties xmlns:cp=' + $Q + 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties' + $Q +
    ' xmlns:dc=' + $Q + 'http://purl.org/dc/elements/1.1/' + $Q +
    ' xmlns:dcterms=' + $Q + 'http://purl.org/dc/terms/' + $Q +
    ' xmlns:xsi=' + $Q + 'http://www.w3.org/2001/XMLSchema-instance' + $Q + '>' + "`n" +
    '<dc:title>' + (Esc $Title) + '</dc:title>' + "`n" +
    '<dc:creator>' + (Esc $Author) + '</dc:creator>' + "`n" +
    '<cp:lastModifiedBy>' + (Esc $Author) + '</cp:lastModifiedBy>' + "`n" +
    '<dcterms:created xsi:type=' + $Q + 'dcterms:W3CDTF' + $Q + '>' + $now + '</dcterms:created>' + "`n" +
    '<dcterms:modified xsi:type=' + $Q + 'dcterms:W3CDTF' + $Q + '>' + $now + '</dcterms:modified>' + "`n" +
    '</cp:coreProperties>'

$appXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>md2docx.ps1</Application>
</Properties>
'@

# ---------- write the ZIP package ----------
if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }

$parts = New-Object System.Collections.Specialized.OrderedDictionary
$parts.Add('[Content_Types].xml', $contentTypesXml)
$parts.Add('_rels/.rels', $relsXml)
$parts.Add('word/document.xml', $documentXml)
$parts.Add('word/styles.xml', $stylesXml)
$parts.Add('word/_rels/document.xml.rels', $docRelsXml)
$parts.Add('docProps/core.xml', $coreXml)
$parts.Add('docProps/app.xml', $appXml)

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$fs = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::CreateNew)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($k in $parts.Keys) {
        $entry = $zip.CreateEntry($k, [System.IO.Compression.CompressionLevel]::Optimal)
        $es = $entry.Open()
        try {
            $bytes = $utf8NoBom.GetBytes([string]$parts[$k])
            $es.Write($bytes, 0, $bytes.Length)
        }
        finally { $es.Dispose() }
    }
}
finally {
    $zip.Dispose()
    $fs.Dispose()
}

# ---------- self-check ----------
$check = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $OutputPath).Path)
$names = @($check.Entries | ForEach-Object { $_.FullName })
$check.Dispose()

Write-Output ('Created: ' + $OutputPath)
Write-Output ('Package parts: ' + $names.Count + ' -> ' + ($names -join ', '))
Write-Output ('document.xml size: ' + $documentXml.Length.ToString('N0') + ' chars')
