# renumber.ps1 - renumber chapter / step headings inside a .docx, in place.
#
# Rules:
#   * "heading 2" paragraphs whose text begins with a CJK numeral followed by
#     the enumeration comma are renumbered sequentially: 1,2,3... rendered as
#     CJK numerals.
#   * "heading 3" paragraphs matching the CJK "Step N: ..." pattern are
#     renumbered sequentially, restarting inside each heading-2 chapter.
#   * everything else (including all run formatting on body paragraphs) is kept.
#
# ASCII-only on purpose: non-ASCII source comments were observed to be
# corrupted on write by some Windows toolchains.
#
# Usage:
#   & .\tools\renumber.ps1 -InputPath .\GIT-manual.docx

param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Q = [string][char]34
$CJK = @([char]0x4E00, [char]0x4E8C, [char]0x4E09, [char]0x56DB, [char]0x4E94,
         [char]0x516D, [char]0x4E03, [char]0x516B, [char]0x4E5D, [char]0x5341)
$IDEOGRAPHIC_COMMA = [char]0x3001    # the CJK enumeration comma
$FULLWIDTH_COLON   = [char]0xFF1A    # the CJK colon
$STEP_WORD         = [char]0x6B65 + [char]0x9AA4        # the CJK word for "step"
$STEP_CHAR         = [char]0x6B65                        # the CJK char "step"

function Get-ZipText($zip, $name) {
    $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
    if (-not $e) { return '' }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    $c = $sr.ReadToEnd(); $sr.Close(); return $c
}

function Read-Para($xml) {
    $style = ''
    if ($xml -match '<w:pStyle w:val=' + $Q + '([^' + $Q + ']*)' + $Q) { $style = $Matches[1] }
    $t = ([regex]::Matches($xml, '<w:t[^>]*>([^<]*)</w:t>') | ForEach-Object { $_.Groups[1].Value }) -join ''
    $t = $t.Replace('&quot;', '"').Replace('&amp;', '&').Replace('&lt;', '<').Replace('&gt;', '>')
    $ppr = ''
    if ($xml -match '<w:pPr>.*?</w:pPr>') { $ppr = $Matches[0] }
    return [pscustomobject]@{ Style = $style; Text = $t; PPr = $ppr }
}

function New-HeadingXml([string]$style, [string]$text, [string]$ppr) {
    if ([string]::IsNullOrEmpty($ppr)) {
        $ppr = '<w:pPr><w:pStyle w:val=' + $Q + $style + $Q + '/></w:pPr>'
    }
    $esc = $text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
    $rpr = '<w:rFonts w:ascii=' + $Q + 'Calibri' + $Q + ' w:hAnsi=' + $Q + 'Calibri' + $Q +
           ' w:eastAsia=' + $Q + 'Microsoft YaHei' + $Q + '/>'
    return '<w:p>' + $ppr + '<w:r><w:rPr>' + $rpr + '</w:rPr><w:t xml:space=' + $Q + 'preserve' + $Q + '>' + $esc + '</w:t></w:r></w:p>'
}

# ---------- load ----------
$src = (Resolve-Path $InputPath).Path
$zip = [System.IO.Compression.ZipFile]::OpenRead($src)
$doc = Get-ZipText $zip 'word/document.xml'
$sty = Get-ZipText $zip 'word/styles.xml'
$zip.Dispose()

$map = @{}
foreach ($m in [regex]::Matches($sty, '<w:style [^>]*w:styleId=' + $Q + '([^' + $Q + ']*)' + $Q + '[^>]*>\s*<w:name w:val=' + $Q + '([^' + $Q + ']*)' + $Q)) {
    $map[$m.Groups[1].Value] = $m.Groups[2].Value
}

# ---------- pass 1: collect headings in document order ----------
$paras = [regex]::Matches($doc, '<w:p[ >].*?</w:p>')
$headings = @()
for ($i = 0; $i -lt $paras.Count; $i++) {
    $p = Read-Para $paras[$i].Value
    $sname = if ($map.ContainsKey($p.Style)) { $map[$p.Style] } else { '' }
    if ($sname -match '^heading\s+(\d)$') {
        $headings += [pscustomobject]@{ Index = $i; Level = [int]$Matches[1]; Text = $p.Text.Trim() }
    }
}

# ---------- pass 2: compute replacements ----------
$chapter = 0
$step = 0
# NOTE: OrderedDictionary in Windows PowerShell 5.1 rejects integer keys
# (it treats them as positional indices), so keys are zero-padded strings to
# keep both assignment and numeric ordering working.
$renames = [ordered]@{}
$log = New-Object System.Collections.Generic.List[string]

foreach ($h in $headings) {

    # chapter headings: CJK numeral + enumeration comma
    if ($h.Level -eq 2 -and $h.Text.Length -ge 2 -and $CJK -contains $h.Text[0] -and $h.Text[1] -eq $IDEOGRAPHIC_COMMA) {
        $chapter++
        $step = 0
        $new = $CJK[$chapter - 1] + $h.Text.Substring(1)
        if ($new -ne $h.Text) {
            $renames[('{0:D5}' -f $h.Index)] = $new
        }
        $log.Add(('chapter ' + $chapter + ': ' + $h.Text + ' -> ' + $new))
        continue
    }

    # step headings: "Step N:" (CJK word), optionally followed by the CJK step char
    $head = $h.Text.Trim()
    if ($h.Level -eq 3 -and $head.StartsWith($STEP_WORD)) {
        $rest = $head.Substring($STEP_WORD.Length).TrimStart()
        if ($rest.StartsWith($STEP_CHAR)) { $rest = $rest.Substring(1).TrimStart() }
        # now expect: digits, then the fullwidth colon
        if ($rest -match '^([0-9]+)\s*' + [regex]::Escape($FULLWIDTH_COLON) + '\s*(.*)$') {
            $step++
            $new = $STEP_WORD + ' ' + $step + $FULLWIDTH_COLON + $Matches[2]
            if ($new -ne $h.Text) {
                $renames[('{0:D5}' -f $h.Index)] = $new
            }
            $log.Add(('step ' + $step + ': ' + $h.Text + ' -> ' + $new))
        }
    }
}

# ---------- pass 3: rebuild document.xml ----------
$sb = New-Object System.Text.StringBuilder
$last = 0
foreach ($key in @($renames.Keys | Sort-Object)) {
    $idx = [int]$key
    $m = $paras[$idx]
    $sb.Append($doc.Substring($last, $m.Index - $last)) | Out-Null
    $p = Read-Para $m.Value
    $sb.Append((New-HeadingXml $p.Style $renames[$key] $p.PPr)) | Out-Null
    $last = $m.Index + $m.Length
}
$sb.Append($doc.Substring($last)) | Out-Null
$newDoc = $sb.ToString()

# ---------- write package, replacing only word/document.xml ----------
if ([string]::IsNullOrEmpty($OutputPath)) { $OutputPath = $src }
if ($OutputPath -ne $src) { Copy-Item -LiteralPath $src -Destination $OutputPath -Force }

$full = (Resolve-Path $OutputPath).Path
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$tmp = $full + '.new'

$fsOut = [System.IO.File]::Open($tmp, [System.IO.FileMode]::CreateNew)
$zipOut = New-Object System.IO.Compression.ZipArchive($fsOut, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    $zipIn = [System.IO.Compression.ZipFile]::OpenRead($full)
    try {
        foreach ($en in $zipIn.Entries) {
            # Preserve directory entries too: some producers emit zero-length
            # folder records, and dropping them would alter the package shape.
            if ($en.FullName.EndsWith('/')) {
                $zipOut.CreateEntry($en.FullName, [System.IO.Compression.CompressionLevel]::NoCompression) | Out-Null
                continue
            }
            $newEntry = $zipOut.CreateEntry($en.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $os = $newEntry.Open()
            try {
                if ($en.FullName -eq 'word/document.xml') {
                    $bytes = $utf8NoBom.GetBytes($newDoc)
                    $os.Write($bytes, 0, $bytes.Length)
                }
                else {
                    $is = $en.Open()
                    try { $is.CopyTo($os) } finally { $is.Dispose() }
                }
            }
            finally { $os.Dispose() }
        }
    }
    finally { $zipIn.Dispose() }
}
finally {
    $zipOut.Dispose()
    $fsOut.Dispose()
}
Move-Item -LiteralPath $tmp -Destination $full -Force

# ---------- report ----------
Write-Output ('Total headings found : ' + $headings.Count)
Write-Output ('Headings renumbered  : ' + $renames.Count)
foreach ($line in $log) { Write-Output ('  ' + $line) }
