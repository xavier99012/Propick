<#
Reconstruye Propick.html a partir de src/app.js + build/template.shell.txt.
Reemplaza unicamente la linea del "__bundler/template" (linea 391) dentro del
bundle original; el resto del archivo (manifest, wrapper, assets) no se toca.
#>
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [switch]$VerifyOnly
)

$htmlPath   = Join-Path $RepoRoot "Propick.html"
$appJsPath  = Join-Path $RepoRoot "src\app.js"
$shellPath  = Join-Path $RepoRoot "build\template.shell.txt"
$marker     = "%%PROPICK_APP_CODE%%"
$templateLineIndex = 390  # linea 391 (0-based)

function ConvertTo-BundlerJsonString {
    param([string]$s)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    foreach ($ch in $s.ToCharArray()) {
        switch ($ch) {
            '"'  { [void]$sb.Append('\"'); continue }
            '\'  { [void]$sb.Append('\\'); continue }
            "`n" { [void]$sb.Append('\n'); continue }
            "`r" { [void]$sb.Append('\r'); continue }
            "`t" { [void]$sb.Append('\t'); continue }
            default {
                if ([int]$ch -lt 0x20) {
                    [void]$sb.Append([string]::Format('\u{0:x4}', [int]$ch))
                } else {
                    [void]$sb.Append($ch)
                }
            }
        }
    }
    [void]$sb.Append('"')
    return ($sb.ToString() -replace '</', ('<' + '\' + 'u002F'))
}

$appCode = Get-Content $appJsPath -Raw -Encoding UTF8
$shell   = Get-Content $shellPath -Raw -Encoding UTF8
if (-not $shell.Contains($marker)) {
    throw "Marker $marker not found in $shellPath"
}

$decoded = $shell.Replace($marker, $appCode)
$newTemplateLine = ConvertTo-BundlerJsonString($decoded)

$lines = Get-Content $htmlPath -Encoding UTF8
if ($VerifyOnly) {
    $match = ($newTemplateLine -ceq $lines[$templateLineIndex])
    Write-Host "Verify only. Line 391 matches current Propick.html: $match"
    return
}

$lines[$templateLineIndex] = $newTemplateLine
$outText = [string]::Join("`n", $lines)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($htmlPath, $outText, $utf8NoBom)
Write-Host "Propick.html regenerado desde src\app.js"
