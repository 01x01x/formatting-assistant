<#
Export one DOCX to PDF through an isolated, hidden Microsoft Word instance.
This is an optional Windows fallback when the packaged renderer is unavailable.
It does not save or update the input DOCX and never overwrites an existing PDF.
Reference: https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputPdf
)
$ErrorActionPreference = 'Stop'
$word = $null
$documents = $null
$document = $null
try {
    $source = (Resolve-Path -LiteralPath $InputDocx).ProviderPath
    $target = [IO.Path]::GetFullPath($OutputPdf)
    if ([IO.Path]::GetExtension($source) -ine '.docx') { throw 'Only DOCX input is supported.' }
    if ([IO.Path]::GetExtension($target) -ine '.pdf') { throw 'OutputPdf must end in .pdf.' }
    if ($source -ieq $target) { throw 'The output must differ from the input.' }
    if (Test-Path -LiteralPath $target) { throw 'Output PDF already exists; choose a new path.' }
    $sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    if ($null -eq ('System.IO.Compression.ZipFile' -as [type])) {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
    }
    $zip = [IO.Compression.ZipFile]::OpenRead($source)
    try {
        if ($null -eq $zip.GetEntry('word/document.xml')) { throw 'Not a supported, unencrypted DOCX package.' }
    } finally { $zip.Dispose() }
    $parent = [IO.Path]::GetDirectoryName($target)
    [IO.Directory]::CreateDirectory($parent) | Out-Null
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3 # msoAutomationSecurityForceDisable, this instance only.
    $documents = $word.Documents
    $document = $documents.Open($source, $false, $true, $false)
    $document.Repaginate()
    $document.ExportAsFixedFormat($target, 17) # wdExportFormatPDF
    $document.Close(0) # wdDoNotSaveChanges
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    $document = $null
    if (-not (Test-Path -LiteralPath $target) -or (Get-Item -LiteralPath $target).Length -eq 0) {
        throw 'Word did not create a non-empty PDF.'
    }
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash) {
        throw 'Input hash changed unexpectedly; do not treat this export as verified.'
    }
    [pscustomobject]@{Input = $source; Pdf = $target; InputUnchanged = $true; Renderer = 'Microsoft Word'} | ConvertTo-Json
} catch {
    Write-Error ('Word PDF export failed: ' + $_.Exception.Message) -ErrorAction Continue
    exit 2
} finally {
    if ($null -ne $document) {
        try { $document.Close(0) } catch {}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    }
    if ($null -ne $documents) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($documents) | Out-Null }
    if ($null -ne $word) {
        try { $word.Quit(0) } catch {}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
}
exit 0
