Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-Tool {
  param([Parameter(Mandatory)][string]$Name)
  $cmd = Get-Command $Name -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Path }
  $paths = switch -Regex ($Name) {
    '^msedge\.exe$' { @("$Env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe","$Env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") }
    '^chrome\.exe$' { @("$Env:ProgramFiles\Google\Chrome\Application\chrome.exe","$Env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe") }
    '^wkhtmltopdf$' { @("$Env:ProgramFiles\wkhtmltopdf\bin\wkhtmltopdf.exe","$Env:ProgramFiles(x86)\wkhtmltopdf\bin\wkhtmltopdf.exe") }
    '^soffice$'     { @("$Env:ProgramFiles\LibreOffice\program\soffice.exe","$Env:ProgramFiles(x86)\LibreOffice\program\soffice.exe") }
    default { @() }
  }
  foreach($p in $paths){ if(Test-Path $p){ return $p } }
  return $null
}
function Ensure-Directory { param([string]$Path) if($Path -and -not(Test-Path $Path)){ New-Item -ItemType Directory -Force -Path $Path | Out-Null } }
function New-CommonPaths { param([string]$CommonPdfDir = (Join-Path $HOME 'Documents/PDF'), [string]$CommonDocxDir = (Join-Path $HOME 'Documents/DOC')); Ensure-Directory $CommonPdfDir; Ensure-Directory $CommonDocxDir; [pscustomobject]@{Pdf=$CommonPdfDir;Docx=$CommonDocxDir} }

# ---------- HTML → PDF ----------
function Convert-HtmlToPdf {
  [CmdletBinding()] param(
    [Parameter(Mandatory)][string]$InputHtml,
    [Parameter(Mandatory)][string]$OutputPdf,
    [switch]$NoHeaderFooter,
    [double]$Scale = 1.0
  )
  if(-not(Test-Path -LiteralPath $InputHtml)){ throw "Missing HTML: $InputHtml" }
  $outDir = Split-Path -LiteralPath $OutputPdf -Parent; Ensure-Directory $outDir
  $edge=Resolve-Tool msedge.exe; $chrome=Resolve-Tool chrome.exe; $wk=Resolve-Tool wkhtmltopdf
  if($edge -or $chrome){
    $browser = $edge ?? $chrome
    $url = "file:///" + ($InputHtml -replace '\\','/')
    $args = @('--headless',"--print-to-pdf=""$OutputPdf""")
    if($NoHeaderFooter){ $args += '--no-pdf-header-footer' }
    if($Scale -ne 1.0){ $args += "--force-device-scale-factor=$Scale" }
    & $browser @args "$url" | Out-Null
    if(-not(Test-Path $OutputPdf)){ throw "Chromium failed: $InputHtml" }
    return Get-Item -LiteralPath $OutputPdf
  }
  if($wk){
    & $wk "$InputHtml" "$OutputPdf" | Out-Null
    if(-not(Test-Path $OutputPdf)){ throw "wkhtmltopdf failed: $InputHtml" }
    return Get-Item -LiteralPath $OutputPdf
  }
  throw "No PDF engine found (Edge/Chrome or wkhtmltopdf)."
}

# ---------- HTML → DOCX ----------
function Convert-HtmlToWord {
  [CmdletBinding()] param(
    [Parameter(Mandatory)][string]$InputHtml,
    [Parameter(Mandatory)][string]$OutputDocx
  )
  if(-not(Test-Path -LiteralPath $InputHtml)){ throw "Missing HTML: $InputHtml" }
  Ensure-Directory (Split-Path -LiteralPath $OutputDocx -Parent)
  $pandoc = Resolve-Tool pandoc
  if($pandoc){
    & $pandoc -s "$InputHtml" -o "$OutputDocx" | Out-Null
    if(-not(Test-Path $OutputDocx)){ throw "Pandoc failed: $InputHtml" }
    return Get-Item -LiteralPath $OutputDocx
  }
  $word=$null
  try{
    $word = New-Object -ComObject "Word.Application"; $word.Visible=$false
    $doc = $word.Documents.Open((Resolve-Path -LiteralPath $InputHtml).Path)
    $doc.SaveAs([ref]$OutputDocx,[ref]16); $doc.Close(); $word.Quit()
    if(-not(Test-Path $OutputDocx)){ throw "Word COM failed." }
    return Get-Item -LiteralPath $OutputDocx
  } catch { if($word){ try{$word.Quit()}catch{} } }
  $soffice = Resolve-Tool soffice
  if($soffice){
    $out = Split-Path -LiteralPath $OutputDocx -Parent
    & $soffice --headless --convert-to docx --outdir "$out" "$InputHtml" | Out-Null
    if(-not(Test-Path $OutputDocx)){ throw "LibreOffice failed." }
    return Get-Item -LiteralPath $OutputDocx
  }
  throw "No DOCX engine found (pandoc, Word, or LibreOffice)."
}

# ---------- DOC/DOCX → PDF ----------
function Convert-DocToPdf {
  [CmdletBinding()] param([Parameter(Mandatory)][string]$InputDoc,[Parameter(Mandatory)][string]$OutputPdf)
  if(-not(Test-Path -LiteralPath $InputDoc)){ throw "Missing DOC/DOCX: $InputDoc" }
  Ensure-Directory (Split-Path -LiteralPath $OutputPdf -Parent)
  $word=$null
  try{
    $word = New-Object -ComObject "Word.Application"; $word.Visible=$false
    $doc = $word.Documents.Open((Resolve-Path -LiteralPath $InputDoc).Path)
    $doc.SaveAs([ref]$OutputPdf,[ref]17); $doc.Close(); $word.Quit()  # 17 = PDF
    if(-not(Test-Path $OutputPdf)){ throw "Word COM failed." }
    return Get-Item -LiteralPath $OutputPdf
  } catch { if($word){ try{$word.Quit()}catch{} } }
  $soffice=Resolve-Tool soffice
  if($soffice){
    $out=Split-Path -LiteralPath $OutputPdf -Parent
    & $soffice --headless --convert-to pdf --outdir "$out" "$InputDoc" | Out-Null
    if(-not(Test-Path $OutputPdf)){ throw "LibreOffice failed." }
    return Get-Item -LiteralPath $OutputPdf
  }
  throw "No engine for DOC→PDF (need Word or LibreOffice)."
}

# ---------- PDF → DOCX ----------
function Convert-PdfToDocx {
  [CmdletBinding()] param([Parameter(Mandatory)][string]$InputPdf,[Parameter(Mandatory)][string]$OutputDocx)
  if(-not(Test-Path -LiteralPath $InputPdf)){ throw "Missing PDF: $InputPdf" }
  Ensure-Directory (Split-Path -LiteralPath $OutputDocx -Parent)
  $word=$null
  try{
    $word = New-Object -ComObject "Word.Application"; $word.Visible=$false
    $doc = $word.Documents.Open((Resolve-Path -LiteralPath $InputPdf).Path,$false,$true)
    $doc.SaveAs([ref]$OutputDocx,[ref]16); $doc.Close(); $word.Quit()
    if(-not(Test-Path $OutputDocx)){ throw "Word COM failed." }
    return Get-Item -LiteralPath $OutputDocx
  } catch { if($word){ try{$word.Quit()}catch{} } }
  $soffice=Resolve-Tool soffice
  if($soffice){
    $out=Split-Path -LiteralPath $OutputDocx -Parent
    & $soffice --headless --convert-to docx --outdir "$out" "$InputPdf" | Out-Null
    if(-not(Test-Path $OutputDocx)){ throw "LibreOffice failed." }
    return Get-Item -LiteralPath $OutputDocx
  }
  $pandoc=Resolve-Tool pandoc
  if($pandoc){
    & $pandoc -s "$InputPdf" -o "$OutputDocx" | Out-Null
    if(-not(Test-Path $OutputDocx)){ throw "Pandoc failed (use Word/LibreOffice for scanned/complex PDFs)." }
    return Get-Item -LiteralPath $OutputDocx
  }
  throw "No engine for PDF→DOCX (need Word or LibreOffice; pandoc last resort)."
}

# ---------- Bulk engine ----------
function Get-HtmlTargets { param([string]$Path,[switch]$Recurse)
  if(-not(Test-Path -LiteralPath $Path)){ throw "Path not found: $Path" }
  $it=Get-Item -LiteralPath $Path
  if($it.PSIsContainer){ Get-ChildItem -LiteralPath $it.FullName -Filter *.html -Recurse:$Recurse }
  else { if($it.Extension -notin '.html','.htm'){ throw "Not an HTML file: $Path" }; ,$it }
}
function Invoke-ConvertHtmlBulk {
  [CmdletBinding()] param(
    [Parameter(Mandatory)][string]$InputPath,
    [string]$OutDir,
    [switch]$ToPdf,
    [switch]$ToDocx,
    [switch]$Recurse,
    [switch]$Parallel,
    [int]$ThrottleLimit = 6,
    [switch]$CopyToCommon,
    [string]$CommonPdfDir = (Join-Path $HOME 'Documents/PDF'),
    [string]$CommonDocxDir = (Join-Path $HOME 'Documents/DOC'),
    [switch]$OpenAfter
  )
  if(-not $ToPdf -and -not $ToDocx){ $ToPdf=$true; $ToDocx=$true }
  $files = Get-HtmlTargets -Path $InputPath -Recurse:$Recurse
  if(-not $files){ Write-Host "No HTML found." -Foreground Yellow; return }
  if(-not $OutDir){
    $root=Get-Item -LiteralPath $InputPath
    $OutDir = if($root.PSIsContainer){ Join-Path $root.FullName "_converted" } else { Split-Path -LiteralPath $root.FullName -Parent }
  }
  Ensure-Directory $OutDir
  if($CopyToCommon){ New-CommonPaths -CommonPdfDir $CommonPdfDir -CommonDocxDir $CommonDocxDir | Out-Null }

  $results=[System.Collections.Concurrent.ConcurrentBag[object]]::new()
  $worker = {
    param($f,$ToPdf,$ToDocx,$OutDir,$CopyToCommon,$CommonPdfDir,$CommonDocxDir)
    try{
      $base=[IO.Path]::GetFileNameWithoutExtension($f.Name)
      if($ToPdf){
        $pdf=Join-Path $OutDir ($base + ".pdf"); Convert-HtmlToPdf -InputHtml $f.FullName -OutputPdf $pdf -NoHeaderFooter
        if($CopyToCommon){ Copy-Item $pdf (Join-Path $CommonPdfDir ([IO.Path]::GetFileName($pdf))) -Force }
        [pscustomobject]@{File=$f.FullName;Output=$pdf;Type='PDF';Status='OK'}
      }
      if($ToDocx){
        $docx=Join-Path $OutDir ($base + ".docx"); Convert-HtmlToWord -InputHtml $f.FullName -OutputDocx $docx
        if($CopyToCommon){ Copy-Item $docx (Join-Path $CommonDocxDir ([IO.Path]::GetFileName($docx))) -Force }
        [pscustomobject]@{File=$f.FullName;Output=$docx;Type='DOCX';Status='OK'}
      }
    } catch { [pscustomobject]@{File=$f.FullName;Output=$null;Type='';Status="FAIL: $($_.Exception.Message)"} }
  }

  $isPS7 = $PSVersionTable.PSVersion.Major -ge 7
  if($Parallel -and $isPS7){
    $files | ForEach-Object -Parallel { & $using:worker $_ $using:ToPdf $using:ToDocx $using:OutDir $using:CopyToCommon $using:CommonPdfDir $using:CommonDocxDir } -ThrottleLimit $ThrottleLimit |
      ForEach-Object { $results.Add($_); if($_.Status -like 'OK*'){ Write-Host "[OK] $($_.Output)" -Foreground Green } else { Write-Host "[FAIL] $($_.File): $($_.Status)" -Foreground Red } }
  } else {
    foreach($f in $files){
      $objs = & $worker $f $ToPdf $ToDocx $OutDir $CopyToCommon $CommonPdfDir $CommonDocxDir
      foreach($o in $objs){ $results.Add($o); if($o.Status -like 'OK*'){ Write-Host "[OK] $($o.Output)" -Foreground Green } else { Write-Host "[FAIL] $($o.File): $($o.Status)" -Foreground Red } }
    }
  }

  $log = Join-Path $OutDir ("conversion_log_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
  $results | Export-Csv -NoTypeInformation -Path $log
  Write-Host "`nOutputs: $OutDir" -Foreground Cyan
  if($CopyToCommon){ if($ToPdf){ Write-Host "Common PDFs:  $CommonPdfDir" -Foreground Cyan } ; if($ToDocx){ Write-Host "Common DOCX: $CommonDocxDir" -Foreground Cyan } }
  Write-Host "Log: $log" -Foreground Cyan
  if($OpenAfter){ Invoke-Item $OutDir; if($CopyToCommon){ if($ToPdf){ Invoke-Item $CommonPdfDir } ; if($ToDocx){ Invoke-Item $CommonDocxDir } } }
}

# ---------- Interactive prompt ----------
function Invoke-Html2DocPrompt {
  [CmdletBinding()] param()
  try{ Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null } catch{}
  $mode = Read-Host "Select input type: (F)ile / (D)irectory [default D]"; if(-not $mode){ $mode='D' }
  if($mode -match '^[Ff]'){
    $ofd=New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter="HTML files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*"; $ofd.Title="Select HTML file"
    if($ofd.ShowDialog() -ne 'OK'){ Write-Host "No file selected." -Foreground Yellow; return }
    $inPath=$ofd.FileName; $recurse=$false
  } else {
    $fbd=New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description="Select folder containing HTML files"
    if($fbd.ShowDialog() -ne 'OK'){ Write-Host "No folder selected." -Foreground Yellow; return }
    $inPath=$fbd.SelectedPath
    $recQ=Read-Host "Recurse subfolders? (Y/N) [default Y]"; $recurse=($recQ -eq '' -or $recQ -match '^[Yy]')
  }
  $outDir=Read-Host "Output directory (blank = default under input)"
  $fmt=Read-Host "Convert to: (1) PDF  (2) DOCX  (3) Both [default 3]"; if(-not $fmt){ $fmt='3' }
  $toPdf=$true; $toDocx=$true; switch($fmt){ '1'{$toDocx=$false} '2'{$toPdf=$false} default{} }
  $parQ=Read-Host "Use parallel mode (PS7+)? (Y/N) [default Y]"; $parallel=($parQ -eq '' -or $parQ -match '^[Yy]')
  $thrIn=Read-Host "Parallel throttle (int) [default 6]"; $thr= if($thrIn){ [int]$thrIn } else { 6 }
  $copyQ=Read-Host "Copy to common PDF/DOC folders? (Y/N) [default Y]"; $copyCommon=($copyQ -eq '' -or $copyQ -match '^[Yy]')
  $openQ=Read-Host "Open folders after conversion? (Y/N) [default Y]"; $openAfter=($openQ -eq '' -or $openQ -match '^[Yy]')

  Invoke-ConvertHtmlBulk -InputPath $inPath -OutDir $outDir -ToPdf:$toPdf -ToDocx:$toDocx -Recurse:$recurse -Parallel:$parallel -ThrottleLimit $thr -CopyToCommon:$copyCommon -OpenAfter:$openAfter
}

function docV { Invoke-Html2DocPrompt }

Export-ModuleMember -Function Convert-HtmlToPdf,Convert-HtmlToWord,Convert-DocToPdf,Convert-PdfToDocx,Invoke-ConvertHtmlBulk,Invoke-Html2DocPrompt,docV


