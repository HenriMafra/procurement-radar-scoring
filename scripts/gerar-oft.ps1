param(
  [string]$JsonPath,
  [string]$OutOftPath
)
$ErrorActionPreference = 'Stop'
if(-not (Test-Path -LiteralPath $JsonPath)){ throw "Arquivo JSON nao encontrado: $JsonPath" }

$raw = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
$to      = [string]$raw.to
$cc      = [string]$raw.cc
$subject = [string]$raw.subject
$html    = [string]$raw.html

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail    = $outlook.CreateItem(0)
    if($to){ $mail.To = $to }
    if($cc){ $mail.CC = $cc }
    if($subject){ $mail.Subject = $subject }
    if($html){ $mail.HTMLBody = $html }

    $mail.SaveAs($OutOftPath, 2)
    $mail.Close(1)
    Write-Host "OFT_GERADO_OK"
} catch {
    Write-Host "ERRO_OFT: $($_.Exception.Message)"
} finally {
    if($outlook){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
}
