# =====================================================================
#  enviar-alerta-auto.ps1 - Disparo automatico com UTF-8 e retry COM
# =====================================================================
param(
    [Parameter(Mandatory = $true)][string]$JsonPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Test-Path -LiteralPath $JsonPath)) {
    Write-Error "Arquivo de dados nao encontrado: $JsonPath"
    exit 1
}

$raw = [System.IO.File]::ReadAllText($JsonPath, [System.Text.Encoding]::UTF8)
$data = $raw | ConvertFrom-Json

$maxRetries = 5
$enviado = $false
$ultimoErro = ""

for ($tentativa = 1; $tentativa -le $maxRetries; $tentativa++) {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail    = $outlook.CreateItem(0)

        if ($data.PSObject.Properties.Name -contains 'to'      -and $data.to)      { $mail.To      = [string]$data.to }
        if ($data.PSObject.Properties.Name -contains 'cc'      -and $data.cc)      { $mail.CC      = [string]$data.cc }
        if ($data.PSObject.Properties.Name -contains 'subject' -and $data.subject) { $mail.Subject = [string]$data.subject }
        if ($data.PSObject.Properties.Name -contains 'html'    -and $data.html)    { $mail.HTMLBody = [string]$data.html }

        $mail.Send()
        $enviado = $true
        Write-Output "OK"
        break
    }
    catch {
        $ultimoErro = $_.Exception.Message
        Start-Sleep -Milliseconds (600 * $tentativa)
    }
}

if (-not $enviado) {
    Write-Error ("FALHA NO ENVIO AUTOMATICO APOS " + $maxRetries + " TENTATIVAS: " + $ultimoErro)
    exit 1
}
