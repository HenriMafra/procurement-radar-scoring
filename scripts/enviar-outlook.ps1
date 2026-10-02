# =====================================================================
#  enviar-outlook.ps1
#  Abre o Outlook classico com um e-mail HTML ja montado (To/CC/Assunto/
#  Corpo), PRONTO PARA REVISAR E ENVIAR. Nao envia sozinho - quem clica
#  "Enviar" e o usuario (usa .Display, nao .Send).
#
#  Recebe os dados por um arquivo JSON (para evitar problemas de escape
#  com HTML grande na linha de comando).
#  JSON esperado: { "to": "...", "cc": "...", "subject": "...", "html": "..." }
# =====================================================================
param(
    [Parameter(Mandatory = $true)][string]$JsonPath
)

$ErrorActionPreference = 'Stop'

$outlook = $null
$mail = $null
$inspector = $null
try {
    if (-not (Test-Path -LiteralPath $JsonPath)) {
        throw "Arquivo de dados nao encontrado: $JsonPath"
    }

    $raw  = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8
    $data = $raw | ConvertFrom-Json

    # olMailItem = 0
    $outlook = New-Object -ComObject Outlook.Application
    $mail    = $outlook.CreateItem(0)

    if ($data.PSObject.Properties.Name -contains 'to'      -and $data.to)      { $mail.To      = [string]$data.to }
    if ($data.PSObject.Properties.Name -contains 'cc'      -and $data.cc)      { $mail.CC      = [string]$data.cc }
    if ($data.PSObject.Properties.Name -contains 'subject' -and $data.subject) { $mail.Subject = [string]$data.subject }
    if ($data.PSObject.Properties.Name -contains 'html'    -and $data.html)    { $mail.HTMLBody = [string]$data.html }

    # Abre a janela de composicao (nao-modal). O usuario revisa e envia.
    $mail.Display($false)

    # BUG achado em 03/09/2026: quando o Node spawna este script, o COM do
    # Outlook por vezes cria uma SEGUNDA instancia (novo OUTLOOK.EXE) em vez de
    # anexar na janela que o usuario ja tem aberta - a janela do e-mail entao
    # existe, mas nunca aparece na tela (usuario reporta "nao abre"). Forcar o
    # Inspector para frente resolve independente de qual instancia o COM usou.
    $inspector = $mail.GetInspector()
    if ($inspector) { $inspector.Activate() }

    Write-Output "OK"
}
catch {
    Write-Error ("FALHA: " + $_.Exception.Message)
    exit 1
}
finally {
    # Libera os objetos COM intermediarios (nao so o Outlook.Application) para
    # nao repetir o mesmo problema de "processo zumbi" ja visto no Excel.
    if ($inspector) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($inspector) | Out-Null }
    if ($mail)      { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail)      | Out-Null }
    if ($outlook)   { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)   | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
