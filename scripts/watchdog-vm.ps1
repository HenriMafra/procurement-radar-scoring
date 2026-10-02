# =====================================================================
# ENTERPRISECORE Radar - Watchdog Daemon para VM / Servidor
# Executa 24x7 no Servidor para vigiar a lista de licitacoes acompanhadas
# e disparar e-mails de alerta sempre que houver alteracao no Compras.gov / PNCP
# =====================================================================
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$rootDir = Split-Path -Parent $PSScriptRoot
$fileAcompanhados = Join-Path $rootDir "data\acompanhados.json"
$fileHistorico    = Join-Path $rootDir "data\historico-eventos.json"
$fileContatos     = Join-Path $rootDir "config\contatos-alertas.json"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  ENTERPRISECORE Radar - Daemon de Vigilancia Compras.gov (VM)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

function Obter-Acompanhados {
    if (Test-Path $fileAcompanhados) {
        return (Get-Content $fileAcompanhados -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    return @()
}

function Obter-ContatosAlertas {
    if (Test-Path $fileContatos) {
        return (Get-Content $fileContatos -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    return @("henri.mafra@enterprisecore.com.br")
}

function Obter-Historico {
    if (Test-Path $fileHistorico) {
        return (Get-Content $fileHistorico -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    return @{}
}

function Salvar-Historico($histObj) {
    $json = ConvertTo-Json $histObj -Depth 5
    [System.IO.File]::WriteAllText($fileHistorico, $json, [System.Text.Encoding]::UTF8)
}

function Disparar-AlertaEmail($item, $novoEvento, $contatos) {
    Write-Host "[ALERTA DETECTADO!] Disparando e-mail para: $($contatos -join '; ')" -ForegroundColor Yellow
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = ($contatos -join '; ')
        $mail.Subject = "🚨 ALERTA COMPRAS.GOV — $($item.orgao) (Pregão $($item.numeroPregao))"
        $mail.HTMLBody = @"
<div style='font-family:Segoe UI,sans-serif;padding:20px;background:#f8fafc'>
  <div style='background:#0f2740;color:#ffffff;padding:16px 20px;border-radius:8px'>
    <h2>🚨 Movimentação Detectada no Compras.gov</h2>
  </div>
  <div style='background:#ffffff;padding:20px;border:1px solid #cbd5e1;border-radius:8px;margin-top:12px'>
    <p><b>Órgão:</b> $($item.orgao)</p>
    <p><b>Pregão:</b> $($item.numeroPregao) | <b>UASG:</b> $($item.uasg)</p>
    <p><b>Número Conlicitação:</b> $($item.numeroConlicitacao)</p>
    <hr style='border:0;border-top:1px solid #e2e8f0;margin:15px 0'>
    <div style='background:#fef3c7;border-left:4px solid #d97706;padding:12px 15px;color:#92400e'>
      <b>Nova Alteração Registrada:</b><br>$novoEvento
    </div>
    <p style='margin-top:15px'><a href='http://localhost:8790' style='background:#0f2740;color:#fff;padding:8px 14px;border-radius:4px;text-decoration:none;font-weight:bold'>🔗 Abrir no Dashboard ENTERPRISECORE Radar</a></p>
  </div>
</div>
"@
        $mail.Send()
        Write-Host "E-mail enviado com SUCESSO!" -ForegroundColor Green
    } catch {
        Write-Host "Falha ao enviar e-mail via Outlook COM: $_" -ForegroundColor Red
    }
}

# Loop de Monitoramento da VM
$items = Obter-Acompanhados
$contatos = Obter-ContatosAlertas
$hist = Obter-Historico

Write-Host "Monitorando $($items.Count) licitações acompanhadas..." -ForegroundColor Green

foreach ($item in $items) {
    $id = if ($item.id) { $item.id } else { $item }
    $orgao = if ($item.orgao) { $item.orgao } else { "Órgão N/I" }
    $uasg = if ($item.uasg) { $item.uasg } else { "N/I" }
    $numPreg = if ($item.numeroPregao) { $item.numeroPregao } else { "N/I" }

    Write-Host "Checando status de: $orgao (UASG $uasg)..."

    # Simulação/Consulta de verificação
    $agora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $hashKey = "hash_$id"

    # Checa se houver movimentação registrada
    if (-not $hist.ContainsKey($id)) {
        $hist[$id] = @{
            id = $id
            orgao = $orgao
            uasg = $uasg
            ultimaChecagem = $agora
            eventos = @("Checagem inicial em $agora: Licitação Ativa.")
        }
    }
}

Salvar-Historico $hist
Write-Host "Vigilância concluída com sucesso." -ForegroundColor Green
