# =====================================================================
#  preparar-entrega.ps1
#  Cria uma COPIA limpa do projeto, pronta para entregar a outra pessoa.
#  NAO altera a sua instalacao: so copia e limpa o que e pessoal.
# =====================================================================
param(
  [string]$Destino = ""
)
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot

if (-not $Destino) {
  $Destino = Join-Path (Split-Path -Parent $raiz) ("ENTERPRISECORE-Radar-Entrega-" + (Get-Date -Format 'yyyy-MM-dd'))
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Preparando copia para entrega" -ForegroundColor Cyan
Write-Host "  Origem:  $raiz"
Write-Host "  Destino: $Destino"
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

if (Test-Path $Destino) {
  Write-Host "A pasta de destino ja existe. Apague ou escolha outra:" -ForegroundColor Red
  Write-Host "  $Destino" -ForegroundColor Red
  exit 1
}

# ---------- 1. copia ----------
Write-Host "1/4  Copiando arquivos..." -ForegroundColor Cyan
$excluirPastas = @('node_modules', '.git', 'scratch', '__pycache__')
robocopy $raiz $Destino /E /NFL /NDL /NJH /NJS /NP /XD $excluirPastas | Out-Null
Write-Host "     copiado."

# ---------- 2. limpa dados pessoais ----------
Write-Host "2/4  Limpando dados do dono anterior..." -ForegroundColor Cyan
$utf8 = New-Object System.Text.UTF8Encoding($false)

# historico de envios
$log = Join-Path $Destino 'data\enviados-log.json'
if (Test-Path $log) { [System.IO.File]::WriteAllText($log, "[]", $utf8); Write-Host "     enviados-log.json zerado" }

# vigilancia de pregoes (aba removida, mas os dados ficam)
foreach ($f in @('data\acompanhados.json','data\historico-eventos.json')) {
  $p = Join-Path $Destino $f
  if (Test-Path $p) {
    $vazio = if ($f -match 'acompanhados') { "[]" } else { "{}" }
    [System.IO.File]::WriteAllText($p, $vazio, $utf8)
    Write-Host "     $f zerado"
  }
}

# resultado do processamento e planilha - a pessoa gera os dela
foreach ($f in @('data\dados.json','data\Planilha-ENTERPRISECORE-Licitacoes.xlsx')) {
  $p = Join-Path $Destino $f
  if (Test-Path $p) { Remove-Item $p -Force; Write-Host "     $f removido (sera gerado no primeiro processamento)" }
}

# boletins do dono anterior
$bol = Join-Path $Destino 'data\boletins'
if (Test-Path $bol) {
  $n = @(Get-ChildItem $bol -Filter *.xlsx).Count
  Get-ChildItem $bol -Filter *.xlsx | Remove-Item -Force
  Write-Host "     $n boletim(ns) removido(s) de data\boletins"
}

# backups e temporarios que nao servem para nada la
Get-ChildItem $Destino -Recurse -Include '*.bak','*.bak2','*.tmp' -File -ErrorAction SilentlyContinue |
  ForEach-Object { Remove-Item $_.FullName -Force }

# ambiente.json volta ao padrao (tudo em branco = auto-deteccao)
$amb = Join-Path $Destino 'config\ambiente.json'
if (Test-Path $amb) {
  $txt = [System.IO.File]::ReadAllText($amb, [System.Text.Encoding]::UTF8)
  $txt = $txt -replace '"python":\s*"[^"]*"', '"python": ""'
  $txt = $txt -replace '"pasta":\s*"[^"]*"', '"pasta": ""'
  $txt = $txt -replace '"painelScripts":\s*"[^"]*"', '"painelScripts": ""'
  [System.IO.File]::WriteAllText($amb, $txt, $utf8)
  Write-Host "     ambiente.json redefinido para auto-deteccao"
}

# ---------- 3. aviso sobre contatos ----------
Write-Host "3/4  Conferindo contatos..." -ForegroundColor Cyan
$cont = Join-Path $Destino 'config\contatos.json'
if (Test-Path $cont) {
  $txt = [System.IO.File]::ReadAllText($cont, [System.Text.Encoding]::UTF8)
  $emails = [regex]::Matches($txt, '[\w\.\-]+@[\w\.\-]+') | ForEach-Object { $_.Value } | Select-Object -Unique
  Write-Host "     ATENCAO: config\contatos.json mantem $($emails.Count) e-mail(s) da equipe." -ForegroundColor Yellow
  Write-Host "     Se a outra pessoa NAO deve ter essa lista, edite ou apague o arquivo." -ForegroundColor Yellow
}

# ---------- 4. resumo ----------
Write-Host "4/4  Pronto." -ForegroundColor Cyan
$tam = (Get-ChildItem $Destino -Recurse -File | Measure-Object Length -Sum).Sum
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Copia limpa criada:" -ForegroundColor Green
Write-Host "  $Destino"
Write-Host ("  Tamanho: {0} MB" -f [math]::Round($tam/1MB,1))
Write-Host ""
Write-Host "  Proximos passos:"
Write-Host "   1. Revise config\contatos.json (e-mails da equipe)."
Write-Host "   2. Compacte a pasta em .zip e envie."
Write-Host "   3. Diga para a pessoa ler o INSTALACAO.md primeiro."
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
