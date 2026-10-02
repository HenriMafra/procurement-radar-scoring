# =====================================================================
#  verificar-ambiente.ps1
#  Roda no PC NOVO e diz, em portugues claro, o que esta pronto e o que
#  falta para o ENTERPRISECORE Radar funcionar 100%.
#  Nao altera nada: so verifica e explica.
# =====================================================================
$ErrorActionPreference = 'SilentlyContinue'
$raiz = Split-Path -Parent $PSScriptRoot

$ok = 0; $aviso = 0; $erro = 0
function Bom($t)  { Write-Host "  [OK]     $t" -ForegroundColor Green;  $script:ok++ }
function Meio($t) { Write-Host "  [AVISO]  $t" -ForegroundColor Yellow; $script:aviso++ }
function Ruim($t) { Write-Host "  [FALTA]  $t" -ForegroundColor Red;    $script:erro++ }
function Titulo($t){ Write-Host ""; Write-Host "== $t ==" -ForegroundColor Cyan }

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  ENTERPRISECORE Radar - verificacao do ambiente" -ForegroundColor Cyan
Write-Host "  Pasta: $raiz"
Write-Host "============================================================" -ForegroundColor Cyan

# ---------------------------------------------------------------- 1
Titulo "OBRIGATORIO - sem isto o Radar nao roda"

$node = Get-Command node -ErrorAction SilentlyContinue
if ($node) {
  $v = (& node --version)
  $maior = 0; if ($v -match 'v(\d+)') { $maior = [int]$Matches[1] }
  if ($maior -ge 18) { Bom "Node.js $v" }
  else { Meio "Node.js $v e antigo. Recomendado 18 ou superior: https://nodejs.org" }
} else {
  Ruim "Node.js nao encontrado. Baixe o instalador LTS em https://nodejs.org e reinstale."
}

if ($PSVersionTable.PSVersion.Major -ge 5) { Bom "PowerShell $($PSVersionTable.PSVersion)" }
else { Ruim "PowerShell muito antigo (precisa 5.1+)" }

if (Test-Path "HKLM:\SOFTWARE\Classes\Excel.Application") { Bom "Excel instalado (usado para LER os boletins)" }
else { Ruim "Excel nao encontrado. O processamento dos boletins depende dele para ler os .xlsx." }

# ---------------------------------------------------------------- 2
Titulo "PARA O E-MAIL (.OFT e abrir no Outlook)"

if (Test-Path "HKLM:\SOFTWARE\Classes\Outlook.Application") {
  Bom "Outlook classico instalado"
  $novo = Get-AppxPackage -Name "Microsoft.OutlookForWindows" -ErrorAction SilentlyContinue
  if ($novo) { Meio "O 'novo Outlook' tambem esta instalado. Ele NAO permite automacao: mantenha o Outlook classico como padrao." }
} else {
  Meio "Outlook classico nao encontrado. Tudo funciona, menos gerar .OFT e abrir e-mail. O 'novo Outlook' (do Windows 11) nao serve."
}

# ---------------------------------------------------------------- 3
Titulo "ESTRUTURA DO PROJETO"

foreach ($p in @('app\server.js','app\ui.html','app\compras-api.js','scripts\processar-boletins.ps1',
                 'scripts\gerar-oft-lote.ps1','scripts\enviar-outlook.ps1','config\escopo-enterprisecore.json',
                 'config\contatos.json','config\ambiente.json')) {
  if (Test-Path (Join-Path $raiz $p)) { Bom $p } else { Ruim "$p (arquivo essencial ausente)" }
}

$bolDir = Join-Path $raiz 'data\boletins'
if (Test-Path $bolDir) {
  $n = @(Get-ChildItem $bolDir -Filter *.xlsx -ErrorAction SilentlyContinue).Count
  if ($n -gt 0) { Bom "data\boletins com $n arquivo(s) .xlsx" }
  else { Meio "data\boletins esta vazia. Solte ali os .xlsx baixados do Conlicitacao." }
} else {
  New-Item -ItemType Directory -Path $bolDir -Force | Out-Null
  Meio "data\boletins nao existia e foi criada agora."
}

if (Test-Path (Join-Path $raiz 'app\logo_b64.txt')) { Bom "Logo ENTERPRISECORE para o e-mail (logo_b64.txt)" }
else { Meio "app\logo_b64.txt ausente: o e-mail sai sem a logo (funciona mesmo assim)." }

# ---------------------------------------------------------------- 4
Titulo "OPCIONAL - Painel de ROs e Painel de Scripts"

$amb = @{}
try { $amb = Get-Content (Join-Path $raiz 'config\ambiente.json') -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}

# Python: descarta o atalho de 0 byte da Microsoft Store
function PythonValido($p) {
  if (-not $p) { return $false }
  if (-not (Test-Path $p)) { return $false }
  if ($p -match 'WindowsApps') { return $false }
  return ((Get-Item $p).Length -gt 0)
}
$cands = @()
if ($amb.python) { $cands += $amb.python }
$cands += @(
  (Join-Path $env:USERPROFILE 'Python\python.exe'),
  (Join-Path $env:USERPROFILE 'AppData\Local\Programs\Python\Python313\python.exe'),
  (Join-Path $env:USERPROFILE 'AppData\Local\Programs\Python\Python312\python.exe'),
  (Join-Path $env:USERPROFILE 'AppData\Local\Programs\Python\Python311\python.exe'),
  'C:\Python313\python.exe','C:\Python312\python.exe','C:\Python311\python.exe'
)
$py = $cands | Where-Object { PythonValido $_ } | Select-Object -First 1
if ($py) { Bom "Python encontrado: $py" }
else { Meio "Python nao encontrado. So afeta o Painel de ROs e o Painel de Scripts; o Radar de Boletins funciona sem ele." }

$stub = Join-Path $env:USERPROFILE 'AppData\Local\Microsoft\WindowsApps\python.exe'
if ((Test-Path $stub) -and ((Get-Item $stub).Length -eq 0)) {
  Meio "Existe o atalho falso do Python da Microsoft Store no PATH. O Radar ja o ignora, mas se voce rodar python na mao vai dar erro 9009."
}

$pastaROs = $amb.painelROs.pasta
if (-not $pastaROs) {
  $pastaROs = @((Join-Path $env:USERPROFILE 'Gestao_ROs_Bitrix'),
                (Join-Path $env:USERPROFILE 'Desktop\Gestao_ROs_Bitrix')) |
              Where-Object { Test-Path (Join-Path $_ 'app.py') } | Select-Object -First 1
}
if ($pastaROs -and (Test-Path (Join-Path $pastaROs 'app.py'))) {
  Bom "Painel de ROs: $pastaROs"
  if ($py) {
    $temStreamlit = & $py -c "import streamlit" 2>&1
    if ($LASTEXITCODE -eq 0) { Bom "Streamlit instalado" }
    else { Meio "Streamlit nao instalado. Rode:  `"$py`" -m pip install streamlit" }
  }
} else {
  Meio "Painel de ROs nao encontrado. A aba 2 vai avisar; o resto funciona. Se a pasta existir com outro nome, preencha 'painelROs.pasta' em config\ambiente.json."
}

# ---------------------------------------------------------------- 5
Titulo "DADOS PESSOAIS - revise antes de usar"

$cont = Join-Path $raiz 'config\contatos.json'
if (Test-Path $cont) {
  $txt = Get-Content $cont -Raw -Encoding UTF8
  $emails = [regex]::Matches($txt, '[\w\.\-]+@[\w\.\-]+') | ForEach-Object { $_.Value } | Select-Object -Unique
  if ($emails) { Meio "config\contatos.json tem $($emails.Count) e-mail(s) do dono anterior. Troque pelos seus: $($emails -join ', ')" }
}
$log = Join-Path $raiz 'data\enviados-log.json'
if (Test-Path $log) {
  $n = 0; try { $n = @((Get-Content $log -Raw -Encoding UTF8 | ConvertFrom-Json)).Count } catch {}
  if ($n -gt 0) { Meio "data\enviados-log.json tem $n envio(s) do dono anterior. Para zerar, troque o conteudo por:  []" }
}

# ---------------------------------------------------------------- 6
Titulo "PORTA DE REDE"
$porta = 8790
if ($amb.porta) { $porta = [int]$amb.porta }
$emUso = Get-NetTCPConnection -LocalPort $porta -State Listen -ErrorAction SilentlyContinue
if ($emUso) {
  $proc = Get-Process -Id $emUso[0].OwningProcess -ErrorAction SilentlyContinue
  Meio "Porta $porta ja esta em uso por '$($proc.ProcessName)' (PID $($proc.Id)). Se nao for o proprio Radar, mude 'porta' em config\ambiente.json."
} else { Bom "Porta $porta livre" }

# ---------------------------------------------------------------- fim
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ("  RESUMO:  {0} OK   {1} avisos   {2} faltando" -f $ok, $aviso, $erro)
if ($erro -eq 0) {
  Write-Host "  Pode usar: de duplo-clique em Abrir-App.cmd" -ForegroundColor Green
} else {
  Write-Host "  Resolva os itens [FALTA] antes de usar." -ForegroundColor Red
}
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
