# =====================================================================
#  publicar-versao.ps1
#  Publica o estado atual no GitHub, com uma versao marcada (tag).
#  A pessoa que usa o Radar baixa sempre a ultima versao de la.
#
#  Rodar:  Publicar-Versao.cmd
#
#  O token e lido do arquivo .env indicado abaixo e NUNCA e gravado no
#  repositorio nem no .git/config.
# =====================================================================
param(
  [string]$Mensagem = '',
  [string]$Versao   = '',
  [string]$ArquivoToken = "$env:USERPROFILE\Documents\ORACLEVMS.env",
  [string]$Repo = 'HenriMafra/enterprisecore-radar'
)
# NAO usar 'Stop' aqui: o git escreve avisos normais no stderr (fim de linha
# CRLF, progresso do push) e no PowerShell 5.1 isso vira erro terminante, fazendo
# o script "falhar" com o push funcionando. Conferimos $LASTEXITCODE na mao.
$ErrorActionPreference = 'Continue'
$raiz = Split-Path -Parent $PSScriptRoot
Set-Location $raiz

function Passo($t) { Write-Host "`n>> $t" -ForegroundColor Cyan }

# ---------- token ----------
if (-not (Test-Path $ArquivoToken)) { Write-Host "Arquivo de token nao encontrado: $ArquivoToken" -ForegroundColor Red; exit 1 }
$tk = ((Get-Content $ArquivoToken | Where-Object { $_ -match '^\s*token github classic\s*=' }) -replace '^\s*token github classic\s*=\s*','').Trim()
if (-not $tk) { Write-Host "Nao achei a linha 'token github classic=' no arquivo." -ForegroundColor Red; exit 1 }

# ---------- o que mudou ----------
Passo "Conferindo o que mudou"
$mudou = git status --porcelain
if (-not $mudou) {
  Write-Host "   Nada mudou desde a ultima publicacao." -ForegroundColor Yellow
  $tagAtual = git describe --tags --abbrev=0 2>$null
  if ($tagAtual) { Write-Host "   Versao publicada continua sendo: $tagAtual" }
  exit 0
}
$mudou | ForEach-Object { Write-Host "   $_" }

# ---------- trava de seguranca ----------
Passo "Conferindo que nao vai vazar dado sensivel"
git add -A
$proibidos = @('config/contatos.json','data/dados.json','data/enviados-log.json','data/acompanhados.json','data/historico-eventos.json')
$vazando = @()
$naFila = git diff --cached --name-only
foreach ($p in $proibidos) { if ($naFila -contains $p) { $vazando += $p } }
# Barra o DADO do boletim (planilha), nao a pasta inteira: o LEIA-ME.txt existe
# so para a pasta vir no zip do GitHub, ja que git nao versiona pasta vazia.
foreach ($f in $naFila) { if ($f -like 'data/boletins/*' -and $f -like '*.xls*' -or $f -like 'data/boletins/*.csv') { $vazando += $f } }
if ($vazando.Count) {
  Write-Host "   ABORTADO. Estes arquivos nao podem ir para o GitHub:" -ForegroundColor Red
  $vazando | ForEach-Object { Write-Host "     $_" -ForegroundColor Red }
  Write-Host "   Confira o .gitignore." -ForegroundColor Red
  git reset --quiet
  exit 1
}
Write-Host "   ok, nada sensivel na fila." -ForegroundColor Green

# ---------- versao ----------
if (-not $Versao) {
  $ultima = git tag --list 'v*' | Sort-Object { [version]($_ -replace '^v','') } -ErrorAction SilentlyContinue | Select-Object -Last 1
  if ($ultima -match '^v(\d+)\.(\d+)\.(\d+)$') {
    $Versao = 'v{0}.{1}.{2}' -f $Matches[1], $Matches[2], ([int]$Matches[3] + 1)
  } else { $Versao = 'v1.0.0' }
}
if (-not $Mensagem) { $Mensagem = "Ajustes de $(Get-Date -Format 'dd/MM/yyyy')" }

Passo "Publicando $Versao"
git commit -q -m @"
$Mensagem

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
"@
git tag -a $Versao -m $Mensagem

# O git escreve progresso no stderr. No PowerShell 5.1 qualquer redirecionamento
# disso vira ErrorRecord e suja a tela (ou aborta o script). Deixar o proprio cmd
# fazer o merge dos fluxos resolve: aqui chega texto puro.
# Empurra SO a tag desta rodada, nunca "--tags". Com "--tags" o git tenta subir
# todas as tags locais de uma vez, e uma unica tag velha em conflito reprova o
# push inteiro - mesmo que o commit e a versao nova estejam corretos.
$saida = cmd /c "git push ""https://$tk@github.com/$Repo.git"" main $Versao 2>&1"
$codigo = $LASTEXITCODE
$saida | ForEach-Object { Write-Host ('   ' + ($_ -replace [regex]::Escape($tk), '***')) }

if ($codigo -ne 0) {
  Write-Host "`n   FALHA no push (codigo $codigo). A tag local $Versao foi criada;" -ForegroundColor Red
  Write-Host "   corrija o problema e rode de novo." -ForegroundColor Red
  exit 1
}

# ---------- Release ----------
# A tag sozinha fica escondida numa aba secundaria do GitHub. A Release e a
# pagina com o botao de download - e o que a pessoa que usa o Radar abre.
Passo "Criando a pagina de download (Release)"
$corpo = @{
  tag_name = $Versao
  name     = "ENTERPRISECORE Radar $Versao"
  body     = @"
$Mensagem

**Como instalar / atualizar**

- Ja tem a pasta: abra o Prompt nela e rode ``git pull``
- Primeira vez: baixe o **Source code (zip)** aqui embaixo, descompacte e leia o ``INSTALACAO.md``

Seus dados nao sao tocados na atualizacao: boletins, planilha gerada, ``contatos.json``
e o historico de envios ficam fora do repositorio de proposito.
"@
  draft      = $false
  prerelease = $false
} | ConvertTo-Json -Depth 4

try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  $rel = Invoke-RestMethod -Method Post `
    -Uri "https://api.github.com/repos/$Repo/releases" `
    -Headers @{ Authorization = "token $tk"; 'User-Agent' = 'enterprisecore-radar' } `
    -ContentType 'application/json; charset=utf-8' `
    -Body ([Text.Encoding]::UTF8.GetBytes($corpo))
  Write-Host "   ok: $($rel.html_url)" -ForegroundColor Green
  $linkDownload = $rel.html_url
} catch {
  # Release e conveniencia: a tag ja subiu, o codigo ja esta la. Nao e falha.
  Write-Host "   Nao consegui criar a Release (a versao $Versao ja subiu do mesmo jeito)." -ForegroundColor Yellow
  Write-Host "   Motivo: $($_.Exception.Message)" -ForegroundColor Yellow
  $linkDownload = "https://github.com/$Repo/archive/refs/tags/$Versao.zip"
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Publicado: $Versao" -ForegroundColor Green
Write-Host ""
Write-Host "  Manda este link para quem usa o Radar:"
Write-Host "     $linkDownload"
Write-Host ""
Write-Host "  Quem ja tem a pasta so precisa rodar, dentro dela:"
Write-Host "     git pull"
Write-Host "============================================================" -ForegroundColor Green
