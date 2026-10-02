# =====================================================================
#  processar-boletins.ps1  (ENTERPRISECORE Radar - Fase 2)
#  Le os boletins .xlsx da Conlicitacao (data/boletins), filtra pelo
#  escopo ENTERPRISECORE (config/escopo-enterprisecore.json) marcando CONFIANCA + MOTIVO,
#  deduplica, MESCLA com o historico preservando colunas manuais, e gera:
#     data/dados.json                    (alimenta o app)
#     data/Planilha-ENTERPRISECORE-Licitacoes.xlsx (17 colunas originais + 5 ajudas)
#
#  Leitura via Excel COM (read-only). Gravacao do .xlsx via OpenXML puro
#  (.NET System.IO.Packaging) com suporte a datas nativas m/d/yyyy e formulas.
# =====================================================================
param(
  [string]$BoletinsDir = "$PSScriptRoot\..\data\boletins",
  [string]$OutJson     = "$PSScriptRoot\..\data\dados.json",
  [string]$OutXlsx     = "$PSScriptRoot\..\data\Planilha-ENTERPRISECORE-Licitacoes.xlsx",
  [string]$ConfigPath  = "$PSScriptRoot\..\config\escopo-enterprisecore.json",
  [string]$FilterIds   = '',
  [switch]$Diagnostico,
  [switch]$Zerar
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName WindowsBase

# Unicode-safe strings (PowerShell 5.1 reads .ps1 without BOM as ANSI, corrupting accented chars)
$STR_Media       = "M$([char]0x00E9)dia"          # Média
$STR_Generico    = "gen$([char]0x00E9)rico"        # genérico
$STR_Excluido    = "Exclu$([char]0x00ED)do"        # Excluído
$STR_Substituicao= "Substitui$([char]0x00E7)$([char]0x00E3)o" # Substituição
$STR_Amapa       = "Amap$([char]0x00E1) (AP)"      # Amapá (AP)
$STR_Ceara       = "Cear$([char]0x00E1) (CE)"      # Ceará (CE)
$STR_Espirito    = "Esp$([char]0x00ED)rito Santo (ES)"  # Espírito Santo (ES)
$STR_Goias       = "Goi$([char]0x00E1)s (GO)"      # Goiás (GO)
$STR_Maranhao    = "Maranh$([char]0x00E3)o (MA)"    # Maranhão (MA)
$STR_Para        = "Par$([char]0x00E1) (PA)"        # Pará (PA)
$STR_Paraiba     = "Para$([char]0x00ED)ba (PB)"     # Paraíba (PB)
$STR_Parana      = "Paran$([char]0x00E1) (PR)"      # Paraná (PR)
$STR_Piaui       = "Piau$([char]0x00ED) (PI)"       # Piauí (PI)
$STR_Rondonia    = "Rond$([char]0x00F4)nia (RO)"    # Rondônia (RO)
$STR_SaoPaulo    = "S$([char]0x00E3)o Paulo (SP)"   # São Paulo (SP)
$STR_Confiaveis  = "Parceiros ENTERPRISECORE (Confi$([char]0x00E1)veis)"  # Confiáveis
$STR_PossConc    = "Poss$([char]0x00ED)veis Concorrentes"        # Possíveis Concorrentes
$STR_Confianca   = "CONFIAN$([char]0x00C7)A"        # CONFIANÇA
$STR_Descricao   = "DESCRI$([char]0x00C7)$([char]0x00C3)O DO OBJETO" # DESCRIÇÃO DO OBJETO
$STR_Insercao    = "DATA DA INSER$([char]0x00C7)$([char]0x00C3)O" # DATA DA INSERÇÃO
$STR_Pregao      = "DATA DO PREG$([char]0x00C3)O"   # DATA DO PREGÃO
$STR_Sitio       = "S$([char]0x00CD)TIO DE COMPRAS" # SÍTIO DE COMPRAS
$STR_NPregao     = "N $([char]0x00BA) DO PREGAO"    # Nº DO PREGAO
$STR_NConlic     = "N$([char]0x00BA) LICITA$([char]0x00C7)$([char]0x00C3)O" # Nº LICITAÇÃO
$STR_Situacao    = "SITUA$([char]0x00C7)$([char]0x00C3)O" # SITUAÇÃO
$STR_Sugestao    = "sugest$([char]0x00E3)o"         # sugestão
$STR_Orgao       = "$([char]0x00D3)RG$([char]0x00C3)O" # ÓRGÃO
$STR_Solucao     = "SOLU$([char]0x00C7)$([char]0x00C3)O" # SOLUÇÃO

# ---------- helpers de texto/data ----------
function Normalize([string]$s){
  if([string]::IsNullOrEmpty($s)){ return ' ' }
  $s = $s.ToLowerInvariant()
  $d = $s.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach($ch in $d.ToCharArray()){
    if([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark){ [void]$sb.Append($ch) }
  }
  $r = ($sb.ToString() -replace '[^a-z0-9]+',' ').Trim()
  return ' ' + $r + ' '
}
function NormKey([string]$s){ (Normalize $s).Trim() }
function ToIso($v){
  if($null -eq $v){ return '' }
  if($v -is [double] -or $v -is [int] -or $v -is [long]){ try{ return ([DateTime]::FromOADate([double]$v)).ToString('yyyy-MM-dd') }catch{ return '' } }
  $s = [string]$v
  $m = [regex]::Match($s,'(\d{2})/(\d{2})/(\d{4})')
  if($m.Success){ return ('{0}-{1}-{2}' -f $m.Groups[3].Value,$m.Groups[2].Value,$m.Groups[1].Value) }
  return ''
}
function ToBr($iso){ if($iso -match '^(\d{4})-(\d{2})-(\d{2})'){ return ('{0}/{1}/{2}' -f $Matches[3],$Matches[2],$Matches[1]) } return '' }
$UfMap = @{ 'AC'='Acre (AC)';'AL'='Alagoas (AL)';'AP'=$STR_Amapa;'AM'='Amazonas (AM)';'BA'='Bahia (BA)';'CE'=$STR_Ceara;'DF'='Distrito Federal (DF)';'ES'=$STR_Espirito;'GO'=$STR_Goias;'MA'=$STR_Maranhao;'MT'='Mato Grosso (MT)';'MS'='Mato Grosso do Sul (MS)';'MG'='Minas Gerais (MG)';'PA'=$STR_Para;'PB'=$STR_Paraiba;'PR'=$STR_Parana;'PE'='Pernambuco (PE)';'PI'=$STR_Piaui;'RJ'='Rio de Janeiro (RJ)';'RN'='Rio Grande do Norte (RN)';'RS'='Rio Grande do Sul (RS)';'RO'=$STR_Rondonia;'RR'='Roraima (RR)';'SC'='Santa Catarina (SC)';'SP'=$STR_SaoPaulo;'SE'='Sergipe (SE)';'TO'='Tocantins (TO)' }
# So aceita UF que exista de verdade. Antes devolvia qualquer par de letras, e o
# "DE" de "Rio DE Janeiro" (coluna desalinhada do boletim) virava um estado no filtro.
function EstadoFmt($uf){
  if([string]::IsNullOrWhiteSpace($uf)){ return '' }
  $k = ([string]$uf).Trim().ToUpper()
  if($UfMap.ContainsKey($k)){ return $UfMap[$k] }
  return ''
}
function RankConf($c){ switch($c){ 'Alta'{4} $STR_Media{3} 'Concorrente'{2} 'Baixa'{1} default{0} } }

# ---------- helpers OpenXML (gravar .xlsx com datas nativas m/d/yyyy e formulas) ----------
function ColLetter([int]$n){ $s=''; while($n -gt 0){ $m=($n-1)%26; $s=[char](65+$m)+$s; $n=[int](($n-$m-1)/26) }; return $s }
function XmlEsc([string]$s){ if($null -eq $s){ return '' }; return ($s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'",'&apos;') }

function ToExcelSerial([string]$v){
  if([string]::IsNullOrWhiteSpace($v)){ return $null }
  $s = ([string]$v).Trim()
  $dt = $null
  try {
    if($s -match '^(\d{4})-(\d{1,2})-(\d{1,2})'){
      $dt = New-Object DateTime ([int]$Matches[1]), ([int]$Matches[2]), ([int]$Matches[3])
    } elseif($s -match '^(\d{1,2})/(\d{1,2})/(\d{4})'){
      # ToBr gera dd/MM/yyyy (padrao BR) - dia vem primeiro, nao o mes
      $dt = New-Object DateTime ([int]$Matches[3]), ([int]$Matches[2]), ([int]$Matches[1])
    }
  } catch { return $null }
  if($dt){
    return [math]::Floor($dt.ToOADate())
  }
  return $null
}

function Write-Xlsx([string]$path,[string[]]$headers,$rows){
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
  [void]$sb.Append('<row r="1">')
  for($c=0;$c -lt $headers.Count;$c++){
    $ref=(ColLetter ($c+1))+'1'
    [void]$sb.Append('<c r="'+$ref+'" t="inlineStr"><is><t xml:space="preserve">'+(XmlEsc $headers[$c])+'</t></is></c>')
  }
  [void]$sb.Append('</row>')
  $rn=1
  foreach($row in $rows){
    $rn++
    [void]$sb.Append('<row r="'+$rn+'">')
    for($c=0;$c -lt $row.Count;$c++){
      $colNum = $c + 1
      $ref = (ColLetter $colNum) + $rn
      $v = $row[$c]
      if($null -eq $v){ continue }

      # Colunas de Data (Formato m/d/yyyy): 7 (G: DATA DA INSERÇÃO), 8 (H: DATA DO PREGÃO), 12 (L: PUBLICADO)
      if(($colNum -eq 7 -or $colNum -eq 8 -or $colNum -eq 12) -and $v){
        $serial = ToExcelSerial [string]$v
        if($serial){
          [void]$sb.Append('<c r="'+$ref+'" s="1"><v>'+$serial+'</v></c>')
          continue
        }
      }

      # Formula para ANO (coluna 15 = O) baseada na Data do Pregão (coluna H)
      if($colNum -eq 15){
        $hRef = 'H' + $rn
        $anoVal = $v
        [void]$sb.Append('<c r="'+$ref+'"><f>IF(ISNUMBER('+$hRef+'),YEAR('+$hRef+'),"")</f><v>'+$anoVal+'</v></c>')
        continue
      }

      # Formula para MES (coluna 16 = P) baseada na Data do Pregão (coluna H)
      if($colNum -eq 16){
        $hRef = 'H' + $rn
        $mesVal = $v
        [void]$sb.Append('<c r="'+$ref+'"><f>IF(ISNUMBER('+$hRef+'),MONTH('+$hRef+'),"")</f><v>'+$mesVal+'</v></c>')
        continue
      }

      # Numericos
      if($v -is [int] -or $v -is [double] -or $v -is [long]){
        [void]$sb.Append('<c r="'+$ref+'"><v>'+([string]$v)+'</v></c>')
      } else {
        $t=[string]$v; if($t.Length -eq 0){ continue }; if($t.Length -gt 32000){ $t=$t.Substring(0,32000) }
        [void]$sb.Append('<c r="'+$ref+'" t="inlineStr"><is><t xml:space="preserve">'+(XmlEsc $t)+'</t></is></c>')
      }
    }
    [void]$sb.Append('</row>')
  }
  [void]$sb.Append('</sheetData></worksheet>')
  $sheetXml=$sb.ToString()
  $workbookXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Licitacoes ENTERPRISECORE" sheetId="1" r:id="rId1"/></sheets></workbook>'
  $stylesXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="165" formatCode="m/d/yyyy"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

  $enc=New-Object System.Text.UTF8Encoding($false)
  $pkg=$null
  try{
    if(Test-Path -LiteralPath $path){ try { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } catch {} }
    $pkg=[System.IO.Packaging.Package]::Open($path,[System.IO.FileMode]::Create)
  } catch {
    Write-Host "AVISO: O arquivo '$path' está aberto no Excel. Feche o Excel para que a planilha seja atualizada." -ForegroundColor Yellow
    return
  }
  try{
    $wbUri  = New-Object System.Uri('/xl/workbook.xml',[System.UriKind]::Relative)
    $wsUri  = New-Object System.Uri('/xl/worksheets/sheet1.xml',[System.UriKind]::Relative)
    $styUri = New-Object System.Uri('/xl/styles.xml',[System.UriKind]::Relative)

    $wbPart  = $pkg.CreatePart($wbUri,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
    $wsPart  = $pkg.CreatePart($wsUri,'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
    $styPart = $pkg.CreatePart($styUri,'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml')

    $b=$enc.GetBytes($workbookXml); $s=$wbPart.GetStream(); $s.Write($b,0,$b.Length); $s.Flush()
    $b=$enc.GetBytes($sheetXml);    $s=$wsPart.GetStream(); $s.Write($b,0,$b.Length); $s.Flush()
    $b=$enc.GetBytes($stylesXml);   $s=$styPart.GetStream(); $s.Write($b,0,$b.Length); $s.Flush()

    $pkg.CreateRelationship($wbUri,[System.IO.Packaging.TargetMode]::Internal,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument','rId1') | Out-Null
    $wbPart.CreateRelationship($wsUri,[System.IO.Packaging.TargetMode]::Internal,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet','rId1') | Out-Null
    $wbPart.CreateRelationship($styUri,[System.IO.Packaging.TargetMode]::Internal,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles','rId2') | Out-Null
  } finally { $pkg.Close() }
}

# ---------- carrega config ----------
if(-not (Test-Path -LiteralPath $ConfigPath)){ throw "Config nao encontrada: $ConfigPath" }
$cfg = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

# ATENCAO: o NOME da marca tambem vira termo de busca, alem dos aliases.
# Para siglas curtas isso causa falso positivo ('IBM' casava em "monitor padrao ibm").
# Use "buscarPeloNome": false no config para procurar SOMENTE pelos aliases.
$fabricantes = New-Object System.Collections.ArrayList
foreach($f in $cfg.fabricantes){
  $nome = $f.nome; $sol = $f.solucao
  $usarNome = $true
  if($f.PSObject.Properties.Name -contains 'buscarPeloNome'){ $usarNome = [bool]$f.buscarPeloNome }
  $terms = @($f.aliases)
  if($usarNome){ $terms = @($nome) + $terms }
  # sinalFraco = marca de catalogo largo (faz de impressora a servidor). O nome
  # sozinho nao estabelece escopo - medido em 100.165 editais: 'Positivo' aparece
  # 12x e ZERO em contexto de TI; 'dell' costuma estar na ficha de um desktop.
  $fraca = $false
  if($f.PSObject.Properties.Name -contains 'sinalFraco'){ $fraca = [bool]$f.sinalFraco }
  foreach($t in $terms){ if($t){ $nk = NormKey $t; if($nk){ [void]$fabricantes.Add([pscustomobject]@{ raw=$nome; norm=$nk; sol=$sol; fraca=$fraca }) } } }
}

$concorrentes = New-Object System.Collections.ArrayList
foreach($c in $cfg.concorrentes){
  $nome = $c.nome; $sol = $c.solucao
  $terms = @($nome) + @($c.aliases)
  foreach($t in $terms){ if($t){ $nk = NormKey $t; if($nk){ [void]$concorrentes.Add([pscustomobject]@{ raw=$nome; norm=$nk; sol=$sol }) } } }
}

# EXPRESSOES DE CONJUNCAO (bloco "expressoes" do config, sintaxe "a & b" da
# fonte do usuario). A categoria so dispara se as DUAS pontas aparecerem no
# mesmo objeto. E o sinal mais preciso que existe aqui: "deteccao" sozinha
# dispara 168 vezes e 89% sao alarme de incendio; "deteccao & resposta"
# dispara 7 vezes e todas sao EDR/XDR.
$expressoes = New-Object System.Collections.ArrayList
foreach($e in @($cfg.expressoes)){
  if(-not $e){ continue }
  $la = New-Object System.Collections.ArrayList
  $lb = New-Object System.Collections.ArrayList
  foreach($x in @($e.a)){ if($x){ $nk = NormKey $x; if($nk){ [void]$la.Add($nk) } } }
  foreach($x in @($e.b)){ if($x){ $nk = NormKey $x; if($nk){ [void]$lb.Add($nk) } } }
  if($la.Count -gt 0 -and $lb.Count -gt 0){
    [void]$expressoes.Add([pscustomobject]@{ sol=$e.solucao; origem=$e.origem; a=$la; b=$lb })
  }
}

$categorias = New-Object System.Collections.ArrayList
foreach($cat in $cfg.categorias){
  $sol = $cat.solucao; $fab = $cat.fabricantePadrao
  foreach($kw in $cat.palavrasChave){ if($kw){ $nk = NormKey $kw; if($nk){ [void]$categorias.Add([pscustomobject]@{ sol=$sol; fab=$fab; norm=$nk; raw=$kw }) } } }
}

$termosExcluidos = @($cfg.termosExcluidos) | ForEach-Object { NormKey $_ }

# Termos que denunciam edital de outra area (audio, obra, frota, saude...)
$contextosFora = New-Object System.Collections.ArrayList
foreach($t in @($cfg.contextosForaEscopo)){ if($t){ $nk = NormKey $t; if($nk -and -not $contextosFora.Contains($nk)){ [void]$contextosFora.Add($nk) } } }

# Situacoes do boletim que entram no radar (config; padrao = so "nova")
$SituacoesAceitas = @()
foreach($s in @($cfg.situacoesAceitas)){ if($s){ $SituacoesAceitas += (NormKey $s) } }
if($SituacoesAceitas.Count -eq 0){ $SituacoesAceitas = @('nova') }
$descartadasPorSituacao = @{}

$palavrasFracas = New-Object System.Collections.ArrayList
foreach($w in @($cfg.palavrasFracas)){ if($w){ $nk = NormKey $w; if($nk -and -not $palavrasFracas.Contains($nk)){ [void]$palavrasFracas.Add($nk) } } }
foreach($w in @('ti','computador','notebook','nobreak','antivirus','servidor','impressora','monitor','desktop')){
  if(-not $palavrasFracas.Contains($w)){ [void]$palavrasFracas.Add($w) }
}

# ESPECIFICIDADE DA PALAVRA-CHAVE (config -> "especificidade").
# Sem isto toda palavra vale o mesmo: "conectividade", que aparece na ficha
# tecnica de um teclado sem fio, pontuava igual a "breach and attack", que so
# aparece num edital que e exatamente o que a ENTERPRISECORE vende. Ambos davam +34.
$ESPEC = $cfg.especificidade
$kwNucleo  = New-Object System.Collections.ArrayList
$kwAmbigua = New-Object System.Collections.ArrayList
if($ESPEC){
  foreach($t in @($ESPEC.nucleo)){   if($t){ $nk = NormKey $t; if($nk -and -not $kwNucleo.Contains($nk)){  [void]$kwNucleo.Add($nk) } } }
  foreach($t in @($ESPEC.ambiguas)){ if($t){ $nk = NormKey $t; if($nk -and -not $kwAmbigua.Contains($nk)){ [void]$kwAmbigua.Add($nk) } } }
}
# MODALIDADE: Registro de Precos com preco por item deixa o fornecedor escolher
# em quais itens propor. Nesses, 'poucos itens nossos' nao e defeito do edital,
# so define o tamanho da fatia - ver bloco 'modalidade' no config.
$MODAL = $cfg.modalidade
$marcRP = New-Object System.Collections.ArrayList
if($MODAL){ foreach($t in @($MODAL.marcadoresRegistroPrecos)){ if($t){ $nk = NormKey $t; if($nk -and -not $marcRP.Contains($nk)){ [void]$marcRP.Add($nk) } } } }
# Termo no proprio objeto que denuncia que NAO e um pregao competitivo de
# verdade (chamamento publico de parceria, sondagem de mercado etc.) - sem
# contrato direto possivel dali, mesmo que o vocabulario bata com o portfolio.
# Tentativa anterior usava o PREFIXO do numero do pregao ("SM/") pensando ser
# Solicitacao de Manifestacao - errado: 18 dos 73 itens de um lote real tinham
# "SM/", e a maioria eram pregoes normais (ex.: SM/52/2026 = CELEPAR comprando
# licenca IBM Storage). O sinal certo e o texto do proprio objeto.
$marcNaoComp = New-Object System.Collections.ArrayList
if($MODAL){ foreach($t in @($MODAL.marcadoresNaoCompetitivos)){ if($t){ $nk = NormKey $t; if($nk -and -not $marcNaoComp.Contains($nk)){ [void]$marcNaoComp.Add($nk) } } } }
function ModalLigado([string]$k){
  if($MODAL -and ($MODAL.PSObject.Properties.Name -contains $k)){ return [bool]$MODAL.$k }
  return $false
}
function EspecPeso([string]$k,[double]$def){
  if($ESPEC -and ($ESPEC.PSObject.Properties.Name -contains $k)){ return [double]$ESPEC.$k }
  return $def
}

# ---------- motor de score (0-100) ----------
# Calibragem fica em config/escopo-enterprisecore.json -> "pesos" e "faixas".
$PESOS  = $cfg.pesos
$FAIXAS = @($cfg.faixas)

function Peso([string]$k,[double]$def){
  if($PESOS -and ($PESOS.PSObject.Properties.Name -contains $k)){ return [double]$PESOS.$k }
  return $def
}
function FaixaDe([int]$s){
  foreach($f in $FAIXAS){ if($s -ge [int]$f.min){ return $f } }
  if($FAIXAS.Count -gt 0){ return $FAIXAS[$FAIXAS.Count-1] }
  return [pscustomobject]@{ nome='?'; cor='#888888'; acao=''; min=0 }
}
function DiasAte([string]$iso){
  if($iso -match '^(\d{4})-(\d{2})-(\d{2})'){
    try{
      $dp = New-Object DateTime ([int]$Matches[1]),([int]$Matches[2]),([int]$Matches[3])
      return [int][Math]::Floor(($dp - (Get-Date).Date).TotalDays)
    }catch{ return $null }
  }
  return $null
}

# Ao contrario da versao anterior, NAO para no primeiro sinal: acumula todos e soma.
function AvaliarItem($item){
  $txtNorm = Normalize ([string]$item.objeto + ' ' + [string]$item.solucao + ' ' + [string]$item.fabricante)

  $score      = 0.0
  $porques    = New-Object System.Collections.ArrayList
  $parcNomes  = New-Object System.Collections.ArrayList
  $concNomes  = New-Object System.Collections.ArrayList
  $solNomes   = New-Object System.Collections.ArrayList
  $kwHit      = New-Object System.Collections.ArrayList
  $catDaPalavra = @{}
  $fracHit    = New-Object System.Collections.ArrayList
  $exHit      = New-Object System.Collections.ArrayList
  $ptsCat     = 0
  $fabTop     = ''
  $solTop     = ''

  # 1. termos fora de escopo (penalizam forte, mas nao descartam sozinhos)
  foreach($ex in $termosExcluidos){
    if($ex -and $txtNorm.Contains(" $ex ") -and -not $exHit.Contains($ex)){ [void]$exHit.Add($ex) }
  }
  # 2. parceiros oficiais ENTERPRISECORE
  $parcFracos = New-Object System.Collections.ArrayList
  $parcValidos = @()
  foreach($f in $fabricantes){
    if($txtNorm.Contains(" $($f.norm) ") -and -not $parcNomes.Contains($f.raw)){
      [void]$parcNomes.Add($f.raw)
      if($f.fraca){ [void]$parcFracos.Add($f.raw) }
      if(-not $fabTop){ $fabTop = $f.raw; $solTop = $f.sol }
    }
  }
  # 3. concorrentes (oportunidade de takeover)
  foreach($c in $concorrentes){
    if($txtNorm.Contains(" $($c.norm) ") -and -not $concNomes.Contains($c.raw)){ [void]$concNomes.Add($c.raw) }
  }
  # 4. categorias de solucao do portfolio
  foreach($cat in $categorias){
    if($txtNorm.Contains(" $($cat.norm) ")){
      $kwLimpa = ([string]$cat.raw).Trim()
      if($kwLimpa -and -not $kwHit.Contains($kwLimpa)){ [void]$kwHit.Add($kwLimpa) }
      if($kwLimpa){ $catDaPalavra[$kwLimpa] = $cat.sol }
      if(-not $solNomes.Contains($cat.sol)){
        [void]$solNomes.Add($cat.sol)
        if(-not $fabTop){ $fabTop = $cat.fab }
        if(-not $solTop){ $solTop = $cat.sol }
      }
    }
  }

  # 4a. expressoes de conjuncao: as duas pontas no mesmo objeto
  $janela = [int](Peso 'expressaoJanelaChars' 40)
  foreach($e in $expressoes){
    # As duas pontas precisam estar PROXIMAS: a fonte quer a locucao
    # ("seguranca da informacao"), nao duas palavras soltas no mesmo texto.
    $perto = $false
    foreach($x in $e.a){
      $pa = $txtNorm.IndexOf(" $x ")
      while($pa -ge 0){
        $ini = [Math]::Max(0, $pa - $janela)
        $fim = [Math]::Min($txtNorm.Length, $pa + $x.Length + 2 + $janela)
        $trecho = $txtNorm.Substring($ini, $fim - $ini)
        foreach($y in $e.b){ if($trecho.Contains(" $y ")){ $perto = $true; break } }
        if($perto){ break }
        $pa = $txtNorm.IndexOf(" $x ", $pa + 1)
      }
      if($perto){ break }
    }
    if(-not $perto){ continue }
    $rot = [string]$e.origem
    if($rot -and -not $kwHit.Contains($rot)){
      [void]$kwHit.Add($rot)
      $catDaPalavra[$rot] = $e.sol
    }
    if(-not $solNomes.Contains($e.sol)){
      [void]$solNomes.Add($e.sol)
      if(-not $solTop){ $solTop = $e.sol }
    }
  }

  # 4b. contexto de outra area (audio, obra, frota, saude, alimentacao...)
  $foraHit = New-Object System.Collections.ArrayList
  foreach($t in $contextosFora){
    if($t -and $txtNorm.Contains(" $t ") -and -not $foraHit.Contains($t)){ [void]$foraHit.Add($t) }
  }
  # 5. termos genericos de TI
  foreach($w in $palavrasFracas){
    if($w -and $txtNorm.Contains(" $w ") -and -not $fracHit.Contains($w)){ [void]$fracHit.Add($w) }
  }
  # 6. periferico/consumo: so conta quando NAO ha nenhum sinal tecnico forte
  $ehPeriferico = ($txtNorm -match ' (smart tv|teclado|mouse|carregador|adaptador|fone de ouvido|microfone|cartucho|toner) ') -and
                  ($parcNomes.Count -eq 0) -and ($concNomes.Count -eq 0) -and ($solNomes.Count -eq 0)

  # ---------- soma ----------
  if($parcNomes.Count -gt 0){
    # Marca de catalogo largo so vale se o edital JA tiver categoria do portfolio.
    # Sem isso, o 'dell' de 'Memoria Ram aplicacao: dell optiplex 3060' valia 26
    # pontos e ainda blindava um RP de projetores contra a penalidade de escopo.
    $parcValidos = @($parcNomes | Where-Object { $solNomes.Count -gt 0 -or -not $parcFracos.Contains($_) })
    if($parcValidos.Count -gt 0){
      $base  = Peso 'parceiroPrimeiro' 38
      $extra = [Math]::Min((($parcValidos.Count-1) * (Peso 'parceiroAdicional' 8)), (Peso 'parceiroMaxAdicional' 16))
      $score += $base + $extra
      $detParc = ($parcValidos -join ', ')
      $ignorados = @($parcNomes | Where-Object { $parcValidos -notcontains $_ })
      if($ignorados.Count -gt 0){ $detParc = "$detParc  [ignorado sem escopo: $($ignorados -join ', ')]" }
      [void]$porques.Add([pscustomobject]@{ sinal='Parceiro ENTERPRISECORE citado'; detalhe=$detParc; pontos=[int]($base+$extra) })
    }
  }
  if($concNomes.Count -gt 0){
    $base  = Peso 'concorrentePrimeiro' 24
    $extra = [Math]::Min((($concNomes.Count-1) * (Peso 'concorrenteAdicional' 6)), (Peso 'concorrenteMaxAdicional' 12))
    $score += $base + $extra
    [void]$porques.Add([pscustomobject]@{ sinal='Concorrente citado (chance de takeover)'; detalhe=($concNomes -join ', '); pontos=[int]($base+$extra) })
  }
  # DENSIDADE: uma palavra achada num objeto de 6000 caracteres nao vale o mesmo
  # que um edital inteiro sobre aquilo. Sem isso, um edital de estudio de radio
  # que cita 'switch' uma vez pontuava igual a um edital de rede.
  $tamObj     = [Math]::Max(1, ([string]$item.objeto).Length)
  $hitsEscopo = $kwHit.Count
  $fatorDens  = 1.0
  $motivoDens = ''

  # ---------------------------------------------------------------------
  # CENTRALIDADE NO LOTE - a medida boa.
  # O objeto do Conlicitacao vem numerado ("1 - ... 2 - ... 3 - ..."). Dá para
  # contar em quantos itens do lote a nossa palavra realmente aparece. Um switch
  # que sai em 3 de 20 itens e acessorio; em 5 de 6 e o edital inteiro.
  # Isso responde direto o que antes era chutado por caracteres.
  # ---------------------------------------------------------------------
  $itensLote = 0; $itensComEscopo = 0; $propLote = -1.0
  if($hitsEscopo -gt 0){
    $segs = @([regex]::Split([string]$item.objeto, '(?=(?:^|\s)\d{1,3}\s*-\s)') | Where-Object { $_.Trim().Length -gt 12 })
    $itensLote = $segs.Count
    if($itensLote -ge 3){
      # A especificidade vale por ITEM, nao so para o edital inteiro. Um item do
      # lote so conta como nosso se casar palavra ESPECIFICA - se casou apenas
      # ambigua, e carona. Casos reais do PE/46/2026 da Aeronautica: o nobreak
      # entrava por 'servidores' (a especificacao dizia 'utilizacao em servidores
      # do Coren-MT', copiada de outro edital) e cinco conectores opticos entravam
      # por 'fibra optica' - material de cabeamento, nao equipamento de rede.
      $minAmb = 2
      if($ESPEC -and ($ESPEC.PSObject.Properties.Name -contains 'minAmbiguasPorItem')){ $minAmb = [int]$ESPEC.minAmbiguasPorItem }
      # Categorias que tiveram palavra ESPECIFICA no edital inteiro. Dentro
      # delas, a palavra ambigua deixa de ser carona: o contexto ja provou que
      # o edital e daquela area.
      $catsFirmes = New-Object System.Collections.ArrayList
      foreach($k in $kwHit){
        $nk = NormKey $k
        if($kwAmbigua.Contains($nk)){ continue }
        $cs = $catDaPalavra[([string]$k)]
        if($cs -and -not $catsFirmes.Contains($cs)){ [void]$catsFirmes.Add($cs) }
      }
      foreach($seg in $segs){
        $segNorm = Normalize $seg
        $temEsp = $false; $qtdAmb = 0
        foreach($k in $kwHit){
          $nk = NormKey $k
          if(-not $nk -or -not $segNorm.Contains(" $nk ")){ continue }
          if($kwAmbigua.Contains($nk)){
            $cs = $catDaPalavra[([string]$k)]
            if($cs -and $catsFirmes.Contains($cs)){ $temEsp = $true; break }  # contexto desambigua
            $qtdAmb++
          } else { $temEsp = $true; break }
        }
        # Uma ambigua sozinha e carona; duas na mesma especificacao curta e sinal.
        if($temEsp -or $qtdAmb -ge $minAmb){ $itensComEscopo++ }
      }
      $propLote = $itensComEscopo / [double]$itensLote
    }
  }

  # E Registro de Precos? Ai o lote e uma cesta de itens com preco individual.
  $ehRP = $false
  foreach($m in $marcRP){ if($m -and $txtNorm.Contains(" $m ")){ $ehRP = $true; break } }

  # A ABERTURA do objeto ja diz que o que se compra e nosso? Entao a divisao em
  # itens nao deve reduzir: num edital de SERVICO os itens sao fases do mesmo
  # entregavel, nao uma cesta de produtos. Caso real: "Contratacao de empresa
  # especializada para prestacao de servicos tecnicos para planejamento,
  # instalacao, configuracao e customizacao do Microsoft Active Directory" -
  # 1 de 7 itens citava AD, fator caia para 0.4 e o edital ficava em 27.
  $cabecaLote = [regex]::Replace([string]$item.objeto, '^(\s*\*[^*]*\*\s*)+', '')
  if($cabecaLote.Length -gt 180){ $cabecaLote = $cabecaLote.Substring(0, 180) }
  $cabecaLoteNorm = Normalize $cabecaLote
  $escopoAbre = $false
  foreach($k in $kwHit){
    $nk = NormKey $k
    if($kwAmbigua.Contains($nk)){ continue }
    if($cabecaLoteNorm.Contains(" $nk ")){ $escopoAbre = $true; break }
  }

  if($escopoAbre -and $propLote -ge 0){
    $motivoDens = "o objeto abre com escopo nosso: $itensComEscopo de $itensLote itens citam, mas a compra e essa"
  }
  elseif($ehRP -and (ModalLigado 'rpIgnoraProporcaoLote') -and $propLote -ge 0){
    # Nao reduz: em RP por item ninguem e obrigado a fornecer o lote inteiro.
    $motivoDens = "registro de precos: $itensComEscopo de $itensLote itens sao nossos, e da para propor so neles"
  }
  elseif($propLote -ge 0){
    # com lote identificado, a proporcao manda
    if($propLote -lt [double](Peso 'loteMarginal' 0.25)){
      $fatorDens  = [double](Peso 'densidadeFatorLongo' 0.4)
      $motivoDens = "item marginal no lote: aparece em $itensComEscopo de $itensLote itens"
    } elseif($propLote -lt [double](Peso 'loteParcial' 0.5)){
      $fatorDens  = [double](Peso 'loteFatorParcial' 0.7)
      $motivoDens = "peso parcial no lote: $itensComEscopo de $itensLote itens"
    } elseif($propLote -ge [double](Peso 'loteDedicado' 0.7)){
      $fatorDens  = [double](Peso 'densidadeFatorConcentrado' 1.35)
      $motivoDens = "lote dedicado: $itensComEscopo de $itensLote itens sao do nosso escopo"
    }
  }
  elseif($tamObj -ge [int](Peso 'densidadeObjetoLongo' 2500) -and $hitsEscopo -le 2){
    $fatorDens  = [double](Peso 'densidadeFatorLongo' 0.4)
    $motivoDens = "sinal fraco: $hitsEscopo palavra(s) em $tamObj caracteres"
  } elseif($tamObj -ge [int](Peso 'densidadeObjetoMedio' 1200) -and $hitsEscopo -le 1){
    $fatorDens  = [double](Peso 'densidadeFatorMedio' 0.6)
    $motivoDens = "sinal fraco: 1 palavra em $tamObj caracteres"
  } else {
    # O contrario tambem importa: edital curto onde a palavra do escopo se repete
    # e edital DEDICADO aquilo. Sem isto, um pregao 100% de backup do TCU tirava a
    # mesma nota de um edital que menciona backup de passagem.
    $ocorr = 0
    foreach($k in $kwHit){
      $nk = NormKey $k
      if($nk){ $ocorr += ([regex]::Matches($txtNorm, [regex]::Escape(" $nk "))).Count }
    }
    $porMil = $ocorr / [Math]::Max(1.0, ($tamObj / 1000.0))
    if($hitsEscopo -ge 2 -and $porMil -ge [double](Peso 'densidadeConcentrada' 4) -and $foraHit.Count -eq 0){
      $fatorDens  = [double](Peso 'densidadeFatorConcentrado' 1.35)
      $motivoDens = "sinal concentrado: $ocorr ocorrencias em $tamObj caracteres"
    }
  }

  # PESO POR ESPECIFICIDADE DA PALAVRA. Antes toda palavra do dicionario valia o
  # mesmo: um edital de teclado sem fio disparava 'conectividade' e tirava os
  # mesmos 34 pontos de um edital de Breach and Attack Simulation. Agora:
  #   so palavra ambigua disparou -> pontos caem para 'fatorSoAmbigua'
  #   alguma palavra do nucleo    -> sobem para 'fatorNucleo' (o edital e nosso)
  $fatorEspec  = 1.0
  $motivoEspec = ''
  $nucHit = New-Object System.Collections.ArrayList
  $ambHit = New-Object System.Collections.ArrayList
  foreach($k in $kwHit){
    $nk = NormKey $k
    if($kwNucleo.Contains($nk)){ [void]$nucHit.Add($k) }
    elseif($kwAmbigua.Contains($nk)){ [void]$ambHit.Add($k) }
  }
  if($hitsEscopo -gt 0){
    # Duas ambiguas da MESMA categoria formam locucao tecnica ("data center" +
    # "virtualizacao"), e ai nao e carona. Uma sozinha continua sendo.
    $minAmbEsp = 2
    if($ESPEC -and ($ESPEC.PSObject.Properties.Name -contains 'ambiguasQueViramEspecifica')){ $minAmbEsp = [int]$ESPEC.ambiguasQueViramEspecifica }
    $ambPorCat = @{}
    foreach($k in $ambHit){ $cs = $catDaPalavra[([string]$k)]; if($cs){ $ambPorCat[$cs] = [int]$ambPorCat[$cs] + 1 } }
    $ambFormaLocucao = $false
    foreach($v in $ambPorCat.Values){ if([int]$v -ge $minAmbEsp){ $ambFormaLocucao = $true; break } }
    if($ambHit.Count -eq $hitsEscopo -and -not $ambFormaLocucao){
      $fatorEspec  = EspecPeso 'fatorSoAmbigua' 0.35
      $motivoEspec = 'sinal generico: so ' + (($ambHit | Select-Object -First 3) -join ', ')
    } elseif($nucHit.Count -gt 0){
      $fatorEspec  = EspecPeso 'fatorNucleo' 1.5
      $motivoEspec = 'termo inequivoco do portfolio: ' + (($nucHit | Select-Object -First 3) -join ', ')
    }
  }

  if($solNomes.Count -gt 0){
    $base  = Peso 'categoriaPrimeira' 20
    $extra = [Math]::Min((($solNomes.Count-1) * (Peso 'categoriaAdicional' 9)), (Peso 'categoriaMaxAdicional' 27))
    $teto  = EspecPeso 'categoriaTeto' 64
    $ptsCat = [int][Math]::Round(($base + $extra) * $fatorDens * $fatorEspec)
    if($ptsCat -gt $teto){ $ptsCat = [int]$teto }
    $score += $ptsCat
    $detCat = ($solNomes -join ', ')
    if($motivoDens){  $detCat = "$detCat  [$motivoDens]" }
    if($motivoEspec){ $detCat = "$detCat  [$motivoEspec]" }
    [void]$porques.Add([pscustomobject]@{ sinal='Solucao do portfolio no objeto'; detalhe=$detCat; pontos=$ptsCat })
  }
  # 'projeto integrado' so vale quando o sinal e denso de verdade
  # ...mas NAO numa cesta de Registro de Precos. Ali 5 categorias nossas quer
  # dizer que a cesta e variada, nao que existe um projeto integrado para
  # entregar. Na Aeronautica isso somava +8 em cima de uma lista de compras.
  $ptsProjetoIntegrado = 0
  if($solNomes.Count -ge 2 -and $fatorDens -ge 1.0 -and -not $ehRP){
    $b = Peso 'projetoIntegrado' 10
    $score += $b
    $ptsProjetoIntegrado = $b
    [void]$porques.Add([pscustomobject]@{ sinal='Projeto integrado'; detalhe=("$($solNomes.Count) solucoes ENTERPRISECORE no mesmo edital"); pontos=[int]$b })
  }
  # PORTE DO CONTRATO: o valor vem no proprio boletim (coluna "Valor Estimado"),
  # preenchido em cerca de metade das linhas. O esforco de disputar e quase o mesmo
  # para R$ 8 mil e para R$ 10 milhoes, entao o porte desempata dentro da mesma
  # faixa tecnica. So pontua com aderencia: valor de edital que nao e nosso nao diz nada.
  $vlr = 0.0
  try { if($null -ne $item.valorEstimado){ $vlr = [double]$item.valorEstimado } } catch { $vlr = 0.0 }
  # O porte so entra se a ADERENCIA se sustentar. Sem isso, uma manutencao predial
  # de sala-cofre de R$ 668 milhoes ganhava +12 por ter casado a palavra 'datacenter'.
  $minAder = EspecPeso 'valorExigeAderencia' 25
  # Em RP o valor estimado e a cesta toda (78 itens da Aeronautica, R$ 479 mi).
  # Nosso porte e a fatia, nao a cesta.
  if($ehRP -and (ModalLigado 'rpValorProporcional') -and $propLote -gt 0 -and $vlr -gt 0){
    $vlr = $vlr * $propLote
  }
  $ptsPorte = 0
  if($vlr -gt 0 -and $solNomes.Count -gt 0 -and $ptsCat -ge $minAder){
    $faixasValor = @(
      @{ min=5000000.0; peso='valorMuitoAlto'; pad=12; rot='acima de R$ 5 mi' },
      @{ min=1000000.0; peso='valorAlto';      pad=8;  rot='R$ 1 mi a 5 mi' },
      @{ min=200000.0;  peso='valorMedio';     pad=4;  rot='R$ 200 mil a 1 mi' },
      @{ min=0.0;       peso='valorBaixo';     pad=-6; rot='abaixo de R$ 200 mil' }
    )
    foreach($fv in $faixasValor){
      if($vlr -ge [double]$fv.min){
        $b = Peso ([string]$fv.peso) ([double]$fv.pad)
        if($b -ne 0){
          $score += $b
          $ptsPorte = $b
          $brl = 'R$ ' + $vlr.ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('pt-BR'))
          [void]$porques.Add([pscustomobject]@{ sinal='Porte do contrato'; detalhe="$brl ($($fv.rot))"; pontos=[int]$b })
        }
        break
      }
    }
  }

  # DESCONTO DE FATIA EM REGISTRO DE PRECOS.
  # A proporcao do lote saiu da aderencia tecnica (em RP da para propor so nos
  # nossos itens), mas ela ainda diz quanto a oportunidade vale: montar proposta
  # custa o mesmo para 12% da cesta e para 50% dela. Graduado, nao degrau.
  if($ehRP -and $propLote -gt 0 -and $ptsCat -gt 0){
    $fatiaPlena = 0.5
    if($MODAL -and ($MODAL.PSObject.Properties.Name -contains 'rpFatiaPlena')){ $fatiaPlena = [double]$MODAL.rpFatiaPlena }
    $penMax = -18.0
    if($MODAL -and ($MODAL.PSObject.Properties.Name -contains 'rpPenalidadeFatiaMax')){ $penMax = [double]$MODAL.rpPenalidadeFatiaMax }
    if($propLote -lt $fatiaPlena -and $fatiaPlena -gt 0){
      $b = [int][Math]::Round($penMax * (1.0 - ($propLote / $fatiaPlena)))
      if($b -ne 0){
        $score += $b
        $pctFatia = [int][Math]::Round($propLote * 100)
        [void]$porques.Add([pscustomobject]@{ sinal='Fatia pequena do registro de precos'; detalhe="so $pctFatia% dos itens sao nossos; a proposta custa o mesmo"; pontos=$b })
      }
    }
  }

  # Edital dominado por outra area.
  # Criterio por PROPORCAO, nao por contagem fixa: se ha tanto ou mais sinal de
  # outra area do que de escopo, o edital e de outro ramo com uns itens de TI no
  # meio. O criterio antigo ('ate 2 palavras de escopo') deixava passar a compra
  # de acessorios do STJ, que tinha 5 de cada lado e pontuava 83.
  #
  # TRAVA: marca PARCEIRA citada protege o edital - ai e nosso mesmo que o texto
  # misture outras coisas (foi assim que um CFTV da Marinha foi zerado por engano).
  # Concorrente NAO protege: TP-Link aparece em roteador domestico de brinde, nao
  # e incumbente para tomar.
  # A protecao da marca parceira e proporcional, nao absoluta. Se os termos de
  # outra area superam as palavras de escopo em 'foraDominanteRazao' vezes, a
  # marca esta citada de passagem (numa ficha de componente) e nao protege nada.
  $razaoFora = 0.0
  if($hitsEscopo -gt 0){ $razaoFora = $foraHit.Count / [double]$hitsEscopo }
  elseif($foraHit.Count -gt 0){ $razaoFora = 99.0 }
  # Contar palavras distintas superestima: as 3 palavras de controle de acesso do
  # RP de projetores vinham todas do MESMO item. Quando ha lote analisavel, a
  # medida honesta e a fatia de ITENS nossos - la eram 2 de 26 (8%).
  # Usa parcValidos, nao parcNomes: marca de catalogo largo citada na ficha de um
  # componente nao blinda o edital contra a penalidade de outra area.
  # O que se compra esta declarado na ABERTURA do objeto; o resto e detalhe.
  # Termo de outra area logo no comeco vale mais que a contagem no texto todo.
  $cabecaChars = [int](Peso 'foraNaCabecaChars' 180)
  $objLimpo = [regex]::Replace([string]$item.objeto, '^(\s*\*[^*]*\*\s*)+', '')
  if($objLimpo.Length -gt $cabecaChars){ $objLimpo = $objLimpo.Substring(0, $cabecaChars) }
  $cabecaNorm = Normalize $objLimpo
  $foraNaCabeca = New-Object System.Collections.ArrayList
  foreach($t in $foraHit){ if($cabecaNorm.Contains(" $t ")){ [void]$foraNaCabeca.Add($t) } }

  $parceiroProtege = ($parcValidos.Count -gt 0)
  if($parceiroProtege){
    if($propLote -ge 0){
      if($propLote -lt [double](Peso 'loteMarginal' 0.25)){ $parceiroProtege = $false }
    } elseif($razaoFora -ge [double](Peso 'foraDominanteRazao' 2.0)){
      $parceiroProtege = $false
    }
  }
  # A cabeca do objeto tem alguma palavra ESPECIFICA nossa? Se tem, o edital e
  # nosso mesmo citando outra area depois. Se nao tem, e a outra area que esta
  # sendo comprada, e basta UM termo dela na abertura.
  $escopoNaCabeca = $false
  foreach($k in $kwHit){
    $nk = NormKey $k
    if($kwAmbigua.Contains($nk)){ continue }
    if($cabecaNorm.Contains(" $nk ")){ $escopoNaCabeca = $true; break }
  }
  if(-not $escopoNaCabeca){
    foreach($p in $parcNomes){ $np = NormKey $p; if($np -and $cabecaNorm.Contains(" $np ")){ $escopoNaCabeca = $true; break } }
  }
  $foraDomina = ($foraHit.Count -ge 2 -and $foraHit.Count -ge $hitsEscopo) -or ($foraNaCabeca.Count -ge 2) -or ($foraNaCabeca.Count -ge 1 -and -not $escopoNaCabeca)
  if($foraDomina -and -not $parceiroProtege){
    if($foraNaCabeca.Count -ge 1 -and -not $escopoNaCabeca -and $ptsCat -gt 0){
      # A ABERTURA diz que o que se compra e de outra area, e nao ha nenhuma
      # palavra especifica nossa nem parceiro ali. Os pontos de categoria vieram
      # de termo AMBIGUO que so calhou de aparecer (ex.: "conectividade",
      # "fibra optica", "monitoramento continuo" descrevendo um link dedicado) -
      # nao sao sinal de verdade, entao sao CANCELADOS, nao apenas multados.
      # Um -45 fixo nao bastava quando a base de pontos de categoria era alta:
      # um link dedicado com esses termos tirava 54 de categoria e sobrava 34
      # mesmo com a penalidade. Cancela tambem os bonus que so existem PORQUE
      # a categoria (fake) foi detectada - "Projeto integrado" (exige 2+
      # categorias) e "Porte do contrato" (exige aderencia) - senao sobra +8 e
      # +12 mesmo com a categoria zerada. Cancelar zera de verdade.
      $b = -$ptsCat - $ptsProjetoIntegrado - $ptsPorte
      $detFora = "e o que se compra: " + (($foraNaCabeca | Select-Object -First 3) -join ', ')
    } else {
      $b = Peso 'contextoForaEscopo' (-45)
      $detFora = ($foraHit | Select-Object -First 5) -join ', '
    }
    $score += $b
    [void]$porques.Add([pscustomobject]@{ sinal='Edital de outra area'; detalhe=$detFora; pontos=[int]$b })
  }
  if($fracHit.Count -gt 0){
    $b = [Math]::Min(($fracHit.Count * (Peso 'palavraFraca' 4)), (Peso 'palavraFracaMax' 8))
    $score += $b
    [void]$porques.Add([pscustomobject]@{ sinal='Termo generico de TI'; detalhe=($fracHit -join ', '); pontos=[int]$b })
  }

  if($ehPeriferico){
    $b = Peso 'periferico' (-30)
    $score += $b
    [void]$porques.Add([pscustomobject]@{ sinal='Periferico / consumo'; detalhe='Item de consumo sem sinal tecnico do portfolio'; pontos=[int]$b })
  }
  if($exHit.Count -gt 0){
    $b = Peso 'termoExcluido' (-70)
    $score += $b
    [void]$porques.Add([pscustomobject]@{ sinal='Termo fora de escopo'; detalhe=($exHit -join ', '); pontos=[int]$b })
  }
  if($marcNaoComp.Count -gt 0){
    $achouNaoComp = $null
    foreach($m in $marcNaoComp){ if($m -and $txtNorm.Contains(" $m ")){ $achouNaoComp = $m; break } }
    if($achouNaoComp){
      $b = -70.0
      if($MODAL -and ($MODAL.PSObject.Properties.Name -contains 'penalidadeNaoCompetitivo')){ $b = [double]$MODAL.penalidadeNaoCompetitivo }
      $score += $b
      [void]$porques.Add([pscustomobject]@{ sinal='Nao e pregao competitivo'; detalhe="`"$achouNaoComp`" - sondagem de mercado ou parceria, sem contrato direto possivel dali"; pontos=[int]$b })
    }
  }

  # ADERENCIA TECNICA: o quanto o edital casa com o portfolio, SEM efeito de calendario.
  # Fica separada para nao se perder a informacao quando o pregao ja passou.
  $scTec = [int][Math]::Round($score)
  if($scTec -lt 0){ $scTec = 0 }
  if($scTec -gt 100){ $scTec = 100 }
  $fxTec = FaixaDe $scTec

  # VIABILIDADE DE PRAZO entra por ultimo: mexe no ranking, nao na aderencia.
  $dias = DiasAte ([string]$item.dataPregao)
  # A flag 'vencido' vale para TODOS (alimenta o filtro "ocultar vencidos" da tela),
  # independente de pontuar ou nao.
  $vencido = ($null -ne $dias -and $dias -lt 0)

  # Ja os PONTOS de prazo so fazem sentido para o que tem alguma aderencia: dar
  # bonus de prazo a um edital que nao e nosso so poluia o fim da lista.
  if($null -ne $dias -and $scTec -gt 0){
    $dApertado = [int](Peso 'diasPrazoApertado' 3)
    $dConf     = [int](Peso 'diasPrazoConfortavel' 7)
    if($dias -lt 0){
      $b = Peso 'prazoVencido' (-25)
      $score += $b
      [void]$porques.Add([pscustomobject]@{ sinal='Prazo vencido'; detalhe=("Pregao ocorreu ha $([Math]::Abs($dias)) dia(s)"); pontos=[int]$b })
    } elseif($dias -le $dApertado){
      $b = Peso 'prazoApertado' (-6)
      $score += $b
      [void]$porques.Add([pscustomobject]@{ sinal='Prazo apertado'; detalhe=("Faltam $dias dia(s) para o pregao"); pontos=[int]$b })
    } elseif($dias -ge $dConf){
      $b = Peso 'prazoConfortavel' 5
      $score += $b
      [void]$porques.Add([pscustomobject]@{ sinal='Prazo confortavel'; detalhe=("Faltam $dias dias para o pregao"); pontos=[int]$b })
    }
  }

  $sc = [int][Math]::Round($score)
  if($sc -lt 0){ $sc = 0 }
  if($sc -gt 100){ $sc = 100 }
  $fx = FaixaDe $sc

  # ---------- compatibilidade com as 4 abas / colunas antigas ----------
  if($exHit.Count -gt 0 -and $parcNomes.Count -eq 0 -and $concNomes.Count -eq 0 -and $solNomes.Count -eq 0){
    $conf='Revisar'; $tipo=$STR_Excluido
  } elseif($parcNomes.Count -gt 0){
    $conf='Alta'; $tipo='Parceiro ENTERPRISECORE'
  } elseif($concNomes.Count -gt 0){
    $conf='Concorrente'; $tipo='Concorrente (Takeover)'
    if(-not $fabTop){ $fabTop = $concNomes[0]; $solTop = "$STR_Substituicao $($concNomes[0])" }
  } elseif($solNomes.Count -gt 0){
    $conf=$STR_Media; $tipo='Escopo Geral (Multiband)'
  } elseif($fracHit.Count -gt 0){
    $conf=$STR_Media; $tipo='Escopo Geral (TI)'
    if(-not $solTop){ $solTop = "TI Geral ($($fracHit[0]))" }
  } else {
    $conf='Baixa'; $tipo='Outros'
    if(-not $solTop){ $solTop='A verificar' }
  }
  if($ehPeriferico){ $conf='Baixa'; $tipo='Perifericos / Consumo'; $solTop='Perifericos / Consumo' }

  $motivo = ($porques | ForEach-Object {
    $sg = ''; if($_.pontos -ge 0){ $sg = '+' }
    "$($_.sinal): $($_.detalhe) ($sg$($_.pontos))"
  }) -join ' | '
  if(-not $motivo){ $motivo = 'Sem sinal direto do escopo ENTERPRISECORE' }

  return [pscustomobject]@{
    confianca          = $conf
    tipoFabricante     = $tipo
    motivo             = $motivo
    fab                = $fabTop
    sol                = $solTop
    score              = $sc
    scoreTecnico       = $scTec
    faixa              = [string]$fx.nome
    faixaCor           = [string]$fx.cor
    faixaAcao          = [string]$fx.acao
    faixaTecnica       = [string]$fxTec.nome
    vencido            = $vencido
    porques            = @($porques)
    parceiros          = @($parcNomes)
    concorrentesDet    = @($concNomes)
    solucoesDet        = @($solNomes)
    palavrasChaveDet   = @($kwHit)
    palavrasFracasDet  = @($fracHit)
    termosExcluidosDet = @($exHit)
    contextoForaDet    = @($foraHit)
    densidadeFator     = $fatorDens
    itensLote          = $itensLote
    itensComEscopo     = $itensComEscopo
    proporcaoLote      = $(if($propLote -ge 0){ [math]::Round($propLote,2) } else { $null })
    valorConsiderado   = $vlr
    diasAtePregao      = $dias
  }
}

# ---------- varre boletins .xlsx (Excel COM read-only) ----------
# Numa copia recem-baixada do GitHub a pasta nao existe: git nao versiona pasta
# vazia. Criar e mais util que abortar - a mensagem seguinte ja explica o que fazer.
if(-not (Test-Path -LiteralPath $BoletinsDir)){
  try { [void](New-Item -ItemType Directory -Path $BoletinsDir -Force) } catch {}
}
if(-not (Test-Path -LiteralPath $BoletinsDir)){ throw "Nao consegui criar o diretorio de boletins: $BoletinsDir" }
$files = Get-ChildItem -LiteralPath $BoletinsDir -Filter "*.xlsx" -File
if($files.Count -eq 0){ Write-Host "Nenhum arquivo .xlsx em $BoletinsDir" -ForegroundColor Yellow; exit 0 }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$brutos = New-Object System.Collections.ArrayList

try {
  foreach($f in $files){
    Write-Host "Lendo boletim: $($f.Name)..." -ForegroundColor Cyan
    $wb = $excel.Workbooks.Open($f.FullName, [type]::Missing, $true)
    $ws = $wb.Sheets.Item(1)
    $used = $ws.UsedRange
    $maxR = $used.Rows.Count
    $maxC = $used.Columns.Count
    $arr  = $used.Value2

    # le cabecalho (linha 1)
    $colIdx = @{}
    for($c=1; $c -le $maxC; $c++){
      $h = [string]$arr[1, $c]
      if($h){ $colIdx[$h.Trim()] = $c }
    }

    if($Diagnostico){
      Write-Host "  Colunas encontradas em $($f.Name):" -ForegroundColor Gray
      $colIdx.Keys | ForEach-Object { Write-Host "   - $_ (col $($colIdx[$_]))" -ForegroundColor Gray }
      $wb.Close($false)
      continue
    }

    function FindCol([string]$pattern){
      foreach($k in $colIdx.Keys){ if($k -match $pattern){ return $colIdx[$k] } }
      return $null
    }

    $cNumLic   = FindCol 'conlic|n.*licit'
    if(-not $cNumLic){ $cNumLic = 1 }

    $cOrgao    = FindCol 'rg|org'
    if(-not $cOrgao){ $cOrgao = 3 }

    # BUG GRAVE ACHADO EM 25/08/2026: 'est|uf' tambem casa "Valor Estimado"
    # (via "Est"imado), e a ordem de enumeracao de um hashtable no PowerShell
    # nao e garantida - as vezes FindCol devolvia a coluna de VALOR como se
    # fosse a de estado. O valor numerico falhava a validacao (tem digito) e
    # a UF ficava vazia mesmo com "Estado"="DF" certinho na planilha. "estado"
    # nao e substring de "estimado" (est-Ado x est-Imado), entao e inequivoco.
    $cUf       = FindCol 'estado|^uf$'
    if(-not $cUf){ $cUf = 6 }

    # o boletim chama de "Site 1"/"Site 2", nao de "sitio de compras"
    $cSitio    = FindCol 'stio|sitio|compras|^site'
    $cSitio2   = FindCol '^site 2'
    $cValor    = FindCol 'valor.*estim|estimado'

    $cNumPreg  = FindCol 'edital|preg'
    if(-not $cNumPreg){ $cNumPreg = 8 }

    $cUasg     = FindCol 'processo|uasg'
    if(-not $cUasg){ $cUasg = 11 }

    $cDtPreg   = FindCol 'aber|preg'
    if(-not $cDtPreg){ $cDtPreg = 16 }

    $cPrazo    = FindCol 'prazo'
    if(-not $cPrazo){ $cPrazo = 17 }

    $cHrPreg   = FindCol 'hora.*preg'
    $cDtIns    = FindCol 'data.*ins'
    $cPub      = FindCol 'publicad|atualiz'

    $cObj      = FindCol 'obj'
    if(-not $cObj){ $cObj = 18 }

    $cItens    = FindCol 'iten'
    if(-not $cItens){ $cItens = 13 }

    $cObs      = FindCol 'obser'
    if(-not $cObs){ $cObs = 19 }

    $cLink     = FindCol 'link|anexo'
    if(-not $cLink){ $cLink = 20 }

    $cSituacao = FindCol 'situa'
    if(-not $cSituacao){ $cSituacao = 14 }

    for($r=2; $r -le $maxR; $r++){
      $numLic = [string]$arr[$r, $cNumLic]
      if([string]::IsNullOrWhiteSpace($numLic)){ continue }

      # Filtro por SITUACAO. A lista vale em config/escopo-enterprisecore.json -> "situacoesAceitas".
      # Padrao = apenas "nova" (filtro estrito pedido pelo usuario). Para tambem considerar
      # URGENTE / PRORROGADA / RETIFICACAO, acrescente na lista do config - nao mexer aqui.
      if($cSituacao){
        $sitVal = NormKey ([string]$arr[$r, $cSituacao])
        if($SituacoesAceitas -notcontains $sitVal){
          $descartadasPorSituacao[$sitVal] = 1 + [int]$descartadasPorSituacao[$sitVal]
          continue
        }
      }

      $obs     = [string]$arr[$r, $cObs]
      $numPreg = [string]$arr[$r, $cNumPreg]
      $uasg    = [string]$arr[$r, $cUasg]

      # Sitio de compras: o boletim traz em "Site 1" e, as vezes, em "Site 2".
      # So vem preenchido em ~4% das linhas; no resto fica em branco mesmo e a
      # pessoa completa a mao.
      $sitio = ([string]$arr[$r, $cSitio]).Trim()
      if([string]::IsNullOrWhiteSpace($sitio) -and $cSitio2){
        $sitio = ([string]$arr[$r, $cSitio2]).Trim()
      }
      if($sitio -match 'informado'){ $sitio = '' }

      # Valor estimado: vem em ~47% das linhas do boletim. Numero puro.
      $valorEstimado = 0.0
      if($cValor){
        $vRaw = ([string]$arr[$r, $cValor]).Trim()
        if($vRaw){
          $vLimpo = $vRaw -replace '[^\d,\.]',''
          # boletim as vezes usa virgula decimal, as vezes ponto
          if($vLimpo -match ','){ $vLimpo = ($vLimpo -replace '\.','') -replace ',','.' }
          $tmpV = 0.0
          # BUG GRAVE ACHADO EM 25/08/2026: TryParse sem cultura usa a cultura da
          # maquina (pt-BR aqui), onde "." e separador de MILHAR, nao decimal.
          # Uma celula numerica nativa do Excel vira string com ponto decimal
          # ("146615.4"), e sem InvariantCulture o TryParse lia isso como
          # "146.615" + "4" solto = 1466154 - dez vezes o valor real. Como o
          # bloco acima ja normalizou qualquer formato BR (virgula) para ponto,
          # a partir daqui o separador SEMPRE e decimal, entao forcar invariante
          # e sempre correto, nunca destrutivo.
          if([double]::TryParse($vLimpo, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$tmpV)){ $valorEstimado = $tmpV }
        }
      }

      # Extrai Órgão com fallback em Observação e Objeto
      $orgao = ([string]$arr[$r, $cOrgao]).Trim()
      if([string]::IsNullOrWhiteSpace($orgao) -or $orgao -match 'informado'){
        $m = [regex]::Match($obs, '(?i)(?:Órgão|Unidade compradora|Fonte):\s*([^\r\n]+)')
        if($m.Success){
          $orgao = ($m.Groups[1].Value -replace '^\d+\s*-\s*','').Trim()
        }
      }
      if([string]::IsNullOrWhiteSpace($orgao) -or $orgao -match 'informado'){
        $mObj = [regex]::Match($obj, '(?i)(?:atendimento|necessidades|serviços)\s+do\s+([A-Z\s]{4,60})')
        if($mObj.Success){ $orgao = $mObj.Groups[1].Value.Trim() }
      }

      # Extrai UF / Estado com fallback em Observação, Órgão, Domínio e Sigla
      $rawObj  = [string]$arr[$r, $cObj]
      $rawItn  = [string]$arr[$r, $cItens]
      $obj     = "$rawObj $rawItn".Trim()

      $uf = ([string]$arr[$r, $cUf]).Trim()
      if($uf -match '\d' -or $uf.Length -gt 2 -or -not $UfMap.ContainsKey($uf.ToUpper())){
        $uf = ''
      }
      if([string]::IsNullOrWhiteSpace($uf)){
        $mLoc = [regex]::Match($obs, 'Local:\s*.*?[/\-\s]+([A-Z]{2})\b')
        if($mLoc.Success){ $uf = $mLoc.Groups[1].Value.ToUpper() }
      }
      if([string]::IsNullOrWhiteSpace($uf)){
        $mGov = [regex]::Match("$orgao $obs $obj", '(?i)\.([a-z]{2})\.gov\.br')
        if($mGov.Success){
          $g = $mGov.Groups[1].Value.ToUpper()
          if($g -ne 'GO' -or "$orgao $obs $obj" -match '(?i)goias|goiás'){ $uf = $g }
        }
      }
      if([string]::IsNullOrWhiteSpace($uf)){
        $mInst = [regex]::Match("$orgao $obs $obj", '\b(IF|UF)(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b')
        if($mInst.Success){ $uf = $mInst.Groups[2].Value.ToUpper() }
      }
      if([string]::IsNullOrWhiteSpace($uf)){
        $mBar = [regex]::Match("$orgao $obs $obj", '[/\-\s](AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b')
        if($mBar.Success){ $uf = $mBar.Groups[1].Value.ToUpper() }
      }
      # TRAVA FINAL. So a leitura da coluna validava contra o $UfMap; os quatro
      # fallbacks gravavam o que achassem. Resultado: 5 dos 72 registros ficaram
      # com UF 'DE', 'DO' e 'OU' - preposicoes. Melhor sem estado do que errado.
      if($uf -and -not $UfMap.ContainsKey($uf.ToUpper())){ $uf = '' }
      $estado = EstadoFmt $uf

      # Extrai Data do Pregão com fallbacks
      $dtPreg = ToIso $arr[$r, $cDtPreg]
      if([string]::IsNullOrWhiteSpace($dtPreg)){ $dtPreg = ToIso $arr[$r, $cPrazo] }
      if([string]::IsNullOrWhiteSpace($dtPreg)){
        $m = [regex]::Match($obs, '(?i)(?:lances|abertura|pregão|publicação):\s*(\d{2}/\d{2}/\d{4})')
        if($m.Success){ $dtPreg = ToIso $m.Groups[1].Value }
      }

      # Extrai Hora do Pregão com fallbacks
      $hrPreg = [string]$arr[$r, $cHrPreg]
      if([string]::IsNullOrWhiteSpace($hrPreg) -and $cDtPreg){
        $vDt = $arr[$r, $cDtPreg]
        if($vDt -is [double]){
          try { $hrPreg = ([DateTime]::FromOADate([double]$vDt)).ToString('HH:mm') } catch {}
        }
      }
      if([string]::IsNullOrWhiteSpace($hrPreg)){
        $m = [regex]::Match($obs, '(\d{2}:\d{2}:\d{2}|\d{2}:\d{2})')
        if($m.Success){ $hrPreg = $m.Groups[1].Value }
      }

      # Extrai Publicado / Data de Inserção com fallbacks
      $pub = ToIso $arr[$r, $cPub]
      if([string]::IsNullOrWhiteSpace($pub)){
        $m = [regex]::Match($obs, '(?i)Publicaç[ãa]o:\s*(\d{2}/\d{2}/\d{4})')
        if($m.Success){ $pub = ToIso $m.Groups[1].Value }
      }
      $dtIns   = $pub
      $rawObj  = [string]$arr[$r, $cObj]
      $rawItn  = [string]$arr[$r, $cItens]
      $obj     = "$rawObj $rawItn".Trim()
      $link    = [string]$arr[$r, $cLink]

      $id = "$numLic-$uasg-$numPreg".Trim('-')
      if([string]::IsNullOrWhiteSpace($id)){ $id = [Guid]::NewGuid().ToString() }

      $ano = ''; $mes = ''
      if($dtPreg -match '^(\d{4})-(\d{2})'){ $ano = [int]$Matches[1]; $mes = [int]$Matches[2] }

      $rec = [pscustomobject]@{
        id                 = $id
        numeroConlicitacao = $numLic
        orgao              = $orgao
        uf                 = $uf
        estado             = EstadoFmt $uf
        sitioCompras       = $sitio
        numeroPregao       = $numPreg
        uasg               = $uasg
        dataPregao         = $dtPreg
        horaPregao         = $hrPreg
        dataInsercao       = $dtIns
        publicado          = $pub
        objeto             = $obj
        link               = $link
        ano                = $ano
        mes                = $mes
        boletimOrigem      = $f.Name
        # colunas manuais da planilha
        valorEstimado      = $valorEstimado
        # colunas que a pessoa preenche na planilha (o boletim nao traz)
        am                 = ''
        produto            = ''
        fabricante         = ''
        status             = 'Novo'
        situacao           = 'Pendente'
      }
      [void]$brutos.Add($rec)
    }
    $wb.Close($false)
    # BUG achado em 03/09/2026: sem liberar wb/ws/used/arr um a um, o EXCEL.EXE
    # deste processo nao morre de verdade nem com $excel.Quit() no finally -
    # o Windows so encerra o processo quando NENHUMA referencia COM sobrevive,
    # e o PowerShell nao coleta essas referencias sozinho. Resultado observado:
    # 41 processos EXCEL.EXE zumbis acumulados desde 01/09/2026, um por rodada.
    if($used){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($used) | Out-Null }
    if($ws){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null }
    if($wb){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    $used = $null; $ws = $null; $wb = $null; $arr = $null
  }
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

if($Diagnostico){ Write-Host "`nModo Diagnostico encerrado." -ForegroundColor Yellow; exit 0 }

# ---------- avalia escopo ENTERPRISECORE + deduplica ----------
$CamposAvaliacao = @(
  @{ prop='confianca';          de='confianca'          },
  @{ prop='tipoFabricante';     de='tipoFabricante'     },
  @{ prop='motivoMatch';        de='motivo'             },
  @{ prop='fabricanteSugerido'; de='fab'                },
  @{ prop='solucao';            de='sol'                },
  @{ prop='score';              de='score'              },
  @{ prop='scoreTecnico';       de='scoreTecnico'       },
  @{ prop='faixa';              de='faixa'              },
  @{ prop='faixaCor';           de='faixaCor'           },
  @{ prop='faixaAcao';          de='faixaAcao'          },
  @{ prop='faixaTecnica';       de='faixaTecnica'       },
  @{ prop='vencido';            de='vencido'            },
  @{ prop='porques';            de='porques'            },
  @{ prop='parceiros';          de='parceiros'          },
  @{ prop='concorrentesDet';    de='concorrentesDet'    },
  @{ prop='solucoesDet';        de='solucoesDet'        },
  @{ prop='palavrasChaveDet';   de='palavrasChaveDet'   },
  @{ prop='palavrasFracasDet';  de='palavrasFracasDet'  },
  @{ prop='termosExcluidosDet'; de='termosExcluidosDet' },
  @{ prop='contextoForaDet';    de='contextoForaDet'    },
  @{ prop='densidadeFator';     de='densidadeFator'     },
  @{ prop='itensLote';          de='itensLote'          },
  @{ prop='itensComEscopo';     de='itensComEscopo'     },
  @{ prop='proporcaoLote';      de='proporcaoLote'      },
  @{ prop='valorConsiderado';   de='valorConsiderado'   },
  @{ prop='diasAtePregao';      de='diasAtePregao'      }
)

$novoById = [ordered]@{}
foreach($rec in $brutos){
  $ev = AvaliarItem $rec
  foreach($c in $CamposAvaliacao){
    $rec | Add-Member -NotePropertyName $c.prop -NotePropertyValue $ev.($c.de) -Force
  }

  if($novoById.Contains($rec.id)){
    $ex=$novoById[$rec.id]
    if($ex.boletimOrigem -notmatch [regex]::Escape($rec.boletimOrigem)){ $ex.boletimOrigem = "$($ex.boletimOrigem); $($rec.boletimOrigem)" }
    # duplicado entre boletins: fica a avaliacao de maior score
    if([int]$rec.score -gt [int]$ex.score){
      foreach($c in $CamposAvaliacao){ $ex.($c.prop) = $rec.($c.prop) }
    }
  } else { $novoById[$rec.id]=$rec }
}

# ---------- mescla com historico (preserva colunas manuais) ----------
$final = [ordered]@{}
if(-not $Zerar -and (Test-Path -LiteralPath $OutJson)){
  try{ $old=(Get-Content -LiteralPath $OutJson -Raw -Encoding UTF8 | ConvertFrom-Json).licitacoes; foreach($m in $old){ $final[[string]$m.id]=$m } }catch{}
}
$novos=0; $atualizados=0
foreach($k in $novoById.Keys){
  $n=$novoById[$k]
  if($final.Contains($k)){
    $o=$final[$k]
    foreach($campo in 'am','produto','fabricante','situacao'){ if($o.$campo){ $n.$campo=$o.$campo } }
    if($o.sitioCompras){ $n.sitioCompras=$o.sitioCompras }
    if($o.dataInsercao){ $n.dataInsercao=$o.dataInsercao }
    $atualizados++
  } else { $novos++ }
  $final[$k]=$n
}
# ranking: do mais propenso ao menos propenso; empate desempata pelo pregao mais proximo
$lista = @($final.Values) | Sort-Object @{e={ -([int]$_.score) }}, @{e={ -(RankConf $_.confianca) }}, @{e={ $_.dataPregao }}

# ---------- grava JSON (app) ----------
$outObj = [pscustomobject]@{ geradoEm=(Get-Date).ToString('yyyy-MM-dd'); fonte='real'; licitacoes=$lista }
[IO.File]::WriteAllText($OutJson, ($outObj | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))

# ---------- grava planilha .xlsx (OpenXML) com a ordem de 17 colunas fornecida ----------
$listaExport = $lista
if(-not [string]::IsNullOrWhiteSpace($FilterIds)){
  $idSet = New-Object 'System.Collections.Generic.HashSet[string]'
  $FilterIds.Split(',') | ForEach-Object { if($_){ [void]$idSet.Add($_.Trim()) } }
  $listaExport = @($lista | Where-Object { $idSet.Contains([string]$_.id) })
}

# ---- rotulos acentuados (PS 5.1 le .ps1 sem BOM como ANSI; por isso via [char]) ----
$STR_Acao      = "A$([char]0x00C7)$([char]0x00C3)O SUGERIDA"
$STR_DiasAte   = "DIAS AT$([char]0x00C9) O PREG$([char]0x00C3)O"
$STR_SolDet    = "SOLU$([char]0x00C7)$([char]0x00D5)ES DETECTADAS"
$STR_PorQue    = "POR QU$([char]0x00CA) (auto)"
$STR_Posicao   = "POSI$([char]0x00C7)$([char]0x00C3)O"
$STR_Aderencia = "ADER$([char]0x00CA)NCIA T$([char]0x00C9)CNICA"

function JoinArr($a){ if($null -eq $a){ return '' }; return ((@($a) | Where-Object { $_ }) -join ', ') }

# ---- layout principal: as 23 colunas originais (ordem preservada) + 8 novas no fim ----
$COLS_MAIN = @(
  @{ h=$STR_Orgao;                          w=42; t='texto'  },  # 1
  @{ h='AM';                                w=10; t='texto'  },  # 2
  @{ h='ESTADO';                            w=20; t='texto'  },  # 3
  @{ h='PRODUTO';                           w=18; t='texto'  },  # 4
  @{ h=$STR_Descricao;                      w=75; t='longo'  },  # 5
  @{ h='FABRICANTE';                        w=18; t='texto'  },  # 6
  @{ h=$STR_Insercao;                       w=14; t='data'   },  # 7
  @{ h=$STR_Pregao;                         w=14; t='data'   },  # 8
  @{ h=$STR_Sitio;                          w=22; t='texto'  },  # 9
  @{ h=$STR_NPregao;                        w=16; t='texto'  },  # 10
  @{ h='UASG';                              w=22; t='texto'  },  # 11
  @{ h='PUBLICADO';                         w=14; t='data'   },  # 12
  @{ h="N$([char]0x00BA) CONLICITA$([char]0x00C7)$([char]0x00C3)O"; w=14; t='texto' }, # 13
  @{ h='LINK';                              w=38; t='link'   },  # 14
  @{ h='ANO';                               w=8;  t='ano'; ref='H' },  # 15
  @{ h='MES';                               w=8;  t='mes'; ref='H' },  # 16
  @{ h=$STR_Situacao;                       w=14; t='texto'  },  # 17
  @{ h=$STR_Confianca;                      w=14; t='texto'  },  # 18
  @{ h='TIPO FABRICANTE';                   w=24; t='texto'  },  # 19
  @{ h='MOTIVO (auto)';                     w=60; t='longo'  },  # 20
  @{ h="FABRICANTE ($STR_Sugestao)";        w=20; t='texto'  },  # 21
  @{ h="$STR_Solucao ($STR_Sugestao)";      w=28; t='texto'  },  # 22
  @{ h='BOLETIM';                           w=24; t='texto'  },  # 23
  @{ h='SCORE';                             w=9;  t='score'  },  # 24
  @{ h=$STR_Aderencia;                      w=13; t='num'    },  # 25
  @{ h='FAIXA';                             w=18; t='faixa'  },  # 26
  @{ h=$STR_Acao;                           w=48; t='longo'  },  # 27
  @{ h=$STR_DiasAte;                        w=12; t='num'    },  # 28
  @{ h='PARCEIROS DETECTADOS';              w=26; t='texto'  },  # 29
  @{ h='CONCORRENTES DETECTADOS';           w=26; t='texto'  },  # 30
  @{ h=$STR_SolDet;                         w=32; t='texto'  },  # 31
  @{ h='PALAVRAS-CHAVE DETECTADAS';         w=40; t='longo'  }   # 32
)

# ---- aba Ranking: visao compacta, do mais propenso ao menos propenso ----
$COLS_RANK = @(
  @{ h=$STR_Posicao;               w=10; t='num'   },
  @{ h='SCORE';                    w=9;  t='score' },
  @{ h=$STR_Aderencia;             w=13; t='num'   },
  @{ h='FAIXA';                    w=18; t='faixa' },
  @{ h=$STR_Orgao;                 w=42; t='texto' },
  @{ h='ESTADO';                   w=20; t='texto' },
  @{ h=$STR_Pregao;                w=14; t='data'  },
  @{ h=$STR_DiasAte;               w=12; t='num'   },
  @{ h=$STR_Descricao;             w=75; t='longo' },
  @{ h=$STR_PorQue;                w=65; t='longo' },
  @{ h=$STR_SolDet;                w=32; t='texto' },
  @{ h='PARCEIROS DETECTADOS';     w=26; t='texto' },
  @{ h='CONCORRENTES DETECTADOS';  w=26; t='texto' },
  @{ h=$STR_Acao;                  w=48; t='longo' },
  @{ h='LINK';                     w=38; t='link'  }
)

function MakeRowArray($x){
  return @(
    $x.orgao,                    # 1: ORGAO
    $x.am,                       # 2: AM
    $x.estado,                   # 3: ESTADO
    $x.produto,                  # 4: PRODUTO
    $x.objeto,                   # 5: DESCRICAO DO OBJETO
    $x.fabricante,               # 6: FABRICANTE
    (ToBr $x.dataInsercao),      # 7: DATA DA INSERCAO
    (ToBr $x.dataPregao),        # 8: DATA DO PREGAO
    $x.sitioCompras,             # 9: SITIO DE COMPRAS
    $x.numeroPregao,             # 10: N DO PREGAO
    $x.uasg,                     # 11: UASG
    (ToBr $x.publicado),         # 12: PUBLICADO
    $x.numeroConlicitacao,       # 13: N CONLICITACAO
    $x.link,                     # 14: LINK
    $x.ano,                      # 15: ANO (formula)
    $x.mes,                      # 16: MES (formula)
    $x.situacao,                 # 17: SITUACAO
    $x.confianca,                # 18: CONFIANCA
    $x.tipoFabricante,           # 19: TIPO FABRICANTE
    $x.motivoMatch,              # 20: MOTIVO (auto)
    $x.fabricanteSugerido,       # 21: FABRICANTE (sugestao)
    $x.solucao,                  # 22: SOLUCAO (sugestao)
    $x.boletimOrigem,            # 23: BOLETIM
    [int]$x.score,               # 24: SCORE
    [int]$x.scoreTecnico,        # 25: ADERENCIA TECNICA
    $x.faixa,                    # 26: FAIXA
    $x.faixaAcao,                # 27: ACAO SUGERIDA
    $x.diasAtePregao,            # 28: DIAS ATE O PREGAO
    (JoinArr $x.parceiros),      # 29
    (JoinArr $x.concorrentesDet),# 30
    (JoinArr $x.solucoesDet),    # 31
    (JoinArr $x.palavrasChaveDet)# 32
  )
}

function MakeRankRow($x,[int]$pos){
  return @(
    $pos,
    [int]$x.score,
    [int]$x.scoreTecnico,
    $x.faixa,
    $x.orgao,
    $x.estado,
    (ToBr $x.dataPregao),
    $x.diasAtePregao,
    $x.objeto,
    $x.motivoMatch,
    (JoinArr $x.solucoesDet),
    (JoinArr $x.parceiros),
    (JoinArr $x.concorrentesDet),
    $x.faixaAcao,
    $x.link
  )
}

$rowsAlta  = New-Object System.Collections.ArrayList
$rowsConc  = New-Object System.Collections.ArrayList
$rowsMedia = New-Object System.Collections.ArrayList
$rowsTodas = New-Object System.Collections.ArrayList
$rowsRank  = New-Object System.Collections.ArrayList

$pos = 0
foreach($x in $listaExport){
  $pos++
  $r = MakeRowArray $x
  [void]$rowsTodas.Add($r)
  [void]$rowsRank.Add((MakeRankRow $x $pos))
  if($x.confianca -eq 'Alta'){ [void]$rowsAlta.Add($r) }
  elseif($x.confianca -eq 'Concorrente'){ [void]$rowsConc.Add($r) }
  elseif($x.confianca -eq $STR_Media){ [void]$rowsMedia.Add($r) }
}

# estilo da celula FAIXA conforme o score da propria linha.
# Deriva das faixas do config (estilos 4..8 seguem a ordem das faixas), para a cor
# nunca divergir do rotulo quando os limites forem recalibrados.
function EstiloFaixa([int]$s){
  $n = [Math]::Min($FAIXAS.Count, 5)
  for($i=0;$i -lt $n;$i++){
    if($s -ge [int]$FAIXAS[$i].min){ return 4 + $i }
  }
  return 8
}

function Build-SheetXml($cols, $rows){
  $nCols = $cols.Count
  # indice da coluna de score (para colorir a celula FAIXA da mesma linha)
  $scoreIdx = -1
  for($i=0;$i -lt $nCols;$i++){ if($cols[$i].t -eq 'score'){ $scoreIdx = $i; break } }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

  # congela a linha de cabecalho
  [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
  [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/>')

  # larguras
  [void]$sb.Append('<cols>')
  for($i=0;$i -lt $nCols;$i++){
    $n = $i + 1
    [void]$sb.Append('<col min="'+$n+'" max="'+$n+'" width="'+([string]$cols[$i].w)+'" customWidth="1"/>')
  }
  [void]$sb.Append('</cols>')

  [void]$sb.Append('<sheetData>')
  [void]$sb.Append('<row r="1" ht="30" customHeight="1">')
  for($c=0;$c -lt $nCols;$c++){
    $ref=(ColLetter ($c+1))+'1'
    [void]$sb.Append('<c r="'+$ref+'" s="2" t="inlineStr"><is><t xml:space="preserve">'+(XmlEsc $cols[$c].h)+'</t></is></c>')
  }
  [void]$sb.Append('</row>')

  $rn=1
  foreach($row in $rows){
    $rn++
    [void]$sb.Append('<row r="'+$rn+'">')
    for($c=0;$c -lt $nCols -and $c -lt $row.Count;$c++){
      $ref  = (ColLetter ($c+1)) + $rn
      $tipo = [string]$cols[$c].t
      $v    = $row[$c]
      if($null -eq $v){ continue }

      if($tipo -eq 'data'){
        $serial = ToExcelSerial ([string]$v)
        if($serial){ [void]$sb.Append('<c r="'+$ref+'" s="1"><v>'+$serial+'</v></c>') }
      }
      elseif($tipo -eq 'ano' -or $tipo -eq 'mes'){
        # <v> de celula sem t="..." precisa ser numerico; sem numero, nao emite a celula
        $num = 0
        if([int]::TryParse(([string]$v).Trim(), [ref]$num)){
          $fn   = 'YEAR'; if($tipo -eq 'mes'){ $fn = 'MONTH' }
          $hRef = [string]$cols[$c].ref + $rn
          [void]$sb.Append('<c r="'+$ref+'" s="9"><f>IF(ISNUMBER('+$hRef+'),'+$fn+'('+$hRef+'),"")</f><v>'+([string]$num)+'</v></c>')
        }
      }
      elseif($tipo -eq 'score'){
        [void]$sb.Append('<c r="'+$ref+'" s="'+(EstiloFaixa ([int]$v))+'"><v>'+([string][int]$v)+'</v></c>')
      }
      elseif($tipo -eq 'faixa'){
        $t=[string]$v
        if($t.Length -gt 0){
          $sIdx = 3
          if($scoreIdx -ge 0 -and $scoreIdx -lt $row.Count){ $sIdx = EstiloFaixa ([int]$row[$scoreIdx]) }
          [void]$sb.Append('<c r="'+$ref+'" s="'+$sIdx+'" t="inlineStr"><is><t xml:space="preserve">'+(XmlEsc $t)+'</t></is></c>')
        }
      }
      elseif($tipo -eq 'num'){
        $num = 0
        if([int]::TryParse(([string]$v).Trim(), [ref]$num)){
          [void]$sb.Append('<c r="'+$ref+'" s="9"><v>'+([string]$num)+'</v></c>')
        }
      }
      elseif($tipo -eq 'link'){
        $t=[string]$v
        if($t.Length -gt 0){
          # HYPERLINK tem limite pratico de argumento; link muito longo vai como texto
          if($t.Length -gt 250 -or $t.Contains('"')){
            [void]$sb.Append('<c r="'+$ref+'" s="10" t="inlineStr"><is><t xml:space="preserve">'+(XmlEsc $t)+'</t></is></c>')
          } else {
            # t="str" e obrigatorio: sem ele o Excel espera numero dentro de <v> e rejeita o arquivo
            [void]$sb.Append('<c r="'+$ref+'" s="10" t="str"><f>HYPERLINK("'+(XmlEsc $t)+'","abrir edital")</f><v>abrir edital</v></c>')
          }
        }
      }
      elseif($v -is [int] -or $v -is [double] -or $v -is [long]){
        [void]$sb.Append('<c r="'+$ref+'" s="9"><v>'+([string]$v)+'</v></c>')
      }
      else {
        $t=[string]$v
        if($t.Length -gt 0){
          if($t.Length -gt 32000){ $t=$t.Substring(0,32000) }
          $st = '11'; if($tipo -eq 'longo'){ $st = '3' }
          [void]$sb.Append('<c r="'+$ref+'" s="'+$st+'" t="inlineStr"><is><t xml:space="preserve">'+(XmlEsc $t)+'</t></is></c>')
        }
      }
    }
    [void]$sb.Append('</row>')
  }
  [void]$sb.Append('</sheetData>')

  # autofiltro no cabecalho
  $ultima = ColLetter $nCols
  $ultLin = [Math]::Max($rn,1)
  [void]$sb.Append('<autoFilter ref="A1:'+$ultima+$ultLin+'"/>')
  [void]$sb.Append('</worksheet>')
  return $sb.ToString()
}

function Write-MultiSheet-Xlsx([string]$path, $sheetsData){
  $enc = New-Object System.Text.UTF8Encoding($false)
  $pkg = $null
  try {
    if(Test-Path -LiteralPath $path){ try { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } catch {} }
    $pkg = [System.IO.Packaging.Package]::Open($path, [System.IO.FileMode]::Create)
  } catch {
    Write-Host "AVISO: O arquivo '$path' está aberto no Excel. Feche o Excel para que a planilha seja atualizada." -ForegroundColor Yellow
    return
  }

  try {
    $sbWb = New-Object System.Text.StringBuilder
    [void]$sbWb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>')
    
    $sIdx = 1
    foreach($sd in $sheetsData){
      $sheetName = XmlEsc $sd.name
      $rId = "rId$sIdx"
      [void]$sbWb.Append('<sheet name="'+$sheetName+'" sheetId="'+$sIdx+'" r:id="'+$rId+'"/>')
      $sIdx++
    }
    [void]$sbWb.Append('</sheets></workbook>')
    $workbookXml = $sbWb.ToString()

    # Indices de estilo usados no Build-SheetXml:
    #  0=normal  1=data  2=cabecalho  3=texto longo (wrap)
    #  4..8=faixas de score (verde/lima/amarelo/laranja/vermelho)
    #  9=numero centralizado  10=link  11=texto com borda
    $stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="165" formatCode="dd/mm/yyyy"/></numFmts>' +
      '<fonts count="5">' +
        '<font><sz val="11"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FF1F2937"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
        '<font><u/><sz val="11"/><color rgb="FF1D4ED8"/><name val="Calibri"/></font>' +
      '</fonts>' +
      '<fills count="8">' +
        '<fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF16A34A"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF84CC16"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFEAB308"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF97316"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFEF4444"/><bgColor indexed="64"/></patternFill></fill>' +
      '</fills>' +
      '<borders count="2">' +
        '<border><left/><right/><top/><bottom/><diagonal/></border>' +
        '<border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>' +
      '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="12">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
        '<xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
        '<xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
        '<xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
        '<xf numFmtId="0" fontId="3" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>' +
        '<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top"/></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>' +
      '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>'

    $wbUri  = New-Object System.Uri('/xl/workbook.xml',[System.UriKind]::Relative)
    $styUri = New-Object System.Uri('/xl/styles.xml',[System.UriKind]::Relative)

    $wbPart  = $pkg.CreatePart($wbUri,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
    $styPart = $pkg.CreatePart($styUri,'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml')

    $b=$enc.GetBytes($workbookXml); $s=$wbPart.GetStream(); $s.Write($b,0,$b.Length); $s.Flush()
    $b=$enc.GetBytes($stylesXml);   $s=$styPart.GetStream(); $s.Write($b,0,$b.Length); $s.Flush()

    $pkg.CreateRelationship($wbUri,[System.IO.Packaging.TargetMode]::Internal,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument','rId1') | Out-Null

    $sIdx = 1
    foreach($sd in $sheetsData){
      $wsUri = New-Object System.Uri("/xl/worksheets/sheet$sIdx.xml", [System.UriKind]::Relative)
      $wsPart = $pkg.CreatePart($wsUri, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
      $sheetXml = Build-SheetXml $sd.cols $sd.rows
      $b = $enc.GetBytes($sheetXml); $s = $wsPart.GetStream(); $s.Write($b,0,$b.Length); $s.Flush()
      $wbPart.CreateRelationship($wsUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', "rId$sIdx") | Out-Null
      $sIdx++
    }

    $wbPart.CreateRelationship($styUri,[System.IO.Packaging.TargetMode]::Internal,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',"rId$sIdx") | Out-Null
  } finally {
    if($pkg){ $pkg.Close() }
  }
}

$sheetsToExport = @(
  @{ name='Ranking';                  cols=$COLS_RANK; rows=$rowsRank  },
  @{ name=$STR_Confiaveis;            cols=$COLS_MAIN; rows=$rowsAlta  },
  @{ name=$STR_PossConc;              cols=$COLS_MAIN; rows=$rowsConc  },
  @{ name='Escopo Geral (Multiband)'; cols=$COLS_MAIN; rows=$rowsMedia },
  @{ name='Todas Oportunidades';      cols=$COLS_MAIN; rows=$rowsTodas }
)

Write-MultiSheet-Xlsx $OutXlsx $sheetsToExport

# ---------- resumo ----------
Write-Host "`n==================== RESUMO ====================" -ForegroundColor Green
Write-Host ("Lidas (brutas):        {0}" -f $brutos.Count)
Write-Host ("No escopo (dedup):     {0}" -f $novoById.Count)
$porConf = $novoById.Values | Group-Object confianca | ForEach-Object { "{0}={1}" -f $_.Name,$_.Count }
Write-Host ("Por confianca:         {0}" -f ($porConf -join '  '))
Write-Host ("Novos: {0}    Atualizados(hist.): {1}    Total no arquivo: {2}" -f $novos,$atualizados,$lista.Count)
Write-Host ("JSON:     {0}" -f $OutJson)
Write-Host ("Planilha: {0}" -f $OutXlsx)
Write-Host "===============================================" -ForegroundColor Green

# linha legivel por maquina (o app le isto)
$alta       =@($novoById.Values | Where-Object { $_.confianca -eq 'Alta'        }).Count
$concorrente=@($novoById.Values | Where-Object { $_.confianca -eq 'Concorrente' }).Count
$media      =@($novoById.Values | Where-Object { $_.confianca -eq $STR_Media    }).Count
$baixa      =@($novoById.Values | Where-Object { $_.confianca -eq 'Baixa'       }).Count
$revisar    =@($novoById.Values | Where-Object { $_.confianca -eq 'Revisar'     }).Count
$porFaixa = [ordered]@{}
foreach($f in $FAIXAS){ $porFaixa[[string]$f.nome] = @($novoById.Values | Where-Object { $_.faixa -eq $f.nome }).Count }
$vencidos = @($novoById.Values | Where-Object { $_.vencido }).Count
Write-Host ("Por faixa:             {0}" -f (($porFaixa.Keys | ForEach-Object { "{0}={1}" -f $_,$porFaixa[$_] }) -join '  '))
Write-Host ("Pregao ja vencido:     {0}" -f $vencidos)
Write-Host ("Situacoes aceitas:     {0}" -f ($SituacoesAceitas -join ', '))
if($descartadasPorSituacao.Count -gt 0){
  $det = ($descartadasPorSituacao.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "{0}={1}" -f $_.Key,$_.Value }) -join '  '
  Write-Host ("Descartadas p/ situacao: {0}" -f $det) -ForegroundColor Yellow
}

$res=[pscustomobject]@{ lidas=$brutos.Count; escopo=($alta+$concorrente+$media+$baixa); revisar=$revisar; alta=$alta; concorrente=$concorrente; media=$media; baixa=$baixa; total=$lista.Count; novos=$novos; atualizados=$atualizados; zerado=[bool]$Zerar; porFaixa=$porFaixa; vencidos=$vencidos }
Write-Output ("RESULTADO_JSON " + ($res | ConvertTo-Json -Compress))
