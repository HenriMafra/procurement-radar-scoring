---
name: calibrar-radar
description: Calibra o motor de score do ENTERPRISECORE Radar contra julgamento humano. Use quando o usuário apontar licitações mal classificadas — mandar números de Conlicitação que considera viáveis, dizer que algo "entra" ou "não entra", ou reclamar do ranking. Também para adicionar lotes de palavras-chave de fonte externa.
---

# Calibrar o motor do ENTERPRISECORE Radar

O usuário aponta casos; você acha a **propriedade geral** que separa certo de errado e corrige o algoritmo. Ele foi explícito: *"tente melhorar o algoritmo de entendimento e não somente alterar direto manualmente"*.

## Regra de ouro

**Nunca conserte o caso. Conserte a propriedade que o caso revela.**

Diante de um erro, pergunte: *qual característica do texto separa este edital dos que estão certos?* Só então escreva regra. Quando um remendo manual for inevitável (listas de vocabulário sempre serão em parte), **crie junto a checagem automática que acharia aquele item sozinha da próxima vez**.

Ao entregar, **separe explicitamente o que foi correção estrutural do que foi remendo**. Ele quer enxergar a diferença.

---

## O ciclo

### 1. Meça antes de tocar em nada

```bash
node .claude/skills/calibrar-radar/avaliar-selecao.js 19281042 19281156 19281648
```

Passe os números de Conlicitação que o usuário aprovou. O script imprime:

- a **posição e a nota** de cada escolhido no ranking atual;
- **recall**: quantos no top 10, 20, 30;
- **o que o motor pôs no topo e ele NÃO escolheu** — é aqui que os erros aparecem;
- para cada escolhido mal ranqueado: palavras casadas, termos fora de escopo, decomposição da nota e o objeto.

Sem número nenhum, ele só lista o topo e os candidatos a falso positivo.

### 2. Leia o objeto de cada erro

Não confie na nota. Abra o texto. As causas reais encontradas até hoje foram quase sempre invisíveis na métrica:

- `detecção de intrusão` num desktop = **sensor de abertura do gabinete**
- `virtualização` num desktop = **VT-x do processador**
- `balanceamento` = **balanceamento de pneus**
- `NOc` = **Navio Oceanográfico** da Marinha
- `PAM` = **Pronto Atendimento Municipal**, e marca de válvulas
- `IPS` = **painel de monitor LCD**
- `nas` = a **preposição** portuguesa
- `empilhamento` = **potes herméticos empilháveis**
- texto colado: `ENDPOINTANTIVÍRUS` não casa `endpoint`

### 3. Corrija a propriedade, não o caso

Onde mexer, em ordem de preferência:

| Sintoma | Onde corrigir |
|---|---|
| Palavra genérica carregando edital sozinha | `especificidade.ambiguas` no config |
| Termo inequívoco valendo pouco | `especificidade.nucleo` |
| Categoria inteira faltando vocabulário | `categorias[].palavrasChave` |
| Área que não é da empresa pontuando | `contextosForaEscopo` |
| Duas palavras que só valem juntas | bloco `expressoes` (conjunção com janela) |
| Regra de pontuação errada | `pesos` no config |
| Comportamento estrutural | `scripts/processar-boletins.ps1` |

**Nada de peso chumbado no `.ps1`** — toda calibragem vive no `config/escopo-enterprisecore.json`.

### 4. Reprocesse e liste o que mudou

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/processar-boletins.ps1
node .claude/skills/calibrar-radar/diff-scores.js
```

`diff-scores.js` compara com o snapshot anterior e lista **todos** que mudaram de nota, com a palavra que disparou.

**Leia o objeto de cada um.** Foi assim que se descobriu que 5 de 6 mudanças eram falso positivo, na rodada em que o operador `&` foi implementado sem janela de proximidade. Sem essa leitura, teria sido publicado "60 expressões novas funcionando".

### 5. Re-meça e publique

Rode o passo 1 de novo. **O recall tem de subir e nenhum caso bom pode ter caído.** Depois:

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/publicar-versao.ps1 -Mensagem "o que mudou"
```

---

## Escopo da ENTERPRISECORE — fatos que o usuário estabeleceu

**Entra:** firewall/NGFW/IPS · WAF/ADC/anti-DDoS · EDR/XDR/antivírus corporativo · SIEM/SOC/SOAR · DLP · PAM/IAM/MFA (inclui Active Directory, Entra ID, LDAP) · backup e DR · storage · rede e wireless corporativo (switch, access point, transceiver, DWDM) · NAC · criptografia/PKI/HSM · cloud security · AppSec · CFTV e controle de acesso · infra/HCI/servidores · big data e Hadoop · **e serviços**: garantia de equipamentos, central de serviços, gerenciamento de redes e de infraestrutura, instalação, parametrização, integração.

> **81% do que a ENTERPRISECORE fatura é SERVIÇO, não produto** — medido em 173 itens que a empresa venceu de fato (CNPJ 09.137.728/0001-34, R$ 507 mi). As maiores linhas: garantia de equipamentos de TIC, licenciamento de software, central de serviços, gerenciamento de redes.

**Não entra** (todos ditos por ele):
- **Link e IP dedicado, internet cabeada** — *"não mexemos com ele cabeado"*
- **Áudio, vídeo, projetor, televisor, videowall** — *"a gente não mexe com isso daí"*
- **Celular, tablet, smartphone**
- **Desktop, notebook, periférico de mesa** (teclado, mouse, monitor)
- Telefonia/PABX · impressora · mobiliário · obra civil · manutenção predial e elétrica · certificado digital commodity (e-CPF/e-CNPJ)

> **CESTA MISTA É EXCEÇÃO IMPORTANTE.** Dito por ele: *"caso tenha coisas como notebook e desktop junto de servidor ou switch, não quer dizer que nós vendemos desktop ou notebook, mas eventualmente pode ser que todos sejam itens separados e assim nós teríamos interesse apenas pela parte de switch e servidor"*. **Nunca zere um edital só porque a abertura cita desktop — se a mesma abertura também cita servidor ou switch, ele entra.** É o que a variável `$escopoNaCabeca` protege.

---

## Regras estruturais já implementadas — entenda antes de mexer

| Regra | O que faz | Config |
|---|---|---|
| **Especificidade** | palavra ambígua sozinha vale 35%; termo de núcleo vale 150% | `especificidade` |
| **Especificidade por item** | no lote, o item só conta com palavra específica **ou** 2 ambíguas | `minAmbiguasPorItem` |
| **Contexto desambigua** | ambígua conta se a categoria dela teve específica no edital | — |
| **Cabeça do objeto** | o que se compra está nos primeiros 180 chars | `foraNaCabecaChars` |
| **Escopo na abertura** | se a abertura é nossa, a proporção de lote não reduz | — |
| **Registro de Preços** | em RP por item não se penaliza proporção; desconta fatia | `modalidade` |
| **Conjunção com janela** | `a & b` exige as duas a ≤40 chars | `expressaoJanelaChars` |
| **Marca de sinal fraco** | Dell/HPE/Lenovo/Positivo/IBM só pontuam com escopo já detectado | `sinalFraco` |

---

## Armadilhas que já custaram caro

- **Ao remover termo, confira singular E plural.** `link dedicado` foi removido e `links dedicados` ficou — link dedicado voltou a pontuar 77.
- **Contagem de palavras engana quando elas vêm do mesmo item do lote.** Já enganou duas vezes. Use a proporção de itens.
- **Fonte confiável não significa termo seguro.** Meça contra o corpus antes de aplicar.
- **Cuidado ao excluir termo ambíguo**: `lente`, `tripé`, `estante`, `armário` zeraram um CFTV legítimo da Marinha.
- **Encoding**: `ui.html` e `.ps1` são UTF-8 **sem BOM**. Use `[IO.File]::ReadAllText/WriteAllText` com `UTF8Encoding($false)`. Nunca `Get-Content` + `Set-Content -Encoding UTF8` — corrompe.
- **JavaScript**: `$'` em `String.replace` insere o texto após o casamento. Use função como substituto.
- **`.ps1` é ASCII puro** no código; acento só em comentário ou vindo do JSON.

## Corpus para medição em larga escala

93.253 contratações federais reais em `scratchpad/coleta/comprasgov-bruto.jsonl`. Se não existir, recolete:

```
https://dadosabertos.compras.gov.br/modulo-contratacoes/1_consultarContratacoes_PNCP_14133
  ?pagina=1&tamanhoPagina=500&dataPublicacaoPncpInicial=AAAA-MM-DD
  &dataPublicacaoPncpFinal=AAAA-MM-DD&codigoModalidade=6
```

`codigoModalidade` é obrigatório. Janela máxima de 365 dias. 429 vem com o número de segundos na mensagem.

Métrica por palavra: **quantos editais dispara**, **em que fração dispara sozinha** (sem outra palavra do dicionário no objeto) e **em que fração vem junto de termo de núcleo**. Dispara muito + quase sempre sozinha + nunca com núcleo = carregadora de ruído.
