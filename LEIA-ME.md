# ENTERPRISECORE Radar

Duas abas **independentes** na mesma janela:

| Aba | Para quê | De onde vêm os dados |
|---|---|---|
| 📢 **Radar de Boletins** | varredura diária do que foi publicado | os `.xlsx` que você baixa do Conlicitação |
| 🛠️ **Painel de ROs** | o painel Streamlit que você já usa | `Gestao_ROs_Bitrix` |

## Os quatro manuais

Cada aba tem dois botões próprios, que abrem PDF numa aba nova — não precisa procurar arquivo:

| Botão | Onde | Para quê |
|---|---|---|
| 📄 Manual de Uso | aba Radar de Boletins | rotina do dia, os dois números da nota, dossiê para IA, planilha |
| 📘 Manual de Construção | aba Radar de Boletins | como este painel foi construído do zero, em fases, para quem for adaptar |
| 📄 Manual de Uso | aba Painel de ROs | o pipeline de 3 scripts, decisão dos AMs, configuração de fabricantes |
| 📘 Manual de Construção | aba Painel de ROs | como o Gestão de ROs foi montado e como o Radar o embute nesta tela |

Fontes em `docs/*.html` (menos o de uso do ROs, que reaproveita o `MANUAL_DE_USO.md` já existente em
`Gestao_ROs_Bitrix`). Para atualizar um manual: edite o `.html`, gere o PDF de novo com o Edge headless
(mesmo comando do `ENTERPRISECORE_Manual_Completo.pdf`, trocando o arquivo de entrada e saída) e publique.

---

## A rotina do dia, do começo ao fim

Esta é a sequência de trabalho. Ela também aparece no topo do painel, no bloco
**🗒️ Sequência do dia** — que fica aberto na primeira vez e pode ser recolhido depois.

### 1. Abrir o painel do Conlicitação

<https://consulteonline.conlicitacao.com.br/painel> — com o seu login.

### 2. Baixar os 3 boletins do dia anterior (D-1)

São sempre os **três**, referentes a **D-1**. O corte em D-1 é o que garante que nada
publicado ontem escape da varredura.

### 3. Salvar os três `.xlsx` em `data\boletins`

Use o botão **📂 Abrir pasta** do painel: ele abre o local certo no Explorer para você
soltar os arquivos. Não precisa renomear nada.

> Se a pasta não existir (cópia recém-baixada do GitHub), ela é criada sozinha no primeiro
> processamento — ou pelo botão.

### 4. Rodar o processamento

Botão **⚡ Processar boletins**. Ele lê os três arquivos, mantém só as linhas com
Situação = **NOVA**, calcula o score de cada licitação e ordena a lista.

Leva alguns segundos: o Excel é aberto em segundo plano, só para leitura.

> A caixa **"Somar ao histórico"** muda o comportamento. Desmarcada (padrão), a tela passa a
> refletir só os boletins que estão na pasta agora. Marcada, acumula com o que já havia.
> Para a rotina diária, deixe desmarcada.

### 5. Selecionar o que é do escopo da empresa

Marque as licitações na lista. A ordenação já traz o mais aderente no topo, e os filtros de
faixa (Parceiros ENTERPRISECORE, Concorrentes, Escopo Geral) ajudam a varrer — mas **a decisão é sua**:
confira o objeto antes de marcar. O score é uma sugestão, não um veredito.

### 6. Baixar o Excel e passar para a planilha de publicações

Botão **📊 Baixar planilha (.xlsx)**. Copie as linhas do que você selecionou para a
**planilha de publicações**.

As colunas **AM**, **PRODUTO** e **FABRICANTE** continuam sendo preenchidas à mão — o
processamento não tem como saber esses três.

### 7. Aprovar no painel de publicações

Aprove lá tudo que foi identificado como sendo do escopo da empresa. É o passo que fecha o
ciclo do dia.

---

### Sobre a antiga aba "Vigilância de Pregões"

Foi **retirada da interface em 12/08/2026**, porque o que ela precisava acompanhar — as
mensagens do pregoeiro na sala de disputa — não existe em API pública: a área é protegida por
hCaptcha invisível.

**O motor continua no disco e funcionando**, caso você queira de volta:

| O quê | Onde |
|---|---|
| Cliente da API do Compras.gov/PNCP | `app/compras-api.js` |
| Rotas | `/api/acompanhar`, `/api/desacompanhar`, `/api/checar-status`, `/api/acompanhados` |
| Dados dos pregões cadastrados | `data/acompanhados.json` |
| Histórico de movimentações | `data/historico-eventos.json` |

Para reativar: descomente o botão `btnModVig` no `ui.html` e as duas linhas do `setInterval`
no fim do `server.js` (procure por `INTERVALO_MIN`).

O que ele entregava, e que você deixa de ter: suspensão, revogação e anulação; resultado
publicado com vencedor e valor homologado; retificação e reabertura **com o motivo escrito**;
e aviso de edital novo publicado.

---

# Radar de Boletins

Central de triagem dos boletins da Conlicitação. Lê os `.xlsx` do dia, mantém só o que está
como **NOVA**, dá uma **nota de 0 a 100** para cada licitação explicando o porquê, e monta o
material para você selecionar, exportar em Excel e gerar o modelo de e-mail `.OFT` do Outlook.

Roda **sem instalar nada**: usa o **Node** que já existe na máquina + **PowerShell** + **Excel/Outlook
clássicos** (automação COM). Não depende de Python nem de pacotes npm.

---

## Como usar (dia a dia)

1. Baixe os boletins do Conlicitação e jogue os `.xlsx` em **`data\boletins`** (1 ou vários).
2. Duplo-clique em **`Abrir-App.cmd`** → abre `http://localhost:8790`.
   Uma janela preta ("ENTERPRISECORE Radar (servidor - NAO FECHE)") fica aberta; **não feche** enquanto usar.
3. Clique **⟳ Processar boletins**.
4. A lista vem ordenada **do mais propenso ao menos propenso**. Em cada card:
   - o **SCORE** grande e a faixa (Prioridade Máxima → Descartar);
   - **"Como esta nota foi formada"**: cada sinal e quantos pontos somou ou tirou;
   - as marcas e soluções detectadas;
   - a ação sugerida.
5. Marque as que interessam e use a barra de baixo:
   - **🤖 Copiar dossiê p/ IA** — joga tudo no clipboard, você cola aqui no Claude;
   - **📊 Exportar Excel** — planilha só com as selecionadas;
   - **✉️ Gerar modelo .OFT** — digita o Para/Cc/Assunto e baixa o modelo do Outlook.

---

## Os dois números

| Coluna | O que é |
|---|---|
| **SCORE** | Propensão final. É a aderência técnica **mais** o efeito do calendário. É por ele que a lista é ordenada. |
| **ADERÊNCIA TÉCNICA** | Só o encaixe no portfólio, **sem** o efeito do prazo. |

Existem separados de propósito: um edital pode ser perfeito para a ENTERPRISECORE e o pregão já ter
passado. Nesse caso a aderência continua alta e o score final cai — a informação não se perde.
O filtro **"Ocultar pregões vencidos"** vem ligado por padrão.

---

## O dossiê para IA

O botão **🤖 Dossiê p/ IA** monta um texto que já vem com:

- a **tarefa** escrita para a IA (o que avaliar, que riscos checar, que formato responder);
- o **escopo comercial completo da ENTERPRISECORE** — todos os parceiros, todos os concorrentes,
  todas as soluções com as palavras-chave que as disparam e as faixas de score;
- para cada licitação: órgão, estado, datas, nº do pregão, UASG, link do edital, o objeto
  inteiro, os sinais detectados e a **decomposição da nota**.

Cole no Claude (ou em qualquer IA) e peça a análise profunda. **Isso exige internet** — o que
roda offline é só o cálculo do score e a montagem do texto.

---

## O que você edita (sem mexer em código)

| Arquivo | Para quê |
|---|---|
| `config/escopo-enterprisecore.json` | O cérebro. Fabricantes parceiros, concorrentes, soluções, palavras-chave, exclusões — **e os `pesos` e `faixas` do score**. |
| `config/contatos.json` | Lista de destinatários (nome, e-mail, papel). |

### Calibrar o ranking

Tudo fica no bloco `pesos` do `escopo-enterprisecore.json`. A lógica: **a categoria de solução é a
condição de entrada** (um edital de firewall é núcleo do negócio mesmo sem marca citada),
por isso pesa mais que a marca; a marca parceira é um multiplicador em cima disso.

- Muita coisa irrelevante subindo? Aumente `min` das `faixas` ou baixe `categoriaPrimeira`.
- Coisa boa afundando? O contrário.
- Pregão vencido incomodando? Mexa em `prazoVencido`.

As palavras-chave devem ser escritas **em minúsculo e sem acento** — o pipeline normaliza o
texto do edital do mesmo jeito antes de comparar. Só os rótulos de tela (`faixas[].nome`,
`acao`) podem ter acento.

---

## A planilha

### As 17 colunas do padrão da equipe

As abas de dados mantêm exatamente esta ordem. O que o Radar preenche e o que sobra
para você digitar:

| # | Coluna | Vem preenchida? |
|---|---|---|
| 1 | ÓRGÃO | ✅ sempre |
| 2 | **AM** | ✍️ **você preenche** |
| 3 | ESTADO | ⚠️ 79% — o resto vem sem UF no boletim |
| 4 | **PRODUTO** | ✍️ **você preenche** |
| 5 | DESCRIÇÃO DO OBJETO | ✅ sempre |
| 6 | **FABRICANTE** | ✍️ **você preenche** (o edital raramente cita marca) |
| 7 | DATA DA INSERÇÃO | ✅ sempre |
| 8 | DATA DO PREGÃO | ✅ 99% — data real, não texto |
| 9 | SÍTIO DE COMPRAS | ⚠️ 6% — o boletim quase nunca traz |
| 10 | Nº DO PREGÃO | ✅ sempre |
| 11 | UASG | ⚠️ 79% — às vezes vem o nº do processo no lugar |
| 12 | PUBLICADO | ✅ sempre |
| 13 | Nº CONLICITAÇÃO | ✅ sempre |
| 14 | LINK | ✅ 97% |
| 15 | ANO | 🔢 fórmula `=IF(ISNUMBER(H2),YEAR(H2),"")` |
| 16 | MES | 🔢 fórmula `=IF(ISNUMBER(H2),MONTH(H2),"")` |
| 17 | SITUAÇÃO | ✅ nasce como "Pendente" — você atualiza conforme trabalha |

**Três colunas são sempre suas:** AM, PRODUTO e FABRICANTE. O resto o Radar tenta
preencher, e o que ficar em branco é porque o boletim não trouxe.

> ANO e MES são fórmulas de verdade, ligadas à coluna H. Se você corrigir a data do
> pregão, os dois recalculam sozinhos.

### As abas

`data/Planilha-ENTERPRISECORE-Licitacoes.xlsx`, com 5 abas:

1. **Ranking** — visão compacta, do mais propenso ao menos, com o porquê e o link.
2. **Parceiros ENTERPRISECORE (Confiáveis)** · 3. **Possíveis Concorrentes** · 4. **Escopo Geral** · 5. **Todas**

As abas 2–5 mantêm **as 23 colunas na ordem original** e acrescentam no fim: SCORE,
ADERÊNCIA TÉCNICA, FAIXA, AÇÃO SUGERIDA, DIAS ATÉ O PREGÃO, PARCEIROS, CONCORRENTES,
SOLUÇÕES e PALAVRAS-CHAVE detectadas.

Vem com cabeçalho congelado, autofiltro, largura de coluna ajustada e a faixa colorida.
É gravada por OpenXML puro (.NET) — **não abre o Excel** para escrever.

---

## Requisitos

- **Outlook clássico** instalado e logado (o "novo Outlook"/web não expõe automação COM).
- **Excel** instalado (a leitura dos boletins usa Excel COM em modo somente-leitura).
- **Node.js** e Windows PowerShell (ambos já validados nesta máquina).

---

## Solução de problemas

- **"A porta 8790 já está em uso"** → já existe um servidor rodando. Use a aba aberta, ou
  feche a janela do servidor e rode de novo.
- **Alterei o `escopo-enterprisecore.json` e nada mudou** → é preciso clicar **⟳ Processar boletins**
  de novo; o score é calculado na hora do processamento, não na hora de exibir.
- **A planilha não atualiza** → ela está aberta no Excel. Feche e reprocesse (o script avisa).
- **Erro ao gerar o .OFT** → confirme que o Outlook clássico está instalado e logado.
- **Selecionei, reprocessei e o .OFT deu erro** → recarregue a página; os IDs mudam quando o
  conjunto de boletins muda.

---

# Painel de ROs

A aba 2 sobe o Streamlit de `Gestao_ROs_Bitrix` (porta 8501) e mostra ele embutido. É o mesmo
painel de sempre, com os mesmos filtros e o mesmo "🔄 Recarregar Dados" — não existe uma segunda
cópia para manter em sincronia. Se você já tiver aberto pelo `.bat`, o Radar detecta e usa o que
está no ar.

---

## Como o filtro evita falso positivo

O filtro casa **palavra-chave no texto do objeto**, e os boletins da Conlicitação frequentemente
empacotam um lote inteiro num objeto de milhares de caracteres. Três regras seguram isso:

**1. Densidade.** Uma palavra achada num objeto de 6.000 caracteres não vale o mesmo que um
edital inteiro sobre aquilo. Com 1–2 palavras acima de 2.500 caracteres, a pontuação de
categoria cai para 40%; 1 palavra acima de 1.200, cai para 60%.

**2. Contexto de outra área** (`contextosForaEscopo` no config). Dois ou mais termos de outro
ramo — estúdio, mesa de som, mobiliário, veículo, obra, medicamento — com sinal de escopo fraco
derrubam o edital em 45 pontos.

**3. Prazo só pontua com aderência.** Edital que não é seu não ganha ponto por ter prazo folgado.

> **Trava importante:** a penalidade de contexto **nunca** se aplica quando há marca parceira ou
> concorrente citada. Se o edital cita Check Point, é seu — mesmo que o texto misture outras coisas.

Cuidado ao acrescentar termos em `contextosForaEscopo`: evite palavras ambíguas. `lente` e
`tripé` já zeraram por engano um edital legítimo de CFTV (câmera tem lente) e tiveram de sair.

A política continua **"melhor sobrar do que faltar"** — quem fecha a conta é a leitura do
edital, e é para isso que serve o dossiê para IA.
