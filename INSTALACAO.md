# ENTERPRISECORE Radar — instalação em um PC novo

Leia isto antes de usar. São 5 minutos.

O Radar **não instala nada** e **não precisa de internet para funcionar** (só para você baixar
os boletins e para colar o dossiê numa IA). Ele roda a partir da própria pasta.

---

## 0. De onde vem o programa

O código fica num repositório **privado** no GitHub:

```
https://github.com/HenriMafra/enterprisecore-radar
```

Você precisa ter sido convidado como colaborador para conseguir acessar.

Você precisa ter sido convidado como colaborador para conseguir acessar. Sem o convite,
os links abaixo dão "404".

**Baixando pela primeira vez** — o jeito simples, sem instalar nada:

1. Abra <https://github.com/HenriMafra/enterprisecore-radar/releases>
2. Na versão do topo, clique em **Source code (zip)**
3. Descompacte onde quiser

Se você tem Git instalado, prefira o clone — assim atualizar depois vira um comando só:

```bash
git clone https://github.com/HenriMafra/enterprisecore-radar.git
```

**Atualizando depois**, quando sair versão nova:

- Se você clonou: abra o Prompt na pasta e rode

```bash
git pull
```

- Se você baixou o `.zip`: baixe o novo em **Releases**, descompacte numa pasta nova e
  copie para lá o seu `config\contatos.json`. Não descompacte por cima da pasta antiga.

> Seus dados **não são sobrescritos** ao atualizar: boletins, planilha gerada,
> `contatos.json` e o histórico ficam de fora do repositório de propósito.

---

## 1. Copie a pasta

Coloque a pasta `ENTERPRISECORE-Radar` onde você quiser — Documentos, Desktop, `C:\`, um pendrive.
**Não importa o caminho nem o nome do usuário do Windows:** o Radar se localiza sozinho.

> Evite só uma coisa: rodar de dentro de pasta sincronizada com OneDrive/SharePoint enquanto
> processa. O Excel trava arquivo durante a leitura e a sincronização pode brigar com isso.

---

## 2. Rode o verificador

Duplo-clique em **`Verificar-Ambiente.cmd`**.

Ele checa tudo e escreve em português o que está pronto e o que falta. Cada linha vem marcada:

| Marca | Significa |
|---|---|
| `[OK]` | pronto, nada a fazer |
| `[AVISO]` | funciona, mas tem detalhe para você saber |
| `[FALTA]` | **precisa resolver antes de usar** |

Se terminar com *"0 faltando"*, pode usar.

---

## 3. O que precisa estar instalado

| Programa | Para quê | Sem ele |
|---|---|---|
| **Node.js 18+** | roda o servidor | ❌ nada funciona — baixe em [nodejs.org](https://nodejs.org) (versão LTS) |
| **Excel** | ler os boletins `.xlsx` | ❌ o processamento não roda |
| **Outlook clássico** | gerar `.OFT` e abrir e-mail | ⚠️ o resto funciona; só o e-mail para de funcionar |
| **PowerShell 5.1** | já vem no Windows 10/11 | — |
| Python + Streamlit | aba "Painel de ROs" | ⚠️ só essa aba deixa de abrir |

### Sobre o Outlook

Tem que ser o **Outlook clássico**. O "novo Outlook" do Windows 11 (aquele que parece o site)
**não permite automação** — não gera `.OFT` nem abre e-mail preenchido. Se os dois estiverem
instalados, deixe o clássico como padrão.

### Sobre o Python

Se você digitar `python` no Prompt e aparecer *"Python was not found; run without arguments to
install from the Microsoft Store"*, isso é um **atalho falso de 0 byte** que a Microsoft coloca
no PATH. O Radar já sabe ignorá-lo e procura o Python de verdade sozinho.

Para tirar o atalho falso do caminho: *Configurações → Aplicativos → Configurações avançadas
de aplicativos → Aliases de execução de aplicativo* → desligue as duas chaves do Python.

---

## 4. Ajustes que talvez você precise fazer

Tudo fica em **`config\ambiente.json`**. Os campos vêm **vazios de propósito**: vazio significa
*"descubra sozinho"*. Só preencha se o verificador reclamar.

```json
{
  "porta": 8790,
  "python": "",
  "painelROs": { "pasta": "", "porta": 8501 },
  "painelScripts": ""
}
```

| Campo | Quando mexer | O que colocar |
|---|---|---|
| `porta` | a porta 8790 já é usada por outro programa | outro número, ex. `8791` |
| `python` | o verificador não achou o Python | caminho completo do `python.exe` |
| `painelROs.pasta` | você usa o painel de ROs e ele está noutro lugar | pasta que contém o `app.py` |
| `painelScripts` | você usa o painel de scripts | caminho do `scripts.json` |

> **Cuidado ao editar:** é um arquivo JSON. Não esqueça vírgulas nem aspas. Se errar, o Radar
> avisa no console com o erro e o link do validador. Barras invertidas precisam ser **dobradas**:
> `"C:\\Users\\fulano\\Python\\python.exe"`.

---

## 5. Troque os dados do dono anterior

| Arquivo | O que tem | O que fazer |
|---|---|---|
| `config\contatos.json` | e-mails da equipe (destinatários) | troque pelos seus |
| `data\enviados-log.json` | histórico de envios | para zerar, deixe só `[]` |
| `data\boletins\` | boletins antigos | apague e ponha os seus |

> Se a cópia foi gerada com o `Preparar-Entrega.cmd`, isso **já vem limpo** — só o
> `contatos.json` é mantido, porque normalmente a lista da equipe continua valendo.

---

## 6. Use

A rotina completa do dia — do Conlicitação até a aprovação no painel de publicações — está
no topo do próprio app, no bloco **🗒️ Sequência do dia**, e detalhada no **`LEIA-ME.md`**.

Em resumo:

1. Baixe os **3 boletins de D-1** em <https://consulteonline.conlicitacao.com.br/painel>
2. Jogue os `.xlsx` em **`data\boletins`**
3. Duplo-clique em **`Abrir-App.cmd`**
4. Clique em **⚡ Processar boletins**
5. Selecione o que é do escopo, baixe a planilha e aprove no painel de publicações

Uma janela preta fica aberta — é o servidor. **Não feche** enquanto usar. Para encerrar, feche
essa janela.

Depois disso, leia o **`LEIA-ME.md`**, que explica o score, o dossiê para IA e a planilha.

---

## Problemas comuns

**"A porta 8790 já está em uso"**
Já tem um Radar aberto. Use a janela existente, ou feche a janela preta e abra de novo.
Se for outro programa ocupando, mude `porta` no `ambiente.json`.

**Cliquei em Processar e deu erro**
Quase sempre é o Excel: ou não está instalado, ou a planilha
`data\Planilha-ENTERPRISECORE-Licitacoes.xlsx` está aberta. Feche o Excel e tente de novo.

**Editei o `escopo-enterprisecore.json` e nada mudou**
O score é calculado no processamento, não na exibição. Clique em **⟳ Processar boletins**.

**Editei um `.json` e o Radar ignorou**
Erro de digitação no arquivo. Olhe a janela preta do servidor: ele escreve o erro e a dica.

**O `.OFT` não gera**
Outlook clássico não instalado, ou está o "novo Outlook". Veja a seção 3.

**A aba "Painel de ROs" não abre**
Python ou Streamlit faltando, ou a pasta do painel está noutro lugar. O verificador diz qual é.
Para instalar o Streamlit: `"caminho\do\python.exe" -m pip install streamlit`

**Acentos aparecem errados (Ã©, Ã§) na planilha**
Não deve acontecer. Se acontecer, avise — é sinal de que algum `.ps1` foi salvo em codificação
errada por um editor.

---

## O que NÃO vem junto

- **Os boletins.** São seus, você baixa do Conlicitação com o seu login.
- **A aba de Vigilância de Pregões.** Foi retirada da interface (o motor está no disco;
  o `LEIA-ME.md` explica como religar).
- **Chave de API.** Não existe nenhuma. O dossiê para IA é texto que você cola onde quiser.

---

## Para quem está entregando

Rode **`Preparar-Entrega.cmd`**. Ele cria uma pasta nova com uma cópia limpa — sem os seus
boletins, sem histórico de envios, sem a planilha gerada e com o `ambiente.json` de volta ao
padrão. **A sua instalação não é alterada.** Depois é só compactar essa pasta e enviar.
