# 🎯 Procurement Radar Scoring — Motor de Triagem & Pontuação de Editais Públicos com IA

Motor de triagem automatizada e inteligência comercial para **licitações públicas e editais governamentais**, implementando um algoritmo de **Scoring Heurístico Multicritério (0 a 100)** para quantificar a aderência técnica de um edital ao portfólio de produtos/serviços de uma empresa fornecedora.

Inclui gerador de **dossiês estruturados para Large Language Models (LLMs)** e exportação de relatórios táticos de certames.

---

## 📌 Que Problema Resolve?

Empresas que disputam licitações públicas enfrentam uma avalanche de diários oficiais e boletins (como PNCP, Compras.gov.br e portais estaduais), com mais de **1.000 novos editais publicados todos os dias**. 

A leitura manual de termos de referência de 150 páginas por analistas humanos gera dois problemas graves:
1. **Perda de oportunidades valiosas** por falta de braço operacional para filtrar tudo a tempo.
2. **Desperdício de centenas de horas** analisando certames onde a empresa não tem atestado de capacidade técnica ou margem competitiva.

O **Procurement Radar Scoring** resolve isso processando os boletins brutos e calculando em segundos uma pontuação preditiva de 0 a 100 para cada certame, priorizando o pipeline comercial apenas nas oportunidades de alta probabilidade de vitória.

---

## ⚙️ Diferencial Técnico & Algoritmo de Scoring

O algoritmo analisa o texto integral do edital, objeto, itens e termos de referência, aplicando um pipeline de ponderação vetorial:

1. **Camada de Filtro Negativo (Hard Stop):**
   Descarta imediatamente certames contendo termos bloqueantes (ex: pregões exclusivos para ME/EPP quando a empresa for de grande porte, ou exigências geográficas restritivas).
2. **Ponderação por Relevância de Termos Técnicos (Soft Scoring):**
   - **Família de Produtos Primária:** Peso $W_1 = 4.0$ (ex: Firewalls NGFW, EDR/XDR, Nuvem, Licenciamento).
   - **Fabricantes Homologados:** Peso $W_2 = 3.0$ (marcas parceiras do canal).
   - **Serviços Especializados:** Peso $W_3 = 2.0$ (implantação, suporte 24x7, migração de dados).
3. **Normalização Sigmoidal / Escala 0-100:**
   $$Score = \min\left(100, \sum_{i=1}^{n} (\text{Frequência}_i \times W_i) \times \text{Fator de Contexto}\right)$$
4. **Gerador de Dossiê Executivo para IA:**
   Estrutura o extrato do edital em JSON/Markdown limpo contendo número do processo, órgão, data da sessão, valor estimado e trechos-chave de habilitação para envio direto a agentes de IA.

---

## 🏗️ Stack Tecnológica

- **Backend & Scripting:** Node.js (v18+) e Python 3.10+.
- **Processamento de Texto & NLP:** Expressões regulares compiladas de alta performance, tokenização e parsing de arquivos `.csv`, `.xlsx` e `.json`.
- **Relatórios:** Geração de planilhas formatadas e templates táticos de e-mail.

---

## 🚀 Como Executar Localmente

```bash
# 1. Clone o repositório
git clone https://github.com/HenriMafra/procurement-radar-scoring.git
cd procurement-radar-scoring

# 2. Instale as dependências
npm install

# 3. Configure seu catálogo de palavras-chave em config/catalogo.json
# 4. Execute a triagem do boletim
node index.js --input ./dados/boletim_exemplo.json
```

---

## 📄 Licença

Distribuído sob a licença **MIT**. Desenvolvido por **Henri Mafra**.
