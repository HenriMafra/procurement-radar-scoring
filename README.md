# Procurement Radar Scoring: Multi-Criteria Heuristic Evaluation and NLP Pipeline for Public Tenders

**Author:** Henri Mafra  
**License:** MIT License  
**Domain:** Natural Language Processing, Information Retrieval, Government Technology (GovTech)  

---

## 1. Overview

Procurement Radar Scoring is an automated intelligence pipeline designed to ingest, normalize, and evaluate public bidding notices published across national procurement repositories (such as the Brazilian National Portal of Public Procurement - PNCP). The engine implements a **Multi-Criteria Heuristic Scoring Algorithm (0 to 100)** to quantify catalog alignment, generate structured executive briefs, and prioritize sales pipeline execution.

---

## 2. Mathematical Formulation of the Scoring Engine

Let document $D$ represent the normalized tender text corpus. The total score $S(D)$ is calculated as:

$$S(D) = \min\left(100, \; \sigma\left(\sum_{k=1}^m w_k \cdot f(t_k, D)\right) \times \prod_{j=1}^p \chi_j(D)\right)$$

### Where:
1. **Hard Constraint Operators ($\chi_j$):**
   $$\chi_j(D) = \begin{cases} 0, & \text{if disqualifying term is present (e.g. exclusive micro-enterprise clause)} \\ 1, & \text{otherwise} \end{cases}$$
2. **Term Frequency Weighting ($w_k$):**
   - Core Product Technology Terms: $w_{\text{core}} = 4.0$
   - Homologated OEM Partner Marks: $w_{\text{oem}} = 3.0$
   - Specialized Deployment Services: $w_{\text{serv}} = 2.0$
3. **Sigmoidal Normalization ($\sigma$):**
   $$\sigma(z) = \frac{100}{1 + e^{-\lambda(z - z_0)}}$$

---

## 3. Pipeline Architecture

1. **Ingestion & Normalization:** Strips HTML/formatting, unifies character encoding to UTF-8, and removes stop words.
2. **Deterministic Tokenization:** Extracts bid numbers, opening dates, estimated values, and contracting agencies.
3. **Scoring Execution:** Applies the heuristic weighting vector to the normalized token stream.
4. **LLM Context Synthesis:** Generates a structured JSON/Markdown executive brief for downstream language model processing.

---

## 4. Setup and Execution

```bash
# 1. Clone repository
git clone https://github.com/HenriMafra/procurement-radar-scoring.git
cd procurement-radar-scoring

# 2. Install dependencies
npm install

# 3. Execute evaluation pipeline
node index.js --input ./data/sample_notice.json
```

---

## 5. References

- Manning, C. D., Raghavan, P., & Schütze, H. (2008). *Introduction to Information Retrieval*. Cambridge University Press.
- Federative Republic of Brazil. (2021). *Federal Law n. 14.133 (Public Bidding and Administrative Contracts Framework)*.

---

## 6. License

Licensed under the MIT License. Copyright (c) Henri Mafra.
