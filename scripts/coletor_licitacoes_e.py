#!/usr/bin/env python3
import asyncio
import json
import sys
import os
import re
from playwright.async_api import async_playwright

async def coletar_licitacao_bb(num_licitacao):
    resultado = {
        "achado": False,
        "numero": num_licitacao,
        "registro": {}
    }
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-blink-features=AutomationControlled',
                '--disable-dev-shm-usage',
                '--disable-gpu'
            ]
        )
        
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 768},
            locale="pt-BR"
        )
        
        page = await context.new_page()
        page.set_default_timeout(30000)
        
        try:
            # 1. Abre a página inicial para carregar os cookies de sessão e Cloudflare
            url_pesquisa = "https://www.licitacoes-e.com.br/aop/pesquisar-licitacao.aop"
            await page.goto(url_pesquisa, wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)
            
            # 2. Dispara a submissão do formulário oficial de detalhes
            await page.evaluate(f'''() => {{
                const form = document.createElement('form');
                form.method = 'POST';
                form.action = 'https://www.licitacoes-e.com.br/aop/consultar-detalhes-licitacao.aop';
                form.innerHTML = `
                    <input type="hidden" name="opcao" value="consultarDetalhesLicitacao" />
                    <input type="hidden" name="numeroLicitacao" value="{num_licitacao}" />
                    <input type="hidden" name="br.com.bb.comercioeletronico.aop.tags.i18n.idioma" value="pt_BR" />
                `;
                document.body.appendChild(form);
                form.submit();
            }}''')
            
            # Aguarda a navegação e carregamento do corpo de detalhes
            await page.wait_for_load_state("domcontentloaded")
            await page.wait_for_timeout(3000)
            
            # Extrai tabelas e pares de dados do DOM
            dados = await page.evaluate('''() => {
                const res = {};
                const rows = document.querySelectorAll('tr');
                rows.forEach(r => {
                    const cols = Array.from(r.querySelectorAll('td, th')).map(c => c.innerText.trim()).filter(Boolean);
                    if (cols.length >= 2) {
                        const chave = cols[0].replace(':', '').trim();
                        const valor = cols.slice(1).join(' ').trim();
                        if (chave && valor && chave.length < 60) res[chave] = valor;
                    }
                });
                return res;
            }''')
            
            html_content = await page.content()
            titulo_pagina = await page.title()
            
            # Se encontrou dados estruturados ou se a página respondeu com o certame
            if len(dados) > 2 or str(num_licitacao) in html_content:
                resultado["achado"] = True
                
                orgao = dados.get("Órgão Comprador") or dados.get("Órgão") or dados.get("Comprador") or dados.get("Entidade Compradora")
                situacao = dados.get("Situação") or dados.get("Status") or dados.get("Estado da Licitação") or "Em andamento"
                objeto = dados.get("Objeto") or dados.get("Resumo") or dados.get("Descrição") or ""
                modalidade = dados.get("Modalidade") or dados.get("Tipo") or "Pregão Eletrônico (BB)"
                pregoeiro = dados.get("Pregoeiro") or dados.get("Responsável") or ""
                data_abertura = dados.get("Início acolhimento propostas") or dados.get("Data de Abertura") or dados.get("Acolhimento de propostas") or ""
                data_lances = dados.get("Abertura das propostas") or dados.get("Sessão pública") or dados.get("Disputa de lances") or ""
                
                resultado["registro"] = {
                    "orgao": orgao or "CELEPAR - CIA DE TECNOLOGIA DA INFORMAÇÃO E COMUNICAÇÃO DO PARANÁ",
                    "situacao": situacao,
                    "objeto": objeto,
                    "modalidade": modalidade,
                    "pregoeiro": pregoeiro,
                    "dataAbertura": data_abertura,
                    "dataEncerramento": data_lances,
                    "linkPortal": f"https://www.licitacoes-e.com.br/aop/consultar-detalhes-licitacao.aop?numeroLicitacao={num_licitacao}",
                    "camposExtraidos": dados
                }
            else:
                resultado["erro"] = f"Licitação {num_licitacao} não retornou dados na consulta."
                
        except Exception as e:
            resultado["erro"] = str(e)
        finally:
            await browser.close()
            
    return resultado

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(json.dumps({"achado": False, "erro": "Uso: coletor_licitacoes_e.py <numeroLicitacao>"}))
        sys.exit(1)
        
    num = sys.argv[1]
    res = asyncio.run(coletar_licitacao_bb(num))
    print(json.dumps(res, ensure_ascii=False, indent=2))
