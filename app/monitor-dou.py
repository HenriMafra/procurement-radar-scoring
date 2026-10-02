import urllib.request, urllib.parse, ssl, re, json, sys

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
}

def consultar_dou_termo(termo):
    if not termo or len(termo.strip()) < 3:
        return []
    
    q = urllib.parse.quote(termo.strip())
    url = f"https://www.in.gov.br/consulta/-/buscar/dou?q={q}&s=todos&exactDate=all"
    
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=12, context=ctx) as r:
            html = r.read().decode('utf-8', errors='ignore')
            
        match = re.search(r'id="_br_com_seatecnologia_in_buscadou_BuscaDouPortlet_params"[^>]*>(.*?)</script>', html, re.DOTALL)
        if match:
            data = json.loads(match.group(1).strip())
            return data.get('jsonArray', [])
            
        match = re.search(r'\{\s*"jsonArray"\s*:\s*(\[.*?\])\s*\}', html, re.DOTALL)
        if match:
            return json.loads(match.group(1))
            
    except Exception:
        return []
        
    return []

def consultar_materias_processo(radical_processo):
    if not radical_processo:
        return []
        
    raw = consultar_dou_termo(f'"{radical_processo}"')
    materias = []
    vistas = set()
    
    for r in raw:
        url_title = r.get('urlTitle')
        if not url_title or url_title in vistas:
            continue
            
        content = r.get('content', '')
        if radical_processo.lower() in content.lower():
            vistas.add(url_title)
            materias.append({
                "data": r.get('pubDate', ''),
                "secao": r.get('pubName', 'DO3'),
                "titulo": r.get('title', ''),
                "edicao": r.get('editionNumber', ''),
                "pagina": r.get('numberPage', ''),
                "url": f"https://www.in.gov.br/en/web/dou/-/{url_title}",
                "resumo": re.sub(r'<[^<]+?>', '', content).strip()
            })
            
    return materias

if __name__ == "__main__":
    termo = sys.argv[1] if len(sys.argv) > 1 else "19.00.6300.0000129"
    res = consultar_materias_processo(termo)
    print(json.dumps(res, ensure_ascii=False, indent=2))
