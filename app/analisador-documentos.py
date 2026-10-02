import urllib.request, zipfile, io, json, re, sys, os
from pypdf import PdfReader

def analisar_url_documento(url, titulo=""):
    headers = {'User-Agent': 'Mozilla/5.0'}
    req = urllib.request.Request(url, headers=headers)
    
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
    except Exception as e:
        return {"ok": False, "erro": str(e)}

    textos = []
    
    # 1. Se for ZIP
    if data[:4] == b'PK\x03\x04' or url.lower().endswith('.zip') or 'zip' in titulo.lower():
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for filename in z.namelist():
                    if filename.lower().endswith('.pdf'):
                        with z.open(filename) as pdf_file:
                            reader = PdfReader(pdf_file)
                            txt = "".join([page.extract_text() or "" for page in reader.pages])
                            textos.append((filename, txt))
                    elif filename.lower().endswith(('.txt', '.html', '.htm')):
                        with z.open(filename) as txt_file:
                            txt = txt_file.read().decode('utf-8', errors='ignore')
                            textos.append((filename, txt))
        except Exception as e:
            pass

    # 2. Se for PDF direto
    elif data[:4] == b'%PDF' or url.lower().endswith('.pdf'):
        try:
            reader = PdfReader(io.BytesIO(data))
            txt = "".join([page.extract_text() or "" for page in reader.pages])
            textos.append((titulo or "documento.pdf", txt))
        except Exception as e:
            pass
            
    # 3. Se for texto puro / HTML
    elif b'<html' in data[:200].lower() or data[:100].isascii():
        try:
            txt = data.decode('utf-8', errors='ignore')
            clean = re.sub(r'<[^<]+?>', ' ', txt)
            textos.append((titulo or "documento.txt", clean))
        except Exception:
            pass

    # Analisar o conteúdo extraído
    decisoes_encontradas = []
    
    palavras_chave_decisao = [
        "decisao em recurso", "decisão em recurso", "recurso administrativo",
        "nego provimento", "dou provimento", "não procede", "nao procede",
        "decisão do pregoeiro", "decisao do pregoeiro", "revisão da autoridade",
        "revisao da autoridade", "homologo o pregão", "homologo o pregao",
        "termo de homologação", "termo de homologacao", "termo de julgamento",
        "ata de realização", "ata de realizacao", "julgamento de impugnação"
    ]

    for nome_arq, txt in textos:
        txt_lower = txt.lower()
        for pc in palavras_chave_decisao:
            matches = [m.start() for m in re.finditer(re.escape(pc), txt_lower)]
            for pos in matches:
                trecho = txt[max(0, pos-150):min(len(txt), pos+350)]
                clean_trecho = " ".join(trecho.split())
                decisoes_encontradas.append({
                    "termo": pc,
                    "arquivo": nome_arq,
                    "trecho": clean_trecho
                })

    return {
        "ok": True,
        "totalArquivos": len(textos),
        "arquivos": [t[0] for t in textos],
        "decisoes": decisoes_encontradas[:10]
    }

if __name__ == "__main__":
    if len(sys.argv) > 1:
        u = sys.argv[1]
        tit = sys.argv[2] if len(sys.argv) > 2 else ""
        res = analisar_url_documento(u, tit)
        print(json.dumps(res, ensure_ascii=False, indent=2))
