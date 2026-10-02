// =====================================================================
//  ENTERPRISECORE Radar Sentinela — Background Service Worker (Manifest V3)
//  Segurança: Comunicação isolada, imune a Mixed-Content e protegida por Token
// =====================================================================

const RADAR_URL = 'http://localhost:8790/api/extensao/capturar';
const AUTH_TOKEN = 'enterprisecore_sentinela_9a7d3f5b8e1c2a4d6f0b8e9a7c5d3b1f';

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'SINCRONIZAR_RADAR') {
    const payload = request.payload || {};
    
    // Adiciona o token de autenticação seguro
    payload.token = AUTH_TOKEN;

    fetch(RADAR_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-EnterpriseCore-Radar-Token': AUTH_TOKEN
      },
      body: JSON.stringify(payload)
    })
      .then(async (res) => {
        const json = await res.json().catch(() => ({ ok: false, erro: 'Resposta inválida do servidor' }));
        sendResponse({ ok: res.ok, status: res.status, dados: json });
      })
      .catch((err) => {
        sendResponse({ ok: false, erro: err.message });
      });

    // Retorna true para indicar resposta assíncrona
    return true;
  }
});
