// =====================================================================
//  auditar-escopo.js
//  Mede a qualidade do filtro contra os boletins JA processados.
//  Nao altera nada: so lê data/dados.json + config/escopo-enterprisecore.json e diz
//  quais palavras estao trabalhando, quais estao mortas e quais estao
//  fazendo barulho. Serve para manter o dicionario com evidencia.
//
//  Rodar:  Auditar-Escopo.cmd     (ou: node scripts/auditar-escopo.js)
// =====================================================================
'use strict';
const fs = require('fs');
const path = require('path');

const RAIZ = path.resolve(__dirname, '..');
const lerJson = (p, f) => { try { let t = fs.readFileSync(p, 'utf8'); if (t.charCodeAt(0) === 0xFEFF) t = t.slice(1); return JSON.parse(t); } catch (_) { return f; } };

const cfg = lerJson(path.join(RAIZ, 'config', 'escopo-enterprisecore.json'), null);
const dados = lerJson(path.join(RAIZ, 'data', 'dados.json'), null);

if (!cfg) { console.error('Nao achei config/escopo-enterprisecore.json'); process.exit(1); }
if (!dados || !dados.licitacoes || !dados.licitacoes.length) {
  console.error('Nao achei data/dados.json com licitacoes. Processe os boletins primeiro.');
  process.exit(1);
}

const arr = (x) => Array.isArray(x) ? x : (x ? [x] : []);
const norm = (s) => ' ' + String(s || '').toLowerCase()
  .normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/[^a-z0-9]+/g, ' ').trim() + ' ';

const itens = dados.licitacoes;
// SO o objeto do edital. Nao incluir 'solucao' nem 'fabricante': esses campos
// carregam a SUGESTAO do proprio motor, e auditar a saida dele contra ela mesma
// da numero falso — o alias "zivasec" casava 11x no rotulo "Rede / Wireless /
// Infra (ZivaSec)" que o motor escreveu, sem existir em edital nenhum.
const textos = itens.map((x) => norm(x.objeto));

function tit(t) { console.log('\n' + '='.repeat(66) + '\n  ' + t + '\n' + '='.repeat(66)); }
function pct(a, b) { return b ? Math.round((a / b) * 100) + '%' : '0%'; }

console.log(`\nAUDITORIA DO ESCOPO ENTERPRISECORE`);
console.log(`Base: ${itens.length} licitacoes processadas em ${dados.geradoEm || '?'}`);

// ---------------------------------------------------------------- 1
tit('1. COMO A BASE FICOU DISTRIBUIDA');
const porFaixa = {};
itens.forEach((x) => { porFaixa[x.faixa || '?'] = (porFaixa[x.faixa || '?'] || 0) + 1; });
Object.entries(porFaixa)
  .sort((a, b) => (itens.find((i) => i.faixa === b[0]).score || 0) - (itens.find((i) => i.faixa === a[0]).score || 0))
  .forEach(([k, v]) => console.log(`  ${String(v).padStart(4)}  ${pct(v, itens.length).padStart(4)}  ${k}`));

const semSinal = itens.filter((x) => !arr(x.solucoesDet).length && !arr(x.parceiros).length).length;
console.log(`\n  Sem nenhum sinal de escopo: ${semSinal} (${pct(semSinal, itens.length)})`);
console.log(`  -> se este numero passar de 70%, o dicionario esta estreito demais para os seus boletins.`);

// ---------------------------------------------------------------- 2
tit('2. PALAVRAS-CHAVE QUE NUNCA DISPARARAM (peso morto)');
const mortas = [];
const vivas = [];
arr(cfg.categorias).forEach((c) => {
  arr(c.palavrasChave).forEach((k) => {
    const t = norm(k).trim();
    if (!t) return;
    const n = textos.filter((x) => x.includes(' ' + t + ' ')).length;
    (n === 0 ? mortas : vivas).push({ cat: c.solucao, k, n });
  });
});
console.log(`  ${mortas.length} de ${mortas.length + vivas.length} palavras nao apareceram em nenhum edital desta base.`);
console.log('  Isso NAO quer dizer que estao erradas: podem so nao ter aparecido ainda.');
console.log('  Vale revisar as que voce nem lembra de ter escrito.\n');
const porCat = {};
mortas.forEach((m) => { (porCat[m.cat] = porCat[m.cat] || []).push(m.k); });
Object.entries(porCat).sort((a, b) => b[1].length - a[1].length).slice(0, 8)
  .forEach(([c, ks]) => console.log(`  ${c} (${ks.length}):\n      ${ks.join(', ')}`));

// ---------------------------------------------------------------- 3
tit('3. PALAVRAS QUE MAIS DISPARAM (candidatas a ruido)');
vivas.sort((a, b) => b.n - a.n).slice(0, 15).forEach((v) => {
  const alerta = v.n > itens.length * 0.25 ? '  <-- dispara demais, conferir se nao e generico' : '';
  console.log(`  ${String(v.n).padStart(3)}x  "${v.k}"  [${v.cat}]${alerta}`);
});

// ---------------------------------------------------------------- 4
tit('4. MARCAS: quais aparecem de fato nos seus boletins');
const marcas = [];
arr(cfg.fabricantes).forEach((f) => {
  const termos = (f.buscarPeloNome === false ? [] : [f.nome]).concat(arr(f.aliases));
  let n = 0;
  const vistos = new Set();
  termos.forEach((t) => {
    const k = norm(t).trim();
    if (!k) return;
    textos.forEach((x, i) => { if (x.includes(' ' + k + ' ')) vistos.add(i); });
  });
  n = vistos.size;
  if (n) marcas.push({ nome: f.nome, n });
});
if (marcas.length) {
  marcas.sort((a, b) => b.n - a.n).forEach((m) => console.log(`  ${String(m.n).padStart(3)}x  ${m.nome}`));
} else {
  console.log('  Nenhuma marca citada nominalmente nesta base.');
  console.log('  Normal: o orgao raramente nomeia fabricante no resumo do objeto.');
}

// ---------------------------------------------------------------- 5
tit('5. CENTRALIDADE NO LOTE');
const comLote = itens.filter((x) => x.proporcaoLote !== null && x.proporcaoLote !== undefined);
console.log(`  ${comLote.length} de ${itens.length} editais tem lote numerado analisavel.\n`);
const faixasLote = { 'dedicado (>=70%)': 0, 'parcial (50-69%)': 0, 'fraco (25-49%)': 0, 'marginal (<25%)': 0 };
comLote.forEach((x) => {
  const p = Number(x.proporcaoLote);
  if (p >= 0.7) faixasLote['dedicado (>=70%)']++;
  else if (p >= 0.5) faixasLote['parcial (50-69%)']++;
  else if (p >= 0.25) faixasLote['fraco (25-49%)']++;
  else faixasLote['marginal (<25%)']++;
});
Object.entries(faixasLote).forEach(([k, v]) => console.log(`  ${String(v).padStart(4)}  ${k}`));

// ---------------------------------------------------------------- 6
tit('6. NA DUVIDA - estes merecem olho humano');
// Só a faixa cinzenta e os que tiveram a nota REDUZIDA por peso.
// Quem ganhou bonus de lote dedicado nao esta em duvida: esse o motor acertou.
const duvida = itens.filter((x) => {
  const s = Number(x.score) || 0;
  const f = Number(x.densidadeFator);
  return (s >= 18 && s <= 55) || (f && f < 1 && s > 0);
}).sort((a, b) => (b.score || 0) - (a.score || 0));
console.log(`  ${duvida.length} edital(is) na faixa cinzenta (score 18-55, ou com peso ajustado).`);
console.log('  Sao os que mais se beneficiam de passar pela IA.\n');
duvida.slice(0, 12).forEach((x) => {
  const lote = (x.proporcaoLote !== null && x.proporcaoLote !== undefined)
    ? `${x.itensComEscopo}/${x.itensLote} do lote` : 'sem lote';
  console.log(`  ${String(x.score).padStart(3)} | ${lote.padEnd(16)} | ${String(x.orgao || '').slice(0, 42)}`);
  console.log(`      ${String(x.objeto || '').replace(/\s+/g, ' ').slice(0, 90)}`);
});

// ---------------------------------------------------------------- 7
tit('7. CONTEXTO DE OUTRA AREA - o que esta sendo barrado');
const fora = {};
itens.forEach((x) => arr(x.contextoForaDet).forEach((t) => { fora[t] = (fora[t] || 0) + 1; }));
const listaFora = Object.entries(fora).sort((a, b) => b[1] - a[1]);
if (listaFora.length) {
  listaFora.slice(0, 15).forEach(([k, v]) => console.log(`  ${String(v).padStart(3)}x  ${k}`));
  const barrados = itens.filter((x) => arr(x.contextoForaDet).length >= 2).length;
  console.log(`\n  ${barrados} edital(is) levaram a penalidade de outra area.`);
  console.log('  Confira a lista acima: se alguma palavra ai tambem aparece em edital seu,');
  console.log('  tire ela de "contextosForaEscopo" - foi assim que "lente" zerou um CFTV.');
} else {
  console.log('  Nenhum termo de outra area disparou nesta base.');
}

// ---------------------------------------------------------------- 8
// CHECAGENS QUE SE MANTEM SOZINHAS.
// Antes eu descobria palavra ruim no olho, caso a caso. Estas tres medem a
// PROPRIEDADE que torna a palavra ruim, entao pegam tambem as que ninguem viu.

// 8a. COLISAO COM O PORTUGUES.
// Esta e a checagem que teria achado o "nas" sozinha. Tentei antes por
// frequencia (termo raro = tecnico) e por co-ocorrencia (aparece junto de
// outra palavra de TI?): as duas falharam. A primeira nao acusou o "nas",
// que dispara em so 3% da base; a segunda acusou o "breach and attack", que
// e justamente o melhor termo do dicionario. A diferenca entre eles nao e
// estatistica, e lexical: "nas" e "poe" sao palavras do portugues e "breach
// and attack" nao e. Por isso a lista abaixo - que e FECHADA: as classes
// gramaticais do portugues (preposicao, artigo, pronome, contracao) nao
// crescem, entao esta checagem nao envelhece.
var PT_COMUM = ('a as o os um uma uns umas de do da dos das dum duma em no na nos nas num numa ' +
  'por pelo pela pelos pelas para pra com sem sob sobre ate apos ante entre contra desde perante tras ' +
  'e ou mas porem contudo todavia entao logo pois porque que se como quando onde quanto qual quais quem ' +
  'cujo cuja este esta estes estas esse essa esses essas aquele aquela aquilo isto isso ' +
  'meu minha seu sua seus suas nosso nossa dele dela deles delas lhe lhes me te vos ' +
  'ser sao foi foram era eram sera serao seja sejam ter tem tinha tera terao tenha ' +
  'haver ha havia houve estar estao estava poder pode podem podia podera ' +
  'fazer faz fazem feito dar dao deve devem dever ir vai vao ver ve vem vir poe poem pos ' +
  'mais menos muito muita pouco todo toda todos todas outro outra mesmo mesma tal tais cada ' +
  'algum alguma nenhum nenhuma ja ainda sempre nunca tambem so apenas bem mal assim aqui ali la ' +
  'hoje ano anos dia dias mes meses hora horas vez vezes caso casos fim forma modo parte ' +
  'nova novo nome numero valor total geral local prazo data lista meio campo linha nivel ' +
  'grupo ordem regra uso via ponto tipo item itens area base').split(' ');
var setPT = {}; PT_COMUM.forEach(function (w) { setPT[w] = 1; });

tit('8. TERMOS QUE COLIDEM COM PALAVRA DO PORTUGUES');
var colide = [];
arr(cfg.categorias).forEach(function (c) {
  arr(c.palavrasChave).forEach(function (k) {
    var nk = norm(k).trim();
    if (!nk || nk.indexOf(' ') >= 0) return;      // so termo de palavra unica colide
    if (!setPT[nk]) return;
    var n = textos.filter(function (t) { return t.indexOf(' ' + nk + ' ') >= 0; }).length;
    colide.push({ k: String(k).trim(), cat: c.solucao, n: n });
  });
});
if (colide.length) {
  colide.sort(function (a, b) { return b.n - a.n; });
  console.log('  Estes termos do dicionario SAO palavras do portugues. O casamento por');
  console.log('  palavra inteira nao protege contra isso - protege sigla, nao idioma.\n');
  colide.forEach(function (x) {
    console.log('  [!] ' + x.k.padEnd(14) + ' dispara em ' + String(x.n).padStart(3) + ' edital(is)   [' + x.cat + ']');
  });
  console.log('\n  Decida um a um: ou tire do dicionario e cubra pelo nome completo');
  console.log('  (foi o que fizemos com "nas" -> "network attached storage"), ou');
  console.log('  marque como ambiguo em especificidade.ambiguas (foi o caso de "poe").');
} else {
  console.log('  Nenhum termo do dicionario colide com palavra do portugues.');
}

// 8b. QUEM SO APARECE SOZINHA. Se a palavra nunca vem acompanhada de outra
// palavra nossa, ela nao esta descrevendo nosso produto - esta pegando carona.
// E o perfil exato de "conectividade" na ficha de um teclado sem fio.
tit('8b. PALAVRAS QUE APARECEM SEMPRE ISOLADAS (candidatas a "ambiguas")');
var esp = cfg.especificidade || {};
var jaAmbigua = {}; arr(esp.ambiguas).forEach(function (t) { jaAmbigua[norm(t).trim()] = 1; });
var jaNucleo = {}; arr(esp.nucleo).forEach(function (t) { jaNucleo[norm(t).trim()] = 1; });
var solo = [];
arr(cfg.categorias).forEach(function (c) {
  arr(c.palavrasChave).forEach(function (k) {
    var nk = norm(k).trim();
    if (!nk || jaAmbigua[nk] || jaNucleo[nk]) return;
    var onde = itens.filter(function (x) { return norm(x.objeto).indexOf(' ' + nk + ' ') >= 0; });
    if (!onde.length) return;
    var sozinha = onde.filter(function (x) { return arr(x.palavrasChaveDet).length === 1; }).length;
    var comFora = onde.filter(function (x) { return arr(x.contextoForaDet).length >= 1; }).length;
    if (sozinha === onde.length && comFora > 0) {
      solo.push({ k: String(k).trim(), cat: c.solucao, n: onde.length, fora: comFora });
    }
  });
});
if (solo.length) {
  solo.sort(function (a, b) { return b.n - a.n; });
  solo.slice(0, 12).forEach(function (x) {
    console.log('  ' + String(x.n).padStart(3) + 'x  ' + x.k.padEnd(28) + ' sempre isolada, ' + x.fora + 'x junto de termo de outra area  [' + x.cat + ']');
  });
  console.log('\n  Sugestao: mover para especificidade.ambiguas no config.');
} else {
  console.log('  Nenhuma candidata nova. As ambiguas ja listadas dao conta desta base.');
}

// 8c. A LISTA ESTA CERTA? Confere contra a evidencia, para nao virar dogma.
// Nucleo tem de puxar nota alta; ambigua, nota baixa. Onde inverter, olhar.
tit('8c. AS LISTAS DE ESPECIFICIDADE BATEM COM A BASE?');
if (cfg.especificidade) {
  var linhas = [];
  ['nucleo', 'ambiguas'].forEach(function (tipo) {
    arr(esp[tipo]).forEach(function (t) {
      var nk = norm(t).trim();
      var onde = itens.filter(function (x) { return norm(x.objeto).indexOf(' ' + nk + ' ') >= 0; });
      if (!onde.length) return;
      var media = Math.round(onde.reduce(function (a, x) { return a + (x.score || 0); }, 0) / onde.length);
      linhas.push({ t: t, n: onde.length, media: media, tipo: tipo === 'nucleo' ? 'nucleo' : 'ambigua' });
    });
  });
  if (linhas.length) {
    linhas.sort(function (a, b) { return b.media - a.media; });
    linhas.forEach(function (l) {
      var alerta = (l.tipo === 'nucleo' && l.media < 40) ? '   <-- nucleo puxando nota baixa, conferir'
        : (l.tipo === 'ambigua' && l.media >= 70) ? '   <-- ambigua em edital bom, talvez promover'
          : '';
      console.log('  ' + l.tipo.padEnd(8) + String(l.n).padStart(2) + 'x  media ' + String(l.media).padStart(3) + '  ' + l.t + alerta);
    });
  } else {
    console.log('  Nenhum termo das duas listas disparou nesta base.');
  }
} else {
  console.log('  Bloco "especificidade" ausente do config.');
}

console.log('\n' + '='.repeat(66));
console.log('  Como usar: ajuste config/escopo-enterprisecore.json e rode o processamento');
console.log('  de novo. Depois rode esta auditoria outra vez e compare.');
console.log('='.repeat(66) + '\n');
