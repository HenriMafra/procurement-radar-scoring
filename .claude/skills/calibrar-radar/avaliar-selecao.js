#!/usr/bin/env node
// Mede o motor do Radar contra o julgamento humano.
//   node avaliar-selecao.js 19281042 19281156 ...   -> recall dos escolhidos + falsos positivos
//   node avaliar-selecao.js                          -> so o topo, para inspecao
// Rodar da raiz do projeto.
const fs = require('fs');
const path = require('path');

const RAIZ = path.resolve(__dirname, '..', '..', '..');
const alvo = path.join(RAIZ, 'data', 'dados.json');
if (!fs.existsSync(alvo)) { console.error('data/dados.json nao existe. Rode o processamento primeiro.'); process.exit(1); }

const B = JSON.parse(fs.readFileSync(alvo, 'utf8').replace(/^﻿/, '')).licitacoes;
const escolhidos = process.argv.slice(2).map(s => s.replace(/\D/g, '')).filter(Boolean);
const sel = new Set(escolhidos);

const ord = B.slice().sort((a, b) => b.score - a.score);
ord.forEach((x, i) => { x.__pos = i + 1; });
const corta = (s, n) => String(s || '').replace(/\s+/g, ' ').replace(/^(\*[^*]*\*\s*)+/, '').slice(0, n);
const orgao = x => String(x.orgao || '').split(' - ')[0];

console.log('base: ' + B.length + ' licitacoes' + (escolhidos.length ? '   |   escolhidos: ' + escolhidos.length : ''));

if (escolhidos.length) {
  console.log('\n=== OS ESCOLHIDOS, na ordem que o motor deu ===');
  console.log('  pos  score  faixa              conlicitacao  objeto');
  const achados = [];
  for (const n of escolhidos) {
    const x = B.find(y => String(y.numeroConlicitacao) === n);
    if (!x) { console.log('   --     --  NAO ESTA NA BASE  ' + n); continue; }
    achados.push(x);
    console.log('  ' + String(x.__pos).padStart(3) + '  ' + String(x.score).padStart(5) + '  '
      + String(x.faixa).slice(0, 17).padEnd(18) + n + '  ' + corta(x.objeto, 56));
  }
  if (achados.length) {
    const pos = achados.map(x => x.__pos).sort((a, b) => a - b);
    const nt = achados.map(x => x.score).sort((a, b) => a - b);
    console.log('\n  posicoes : ' + pos.join(', '));
    console.log('  top 10   : ' + pos.filter(p => p <= 10).length + '/' + achados.length
      + '   top 20: ' + pos.filter(p => p <= 20).length + '/' + achados.length
      + '   top 30: ' + pos.filter(p => p <= 30).length + '/' + achados.length);
    console.log('  score    : min ' + nt[0] + '  mediana ' + nt[Math.floor(nt.length / 2)] + '  max ' + nt[nt.length - 1]);

    const ruins = achados.filter(x => x.__pos > 15).sort((a, b) => b.__pos - a.__pos);
    if (ruins.length) {
      console.log('\n=== ESCOLHIDOS MAL RANQUEADOS - o que o motor nao viu ===');
      for (const x of ruins) {
        console.log('\n  #' + x.__pos + '  score ' + x.score + '   ' + orgao(x).slice(0, 46));
        console.log('     kw    : ' + JSON.stringify(x.palavrasChaveDet || []));
        console.log('     fora  : ' + JSON.stringify(x.contextoForaDet || []));
        console.log('     lote  : ' + (x.itensComEscopo || 0) + '/' + (x.itensLote || 0)
          + '   fator ' + (x.densidadeFator !== undefined ? x.densidadeFator : '-'));
        console.log('     pts   : ' + (x.porques || []).map(p => p.sinal + ' ' + (p.pontos > 0 ? '+' : '') + p.pontos).join(' | '));
        console.log('     objeto: ' + corta(x.objeto, 260));
      }
    } else {
      console.log('\n  Nenhum escolhido ficou abaixo da 15a posicao.');
    }
  }
}

console.log('\n=== TOPO 20' + (escolhidos.length ? ' QUE VOCE NAO ESCOLHEU (candidatos a falso positivo)' : '') + ' ===');
ord.slice(0, 20).filter(x => !sel.has(String(x.numeroConlicitacao))).forEach(x => {
  console.log('  ' + String(x.__pos).padStart(3) + '  ' + String(x.score).padStart(5) + '  '
    + orgao(x).slice(0, 26).padEnd(27) + corta(x.objeto, 54));
});

console.log('\n=== SINAIS QUE MAIS DISPARAM NA BASE ===');
const cont = {};
B.forEach(x => (x.palavrasChaveDet || []).forEach(k => { cont[k] = (cont[k] || 0) + 1; }));
Object.entries(cont).sort((a, b) => b[1] - a[1]).slice(0, 15)
  .forEach(([k, v]) => console.log('  ' + String(v).padStart(3) + 'x  ' + k));
