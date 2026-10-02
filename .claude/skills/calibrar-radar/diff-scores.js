#!/usr/bin/env node
// Compara o dados.json atual com o snapshot da ultima execucao e lista TUDO que
// mudou de nota, com o objeto — para ler um por um antes de dizer que funcionou.
//   node diff-scores.js            -> compara com o snapshot e regrava o snapshot
//   node diff-scores.js --marcar   -> so grava o snapshot (antes de mexer no config)
// Rodar da raiz do projeto.
const fs = require('fs');
const path = require('path');

const RAIZ = path.resolve(__dirname, '..', '..', '..');
const ATUAL = path.join(RAIZ, 'data', 'dados.json');
const SNAP = path.join(RAIZ, 'data', '.snapshot-scores.json');

const ler = p => JSON.parse(fs.readFileSync(p, 'utf8').replace(/^﻿/, '')).licitacoes;
if (!fs.existsSync(ATUAL)) { console.error('data/dados.json nao existe.'); process.exit(1); }
const B = ler(ATUAL);

const gravar = () => {
  const mapa = {};
  B.forEach(x => { mapa[x.id] = { s: x.score, kw: x.palavrasChaveDet || [] }; });
  fs.writeFileSync(SNAP, JSON.stringify(mapa), 'utf8');
};

if (process.argv.includes('--marcar')) { gravar(); console.log('snapshot gravado: ' + B.length + ' licitacoes'); process.exit(0); }
if (!fs.existsSync(SNAP)) { gravar(); console.log('sem snapshot anterior; gravei o de agora (' + B.length + '). Mexa no config e rode de novo.'); process.exit(0); }

const A = JSON.parse(fs.readFileSync(SNAP, 'utf8'));
const corta = s => String(s || '').replace(/\s+/g, ' ').replace(/^(\*[^*]*\*\s*)+/, '');
const mud = [];
for (const x of B) {
  const a = A[x.id];
  if (!a) { mud.push({ x, de: null, para: x.score, novos: x.palavrasChaveDet || [] }); continue; }
  if (a.s !== x.score) {
    const novos = (x.palavrasChaveDet || []).filter(k => !a.kw.includes(k));
    mud.push({ x, de: a.s, para: x.score, novos });
  }
}
mud.sort((p, q) => (q.para - (q.de || 0)) - (p.para - (p.de || 0)));

console.log('mudaram de nota: ' + mud.filter(m => m.de !== null).length
  + '   |   novos na base: ' + mud.filter(m => m.de === null).length
  + '   |   total: ' + B.length + '\n');

if (!mud.length) { console.log('nada mudou.'); gravar(); process.exit(0); }

console.log('LEIA O OBJETO DE CADA UM. E aqui que o falso positivo aparece.\n');
for (const m of mud) {
  const seta = m.de === null ? '  (novo) ->' : String(m.de).padStart(4) + ' ->';
  console.log(seta + ' ' + String(m.para).padStart(4) + '   ' + String(m.x.orgao || '').split(' - ')[0].slice(0, 34));
  if (m.novos.length) console.log('        casou agora: ' + JSON.stringify(m.novos));
  const fora = m.x.contextoForaDet || [];
  if (fora.length) console.log('        fora escopo : ' + JSON.stringify(fora.slice(0, 4)));
  console.log('        objeto      : ' + corta(m.x.objeto).slice(0, 150));
  console.log('');
}
gravar();
console.log('snapshot atualizado.');
