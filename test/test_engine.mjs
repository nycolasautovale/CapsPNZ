// Testa o motor (parse + métricas) do app.js com a planilha real, via Node.
import XLSX from '../vendor/xlsx.full.min.js';
import fs from 'node:fs';
import vm from 'node:vm';

// Carrega app.js num contexto com stubs de DOM/echarts para extrair parseWorkbook/computeMetrics.
const appSrc = fs.readFileSync(new URL('../src/app.js', import.meta.url), 'utf8');

// Em vez de rodar o IIFE inteiro (que depende de DOM), extraímos as funções puras
// recriando-as: carregamos o arquivo e expomos via hack — mais simples: reimplementar o teste
// chamando o mesmo algoritmo. Para garantir paridade, injetamos as funções no escopo.

// Estratégia: envolver o IIFE expondo as funções no globalThis quando rodando em teste.
globalThis.XLSX = XLSX;
const _store = {};
globalThis.localStorage = { getItem:k=>_store[k]??null, setItem:(k,v)=>{_store[k]=String(v);} };
globalThis.window = { addEventListener(){}, matchMedia:()=>({matches:false}) };
globalThis.document = {
  addEventListener(){}, querySelector(){return null;}, createElement(){return {};},
  documentElement: { setAttribute(){}, getAttribute(){return 'light';} }
};
globalThis.getComputedStyle = () => ({ getPropertyValue:()=> '' });
globalThis.echarts = { init(){return {setOption(){},resize(){}};}, graphic:{LinearGradient:function(){}} };
globalThis.ICON = new Proxy({}, { get:()=> '' });

// injeta um gancho no fim do app.js para exportar as internals
const patched = appSrc.replace(
  'document.addEventListener(\'DOMContentLoaded\', initUpload);',
  'globalThis.__TEST__ = { parseWorkbook, computeMetrics, shortDesc };'
);
vm.runInThisContext(patched);

const { parseWorkbook, computeMetrics } = globalThis.__TEST__;

const buf = fs.readFileSync(new URL('../data/base_final.xlsx', import.meta.url));
const wb = XLSX.read(buf, { type:'buffer' });
const data = parseWorkbook(wb);
computeMetrics(data);

console.log('=== TESTE DO MOTOR (Node + XLSX real) ===\n');
console.log('Unidades:', data.units.map(u=>u.name).join(' | '));
console.log('Avisos:', data.warnings.length, data.warnings);
console.log('Erros:', data.errors.length, data.errors);
console.log('\nMétricas por unidade:');
let pass = true;
const expect = {
  'CAPSIJ':   { real: 22947, taxaG: 40.6 },
  'CAPS II':  { real: 13098, taxaG: 52.1 },
  'CAPS AD III': { real: 20625, taxaG: 11.1 },
};
data.units.forEach(u => {
  const line = `  ${u.name.padEnd(38)} real=${u.totalReal} taxaGeral=${u.taxaGeral==null?'—':u.taxaGeral.toFixed(1)}%`;
  console.log(line);
  const ex = expect[u.sheet];
  if (ex) {
    if (u.totalReal !== ex.real) { console.log(`     [FALHA] realizado esperado ${ex.real}`); pass=false; }
    if (Math.abs(u.taxaGeral - ex.taxaG) > 0.15) { console.log(`     [FALHA] taxa geral esperada ~${ex.taxaG}`); pass=false; }
  }
});
console.log('\nTotal rede:', data.network.totalReal, '(esperado 59456)');
if (data.network.totalReal !== 59456) pass=false;
console.log('Meses rede:', data.network.months.length, '| meses TAB:', data.network.tabMonths.length);

// ---- valida os campos derivados usados pelas novas seções ----
console.log('\n=== CAMPOS DERIVADOS (novas seções) ===');
console.log('Mix top5:', data.network.mix.slice(0,5).map(m=>m.desc+'='+m.real));
console.log('tabGtCount:', data.network.tabGtCount);
console.log('alertas:', data.alerts.map(a=>`[${a.nivel}] ${a.unidade}: ${a.texto}`));
data.units.filter(u=>u.hasTab).forEach(u=>{
  console.log(`  ${u.name}: naoTabulados=${u.naoTabulados.length}, maioresPerdas=${u.maioresPerdas.length}`);
});
// checagens básicas de integridade
if (!data.network.mix.length) { console.log('[FALHA] mix vazio'); pass=false; }
if (data.network.tabGtCount !== 32) { console.log('[AVISO] tabGtCount != 32 (esperado 32), veio', data.network.tabGtCount); }
data.units.filter(u=>u.hasTab).forEach(u=>{ if(!Array.isArray(u.naoTabulados)){console.log('[FALHA] naoTabulados ausente em',u.name);pass=false;} });

// ---- novos campos (melhorias) ----
console.log('\n=== NOVOS CAMPOS (melhorias) ===');
data.units.filter(u=>u.hasTab).forEach(u=>{
  const serie = u.taxaMensal.map(x=>`${x.mes.slice(0,3)}=${x.taxa==null?'—':x.taxa.toFixed(0)}%`).join(' ');
  console.log(`  ${u.name}: taxaGeral=${u.taxaGeral?.toFixed(1)}% | lacunaRegistro=${u.lacunaRegistro} | mensal: ${serie}`);
});
console.log('prodMensal varPct:', data.network.prodMensal.map(x=>`${x.mes.slice(0,3)}:${x.varPct==null?'—':x.varPct.toFixed(0)+'%'}`).join(' '));
console.log('maiorQueda:', data.network.maiorQueda ? `${data.network.maiorQueda.mes} ${data.network.maiorQueda.varPct.toFixed(1)}%` : 'nenhuma');
console.log('maiorAlta:', data.network.maiorAlta ? `${data.network.maiorAlta.mes} ${data.network.maiorAlta.varPct.toFixed(1)}%` : 'nenhuma');
console.log('tabGtAgg (top3):', data.network.tabGtAgg.slice(0,3).map(x=>`${x.desc} x${x.ocorrencias}`));

if (!Array.isArray(data.network.prodMensal) || !data.network.prodMensal.length) { console.log('[FALHA] prodMensal vazio'); pass=false; }
data.units.filter(u=>u.hasTab).forEach(u=>{
  if (u.lacunaRegistro!=null && u.realInTab!=null && u.lacunaRegistro > u.realInTab) { console.log('[FALHA] lacunaRegistro > realInTab em',u.name); pass=false; }
});

console.log('\n' + (pass ? '>>> TODOS OS TESTES PASSARAM' : '>>> HÁ FALHAS'));
process.exit(pass?0:1);
