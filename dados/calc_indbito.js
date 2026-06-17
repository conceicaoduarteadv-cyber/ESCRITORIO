// MEMORIAL DE CALCULOS — REVISIONAL MIDWAY 13864630
// SIMONE MARQUES MARTINS SEPEDA | CPF 324.863.672-68

const pv   = 5163.35;
const pv0  = 5000.00;
const n    = 21;
const tc   = 0.0999;
const tb   = 0.0510;
const ts   = 0.0700;
const pmtContrato = 636.64; // PMT efetivamente cobrado pelo banco

const f2 = v =>
  'R$ ' + v.toFixed(2)
    .replace('.', ',')
    .replace(/\B(?=(\d{3})+(?!\d))/g, '.');

const pagos = [
  626.62, 669.60, 662.79, 630.61, 636.64,
  660.58, 634.62, 636.64, 636.64, 636.64,
  636.64, 634.62, 636.64, 632.61, 628.61,
  628.61, 626.62, 630.61, 626.62, 584.35, 578.98
];

function pmt(pv, i, n) {
  return pv * i / (1 - Math.pow(1 + i, -n));
}

function simular(pv, i, n) {
  const p = pmt(pv, i, n);
  let s = pv, tj = 0;
  for (let k = 0; k < n; k++) {
    const j = s * i;
    tj += j;
    s -= (p - j);
  }
  return { p, total: +(p * n).toFixed(2), tj: +tj.toFixed(2) };
}

// Calcula TIR por bisseção (robusto, sem conflito de nomes)
function calcTIR(pvLoan, pmtVal, nPer) {
  // pv_calc(i) = pmtVal * (1 - (1+i)^-nPer) / i
  // TIR é o i tal que pv_calc(i) = pvLoan
  // pv_calc decresce conforme i aumenta
  const pvCalc = i => pmtVal * (1 - Math.pow(1 + i, -nPer)) / i;
  let lo = 0.001, hi = 2.0;
  for (let k = 0; k < 200; k++) {
    const mid = (lo + hi) / 2;
    if (pvCalc(mid) > pvLoan) lo = mid;
    else hi = mid;
    if (hi - lo < 1e-9) break;
  }
  return (lo + hi) / 2;
}

function amortTable(pv, ic, ib, n) {
  const pc = pmt(pv, ic, n);
  const pb = pmt(pv, ib, n);
  let sc = pv, sb = pv;
  const rows = [];
  for (let i = 1; i <= n; i++) {
    const jc = sc * ic, ac = pc - jc;
    const jb = sb * ib, ab = pb - jb;
    rows.push({ i, sc, pc, jc, sb, pb, jb, exc: pc - pb });
    sc -= ac;
    sb -= ab;
  }
  return rows;
}

const rC  = simular(pv,  tc, n);
const rB  = simular(pv,  tb, n);
const rS  = simular(pv,  ts, n);
const rI  = simular(pv0, tb, n);
const real = +pagos.reduce((a, b) => a + b, 0).toFixed(2);

const sep  = '='.repeat(72);
const dash = '-'.repeat(108);
const P = (v, w) => String(v).padStart(w);
const L = (v, w) => String(v).padEnd(w);

console.log(sep);
console.log('MEMORIAL DE CALCULOS - REVISIONAL MIDWAY/RIACHUELO no 13864630');
console.log('SIMONE MARQUES MARTINS SEPEDA | CPF 324.863.672-68');
console.log(sep);

// TIR real
const tirReal = calcTIR(pv, pmtContrato, n);
const tirAnual = Math.pow(1 + tirReal, 12) - 1;

console.log('\n[1] DADOS DO CONTRATO');
console.log('  Principal liberado (especie)   : ' + f2(pv0));
console.log('  IOF financiado                 :        R$ 163,35');
console.log('  Principal financiado           : ' + f2(pv));
console.log('  Taxa contratada                : 9,99% a.m. / 213,50% a.a.');
console.log('  CET declarado                  : 10,52% a.m. / 231,95% a.a.');
console.log('  PMT cobrado pelo banco         : ' + f2(pmtContrato) + '  (fluxo de vencimento)');
console.log('  PMT PRICE a 9,99% (calculado)  : ' + f2(rC.p) + '  << DIVERGENCIA de ' + f2(pmtContrato - rC.p) + '/parcela >>');
console.log('  Taxa efetiva real (TIR)        : ' + (tirReal*100).toFixed(4).replace('.',',') + '% a.m. / ' + (tirAnual*100).toFixed(2).replace('.',',') + '% a.a.');
console.log('  Total a 9,99% PRICE calculado  : ' + f2(rC.total));
console.log('  Total contratual (636,64 x 21) : ' + f2(pmtContrato * n));
console.log('  Excesso por carencia/TIR real  : ' + f2(pmtContrato * n - rC.total));
console.log('  Parcelas                       : 21  |  Carencia: 48 dias');

console.log('\n[2] REFERENCIA BACEN — SERIE 25479');
console.log('  Modalidade : Credito pessoal nao consignado (outros)');
console.log('  Jul/2024   : 5,10% a.m.');
console.log('  Equiv. ano : ' + ((Math.pow(1 + tb, 12) - 1) * 100).toFixed(2).replace('.', ',') + '% a.a.');
console.log('  Fonte      : api.bcb.gov.br/dados/serie/bcdata.sgs.25479');

const excRel = ((tc / tb) - 1) * 100;
console.log('\n[3] ABUSIVIDADE — STJ TEMA 27 / REsp 1.061.530/RS');
console.log('  Taxa contratada  : 9,99% a.m.  |  213,50% a.a.');
console.log('  Taxa media BACEN : 5,10% a.m.  |  ' + ((Math.pow(1 + tb, 12) - 1) * 100).toFixed(2).replace('.', ',') + '% a.a.');
console.log('  Excesso absoluto : ' + ((tc - tb) * 100).toFixed(2).replace('.', ',') + ' p.p. ao mes');
console.log('  Excesso relativo : ' + excRel.toFixed(1).replace('.', ',') + '% acima da media de mercado');
console.log('  -> TAXA CONTRATADA E ' + excRel.toFixed(0) + '% SUPERIOR A MEDIA DE MERCADO.');
console.log('  -> ABUSIVIDADE CONFIGURADA PARA FINS DE REVISAO JUDICIAL.');

console.log('\n[4] TABELA DE AMORTIZACAO COMPARATIVA (PRICE)');
console.log('  PV = R$ 5.163,35  |  9,99% a.m. (contratada)  x  5,10% a.m. (BACEN)');
console.log();
console.log('  ' + [
  P('Parc', 4), P('Saldo(9,99%)', 13), P('PMT(9,99%)', 11),
  P('Juros(9,99%)', 12), P('Saldo(5,10%)', 13), P('PMT(5,10%)', 11),
  P('Juros(5,10%)', 12), P('Excesso', 10)
].join('  '));
console.log('  ' + dash);

const rows = amortTable(pv, tc, tb, n);
let totExc = 0;
rows.forEach(r => {
  totExc += r.exc;
  console.log('  ' + [
    P(r.i, 4),
    P(f2(r.sc), 13), P(f2(r.pc), 11), P(f2(r.jc), 12),
    P(f2(r.sb), 13), P(f2(r.pb), 11), P(f2(r.jb), 12),
    P(f2(r.exc), 10)
  ].join('  '));
});
console.log('  ' + dash);
console.log('  ' + [
  P('TOTAL', 4),
  P('', 13), P(f2(rC.total), 11), P(f2(rC.tj), 12),
  P('', 13), P(f2(rB.total), 11), P(f2(rB.tj), 12),
  P(f2(totExc), 10)
].join('  '));

console.log('\n[5] PAGAMENTOS REAIS EFETUADOS (conforme extrato Midway)');
pagos.forEach((v, i) => {
  console.log('  Parcela ' + String(i + 1).padStart(2, '0') + ':  ' + f2(v));
});
console.log('  ' + '-'.repeat(32));
console.log('  TOTAL EFETIVAMENTE PAGO:  ' + f2(real));

const indA = +(real - rB.total).toFixed(2);
const indB = +(real - rS.total).toFixed(2);
const indC = +(real - rI.total).toFixed(2);

console.log('\n[6] QUADRO DO INDBITO — CC art. 876 + CDC art. 42');
console.log('  Total efetivamente pago (extrato)    : ' + f2(real));
console.log('');
console.log('  CENARIO A — BACEN 5,10% a.m. (pedido principal)');
console.log('  PMT revisado                         : ' + f2(rB.p));
console.log('  Total devido a 5,10%                 : ' + f2(rB.total));
console.log('  Juros em excesso                     : ' + f2(rC.tj - rB.tj));
console.log('  INDBITO (A)                          : ' + f2(indA));
console.log('  Restituicao em dobro — CDC 42 p.un.  : ' + f2(2 * indA));
console.log('  (Dobro: somente se provada ma-fe — STJ Tema 929)');
console.log('');
console.log('  CENARIO B — 7,00% a.m. com spread (pedido subsidiario)');
console.log('  PMT revisado                         : ' + f2(rS.p));
console.log('  Total devido a 7,00%                 : ' + f2(rS.total));
console.log('  INDBITO (B)                          : ' + f2(indB));
console.log('');
console.log('  CENARIO C — BACEN 5,10% sem IOF no principal (maximo)');
console.log('  Base R$ 5.000,00  | PMT revisado     : ' + f2(rI.p));
console.log('  Total devido                         : ' + f2(rI.total));
console.log('  INDBITO (C)                          : ' + f2(indC));
console.log('');
console.log('  RESUMO CONSOLIDADO');
console.log('  Cenario A — pedido principal          : ' + f2(indA));
console.log('  Cenario B — subsidiario conservador   : ' + f2(indB));
console.log('  Cenario C — maximo (incl. IOF)        : ' + f2(indC));
console.log('  Em dobro Cenario A (se ma-fe)         : ' + f2(2 * indA));

console.log('\n[7] PRESCRICAO — CC art. 206, par. 5, I (5 anos por pagamento)');
console.log('  Primeira parcela : 11/09/2024  ->  prescreve 11/09/2029');
console.log('  Ultima parcela   : 16/04/2026  ->  prescreve 16/04/2031');
console.log('  Data do calculo  : 17/06/2026');
console.log('  -> TODOS OS PAGAMENTOS DENTRO DO PRAZO. RESTITUICAO INTEGRAL. OK');

const iof = 163.35;
const jurosIof = +(iof * rC.tj / pv).toFixed(2);
console.log('\n[8] IMPACTO DO IOF FINANCIADO — CDC art. 51, IV');
console.log('  IOF embutido no principal            : ' + f2(iof));
console.log('  Juros pagos sobre o IOF (9,99%)      : ' + f2(jurosIof));
console.log('  -> Cliente pagou 9,99%/mes sobre R$ 163,35 que nunca recebeu.');
console.log('  -> Pratica abusiva: custo do tributo transferido ao consumidor.');

console.log('\n' + sep);
console.log('FIM DO MEMORIAL DE CALCULOS');
console.log(sep);
