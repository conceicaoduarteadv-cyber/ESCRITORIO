import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, AlignmentType, BorderStyle,
  WidthType, ShadingType
} from 'docx';
import { writeFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';

const OUT = path.join(
  'C:\\Users\\CONCEIÇÃO DUARTE\\teste\\peticoes',
  'memorial-calculos-simone-sepeda-midway-2026-06-17.docx'
);

// helpers
const t  = (txt, opts = {}) => new TextRun({ text: txt, font: 'Calibri', size: 22, ...opts });
const tb = (txt) => t(txt, { bold: true });
const ti = (txt) => t(txt, { italics: true });

const p  = (...runs) => new Paragraph({ children: runs, spacing: { after: 100 } });
const pb = (txt) => p(tb(txt));
const pi = (txt) => p(ti(txt));
const nl = () => p(t(''));

const h1 = (txt) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [t(txt, { bold: true, size: 28 })],
  spacing: { after: 200 }
});
const h2 = (txt) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [t(txt, { bold: true, size: 24 })],
  spacing: { before: 200, after: 100 }
});
const h3 = (txt) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [t(txt, { bold: true, size: 22 })],
  spacing: { before: 140, after: 80 }
});

const GRAY = 'D9D9D9';
const borderOpts = {
  top:    { style: BorderStyle.SINGLE, size: 4, color: '999999' },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
  left:   { style: BorderStyle.SINGLE, size: 4, color: '999999' },
  right:  { style: BorderStyle.SINGLE, size: 4, color: '999999' },
};

function cell(text, isHeader = false, width = null) {
  const opts = {
    children: [new Paragraph({
      children: [t(text, { bold: isHeader, size: 20 })],
      spacing: { before: 40, after: 40 }
    })],
    borders: borderOpts,
  };
  if (isHeader) opts.shading = { type: ShadingType.CLEAR, fill: GRAY };
  if (width) opts.width = { size: width, type: WidthType.DXA };
  return new TableCell(opts);
}

function table2(h1txt, h2txt, keys, vals) {
  const rows = [
    new TableRow({ children: [cell(h1txt, true), cell(h2txt, true)], tableHeader: true })
  ];
  for (let i = 0; i < keys.length; i++) {
    rows.push(new TableRow({ children: [cell(keys[i]), cell(vals[i])] }));
  }
  return new Table({ rows, width: { size: 9000, type: WidthType.DXA } });
}

function table3(c1, c2, c3, col1, col2, col3) {
  const rows = [
    new TableRow({ children: [cell(c1, true), cell(c2, true), cell(c3, true)], tableHeader: true })
  ];
  for (let i = 0; i < col1.length; i++) {
    rows.push(new TableRow({ children: [cell(col1[i]), cell(col2[i]), cell(col3[i])] }));
  }
  return new Table({ rows, width: { size: 9000, type: WidthType.DXA } });
}

function tableAmort() {
  const headers = ['Parc', 'Saldo(9,99%)', 'PMT(9,99%)', 'Juros(9,99%)', 'Saldo(5,10%)', 'PMT(5,10%)', 'Juros(5,10%)', 'Excesso'];
  const parcs = ['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','TOTAL'];
  const sd99  = ['R$ 5.163,35','R$ 5.082,58','R$ 4.993,74','R$ 4.896,02','R$ 4.788,54','R$ 4.670,33','R$ 4.540,30','R$ 4.397,29','R$ 4.239,99','R$ 4.066,97','R$ 3.876,68','R$ 3.667,36','R$ 3.437,14','R$ 3.183,92','R$ 2.905,41','R$ 2.599,07','R$ 2.262,12','R$ 1.891,52','R$ 1.483,89','R$ 1.035,54','R$ 542,40','--'];
  const pmt99 = Array(21).fill('R$ 596,59').concat(['R$ 12.528,40']);
  const j99   = ['R$ 515,82','R$ 507,75','R$ 498,87','R$ 489,11','R$ 478,38','R$ 466,57','R$ 453,58','R$ 439,29','R$ 423,57','R$ 406,29','R$ 387,28','R$ 366,37','R$ 343,37','R$ 318,07','R$ 290,25','R$ 259,65','R$ 225,99','R$ 188,96','R$ 148,24','R$ 103,45','R$ 54,19','R$ 7.365,05'];
  const sd51  = ['R$ 5.163,35','R$ 5.020,41','R$ 4.870,18','R$ 4.712,28','R$ 4.546,33','R$ 4.371,92','R$ 4.188,62','R$ 3.995,96','R$ 3.793,49','R$ 3.580,68','R$ 3.357,02','R$ 3.121,96','R$ 2.874,90','R$ 2.615,25','R$ 2.342,35','R$ 2.055,54','R$ 1.754,10','R$ 1.437,28','R$ 1.104,31','R$ 754,36','R$ 386,56','--'];
  const pmt51 = Array(21).fill('R$ 406,27').concat(['R$ 8.531,74']);
  const j51   = ['R$ 263,33','R$ 256,04','R$ 248,38','R$ 240,33','R$ 231,86','R$ 222,97','R$ 213,62','R$ 203,79','R$ 193,47','R$ 182,61','R$ 171,21','R$ 159,22','R$ 146,62','R$ 133,38','R$ 119,46','R$ 104,83','R$ 89,46','R$ 73,30','R$ 56,32','R$ 38,47','R$ 19,71','R$ 3.368,39'];
  const exc   = Array(21).fill('R$ 190,32').concat(['R$ 3.996,66']);

  const rows = [
    new TableRow({ children: headers.map(h => cell(h, true)), tableHeader: true })
  ];
  for (let i = 0; i < parcs.length; i++) {
    rows.push(new TableRow({ children: [
      cell(parcs[i]), cell(sd99[i]), cell(pmt99[i]), cell(j99[i]),
      cell(sd51[i]), cell(pmt51[i]), cell(j51[i]), cell(exc[i])
    ]}));
  }
  return new Table({ rows, width: { size: 9000, type: WidthType.DXA } });
}

// ── Documento ────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: {
      page: {
        margin: { top: 1134, bottom: 1134, left: 1701, right: 1134 }
      }
    },
    children: [
      h1('MEMORIAL DE CALCULOS'),
      h2('Revisional -- Midway Financial S.A. | Credito Pessoal no 13864630'),
      p(t('Cliente: SIMONE MARQUES MARTINS SEPEDA | CPF: 324.863.672-68')),
      p(t('Elaborado em: 17/06/2026  |  Advogada: Conceicao Maria Duarte Portilho')),
      p(t('Fonte BACEN: Serie 25479 -- api.bcb.gov.br/dados/serie/bcdata.sgs.25479')),
      nl(),

      // 1
      h2('1. Dados do Contrato'),
      table2('Item', 'Valor',
        ['Principal liberado (especie)','IOF financiado','Principal total financiado',
         'Taxa nominal declarada','CET declarado','PMT cobrado pelo banco',
         'PMT PRICE calculado a 9,99%','DIVERGENCIA por parcela',
         'Taxa efetiva real (TIR calculada)','Total PRICE a 9,99% (calculado)',
         'Total contratual (636,64 x 21)','Excesso pela carencia/TIR real',
         'Carencia','Parcelas'],
        ['R$ 5.000,00','R$ 163,35','R$ 5.163,35',
         '9,99% a.m. / 213,50% a.a.','10,52% a.m. / 231,95% a.a.','R$ 636,64',
         'R$ 596,59','R$ 40,05 a mais/mes',
         '10,9352% a.m. / 247,40% a.a.','R$ 12.528,40',
         'R$ 13.369,44','R$ 841,04',
         '48 dias','21']
      ),
      pi('NOTA: O PMT de R$ 636,64 NAO corresponde a formula PRICE a 9,99% a.m. (que produziria R$ 596,59). A TIR real e 10,9352% a.m. -- 0,94 p.p. acima do declarado. Fundamento: CDC art. 6, III; Res. BACEN 3.517/2007.'),
      nl(),

      // 2
      h2('2. Referencia BACEN -- Serie 25479'),
      table2('Dado', 'Valor',
        ['Serie','Mes de contratacao (jul/2024)','Equivalente anual','Fonte'],
        ['25479 -- Credito pessoal nao consignado (outros)','5,10% a.m.','81,65% a.a.','api.bcb.gov.br/dados/serie/bcdata.sgs.25479']
      ),
      nl(),

      // 3
      h2('3. Abusividade -- STJ Tema 27 / REsp 1.061.530/RS'),
      table3('Indice','Mensal','Anual',
        ['Taxa declarada no contrato','Taxa efetiva real (TIR)','Taxa media BACEN -- jul/2024','Excesso absoluto (TIR vs BACEN)','Excesso relativo (TIR vs BACEN)'],
        ['9,99%','10,9352%','5,10%','5,84 p.p./mes','114,4% acima'],
        ['213,50%','247,40%','81,65%','--','--']
      ),
      pb('-> Taxa efetiva real e mais que o DOBRO da taxa media de mercado.'),
      pb('-> Abusividade configurada -- STJ REsp 1.061.530/RS (recurso repetitivo).'),
      nl(),

      // 4
      h2('4. Tabela de Amortizacao Comparativa (PRICE)'),
      p(t('Principal: R$ 5.163,35 | Contratada 9,99% a.m. x BACEN 5,10% a.m.')),
      tableAmort(),
      nl(),

      // 5
      h2('5. Pagamentos Reais Efetuados (extrato Midway)'),
      table2('Parcela','Valor pago',
        ['01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','TOTAL PAGO'],
        ['R$ 626,62','R$ 669,60','R$ 662,79','R$ 630,61','R$ 636,64','R$ 660,58','R$ 634,62','R$ 636,64','R$ 636,64','R$ 636,64','R$ 636,64','R$ 634,62','R$ 636,64','R$ 632,61','R$ 628,61','R$ 628,61','R$ 626,62','R$ 630,61','R$ 626,62','R$ 584,35','R$ 578,98','R$ 13.276,29']
      ),
      pi('Contrato quitado em 16/04/2026. Saldo devedor: R$ 0,00.'),
      nl(),

      // 6
      h2('6. Quadro do Indbito -- CC art. 876 + CDC art. 42'),
      pb('Total efetivamente pago (extrato Midway): R$ 13.276,29'),
      nl(),
      h3('Cenario A -- Taxa BACEN 5,10% a.m. (pedido principal)'),
      table2('Item','Valor',
        ['PMT revisado','Total devido a 5,10%','Juros pagos em excesso','INDBITO (A)','Restituicao em dobro -- CDC 42 par. unico','(dobro somente se provada ma-fe -- STJ Tema 929)'],
        ['R$ 406,27','R$ 8.531,74','R$ 3.996,66','R$ 4.744,55','R$ 9.489,10','']
      ),
      nl(),
      h3('Cenario B -- 7,00% a.m. com spread (pedido subsidiario)'),
      table2('Item','Valor',
        ['PMT revisado','Total devido a 7,00%','INDBITO (B)'],
        ['R$ 476,52','R$ 10.006,93','R$ 3.269,36']
      ),
      nl(),
      h3('Cenario C -- BACEN 5,10% sem IOF financiado (maximo)'),
      table2('Item','Valor',
        ['Base: R$ 5.000,00 | PMT revisado','Total devido','INDBITO (C)'],
        ['R$ 393,42','R$ 8.261,83','R$ 5.014,46']
      ),
      nl(),
      h3('Resumo Consolidado'),
      table3('Cenario','Indbito','Base',
        ['A -- Pedido principal','B -- Subsidiario conservador','C -- Maximo (com IOF)','Em dobro -- Cenario A (se ma-fe)'],
        ['R$ 4.744,55','R$ 3.269,36','R$ 5.014,46','R$ 9.489,10'],
        ['BACEN 5,10% s/ R$ 5.163,35','7,00% s/ R$ 5.163,35','BACEN 5,10% s/ R$ 5.000,00','Somente se ma-fe provada']
      ),
      nl(),

      // 7
      h2('7. Prescricao -- CC art. 206, par. 5, I (5 anos por pagamento)'),
      table3('Marco','Data','Prescricao',
        ['1a parcela paga','Ultima parcela paga','Data do calculo'],
        ['11/09/2024','16/04/2026','17/06/2026'],
        ['11/09/2029','16/04/2031','--']
      ),
      pb('-> Todos os pagamentos estao dentro do prazo prescricional. Restituicao integral cabivel.'),
      nl(),

      // 8
      h2('8. IOF Financiado -- CDC art. 51, IV'),
      table2('Item','Valor',
        ['IOF embutido no principal','Juros pagos sobre o IOF (9,99% a.m.)'],
        ['R$ 163,35','R$ 233,00 (estimado proporcional)']
      ),
      pi('Cliente pagou juros de 9,99% a.m. sobre R$ 163,35 de tributo que nunca recebeu em especie. Pratica abusiva -- CDC art. 51, IV.'),
      nl(),

      // 9
      h2('9. Achado Adicional -- Divergencia entre Taxa Declarada e TIR Real'),
      pi('O banco declara 9,99% a.m., mas o PMT cobrado (R$ 636,64) implica TIR de 10,9352% a.m. -- 0,94 p.p. acima do declarado. A diferenca de R$ 40,05/parcela x 21 = R$ 841,04 cobrados a mais do que o PRICE a 9,99% produziria. Causa provavel: capitalizacao dos juros da carencia de 48 dias sem informacao clara ao consumidor. Fundamento: CDC art. 6, III; Res. BACEN 3.517/2007; CDC art. 51, IV.'),
      nl(),

      p(t('Calculos elaborados com base no extrato Midway emitido em 17/06/2026.')),
      p(t('Serie BACEN 25479 consultada em 17/06/2026 via API oficial do Banco Central.')),
      p(t('Sujeito a conferencia por perito judicial.')),
    ]
  }]
});

const buffer = await Packer.toBuffer(doc);
writeFileSync(OUT, buffer);
console.log('CONCLUIDO: ' + OUT);
