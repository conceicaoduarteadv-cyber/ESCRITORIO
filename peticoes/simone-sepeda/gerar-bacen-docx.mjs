import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, convertInchesToTwip, BorderStyle, ShadingType
} from 'docx';
import { writeFileSync } from 'fs';

const FONTE = 'Times New Roman';
const pt = (n) => n * 20;
const txt = (t, o = {}) => new TextRun({ text: t, font: FONTE, size: 24, ...o });
const bold = (t, o = {}) => txt(t, { bold: true, ...o });
const par = (children, opts = {}) => new Paragraph({ children, ...opts });

// Cores
const COR_HEADER = 'D0D0D0';
const COR_DESTAQUE = 'FFE0E0';
const COR_BACEN = 'E0F0E0';

const cellOpts = (text, isBold = false, shade = null) => new TableCell({
  children: [par(
    [isBold ? bold(text) : txt(text)],
    { alignment: AlignmentType.CENTER, spacing: { before: pt(3), after: pt(3) } }
  )],
  shading: shade ? { type: ShadingType.SOLID, color: shade } : undefined,
  margins: { top: pt(2), bottom: pt(2), left: pt(4), right: pt(4) },
});

const tabelaBACEN = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        cellOpts('Mês de referência', true, COR_HEADER),
        cellOpts('Taxa média ao mês', true, COR_HEADER),
        cellOpts('Taxa média ao ano', true, COR_HEADER),
      ],
    }),
    new TableRow({ children: [cellOpts('Fevereiro/2026', false, COR_BACEN), cellOpts('6,47% a.m.', true, COR_BACEN), cellOpts('112,15% a.a.', true, COR_BACEN)] }),
    new TableRow({ children: [cellOpts('Março/2026', false, COR_BACEN), cellOpts('6,67% a.m.', true, COR_BACEN), cellOpts('117,05% a.a.', true, COR_BACEN)] }),
    new TableRow({ children: [cellOpts('Abril/2026', false, COR_BACEN), cellOpts('6,99% a.m.', true, COR_BACEN), cellOpts('125,06% a.a.', true, COR_BACEN)] }),
  ],
});

const tabelaCCB = (ccb, data, bacenMes, bacenAno, contratoMes, contratoAno, excMes, excAno, fatorMes, fatorAno) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ tableHeader: true, children: [cellOpts('', true, COR_HEADER), cellOpts('Taxa ao mês', true, COR_HEADER), cellOpts('Taxa ao ano', true, COR_HEADER)] }),
      new TableRow({ children: [cellOpts(`Média BACEN (${data.slice(3)})`, false, COR_BACEN), cellOpts(`${bacenMes}% a.m.`, true, COR_BACEN), cellOpts(`${bacenAno}% a.a.`, true, COR_BACEN)] }),
      new TableRow({ children: [cellOpts('Taxa contratada pela Ré', false, COR_DESTAQUE), cellOpts(`${contratoMes}% a.m.`, true, COR_DESTAQUE), cellOpts(`${contratoAno}% a.a.`, true, COR_DESTAQUE)] }),
      new TableRow({ children: [cellOpts('Excesso', false), cellOpts(`+${excMes} p.p.`, true), cellOpts(`+${excAno} p.p.`, true)] }),
      new TableRow({ children: [cellOpts('Fator de desvio', false, 'FFD700'), cellOpts(`${fatorMes}× a taxa de mercado`, true, 'FFD700'), cellOpts(`${fatorAno}× a taxa de mercado`, true, 'FFD700')] }),
    ],
  });

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONTE, size: 24 } } },
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertInchesToTwip(1.18),
          bottom: convertInchesToTwip(0.98),
          left: convertInchesToTwip(1.57),
          right: convertInchesToTwip(0.98),
        },
      },
    },
    children: [
      // Cabeçalho
      par([bold('CONSULTA BACEN — TAXA MÉDIA DE MERCADO', { size: 28 })], { alignment: AlignmentType.CENTER, spacing: { after: pt(4) } }),
      par([bold('Crédito Pessoal Não Consignado — Recursos Livres — Pessoas Físicas', { size: 24 })], { alignment: AlignmentType.CENTER, spacing: { after: pt(12) } }),

      par([txt('Fonte: '), bold('Banco Central do Brasil — Sistema Gerenciador de Séries Temporais (SGS)')], { spacing: { after: pt(3) } }),
      par([txt('Séries: 25464 (taxa mensal) e 20742 (taxa anual)')], { spacing: { after: pt(3) } }),
      par([txt('Consulta disponível em: '), txt('https://dadosabertos.bcb.gov.br/dataset/25464-taxa-media-mensal-de-juros-das-operacoes-de-credito-com-recursos-livres---pessoas-fisicas---c')], { spacing: { after: pt(3) } }),
      par([txt('Data da consulta: 15/06/2026')], { spacing: { after: pt(16) } }),

      // Tabela BACEN
      par([bold('I — TAXAS MÉDIAS DE MERCADO APURADAS PELO BACEN')], { spacing: { before: pt(8), after: pt(8) } }),
      tabelaBACEN,
      par([txt('')], { spacing: { after: pt(16) } }),

      // Comparativo
      par([bold('II — COMPARATIVO COM AS TAXAS COBRADAS PELA RÉ (MIDWAY S.A.)')], { spacing: { before: pt(8), after: pt(8) } }),

      par([bold('CCB nº 75805602 — emitida em 18/02/2026')], { spacing: { before: pt(8), after: pt(6) } }),
      tabelaCCB('75805602', '18/02/2026', '6,47', '112,15', '13,75', '369,26', '7,28', '257,11', '2,13', '3,29'),
      par([txt('')], { spacing: { after: pt(14) } }),

      par([bold('CCB nº 75825597 — emitida em 09/03/2026')], { spacing: { before: pt(4), after: pt(6) } }),
      tabelaCCB('75825597', '09/03/2026', '6,67', '117,05', '13,75', '369,26', '7,08', '252,21', '2,06', '3,15'),
      par([txt('')], { spacing: { after: pt(14) } }),

      par([bold('CCB nº 75879070 — emitida em 27/04/2026')], { spacing: { before: pt(4), after: pt(6) } }),
      tabelaCCB('75879070', '27/04/2026', '6,99', '125,06', '14,99', '434,47', '8,00', '309,41', '2,14', '3,47'),
      par([txt('')], { spacing: { after: pt(16) } }),

      // Conclusão
      par([bold('III — CONCLUSÃO')], { spacing: { before: pt(8), after: pt(8) } }),
      par([txt('As três CCBs celebradas entre a Autora e a Ré apresentam taxas de juros remuneratórios entre '), bold('2,06 e 2,14 vezes'), txt(' a taxa média de mercado apurada pelo Banco Central do Brasil para crédito pessoal não consignado nas respectivas datas de contratação.')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(8) } }),
      par([txt('Nos termos do '), bold('Tema 27/STJ (REsp 1.061.530/RS)'), txt(', o simples fato de a taxa contratada superar significativamente a taxa média de mercado é suficiente para configurar abusividade e autorizar a revisão judicial. Nos três contratos, as taxas cobradas pela Ré superam o dobro da média de mercado (+100%), configurando abusividade manifesta.')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(16) } }),

      // Citação jurisprudencial
      par([txt('"A cobrança de juros em percentual superior à média do mercado é considerada abusiva quando há significativa discrepância, cabendo ao credor a prova de que as peculiaridades do negócio justificam essa diferença." — REsp 1.061.530/RS, Rel. Min. Nancy Andrighi, Tema 27/STJ')],
        { alignment: AlignmentType.JUSTIFIED, spacing: { before: pt(4), after: pt(4) }, indent: { left: convertInchesToTwip(0.5), right: convertInchesToTwip(0.5) } }),

      par([txt('')], { spacing: { after: pt(32) } }),
      par([txt('Macapá/AP, 15 de junho de 2026.')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(60) } }),
      par([txt('___________________________________________')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
      par([bold('CONCEIÇÃO MARIA DUARTE PORTILHO')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
      par([txt('OAB/AP nº 3576')], { alignment: AlignmentType.CENTER }),
    ],
  }],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync('peticoes/simone-sepeda/consulta-bacen-taxas.docx', buffer);
console.log('Arquivo gerado: peticoes/simone-sepeda/consulta-bacen-taxas.docx');
