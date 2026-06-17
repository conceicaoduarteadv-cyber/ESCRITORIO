import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, AlignmentType, BorderStyle,
  WidthType, ShadingType
} from 'docx';
import { writeFileSync } from 'fs';
import path from 'path';

const OUT = path.join(
  'C:\\Users\\CONCEIÇÃO DUARTE\\teste\\peticoes',
  'revisional-simone-sepeda-midway-13864630-2026-06-17.docx'
);

// helpers
const t  = (txt, opts = {}) => new TextRun({ text: txt, font: 'Times New Roman', size: 24, ...opts });
const tb = (txt) => t(txt, { bold: true });
const ti = (txt) => t(txt, { italics: true });

const centro = (...runs) => new Paragraph({ children: runs, alignment: AlignmentType.CENTER, spacing: { after: 80 } });
const just   = (...runs) => new Paragraph({ children: runs, alignment: AlignmentType.JUSTIFIED, spacing: { after: 180 }, indent: { firstLine: 720 } });
const justSR = (...runs) => new Paragraph({ children: runs, alignment: AlignmentType.JUSTIFIED, spacing: { after: 180 } });
const nl = () => new Paragraph({ children: [t('')], spacing: { after: 80 } });

const h1 = (txt) => new Paragraph({
  children: [t(txt, { bold: true, size: 28, allCaps: true })],
  alignment: AlignmentType.CENTER,
  spacing: { before: 240, after: 240 }
});
const h2 = (txt) => new Paragraph({
  children: [t(txt, { bold: true, size: 24, underline: {} })],
  alignment: AlignmentType.LEFT,
  spacing: { before: 280, after: 120 }
});
const h3 = (txt) => new Paragraph({
  children: [t(txt, { bold: true, size: 24 })],
  alignment: AlignmentType.LEFT,
  spacing: { before: 180, after: 80 }
});

const GRAY = 'D9D9D9';
const borderOpts = {
  top:    { style: BorderStyle.SINGLE, size: 4, color: '555555' },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: '555555' },
  left:   { style: BorderStyle.SINGLE, size: 4, color: '555555' },
  right:  { style: BorderStyle.SINGLE, size: 4, color: '555555' },
};

function cell(text, isHeader = false) {
  const opts = {
    children: [new Paragraph({
      children: [t(text, { bold: isHeader, size: 20 })],
      spacing: { before: 40, after: 40 },
      alignment: AlignmentType.CENTER
    })],
    borders: borderOpts,
  };
  if (isHeader) opts.shading = { type: ShadingType.CLEAR, fill: GRAY };
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

const doc = new Document({
  sections: [{
    properties: {
      page: {
        margin: { top: 1701, bottom: 1134, left: 1701, right: 1134 }
      }
    },
    children: [
      // cabecalho
      centro(t('EXCELENTISSIMO(A) SENHOR(A) DOUTOR(A) JUIZ(A) DE DIREITO DA', { bold: true })),
      centro(tb('VARA CIVEL DA COMARCA DE MACAPA/AP')),
      nl(),
      nl(),
      nl(),

      justSR(tb('SIMONE MARQUES MARTINS SEPEDA'), t(', brasileira, estado civil [A PREENCHER], profissao [A PREENCHER], portadora do RG no 040825 e CPF no 324.863.672-68, residente e domiciliada na Rua Carlos Gomes, no 301, Bairro Jesus de Nazare, Macapa/AP, CEP 68.908-125, por meio de sua advogada que esta subscreve, vem, respeitosamente, a presenca de Vossa Excelencia, propor a presente')),
      nl(),

      h1('ACAO DE REVISAO DE CONTRATO BANCARIO CUMULADA COM REPETICAO DE INDBITO'),

      justSR(t('em face de '), tb('MIDWAY FINANCIAL S.A. - CREDITO, FINANCIAMENTO E INVESTIMENTO'), t(', instituicao financeira com sede na Rua Lemos Monteiro, 120, 15o Andar, Edificio Pinheiros One, Bairro Butanta, Sao Paulo/SP, CEP 05.501-050, inscrita no CNPJ/ME sob o no 09.464.032/0001-12, pelos fatos e fundamentos juridicos a seguir expostos.')),
      nl(),

      // I - FATOS
      h2('I - DOS FATOS'),

      just(t('Em 29 de julho de 2024, a Autora celebrou com a Re o '), tb('Contrato de Credito Pessoal no 13864630'), t(' (Produto 302 - Emprestimo Pessoal Pre-Fixado), por meio do qual tomou emprestado o valor liquido de '), tb('R$ 5.000,00 (cinco mil reais)'), t(', a ser restituido em '), tb('21 (vinte e uma) parcelas mensais de R$ 636,64 (seiscentos e trinta e seis reais e sessenta e quatro centavos)'), t(', com vencimento da primeira parcela em 15 de setembro de 2024 e da ultima em 15 de maio de 2026.')),

      just(t('Alem do valor recebido em especie, a Re incorporou ao principal financiado o montante de '), tb('R$ 163,35 (cento e sessenta e tres reais e trinta e cinco centavos)'), t(' a titulo de IOF - Imposto sobre Operacoes Financeiras -, fazendo com que o principal total financiado correspondesse a '), tb('R$ 5.163,35 (cinco mil, cento e sessenta e tres reais e trinta e cinco centavos)'), t(', sobre o qual passou a cobrar integralmente os juros contratuais de '), tb('9,99% ao mes'), t(', embora a Autora jamais tenha recebido em especie essa parcela tributaria.')),

      just(t('O contrato declarou a taxa nominal de '), tb('9,99% ao mes (213,50% ao ano)'), t(' e o Custo Efetivo Total - CET - de '), tb('10,52% ao mes (231,95% ao ano)'), t('. Entretanto, a aplicacao da formula matematica PRICE ao principal financiado de R$ 5.163,35, a taxa declarada de 9,99% ao mes e ao prazo de 21 parcelas, produz uma parcela mensal de '), tb('R$ 596,59 (quinhentos e noventa e seis reais e cinquenta e nove centavos)'), t(' - e nao os R$ 636,64 cobrados pela Re. A Taxa Interna de Retorno (TIR) efetiva apurada a partir da parcela efetivamente cobrada e de '), tb('10,9352% ao mes (247,40% ao ano)'), t(', valor '), tb('0,94 ponto percentual superior'), t(' a taxa declarada no contrato, o que evidencia que a Re cobrou sistematicamente mais do que o proprio instrumento contratual autorizava.')),

      just(t('A diferenca de '), tb('R$ 40,05 por parcela'), t(', multiplicada pelas 21 parcelas do contrato, representa '), tb('R$ 841,04 (oitocentos e quarenta e um reais e quatro centavos)'), t(' cobrados a mais do que a taxa nominal declarada justificaria, sem qualquer fundamento contratual ou legal. A causa mais provavel e o tratamento conferido a carencia de 48 dias: ao capitalizar os juros do periodo de graca sobre o saldo devedor antes de iniciar a amortizacao, a Re elevou a taxa efetiva sem informar o consumidor de forma clara e precisa, em afronta ao artigo 6o, inciso III, do Codigo de Defesa do Consumidor e a Resolucao BACEN no 3.517/2007.')),

      just(t('Ao longo dos 21 meses de vigencia do contrato - de setembro de 2024 a abril de 2026 -, a Autora efetuou todos os pagamentos, conforme demonstram os registros de boleto bancario constantes do extrato emitido pela propria Re, totalizando '), tb('R$ 13.276,29 (treze mil, duzentos e setenta e seis reais e vinte e nove centavos)'), t(' efetivamente pagos. O contrato foi integralmente quitado em '), tb('16 de abril de 2026'), t(', com saldo devedor zerado (R$ 0,00), conforme comprovam os registros do extrato ora anexado.')),

      just(t('Em sintese: a Autora recebeu '), tb('R$ 5.000,00'), t(' e devolveu a Re '), tb('R$ 13.276,29'), t(' - custo total do credito de '), tb('R$ 8.276,29'), t(', equivalente a '), tb('165,5% do valor principal recebido'), t(', em apenas 21 meses.')),

      just(t('Tal disparidade e produto direto da taxa de juros abusiva praticada pela Re. Conforme dados publicos do Banco Central do Brasil, a taxa media de mercado para operacoes de '), tb('credito pessoal nao consignado'), t(' (Serie SGS no 25.479) no mes de julho de 2024 - mes da contratacao - era de '), tb('5,10% ao mes (81,65% ao ano)'), t('. A taxa cobrada pela Re corresponde a '), tb('96% acima da media de mercado'), t(' para a mesma modalidade, o que configura abusividade na acepcao fixada pelo Superior Tribunal de Justica no julgamento do REsp 1.061.530/RS.')),

      just(t('Recalculando-se as 21 parcelas contratuais a taxa justa de 5,10% ao mes sobre o mesmo principal de R$ 5.163,35, a parcela mensal seria de '), tb('R$ 406,27'), t(' e o total do contrato seria de '), tb('R$ 8.531,74 (oito mil, quinhentos e trinta e um reais e setenta e quatro centavos)'), t('. A diferenca entre o que foi efetivamente pago (R$ 13.276,29) e o que deveria ter sido cobrado a taxa de mercado (R$ 8.531,74) corresponde ao '), tb('indbito de R$ 4.744,55 (quatro mil, setecentos e quarenta e quatro reais e cinquenta e cinco centavos)'), t(' que a Re recebeu indevidamente e que ora se requer a restituicao, conforme demonstrado no Memorial de Calculos que acompanha esta peca (Doc. 5).')),

      nl(),

      // II - DIREITO
      h2('II - DO DIREITO'),

      h3('2.1 Da aplicabilidade do Codigo de Defesa do Consumidor'),

      just(t('A relacao juridica estabelecida entre a Autora e a Re e de consumo, porquanto a Re e fornecedora de servicos financeiros (CDC, art. 3o, paragrafo 2o) e a Autora e destinataria final do credito. A incidencia do Codigo de Defesa do Consumidor as instituicoes financeiras e materia pacificada pelo Superior Tribunal de Justica:')),

      justSR(ti('"O Codigo de Defesa do Consumidor e aplicavel as instituicoes financeiras."'), t(' (STJ, Sumula 297)')),

      just(t('O STF, no julgamento da ADI 2.591, confirmou a plena aplicacao do CDC aos contratos bancarios. Nessa condicao, a Autora faz jus a revisao das clausulas abusivas (CDC, art. 6o, V), a nulidade das disposicoes que estabelecam obrigacoes inequas (CDC, art. 51, IV) e a restituicao dos valores pagos indevidamente (CDC, art. 42, paragrafo unico).')),

      nl(),
      h3('2.2 Da abusividade dos juros remuneratorios - STJ Tema 27'),

      just(t('A taxa de juros remuneratorios contratada corresponde a 9,99% ao mes (213,50% ao ano) - declarada -, embora a taxa efetiva real apurada pela TIR seja de 10,9352% ao mes (247,40% ao ano). Ambas as taxas superam em quase o dobro a taxa media de mercado aferida pelo Banco Central do Brasil para a mesma modalidade crediticia no mes da contratacao:')),

      nl(),
      table3('Indice', 'Mensal', 'Anual',
        ['Taxa declarada (contrato)', 'Taxa efetiva real (TIR)', 'Taxa media BACEN - jul/2024', 'Excesso (TIR vs BACEN)', 'Excesso relativo (TIR vs BACEN)'],
        ['9,99%', '10,9352%', '5,10%', '5,84 p.p./mes', '114,4% acima'],
        ['213,50%', '247,40%', '81,65%', '--', '--']
      ),
      nl(),
      justSR(ti('Fonte: Serie SGS no 25.479 - Banco Central do Brasil. Consulta realizada em 17/06/2026.')),
      nl(),

      just(t('O Superior Tribunal de Justica, no julgamento do REsp 1.061.530/RS (Tema Repetitivo no 27), fixou a seguinte tese:')),

      justSR(ti('"E admitida a revisao das taxas de juros remuneratorios em situacoes excepcionais, desde que caracterizada a relacao de consumo e que a abusividade (juros excessivos) fique cabalmente demonstrada, ante as peculiaridades do julgamento em concreto."')),

      just(t('O parametro de afeicao da abusividade e a taxa media de mercado para a mesma modalidade de credito, apurada pelo Banco Central (Res. BACEN no 4.559/2017). A taxa efetiva real cobrada pela Re supera em '), tb('114,4%'), t(' a taxa media de mercado - nao se trata de excesso marginal, mas de cobranca sistematicamente mais que o dobro do praticado pelo mercado, situacao que configura abusividade manifesta e autoriza a revisao judicial com fundamento no CDC (art. 6o, V) e no CC (art. 884 - vedacao ao enriquecimento sem causa).')),

      nl(),
      h3('2.3 Da taxa efetiva real superior a taxa declarada - violacao ao dever de transparencia'),

      just(t('O contrato declara taxa de 9,99% ao mes, mas a parcela cobrada (R$ 636,64) implica uma TIR de 10,9352% ao mes - divergencia de 0,94 ponto percentual que, projetada sobre 21 parcelas, corresponde a '), tb('R$ 841,04'), t(' cobrados a mais do que a propria taxa contratual autorizava.')),

      just(t('A Resolucao CMN no 3.517/2007, que disciplina a divulgacao do Custo Efetivo Total - CET -, exige que o consumidor seja informado de forma clara, precisa e em destaque sobre o custo total da operacao de credito. A divergencia entre a taxa declarada (9,99% a.m.) e a taxa efetivamente praticada (10,9352% a.m.) viola diretamente o artigo 6o, inciso III, do CDC - que assegura ao consumidor "a informacao adequada e clara sobre os diferentes produtos e servicos" - e configura clausula contratual enganosa, nula de pleno direito (CDC, art. 51, I e XV).')),

      nl(),
      h3('2.4 Do IOF financiado - juros sobre tributo nao recebido'),

      just(t('A Re incorporou ao principal financiado o valor de '), tb('R$ 163,35 (cento e sessenta e tres reais e trinta e cinco centavos)'), t(' a titulo de IOF, fazendo com que a Autora pagasse juros de 9,99% ao mes sobre essa parcela tributaria durante os 21 meses do contrato - gerando um custo estimado de '), tb('R$ 233,00 (duzentos e trinta e tres reais)'), t(' em juros sobre tributo que a Autora jamais recebeu em especie.')),

      just(t('O IOF e obrigacao tributaria do contribuinte perante a Receita Federal, e seu recolhimento e de responsabilidade do agente financeiro. Ao financia-lo e cobrar juros sobre seu montante, a Re transferiu ao consumidor o custo de uma obrigacao propria, caracterizando vantagem manifestamente excessiva vedada pelo artigo 51, inciso IV, do CDC.')),

      nl(),
      h3('2.5 Da repeticao de indbito - CC art. 876 e CDC art. 42'),

      just(t('Sendo o contrato integralmente quitado em 16 de abril de 2026, a presente acao tem por objeto nao apenas a declaracao de abusividade das clausulas, mas a '), tb('restituicao dos valores pagos a maior'), t(' em razao dos encargos ilegais.')),

      just(t('O artigo 876 do Codigo Civil estabelece que "todo aquele que recebeu o que lhe nao era devido fica obrigado a restituir". O artigo 884 do mesmo Diploma veda o enriquecimento sem causa as custas de outrem. A diferenca entre o que a Autora pagou (R$ 13.276,29) e o que deveria ter pago a taxa de mercado (R$ 8.531,74, a 5,10% ao mes) corresponde ao '), tb('indbito de R$ 4.744,55'), t(' ora reclamado.')),

      just(t('A pretensao de repeticao de indbito e de natureza pessoal e prescreve em 5 (cinco) anos, nos termos do artigo 206, paragrafo 5o, inciso I, do Codigo Civil. O primeiro pagamento data de 11/09/2024 e o ultimo de 16/04/2026, todos dentro do prazo prescricional, que somente se esgotara entre 11/09/2029 e 16/04/2031, conforme detalhado no Memorial de Calculos (Doc. 5).')),

      nl(),

      // III - PEDIDOS
      h2('III - DOS PEDIDOS'),

      just(t('Ante o exposto, requer a Autora:')),

      justSR(tb('a)'), t(' A '), tb('citacao da Re'), t(' para, querendo, contestar o presente feito, sob pena de revelia;')),

      justSR(tb('b)'), t(' A procedencia do pedido, com julgamento de merito, para:')),

      justSR(t('      b.1) '), tb('Declarar a abusividade da taxa de juros remuneratorios'), t(' praticada no Contrato de Credito Pessoal no 13864630, por ser significativamente superior a taxa media de mercado para a modalidade de credito pessoal nao consignado apurada pelo Banco Central do Brasil na data da contratacao (Serie SGS no 25.479 - julho/2024: 5,10% ao mes), nos termos do REsp 1.061.530/RS (STJ, Tema 27);')),

      justSR(t('      b.2) Declarar a '), tb('nulidade da divergencia entre a taxa declarada (9,99% a.m.) e a taxa efetivamente cobrada (TIR: 10,9352% a.m.)'), t(', por violacao ao CDC, art. 6o, III, e a Res. CMN no 3.517/2007, reconhecendo que a diferenca de R$ 841,04 cobrada acima do proprio contrato e indevida;')),

      justSR(t('      b.3) Declarar a '), tb('abusividade do IOF financiado'), t(' (R$ 163,35), sobre o qual a Re indevidamente cobrou juros remuneratorios durante toda a vigencia do contrato, nos termos do CDC, art. 51, IV;')),

      justSR(t('      b.4) Condenar a Re a '), tb('restituicao do indbito apurado'), t(', calculado como a diferenca entre o total efetivamente pago pela Autora (R$ 13.276,29) e o total que deveria ter sido cobrado a taxa media de mercado de 5,10% ao mes (R$ 8.531,74), perfazendo o valor de '), tb('R$ 4.744,55 (quatro mil, setecentos e quarenta e quatro reais e cinquenta e cinco centavos)'), t(' - pedido principal -, ou, alternativamente:')),

      justSR(t('           No Cenario B (subsidiario): restituicao de '), tb('R$ 3.269,36'), t(', calculada com base na taxa maxima de 7,00% ao mes;')),
      justSR(t('           No Cenario C (maximo): restituicao de '), tb('R$ 5.014,46'), t(', calculada com base na taxa BACEN de 5,10% ao mes sobre o principal efetivamente recebido em especie (R$ 5.000,00), excluindo o IOF do principal financiado;')),

      justSR(t('      Todos os valores acima deverao ser corrigidos monetariamente pelo INPC desde a data de cada pagamento indevido e acrescidos de juros de mora de 1% ao mes a partir da citacao, nos termos da Sumula 54 do STJ;')),

      justSR(t('      b.5) Alternativamente ao pedido b.4, na hipotese de restar comprovada a ma-fe da Re, condenar a '), tb('restituicao em dobro do indbito'), t(' no valor de '), tb('R$ 9.489,10 (nove mil, quatrocentos e oitenta e nove reais e dez centavos)'), t(', nos termos do artigo 42, paragrafo unico, do CDC e do STJ, Tema 929;')),

      justSR(tb('c)'), t(' A condenacao da Re ao pagamento das '), tb('custas processuais e honorarios advocaticios'), t(', nos termos do artigo 85 do Codigo de Processo Civil;')),

      justSR(tb('d)'), t(' A '), tb('producao de todos os meios de prova'), t(' admitidos em direito, especialmente:')),
      justSR(t('    - Documental (contrato, extrato completo da operacao, comprovantes de pagamento);')),
      justSR(t('    - Pericial contabil, para conferencia independente dos calculos do indbito e apuracao da TIR efetiva praticada;')),
      justSR(t('    - Juntada de informacoes pelo Banco Central do Brasil sobre a Serie SGS no 25.479 para o mes de julho de 2024.')),

      nl(),
      justSR(tb('Da-se a causa o valor de R$ 4.744,55 (quatro mil, setecentos e quarenta e quatro reais e cinquenta e cinco centavos)'), t(', correspondente ao indbito apurado no Cenario A (pedido principal).')),

      nl(),
      justSR(t('Nestes termos,')),
      justSR(t('Pede deferimento.')),
      nl(),
      justSR(t('Macapa/AP, 17 de junho de 2026.')),
      nl(),
      nl(),
      nl(),

      centro(t('_______________________________________')),
      centro(tb('Conceicao Maria Duarte Portilho')),
      centro(t('Advogada - OAB/AP no 3.576')),
      centro(t('Conceicao Duarte Advocacia e Consultoria Juridica')),
      centro(t('Av. Aimores, 67, Beirol, Macapa/AP - CEP 68.902-140')),
      centro(t('conceicaoduarte.adv@gmail.com')),

      nl(),
      nl(),
      justSR(ti('Documentos que acompanham esta inicial:')),
      justSR(ti('Doc. 1 - Extrato do Contrato de Credito Pessoal no 13864630 (Midway/Riachuelo)')),
      justSR(ti('Doc. 2 - Comprovantes de pagamento das 21 parcelas (boletos bancarios)')),
      justSR(ti('Doc. 3 - Procuracao ad judicia')),
      justSR(ti('Doc. 4 - Serie SGS no 25.479 - BACEN (taxa media credito pessoal nao consignado - jul/2024: 5,10% a.m.)')),
      justSR(ti('Doc. 5 - Memorial de Calculos com Tabela PRICE Comparativa e Quadro do Indbito')),
      justSR(ti('Doc. 6 - [Comprovante de renda - A JUNTAR]')),
      justSR(ti('Doc. 7 - [Comprovante de residencia - A JUNTAR]')),
    ]
  }]
});

const buffer = await Packer.toBuffer(doc);
writeFileSync(OUT, buffer);
console.log('CONCLUIDO: ' + OUT);
