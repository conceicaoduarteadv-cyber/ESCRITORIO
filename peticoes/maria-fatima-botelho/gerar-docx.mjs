import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, Table, TableRow, TableCell,
  WidthType, convertInchesToTwip
} from 'docx';
import { writeFileSync } from 'fs';

const FONTE = 'Times New Roman';
const TAMANHO = 24; // 12pt em half-points
const pt = (n) => n * 20;

const par = (children, opts = {}) => new Paragraph({ children, ...opts });
const txt = (text, opts = {}) => new TextRun({ text, font: FONTE, size: TAMANHO, ...opts });
const bold = (text, opts = {}) => txt(text, { bold: true, ...opts });
const espaco = () => par([txt('')], { spacing: { after: 0 } });

const celula = (text, isBold = false) => new TableCell({
  children: [par([isBold ? bold(text) : txt(text)], { spacing: { after: pt(2) } })],
  margins: { top: pt(2), bottom: pt(2), left: pt(4), right: pt(4) },
});

const tabelaContrato = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [celula('Item', true), celula('Valor', true)] }),
    new TableRow({ children: [celula('Valor líquido do crédito'), celula('R$ 27.556,49')] }),
    new TableRow({ children: [celula('IOF (financiado)'), celula('R$ 872,86')] }),
    new TableRow({ children: [celula('Valor total do crédito (principal)', true), celula('R$ 28.429,35', true)] }),
    new TableRow({ children: [celula('Quantidade de parcelas'), celula('96')] }),
    new TableRow({ children: [celula('Valor de cada parcela'), celula('R$ 656,12')] }),
    new TableRow({ children: [celula('Valor total a ser pago', true), celula('R$ 62.987,52', true)] }),
    new TableRow({ children: [celula('Taxa de juros mensal'), celula('1,87%')] }),
    new TableRow({ children: [celula('Taxa de juros anual'), celula('24,97%')] }),
    new TableRow({ children: [celula('CET mensal'), celula('1,96%')] }),
    new TableRow({ children: [celula('CET anual'), celula('26,70%')] }),
    new TableRow({ children: [celula('Vencimento da 1ª parcela'), celula('10/12/2016')] }),
    new TableRow({ children: [celula('Vencimento da última parcela'), celula('10/11/2024')] }),
  ],
});

const tabelaCusto = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [celula('', true), celula('Valor', true)] }),
    new TableRow({ children: [celula('Valor líquido do crédito'), celula('R$ 27.556,49')] }),
    new TableRow({ children: [celula('Total pago (96 × R$ 656,12)'), celula('R$ 62.987,52')] }),
    new TableRow({ children: [celula('Custo do crédito', true), celula('R$ 35.431,03', true)] }),
    new TableRow({ children: [celula('Custo como % do valor recebido', true), celula('128,7%', true)] }),
  ],
});

const cabecalho = par(
  [bold('EXCELENTÍSSIMO(A) SENHOR(A) DOUTOR(A) JUIZ(A) DE DIREITO DA VARA CÍVEL DA COMARCA DE MACAPÁ — ESTADO DO AMAPÁ')],
  { alignment: AlignmentType.CENTER, spacing: { after: pt(18) } }
);

const tituloAcao = par(
  [bold('AÇÃO REVISIONAL DE CONTRATO BANCÁRIO COM PEDIDO DE PRODUÇÃO ANTECIPADA DE PROVAS', { size: 28 })],
  { alignment: AlignmentType.CENTER, spacing: { before: pt(6), after: pt(18) } }
);

const qualificacao = [
  par([
    bold('MARIA DE FÁTIMA DIAS BOTELHO'),
    txt(', brasileira, solteira, pensionista, portadora do RG nº 740893 (POLITEC/AP) e inscrita no CPF sob o nº 066.796.172-00, residente e domiciliada na Rua Santos Dumont, nº 3820, Bairro Muca, Macapá/AP, CEP 68902-230, por meio de sua advogada que esta subscreve, nos termos do instrumento procuratório em anexo, vem, respeitosamente, perante Vossa Excelência, propor'),
  ], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),
  par([
    txt('em face de '),
    bold('BANCO PAN S/A'),
    txt(', instituição financeira inscrita no CNPJ/MF sob o nº 59.285.411/0001-13, com sede na Avenida Paulista, nº 1.374, 12º andar, Bela Vista, São Paulo/SP, CEP 01310-300, pelos fatos e fundamentos jurídicos a seguir expostos:'),
  ], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(18) } }),
];

const secaoI = [
  par([bold('I — DOS FATOS')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(8) } }),

  par([bold('1. A contratação')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('Em 17 de outubro de 2016, a Autora celebrou com o Réu a '), bold('Cédula de Crédito Bancário nº 712228840-5'), txt(', modalidade consignado público — refinanciamento, vinculada à fonte pagadora SIAPE PENSION (matrícula nº 05018714), com as seguintes características declaradas pelo próprio Réu:')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(8) } }),
  tabelaContrato,
  espaco(),
  par([txt('Intermediou a operação a correspondente bancária AY SOLUÇÕES LTDA ME (CNPJ 16.781.906/0001-20), localizada em Macapá/AP.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { before: pt(6), after: pt(12) } }),

  par([bold('2. O que a Autora efetivamente recebeu')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('Embora o "Valor Líquido do Crédito" declarado na CCB seja de R$ 27.556,49, o Réu destinou R$ 25.136,66 desse montante à liquidação de contrato anterior celebrado com o próprio '), bold('Banco PAN S/A'), txt(' — refinanciamento interno, portanto. O valor efetivamente creditado na conta bancária da Autora foi de apenas '), bold('R$ 2.419,83'), txt('.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(8) } }),
  par([txt('A Autora comprometeu sua renda de pensionista ao longo de '), bold('8 anos'), txt(', em 96 parcelas descontadas diretamente em folha, para receber em conta a quantia de '), bold('R$ 2.419,83'), txt(' — e pagar ao Réu o total de '), bold('R$ 62.987,52'), txt('.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('3. O IOF embutido no principal da dívida')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('O IOF apurado na operação, no valor de '), bold('R$ 872,86'), txt(', foi incorporado ao principal da dívida, elevando-o de R$ 27.556,49 para R$ 28.429,35. A consequência direta é que a Autora pagou '), bold('juros remuneratórios de 1,87% ao mês sobre um tributo'), txt(' durante 96 meses — prática que onera ilegalmente o custo do crédito ao converter imposto em base de incidência de juros bancários.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('4. O refinanciamento interno sem demonstrativo do contrato anterior')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('O Réu liquidou, com recursos do novo contrato, dívida anterior da própria Autora '), bold('com o próprio Banco PAN S/A'), txt(', no montante de R$ 25.136,66. Não foi apresentado à Autora, no momento da contratação, qualquer demonstrativo do saldo devedor do contrato original que justificasse esse valor.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('Não se sabe, portanto: (i) qual era o valor original do contrato anterior; (ii) qual a taxa de juros que nele incidia; (iii) se o saldo devedor de R$ 25.136,66 foi calculado corretamente; (iv) se cobranças indevidas do contrato anterior foram incorporadas ao saldo e, com ele, ao novo principal.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('5. O formulário com data inconsistente — indício de formulário em branco')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('A CCB apresenta três datas que se contradizem: (i) "Data do Lote" impressa no rodapé do formulário: '), bold('01/06/2015'), txt('; (ii) "Data de Emissão" no quadro: '), bold('25/04/2016'), txt('; (iii) data da assinatura e liberação efetiva: '), bold('17/10/2016'), txt('. O formulário foi fabricado em junho de 2015, preenchido com data de emissão em abril de 2016 e assinado em outubro de 2016 — divergência de mais de 16 meses que aponta para utilização de formulário em branco, viciando o consentimento.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('6. O custo total desproporcional')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  tabelaCusto,
  espaco(),
  par([txt('A Autora pagou, a título de juros e encargos, valor superior ao dobro do que tomou emprestado.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { before: pt(6), after: pt(12) } }),
];

const secaoII = [
  par([bold('II — DO DIREITO')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(8) } }),

  par([bold('7. Aplicabilidade do Código de Defesa do Consumidor')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('As instituições financeiras sujeitam-se ao Código de Defesa do Consumidor (ADI 2.591/DF — STF; '), bold('Súmula 297 do STJ'), txt(': "O Código de Defesa do Consumidor é aplicável às instituições financeiras"). A Autora é consumidora final e hipossuficiente técnica em matéria bancária. A relação é inequivocamente consumerista.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('8. Da ilegalidade do IOF financiado')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('O Imposto sobre Operações Financeiras (IOF) é tributo regido pelo Decreto nº 6.306/2007. Ao incluí-lo no principal da dívida, o Réu transformou tributo em base de incidência de juros remuneratórios. A prática: (i) afronta a vedação ao anatocismo ('), bold('Súmula 121 do STF'), txt('; CC, art. 591); (ii) viola a transparência exigida pelo CDC (art. 46); (iii) configura enriquecimento ilícito do Réu (CC, art. 884). O financiamento do IOF não encontra respaldo em norma do Banco Central ou do Conselho Monetário Nacional.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('9. Da nulidade das cláusulas por ausência de transparência no refinanciamento')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('A '), bold('Resolução CMN nº 3.954/2011'), txt(' impõe às instituições financeiras o dever de fornecer informações claras sobre operações de crédito consignado, incluindo a demonstração do saldo devedor do contrato refinanciado. O mesmo dever decorre do '), bold('CDC, art. 6º, III'), txt('. No caso, o Réu liquidou contrato anterior com o próprio banco sem apresentar o demonstrativo de saldo devedor, impedindo a Autora de verificar se o refinanciamento era economicamente vantajoso.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('A omissão de informações essenciais torna as cláusulas correspondentes nulas, nos termos do '), bold('CDC, art. 46'), txt('. Aplica-se o '), bold('STJ, REsp 1.061.530/RS (Tema 27 — Recurso Repetitivo)'), txt(', que autoriza a revisão de cláusulas bancárias abusivas independentemente de prova de desproporção econômica.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('10. Da abusividade decorrente do formulário pré-datado')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('A utilização de formulário com "Data do Lote" de junho de 2015 em contrato assinado em outubro de 2016 indica preenchimento em data não contemporânea à assinatura. Trata-se de prática que: (i) vicia o consentimento (CC, arts. 138 e 147); (ii) configura indício de formulário em branco preenchido após a assinatura, tornando o instrumento nulo na forma do '), bold('CC, art. 166, II'), txt(' e do '), bold('CDC, art. 51, XIII'), txt('; (iii) impede à consumidora a verificação das condições contratadas.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('11. Da onerosidade excessiva e da função social do contrato')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('O contrato impôs à Autora — pensionista com rendimentos fixos — o custo de R$ 62.987,52 sobre um crédito líquido de R$ 27.556,49, dos quais apenas R$ 2.419,83 ingressaram em sua conta. A vantagem obtida pelo Réu é manifestamente excessiva, caracterizando '), bold('lesão (CC, art. 157)'), txt(' e '), bold('onerosidade excessiva (CC, arts. 478 e 480)'), txt('. A função social do contrato (CC, art. 421) e a boa-fé objetiva (CC, art. 422) impõem a revisão das cláusulas que conduziram a esse resultado.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('12. Da devolução dos valores cobrados indevidamente')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('Reconhecida a nulidade das cláusulas abusivas, o Réu deverá devolver os valores cobrados a maior, devidamente corrigidos. A devolução deve ser em dobro, nos termos do '), bold('CDC, art. 42, parágrafo único'), txt(', pois não há engano justificável: o Réu é instituição financeira de grande porte, e as irregularidades apontadas são estruturais e deliberadas ('), bold('Súmula 322/STJ'), txt(' — a devolução simples aplica-se apenas à cobrança de boa-fé).')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),
];

const secaoIII = [
  par([bold('III — DO PEDIDO DE PRODUÇÃO ANTECIPADA DE PROVAS')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(8) } }),
  par([txt('Para a completa apuração dos fatos narrados — em especial a regularidade do saldo devedor do contrato anterior e a composição exata do principal refinanciado —, a Autora requer, com fundamento no '), bold('CPC, art. 381, II'), txt(', a determinação ao Réu de exibição dos seguintes documentos, no prazo de 15 (quinze) dias, sob pena de multa diária de R$ 500,00 (CPC, art. 400):')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(8) } }),
  par([txt('a) Contrato original ou CCB anterior à nº 712228840-5, celebrado entre a Autora e o Banco PAN S/A, que originou o saldo de R$ 25.136,66 liquidado em 17/10/2016;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('b) Extrato completo do contrato anterior, com histórico de pagamentos, saldo devedor, encargos e cálculo do valor de quitação;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('c) Demonstrativo detalhado do cálculo do IOF (R$ 872,86) cobrado na CCB nº 712228840-5;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('d) Planilha de amortização do contrato objeto desta ação, com discriminação de principal, juros e encargos em cada parcela.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) }, indent: { left: convertInchesToTwip(0.3) } }),
];

const secaoIV = [
  par([bold('IV — DOS PEDIDOS')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(8) } }),
  par([txt('Ante o exposto, requer a Vossa Excelência:')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(8) } }),

  par([bold('a)'), txt(' O recebimento e processamento da presente ação;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),

  par([bold('b)'), txt(' A citação do Réu '), bold('BANCO PAN S/A'), txt(' para contestar, sob pena de revelia (CPC, art. 344);')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),

  par([bold('c)'), txt(' A '), bold('produção antecipada de provas'), txt(' mediante determinação ao Réu para exibir os documentos listados no item III, no prazo de 15 dias, sob pena de multa diária e busca e apreensão (CPC, arts. 381 e 400);')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),

  par([bold('d)'), txt(' No mérito, a procedência integral dos pedidos para:')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.3) } }),

  par([txt('d.1) Declarar a '), bold('nulidade da cláusula de financiamento do IOF'), txt(' (R$ 872,86), determinando o expurgo desse valor do principal da dívida e o recálculo de todas as parcelas, com devolução dos valores pagos a maior;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.6) } }),

  par([txt('d.2) Declarar a '), bold('nulidade das cláusulas relativas ao refinanciamento interno'), txt(' do contrato anterior, por ausência de demonstrativo do saldo devedor e violação ao dever de informação (CDC, art. 6º, III; Res. CMN 3.954/2011), determinando a revisão do principal com base no saldo efetivamente devido;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.6) } }),

  par([txt('d.3) Determinar o '), bold('recálculo de toda a operação'), txt(' com expurgo das ilegalidades identificadas, apurado por perito contábil nomeado pelo Juízo;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.6) } }),

  par([txt('d.4) Condenar o Réu à '), bold('devolução em dobro'), txt(' de todos os valores cobrados e pagos indevidamente, nos termos do CDC, art. 42, parágrafo único, devidamente atualizados pelo IPCA desde cada desembolso e acrescidos de juros moratórios de 1% ao mês desde a citação (CC, art. 406; Súm. 54/STJ);')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.6) } }),

  par([txt('d.5) Condenar o Réu ao pagamento de '), bold('danos morais'), txt(' em valor a ser arbitrado por Vossa Excelência, em razão da utilização de formulário pré-datado, da falta de transparência no refinanciamento e da submissão da Autora a obrigação financeira desproporcional;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.6) } }),

  par([bold('e)'), txt(' A condenação do Réu ao pagamento das custas processuais e honorários advocatícios sucumbenciais (CPC, art. 85), fixados em valor não inferior a 20% sobre o proveito econômico obtido pela Autora;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),

  par([bold('f)'), txt(' A concessão dos benefícios da '), bold('justiça gratuita'), txt(' à Autora (CPC, art. 98);')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),

  par([bold('g)'), txt(' A '), bold('inversão do ônus da prova'), txt(' em favor da Autora (CDC, art. 6º, VIII), determinando ao Réu que comprove a legalidade de cada encargo cobrado.'),],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) }, indent: { left: convertInchesToTwip(0.3) } }),
];

const secaoV = [
  par([bold('V — DAS PROVAS')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(8) } }),
  par([txt('A Autora pretende provar os fatos alegados por todos os meios de prova admitidos em direito, em especial: (i) '), bold('documental'), txt(' — CCB nº 712228840-5, planilha de proposta simplificada e CET em anexo, bem como os documentos a serem exibidos pelo Réu; (ii) '), bold('pericial contábil'), txt(' — para apuração do excesso de juros cobrado, recálculo da dívida com expurgo das ilegalidades e quantificação do valor a ser devolvido; (iii) quaisquer '), bold('outras provas'), txt(' que se fizerem necessárias no curso do processo.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),
];

const valorEFechamento = [
  par([txt('Dá-se à causa o valor de '), bold('R$ 62.987,52 (sessenta e dois mil, novecentos e oitenta e sete reais e cinquenta e dois centavos)'), txt(', correspondente ao total contratual cobrado pelo Réu.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(36) } }),

  par([txt('Termos em que,')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) } }),
  par([txt('Pede deferimento.')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(36) } }),

  par([txt('Macapá/AP, _____ de __________________ de 2026.')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(60) } }),

  par([txt('___________________________________________')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([bold('CONCEIÇÃO MARIA DUARTE PORTILHO')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([txt('OAB/AP nº 3576')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([txt('conceicaoduarte.adv@gmail.com')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([txt('Av. Aimorés, nº 67, Beirol — Macapá/AP — CEP 68902-140')], { alignment: AlignmentType.CENTER, spacing: { after: pt(24) } }),
];

const documentos = [
  par([bold('DOCUMENTOS QUE ACOMPANHAM A INICIAL')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(8) } }),
  par([txt('1. Procuração ad judicia et extra')], { spacing: { after: pt(4) } }),
  par([txt('2. Declaração de hipossuficiência econômica')], { spacing: { after: pt(4) } }),
  par([txt('3. CCB nº 712228840-5 (3 páginas) — Banco PAN S/A')], { spacing: { after: pt(4) } }),
  par([txt('4. CET — Custo Efetivo Total (planilha Banco PAN)')], { spacing: { after: pt(4) } }),
  par([txt('5. Planilha de Proposta Simplificada (Banco PAN)')], { spacing: { after: pt(4) } }),
  par([txt('6. Cópia do RG e CPF da Autora')], { spacing: { after: pt(4) } }),
];

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONTE, size: TAMANHO } },
    },
    paragraphStyles: [
      {
        id: 'Heading1',
        name: 'Heading 1',
        basedOn: 'Normal',
        run: { bold: true, font: FONTE, size: TAMANHO, color: '000000' },
        paragraph: { spacing: { before: pt(12), after: pt(6) } },
      },
      {
        id: 'Heading2',
        name: 'Heading 2',
        basedOn: 'Normal',
        run: { bold: true, font: FONTE, size: TAMANHO, color: '000000' },
        paragraph: { spacing: { before: pt(8), after: pt(4) } },
      },
    ],
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
      cabecalho,
      espaco(),
      tituloAcao,
      espaco(),
      ...qualificacao,
      ...secaoI,
      ...secaoII,
      ...secaoIII,
      ...secaoIV,
      ...secaoV,
      ...valorEFechamento,
      ...documentos,
    ],
  }],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync('peticoes/maria-fatima-botelho/peticao-revisional-banco-pan-2026.docx', buffer);
console.log('Arquivo gerado: peticoes/maria-fatima-botelho/peticao-revisional-banco-pan-2026.docx');
