import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, convertInchesToTwip, UnderlineType
} from 'docx';
import { writeFileSync } from 'fs';

const FONTE = 'Times New Roman';
const TAMANHO = 24; // 12pt em half-points

const pt = (n) => n * 20; // pontos para twips

// Helpers
const par = (children, opts = {}) => new Paragraph({ children, ...opts });
const txt = (text, opts = {}) => new TextRun({ text, font: FONTE, size: TAMANHO, ...opts });
const bold = (text, opts = {}) => txt(text, { bold: true, ...opts });
const espaco = () => par([txt('')], { spacing: { after: 0 } });

// Cabeçalho centralizado
const cabecalho = par(
  [bold('EXCELENTÍSSIMO(A) SENHOR(A) DOUTOR(A) JUIZ(A) DE DIREITO DA VARA CÍVEL DA COMARCA DE MACAPÁ — ESTADO DO AMAPÁ')],
  { alignment: AlignmentType.CENTER, spacing: { after: pt(12) } }
);

// Título da ação
const tituloAcao = par(
  [bold('AÇÃO REVISIONAL DE CONTRATOS BANCÁRIOS COM PEDIDO DE TUTELA DE URGÊNCIA', { size: 28 })],
  { alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1, spacing: { before: pt(6), after: pt(12) } }
);

// Qualificação da autora
const qualificacao = [
  par([
    bold('SIMONE MARQUES MARTINS SEPEDA'),
    txt(', brasileira, nascida em 09 de dezembro de 1971, natural de Macapá/AP, portadora do RG nº 040825 (POLITEC/AP) e inscrita no CPF sob o nº 324.863.672-68, residente e domiciliada na Rua Carlos Gomes, nº 301, Bairro Jesus de Nazaré, Macapá/AP, CEP 68908-125, e-mail symonemartins2002@yahoo.com.br, por meio de sua advogada que esta subscreve, nos termos do instrumento procuratório em anexo, vem, respeitosamente, perante Vossa Excelência, propor'),
  ], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),
  par([txt('em face de '), bold('MIDWAY S.A. — CRÉDITO, FINANCIAMENTO E INVESTIMENTO'), txt(', instituição financeira inscrita no CNPJ/ME sob o nº 09.464.032/0001-12, com sede na Rua Lemos Monteiro, nº 120, 15º Andar, Edifício Pinheiros One, Butantã, São Paulo/SP, CEP 05501-050, pelos fatos e fundamentos a seguir expostos:')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(18) } }),
];

// Seção I
const secaoI = [
  par([bold('I — DOS FATOS')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(6) } }),

  par([bold('1. A contratação')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),

  par([txt('Entre os meses de fevereiro e abril de 2026, a Autora celebrou três Cédulas de Crédito Bancário (CCBs) com a Ré, intermediadas pelas Lojas Riachuelo S.A. na condição de correspondente bancário, conforme discriminado abaixo:')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(10) } }),

  // CCB 1
  par([bold('CCB nº 75805602 — emitida em 18/02/2026')], { spacing: { before: pt(6), after: pt(4) } }),
  par([txt('Valor líquido liberado: R$ 5.000,00')], { spacing: { after: pt(2) } }),
  par([txt('Valor total do empréstimo (com IOF e seguro): R$ 5.529,71')], { spacing: { after: pt(2) } }),
  par([txt('IOF embutido e financiado: R$ 179,71')], { spacing: { after: pt(2) } }),
  par([txt('Seguro Prestamista embutido: R$ 350,00')], { spacing: { after: pt(2) } }),
  par([txt('Taxa de juros remuneratórios: '), bold('13,75% ao mês / 369,26% ao ano')], { spacing: { after: pt(2) } }),
  par([txt('Custo Efetivo Total (CET): 14,41% ao mês / 402,79% ao ano')], { spacing: { after: pt(2) } }),
  par([txt('Número de parcelas: 30 (trinta)')], { spacing: { after: pt(2) } }),
  par([txt('Valor de cada parcela: R$ 775,06')], { spacing: { after: pt(2) } }),
  par([bold('Valor total a pagar: R$ 23.251,80')], { spacing: { after: pt(10) } }),

  // CCB 2
  par([bold('CCB nº 75825597 — emitida em 09/03/2026')], { spacing: { before: pt(6), after: pt(4) } }),
  par([txt('Valor líquido liberado: R$ 6.372,78')], { spacing: { after: pt(2) } }),
  par([txt('Valor total do empréstimo (com IOF e seguro): R$ 7.049,46')], { spacing: { after: pt(2) } }),
  par([txt('IOF embutido e financiado: R$ 230,59')], { spacing: { after: pt(2) } }),
  par([txt('Seguro Prestamista embutido: R$ 446,09')], { spacing: { after: pt(2) } }),
  par([txt('Taxa de juros remuneratórios: '), bold('13,75% ao mês / 369,26% ao ano')], { spacing: { after: pt(2) } }),
  par([txt('Custo Efetivo Total (CET): 14,41% ao mês / 402,80% ao ano')], { spacing: { after: pt(2) } }),
  par([txt('Número de parcelas: 30 (trinta)')], { spacing: { after: pt(2) } }),
  par([txt('Valor de cada parcela: R$ 994,60')], { spacing: { after: pt(2) } }),
  par([bold('Valor total a pagar: R$ 29.838,00')], { spacing: { after: pt(10) } }),

  // CCB 3
  par([bold('CCB nº 75879070 — emitida em 27/04/2026')], { spacing: { before: pt(6), after: pt(4) } }),
  par([txt('Valor líquido liberado: R$ 1.600,00')], { spacing: { after: pt(2) } }),
  par([txt('Valor total do empréstimo (com IOF e seguro): R$ 1.740,92')], { spacing: { after: pt(2) } }),
  par([txt('IOF embutido e financiado: R$ 28,92')], { spacing: { after: pt(2) } }),
  par([txt('Seguro Prestamista embutido: R$ 112,00')], { spacing: { after: pt(2) } }),
  par([txt('Taxa de juros remuneratórios: '), bold('14,99% ao mês / 434,47% ao ano')], { spacing: { after: pt(2) } }),
  par([txt('Custo Efetivo Total (CET): 15,62% ao mês / 470,41% ao ano')], { spacing: { after: pt(2) } }),
  par([txt('Número de parcelas: 8 (oito)')], { spacing: { after: pt(2) } }),
  par([txt('Valor de cada parcela: R$ 387,50')], { spacing: { after: pt(2) } }),
  par([bold('Valor total a pagar: R$ 3.100,00')], { spacing: { after: pt(12) } }),

  par([bold('2. O descompasso entre o que foi recebido e o que se cobra')],
    { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('A Autora recebeu, nos três contratos, o total líquido de '), bold('R$ 12.972,78 (doze mil, novecentos e setenta e dois reais e setenta e oito centavos)'), txt('. Em contrapartida, o montante total cobrado pela Ré nos três instrumentos soma '), bold('R$ 56.189,80 (cinquenta e seis mil, cento e oitenta e nove reais e oitenta centavos)')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('Trata-se de sobrecusto de '), bold('R$ 43.217,02'), txt(' — o equivalente a mais de '), bold('333%'), txt(' do valor efetivamente recebido, ou seja, a Autora pagaria mais de quatro vezes o valor que ingressou em sua conta.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('3. A prática abusiva do seguro prestamista')],
    { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('Nos três contratos, a Ré inseriu automaticamente seguros prestamistas nos valores de R$ 350,00, R$ 446,09 e R$ 112,00, totalizando '), bold('R$ 908,09'), txt(', administrados pela Zurich Seguros, através da intermediação das Lojas Riachuelo.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('Em nenhum dos contratos foi concedida à Autora a possibilidade real de recusar o seguro ou de contratar livremente outra seguradora de sua escolha. O prêmio foi diretamente incorporado ao valor financiado, sobre o qual ainda incidem juros remuneratórios — onerando duplamente a consumidora.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('4. O IOF embutido no montante financiado')],
    { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('Nos três contratos, o IOF (R$ 179,71 + R$ 230,59 + R$ 28,92 = '), bold('R$ 439,22'), txt(') foi incorporado ao valor total financiado, compondo a base de cálculo sobre a qual incidiram os juros remuneratórios. A Autora, assim, pagou juros sobre um tributo — o que desvirtua a base de cálculo das parcelas e gera enriquecimento ilícito da Ré.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('5. A capitalização mensal de juros')],
    { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('A cláusula 2 das três CCBs estabelece expressamente que os juros são '), bold('"capitalizados mensalmente"'), txt('. Combinada com as taxas praticadas (13,75% e 14,99% ao mês), a capitalização mensal gera um custo anual de 369% a 434%, transformando dívidas de R$ 12.972,78 em obrigações de R$ 56.189,80 — prática que viola frontalmente a função social dos contratos e os princípios da boa-fé objetiva.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('6. Cláusulas que suprimem direitos do consumidor')],
    { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('O item 9 das Condições Gerais de cada CCB estabelece que a Autora não poderia invocar os arts. 317, 393, 478, 479 e 480 do Código Civil — dispositivos que tratam da revisão por onerosidade excessiva, caso fortuito, força maior e teoria da imprevisão. Cláusulas que suprimem direitos legais do consumidor são nulas de pleno direito, independentemente de aceitação.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),
];

// Seção II
const secaoII = [
  par([bold('II — DO DIREITO')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(6) } }),

  par([bold('7. Aplicabilidade do Código de Defesa do Consumidor')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('A relação jurídica entre a Autora e a Ré é de consumo, nos termos do art. 2º e 3º da Lei nº 8.078/90. A Súmula 297 do STJ consagra que "o Código de Defesa do Consumidor é aplicável às instituições financeiras", afastando qualquer argumento em contrário.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('8. Abusividade das taxas de juros')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('Os contratos bancários, embora não se submetam à limitação da Lei de Usura (Decreto nº 22.626/33), estão sujeitos ao controle de abusividade pelo Poder Judiciário, com base no CDC e no Código Civil.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('O Superior Tribunal de Justiça pacificou que, para configurar a abusividade, basta demonstrar que a taxa contratada supera significativamente a taxa média de mercado apurada pelo Banco Central do Brasil para a mesma modalidade de crédito (REsp 1.061.530/RS — Tema 27 STJ).')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('A taxa média BACEN para crédito pessoal não consignado no período dos contratos (fevereiro a abril de 2026) situou-se em patamar substancialmente inferior às taxas de 13,75% e 14,99% ao mês praticadas pela Ré, o que demonstra a abusividade e autoriza a revisão judicial.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('Requer-se a juntada de planilha BACEN como prova documental. Autoriza-se o Juízo a determinar a consulta ao Banco Central para comprovação da taxa média vigente nas datas de 18/02/2026, 09/03/2026 e 27/04/2026.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12), before: pt(6) }, indent: { left: convertInchesToTwip(0.5), right: convertInchesToTwip(0.5) } }),

  par([bold('9. Ilegalidade do seguro prestamista — venda casada')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('A Resolução CNSP nº 382/2020 e a Circular SUSEP nº 642/2021 garantem ao segurado o direito à livre escolha da seguradora, vedando expressamente que a instituição financeira condicione a concessão do crédito à contratação de seguro junto à seguradora por ela indicada.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('No presente caso, a Ré inseriu o seguro prestamista diretamente no valor financiado, sem que a Autora tivesse a possibilidade real de recusar ou de contratar o seguro com outra empresa. Trata-se de venda casada, vedada pelo art. 39, inciso I, do CDC, e da qual decorre o direito à restituição dos prêmios indevidamente cobrados, nos termos do art. 42, parágrafo único, do CDC (repetição em dobro quando a cobrança é indevida e dolosa).')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('A Súmula 473 do STJ não impede a contestação do seguro prestamista quando demonstrada a venda casada:')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('"O mutuário do SFH não pode ser compelido a contratar o seguro habitacional obrigatório com a instituição financeira mutuante ou com a seguradora por ela indicada." (Súm. 473/STJ — aplicação analógica ao crédito pessoal)')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12), before: pt(6) }, indent: { left: convertInchesToTwip(0.5), right: convertInchesToTwip(0.5) } }),

  par([bold('10. IOF embutido no capital financiado')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('O IOF é tributo de responsabilidade do tomador do crédito e deve ser recolhido à Receita Federal. Ao incorporá-lo ao valor financiado, a Ré passa a cobrar juros sobre o próprio tributo, gerando enriquecimento sem causa (CC art. 884) e majorando indevidamente as parcelas.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('11. Nulidade das cláusulas que afastam o CDC e o Código Civil')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('O art. 51, inciso I, do CDC estabelece que são nulas de pleno direito as cláusulas contratuais que impossibilitem, exonerem ou atenuem a responsabilidade do fornecedor por vícios, ou que impliquem renúncia ou disposição de direitos. O art. 51, inciso IV, acrescenta a nulidade das cláusulas incompatíveis com a boa-fé e a equidade.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('A cláusula que veda à Autora invocar os arts. 317, 393, 478, 479 e 480 do Código Civil é nula de pleno direito, operando a nulidade independentemente de declaração judicial, nos termos do art. 51, § 2º, do CDC.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),

  par([bold('12. Boa-fé objetiva e função social do contrato')], { heading: HeadingLevel.HEADING_2, spacing: { after: pt(6) } }),
  par([txt('O art. 421 do Código Civil estabelece que a liberdade contratual será exercida nos limites da função social do contrato. O art. 422 impõe às partes o dever de guardar a boa-fé objetiva tanto na conclusão como na execução do contrato. Contratos que geram enriquecimento desproporcionado em detrimento do consumidor vulnerável violam esses princípios e devem ser revisados.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) } }),
];

// Seção III
const secaoIII = [
  par([bold('III — DA TUTELA DE URGÊNCIA (CPC art. 300)')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(6) } }),
  par([txt('A probabilidade do direito (fumus boni iuris) está demonstrada pelos próprios contratos, que evidenciam taxas superiores à média de mercado, inserção compulsória de seguro e cláusulas abusivas — vícios objetivos verificáveis de plano.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('O perigo de dano (periculum in mora) decorre do risco iminente de:')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) } }),
  par([txt('— Inclusão do nome da Autora em cadastros de inadimplentes (SPC/Serasa) caso as parcelas, ora discutidas em sua legalidade, não sejam pagas no valor contratado;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('— Continuidade do débito automático nas contas bancárias da Autora durante toda a tramitação do processo, consolidando o pagamento de valores cuja legalidade é questionada;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('— Dano de difícil reparação ao crédito da Autora, afetando sua capacidade financeira.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(10) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([bold('Requer-se, em sede de tutela de urgência antecipada (CPC art. 300 c/c art. 303):')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('a) A suspensão de qualquer inclusão do nome da Autora em serviços de proteção ao crédito (SPC, Serasa, SCR) em razão das CCBs nºs 75805602, 75825597 e 75879070, enquanto pendente o presente processo;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('b) A determinação à Ré de que se abstenha de realizar débitos automáticos nas contas da Autora (Banco Midway, Ag. 1001, CC 53676499; e Banco do Brasil, Ag. 2825, CC 146602) em valores superiores às parcelas que vierem a ser fixadas pelo Juízo após a revisão dos contratos, ou, alternativamente, que os débitos fiquem suspensos até decisão de mérito, mediante depósito judicial da diferença;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('c) A expedição de ofício às instituições financeiras Banco Midway e Banco do Brasil para ciência da decisão liminar.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) }, indent: { left: convertInchesToTwip(0.3) } }),
];

// Seção IV
const secaoIV = [
  par([bold('IV — DOS PEDIDOS')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(6) } }),
  par([txt('Diante do exposto, requer a Autora:')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(8) } }),
  par([bold('I — DA TUTELA DE URGÊNCIA'), txt(', nos termos do item III acima, com fixação de multa diária (astreinte) de R$ 1.000,00 (mil reais) por descumprimento, nos termos do art. 537 do CPC.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(8) } }),
  par([bold('II — DO MÉRITO:')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('a) A revisão das taxas de juros remuneratórios das três CCBs, limitando-as à taxa média de mercado apurada pelo Banco Central do Brasil para crédito pessoal não consignado nas datas de cada contratação;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('b) O recálculo das parcelas vencidas e vincendas com a nova taxa fixada, compensando os valores eventualmente pagos a maior pela Autora, com correção monetária pelo IPCA e juros legais de 1% ao mês a partir de cada pagamento indevido (Súm. 54/STJ);')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('c) A declaração de nulidade da cláusula de capitalização mensal de juros quando combinada com taxas superiores à média de mercado, ou, alternativamente, a limitação dos efeitos da capitalização à taxa revisada;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('d) A declaração de nulidade da cobrança do Seguro Prestamista nas três CCBs, por caracterizar venda casada (CDC art. 39, I), com a condenação da Ré à restituição dos valores pagos a esse título (R$ 908,09 comprovados, mais os que se vencerem até a sentença), corrigidos monetariamente e acrescidos de juros legais, podendo a restituição ser simples ou em dobro conforme apuração em liquidação (CDC art. 42, parágrafo único);')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('e) A declaração de nulidade da incorporação do IOF ao montante financiado para fins de incidência de juros, com recálculo das parcelas excluindo o tributo da base de cálculo dos juros remuneratórios e restituição do excesso;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('f) A declaração de nulidade das cláusulas que afastam a aplicação dos arts. 317, 393, 478, 479 e 480 do Código Civil (CDC art. 51, I e IV);')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('g) A condenação da Ré ao pagamento das custas processuais e honorários advocatícios sucumbenciais, nos termos do art. 85 do CPC;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('h) A concessão dos benefícios da gratuidade da justiça à Autora, nos termos do art. 98 do CPC e da Lei nº 1.060/50, pela declaração de hipossuficiência que instrui esta inicial.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) }, indent: { left: convertInchesToTwip(0.3) } }),

  par([bold('III — DAS PROVAS')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(6) } }),
  par([txt('A Autora pretende provar o alegado por todos os meios de prova em direito admitidos, em especial:')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) } }),
  par([txt('— Documentais: CCBs nºs 75805602, 75825597 e 75879070 (documentos em anexo); CNH; extratos bancários demonstrando os débitos das parcelas;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('— Pericial: requer desde já a realização de perícia contábil para apuração do montante pago a maior, com base na taxa revisada a ser fixada pelo Juízo;')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) }, indent: { left: convertInchesToTwip(0.3) } }),
  par([txt('— Prova documental a ser produzida: dados da taxa média BACEN para crédito pessoal não consignado nas datas dos contratos (a ser obtida pelo próprio Juízo ou por diligência da Autora).')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(12) }, indent: { left: convertInchesToTwip(0.3) } }),
];

// Seção V
const secaoV = [
  par([bold('V — DO VALOR DA CAUSA')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(6) } }),
  par([txt('Atribui-se à presente causa o valor de '), bold('R$ 56.189,80 (cinquenta e seis mil, cento e oitenta e nove reais e oitenta centavos)'), txt(', correspondente à soma dos valores totais a pagar nas três CCBs (CPC art. 292, II), sem prejuízo da liquidação posterior dos valores a restituir.')],
    { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(24) } }),
];

// Fechamento
const fechamento = [
  par([txt('Nestes termos,')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(4) } }),
  par([txt('Pede deferimento.')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(36) } }),
  par([txt('Macapá/AP, _____ de __________________ de 2026.')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(60) } }),
  par([txt('___________________________________________')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([bold('CONCEIÇÃO MARIA DUARTE PORTILHO')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([txt('OAB/AP nº 3576')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([txt('conceicaoduarte.adv@gmail.com')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([txt('Av. Aimorés, nº 67, Beirol — Macapá/AP — CEP 68902-140')], { alignment: AlignmentType.CENTER, spacing: { after: pt(24) } }),
];

// Documentos a juntar
const documentos = [
  par([bold('DOCUMENTOS A JUNTAR')], { heading: HeadingLevel.HEADING_1, spacing: { before: pt(12), after: pt(8) } }),
  par([txt('☐  Procuração ad judicia et extra assinada pela Autora')], { spacing: { after: pt(4) } }),
  par([txt('☐  Declaração de hipossuficiência / gratuidade de justiça')], { spacing: { after: pt(4) } }),
  par([txt('☐  CCB nº 75805602 (18/02/2026)')], { spacing: { after: pt(4) } }),
  par([txt('☐  CCB nº 75825597 (09/03/2026)')], { spacing: { after: pt(4) } }),
  par([txt('☐  CCB nº 75879070 (27/04/2026)')], { spacing: { after: pt(4) } }),
  par([txt('☐  CNH da Autora (qualificação)')], { spacing: { after: pt(4) } }),
  par([txt('☐  Comprovante de residência atualizado')], { spacing: { after: pt(4) } }),
  par([txt('☐  Extratos bancários (débitos das parcelas)')], { spacing: { after: pt(4) } }),
  par([txt('☐  Consulta BACEN — taxa média crédito pessoal não consignado (fev/mar/abr 2026)')], { spacing: { after: pt(4) } }),
];

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONTE, size: TAMANHO },
      },
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
          top: convertInchesToTwip(1.18),    // 3cm
          bottom: convertInchesToTwip(0.98), // 2.5cm
          left: convertInchesToTwip(1.57),   // 4cm
          right: convertInchesToTwip(0.98),  // 2.5cm
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
      ...fechamento,
      ...documentos,
    ],
  }],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync('peticoes/simone-sepeda/peticao-inicial-revisional-midway.docx', buffer);
console.log('Arquivo gerado: peticoes/simone-sepeda/peticao-inicial-revisional-midway.docx');
