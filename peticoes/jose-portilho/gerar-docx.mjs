import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, Table, TableRow, TableCell,
  WidthType, convertInchesToTwip, BorderStyle
} from 'docx';
import { writeFileSync } from 'fs';

const FONTE = 'Times New Roman';
const TAM = 24; // 12pt em half-points
const pt = (n) => n * 20;

const par = (children, opts = {}) => new Paragraph({ children, spacing: { after: pt(6) }, ...opts });
const txt = (t, opts = {}) => new TextRun({ text: t, font: FONTE, size: TAM, ...opts });
const bold = (t, opts = {}) => txt(t, { bold: true, ...opts });
const italic = (t, opts = {}) => txt(t, { italics: true, ...opts });
const espaco = () => new Paragraph({ children: [txt('')], spacing: { after: 0 } });

const h1 = (t) => par([bold(t)], {
  heading: HeadingLevel.HEADING_1,
  spacing: { before: pt(12), after: pt(6) }
});

const h2 = (t) => par([bold(t)], {
  heading: HeadingLevel.HEADING_2,
  spacing: { before: pt(8), after: pt(4) }
});

const h3 = (t) => par([bold(t)], {
  heading: HeadingLevel.HEADING_3,
  spacing: { before: pt(6), after: pt(4) }
});

const j = (children, opts = {}) =>
  par(children, { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(10) }, ...opts });

const citacao = (t) =>
  par([italic(t)], {
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: pt(6), after: pt(10) },
    indent: { left: convertInchesToTwip(0.6), right: convertInchesToTwip(0.6) }
  });

// ─── TABELA FINANCEIRA ───────────────────────────────────────────────
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const linBorder = { style: BorderStyle.SINGLE, size: 4, color: '000000' };

const celTit = (t) => new TableCell({
  children: [new Paragraph({ children: [bold(t, { size: 22 })], alignment: AlignmentType.LEFT, spacing: { before: pt(2), after: pt(2) } })],
  margins: { top: pt(1), bottom: pt(1), left: pt(4), right: pt(4) },
  shading: { fill: 'DDDDDD' },
  borders: { top: linBorder, bottom: linBorder, left: linBorder, right: linBorder }
});

const celVal = (t, negrito = false) => new TableCell({
  children: [new Paragraph({ children: [negrito ? bold(t, { size: 22 }) : txt(t, { size: 22 })], alignment: AlignmentType.RIGHT, spacing: { before: pt(2), after: pt(2) } })],
  margins: { top: pt(1), bottom: pt(1), left: pt(4), right: pt(4) },
  borders: { top: linBorder, bottom: linBorder, left: linBorder, right: linBorder }
});

const linha = (label, valor, destaque = false) => new TableRow({
  children: [celTit(label), celVal(valor, destaque)]
});

const tabelaFinanceira = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    linha('Valor do financiamento (crédito líquido liberado)', 'R$ 128.668,00'),
    linha('IOF financiado', 'R$ 4.856,45'),
    linha('Tarifa de Cadastro financiada', 'R$ 23.481,91'),
    linha('VALOR TOTAL FINANCIADO (capital + encargos)', 'R$ 157.006,36', true),
    linha('Número de parcelas', '60'),
    linha('Valor de cada parcela', 'R$ 4.709,60'),
    linha('TOTAL A SER PAGO PELO AUTOR', 'R$ 282.576,00', true),
    linha('Taxa nominal mensal', '2,01%'),
    linha('Taxa nominal anual', '26,95%'),
    linha('CET mensal', '2,77%'),
    linha('CET ANUAL', '38,75%', true),
  ]
});

// ─── CONTEÚDO ────────────────────────────────────────────────────────

const cabecalho = par(
  [bold('EXCELENTÍSSIMO SENHOR DOUTOR JUIZ DE DIREITO DA ___ VARA CÍVEL DA COMARCA DE MACAPÁ — ESTADO DO AMAPÁ')],
  { alignment: AlignmentType.CENTER, spacing: { before: pt(24), after: pt(24) } }
);

const titulo = par(
  [bold('AÇÃO REVISIONAL DE CONTRATO BANCÁRIO COM TUTELA DE URGÊNCIA DE NATUREZA INIBITÓRIA', { size: 26 })],
  { alignment: AlignmentType.CENTER, spacing: { before: pt(6), after: pt(18) } }
);

const qualificacao = [
  j([
    bold('JOSE ANTONIO PORTILHO DUARTE'),
    txt(', brasileiro, nascido em 08/02/1975, portador do CPF nº 400.822.532-34, residente e domiciliado na Avenida dos Aimorés, nº 67, Bairro Beirol, Macapá-AP, CEP 68.902-140, e-mail: presidentetony@gmail.com, por meio de sua advogada que esta subscreve (procuração em anexo — Documento 01), vem, respeitosamente, à presença de Vossa Excelência, com fundamento nos artigos 300, 303, 497, parágrafo único, e 536 do Código de Processo Civil, nos artigos 6º, VI, 51, IV, e 101, I, do Código de Defesa do Consumidor, e no artigo 5º, LIV e LV, da Constituição Federal, propor'),
  ]),
  j([txt('em face de:')]),
  j([
    bold('1. BMP MONEY PLUS SOCIEDADE DE CRÉDITO DIRETO S.A.'),
    txt(', pessoa jurídica de direito privado, inscrita no CNPJ/MF nº 34.337.707/0001-00, com sede na Avenida Paulista, nº 1.765, 1º andar, Bela Vista, São Paulo-SP, CEP 01.311-200, doravante denominada '),
    bold('"BMP"'),
    txt(' ou '),
    bold('"1ª Ré"'),
    txt('; e'),
  ]),
  j([
    bold('2. SOLFÁCIL ENERGIA SOLAR TECNOLOGIA & SERVIÇOS FINANCEIROS LTDA.'),
    txt(', pessoa jurídica de direito privado, inscrita no CNPJ/MF nº 31.931.053/0001-50, com sede na Rua Cardeal Arcoverde, nº 2.450, Pinheiros, São Paulo-SP, CEP 05.408-003, doravante denominada '),
    bold('"SOLFÁCIL"'),
    txt(' ou '),
    bold('"2ª Ré"'),
    txt(';'),
  ]),
  j([txt('pelos fatos e fundamentos jurídicos que passa a expor.')]),
];

// ── SEÇÃO I ──
const secaoI = [
  h1('I — DOS FATOS'),

  j([
    bold('1. '),
    txt('Em 23 de fevereiro de 2024, o Autor firmou com as Rés uma Cédula de Crédito Bancário — CCB nº f488962a-692e-4656-8dff-6926fe7df5d5, por meio eletrônico via plataforma Clicksign, para aquisição e instalação de sistema fotovoltaico em sua residência sita à Avenida dos Aimorés, nº 67, Macapá-AP (Documento 02 — cópia integral da CCB).'),
  ]),

  j([
    bold('2. '),
    txt('A operação envolve as seguintes condições financeiras, conforme constam da própria CCB:'),
  ]),

  new Paragraph({ children: [], spacing: { after: pt(4) } }),
  tabelaFinanceira,
  new Paragraph({ children: [], spacing: { after: pt(10) } }),

  j([
    bold('3. '),
    txt('O crédito foi integralmente destinado à empresa '),
    bold('Grupo Cristal Solar Ltda.'),
    txt(', CNPJ nº 08.645.363/0001-96, fornecedora indicada pela própria Solfácil, para aquisição e instalação de sistema fotovoltaico de '),
    bold('55,55 kWp'),
    txt(' de potência (módulos JA Solar JAM72D30-550/MB e inversor Growatt MAC 50KTL3-X LV), o qual foi devidamente instalado no imóvel do Autor.'),
  ]),

  j([
    bold('4. '),
    txt('Junto ao sistema fotovoltaico, foi instalado o dispositivo eletrônico denominado '),
    bold('"Ampera"'),
    txt(', fabricado pela Solfácil ou empresa do seu grupo societário, e alienado fiduciariamente à 1ª Ré (BMP). Trata-se de dispositivo de monitoramento remoto que, nos termos da própria CCB, '),
    bold('permite o desligamento remoto dos equipamentos em caso de inadimplência'),
    txt(', conforme expressamente previsto na cláusula transcrita abaixo:'),
  ]),

  citacao('"A Credora ou terceiro por ela contratado poderá desligar os Equipamentos remotamente, fazendo com que eles parem de produzir energia fotovoltaica, de modo que Você passará a consumir integralmente a energia elétrica fornecida pela concessionária local."'),

  j([
    bold('5. '),
    txt('O contrato originalmente estipulava 60 parcelas mensais de R$ 4.709,60, com vencimento a partir de 23/05/2024. Diante das condições onerosas impostas pelas Rés — CET de 38,75% ao ano e encargos que elevam o total pago a R$ 282.576,00, mais que o dobro do valor liberado —, o Autor buscou '),
    bold('renegociação'),
    txt(' junto à credora, encontrando-se atualmente adimplente nos termos do acordo celebrado.'),
  ]),

  j([
    bold('6. '),
    txt('Não obstante a adimplência atual, o dispositivo eletrônico "Ampera" '),
    bold('permanece instalado'),
    txt(' no sistema fotovoltaico do Autor e '),
    bold('operacional'),
    txt(' — apto a executar o desligamento remoto a qualquer momento em que as Rés, a seu exclusivo critério, assim decidirem, independentemente de qualquer provimento judicial.'),
  ]),

  j([
    bold('7. '),
    txt('O sistema fotovoltaico está em pleno funcionamento, gerando energia e produzindo a compensação prevista perante a concessionária local. A eventual ativação do Ampera pelas Rés cessaria imediatamente toda a geração, causando ao Autor '),
    bold('prejuízo diário e contínuo'),
    txt(' — sem que haja qualquer mecanismo legal que obrigue as Rés a prévia notificação ou autorização judicial para tanto.'),
  ]),

  j([
    bold('8. '),
    txt('É exatamente para prevenir esse ato ilícito iminente — consubstanciado em cláusula contratual nula que permanece ativa — que o Autor busca a tutela jurisdicional, antes que o dano se consume e torne inútil qualquer providência posterior.'),
  ]),
];

// ── SEÇÃO II ──
const secaoII = [
  h1('II — DO DIREITO'),

  h2('2.1. Da Aplicabilidade do Código de Defesa do Consumidor'),

  j([
    bold('9. '),
    txt('O Autor é pessoa física que contratou financiamento bancário como destinatário final, caracterizando relação de consumo nos termos do artigo 2º do CDC. As Rés são fornecedoras de serviços financeiros, subsumindo-se ao artigo 3º, § 2º, do CDC.'),
  ]),

  j([
    bold('10. '),
    txt('A aplicação do CDC às instituições financeiras é pacífica no Superior Tribunal de Justiça, conforme a '),
    bold('Súmula nº 297'),
    txt(': '),
    italic('"O Código de Defesa do Consumidor é aplicável às instituições financeiras."'),
  ]),

  j([
    bold('11. '),
    txt('A competência territorial desta Comarca de Macapá decorre do artigo 101, I, do CDC, que assegura ao consumidor o direito de propor ação no foro de seu domicílio. A cláusula contratual que elege o foro de São Paulo-SP (item 13 da CCB) é '),
    bold('ineficaz'),
    txt(', por colocar o Autor em manifesta desvantagem de acesso à justiça (CDC, art. 51, IV e § 1º, II; STJ, Súmula 381).'),
  ]),

  h2('2.2. Da Tutela de Urgência de Natureza Inibitória — Requisitos do CPC art. 300'),

  j([
    bold('12. '),
    txt('O artigo 300 do CPC autoriza a concessão de tutela de urgência quando presentes '),
    bold('probabilidade do direito'),
    txt(' e '),
    bold('perigo de dano ou de risco ao resultado útil do processo'),
    txt('. Ambos os requisitos estão manifestamente presentes na espécie.'),
  ]),

  h2('2.3. Da Probabilidade do Direito — Fumus Boni Iuris'),

  h3('2.3.1. A Cláusula de Desligamento Remoto como Autotutela Ilegal'),

  j([
    bold('13. '),
    txt('A cláusula que autoriza o desligamento remoto do sistema fotovoltaico configura '),
    bold('autotutela privada'),
    txt(' não autorizada pelo ordenamento jurídico brasileiro, em manifesta violação ao:'),
  ]),

  j([
    bold('a) Artigo 5º, LIV, da Constituição Federal'),
    txt(': '),
    italic('"ninguém será privado da liberdade ou de seus bens sem o devido processo legal"'),
    txt(' — o sistema fotovoltaico constitui bem do Autor e sua neutralização operacional é equiparável à privação de uso do bem sem processo judicial;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('b) Artigo 5º, XXXV, da Constituição Federal'),
    txt(': a lei não excluirá da apreciação do Poder Judiciário lesão ou ameaça a direito — a cláusula permite que a credora aplique sanção patrimonial grave sem qualquer intervenção jurisdicional;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('c) Decreto-Lei nº 911/1969, artigo 3º'),
    txt(': a lei que regula a alienação fiduciária de bens móveis exige o '),
    bold('ajuizamento de ação de busca e apreensão'),
    txt(' para retomada do bem — não admite o desligamento remoto como substitutivo da via judicial.'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('14. '),
    txt('Ademais, a '),
    bold('Lei nº 14.300/2022'),
    txt(' — Marco Legal da Micro e Minigeração Distribuída — reconhece direitos do consumidor titular de sistema de geração própria conectado à rede. O desligamento remoto interfere em relação jurídica regulada perante a ANEEL, sem que haja autorização legal ou regulatória para tanto.'),
  ]),

  h3('2.3.2. Nulidade da Cláusula pelo CDC'),

  j([
    bold('15. '),
    txt('A cláusula de desligamento remoto é '),
    bold('nula de pleno direito'),
    txt(' por múltiplas razões:'),
  ]),

  j([
    bold('a) CDC, art. 51, IV'),
    txt(' — é nula a cláusula que '),
    italic('"estabeleça obrigações consideradas iníquas, abusivas, que coloquem o consumidor em desvantagem exagerada"'),
    txt(': a cláusula autoriza sanção grave sem contraditório, sem notificação judicial e sem qualquer proporcionalidade;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('b) CDC, art. 51, VIII'),
    txt(' — é nula a cláusula que imponha representante para concluir negócio jurídico pelo consumidor: a CCB outorga procuração irrevogável à Credora para praticar, em nome do Autor, atos de execução — autocontrato vedado pelo art. 117 do Código Civil;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('c) CDC, art. 54, § 4º'),
    txt(' — em contratos de adesão, cláusulas limitativas de direito do consumidor devem ser redigidas com destaque: a cláusula de desligamento remoto está inserida sem qualquer destaque formal compatível com sua gravidade.'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('16. '),
    txt('A abusividade é reforçada pela '),
    bold('dupla penalidade'),
    txt(' que a cláusula impõe: ao ser desligado, o Autor perde a geração de energia e '),
    bold('continua devendo as parcelas mensais'),
    txt(' — arcando simultaneamente com a dívida do financiamento e com a conta integral da concessionária, sem qualquer proporcionalidade.'),
  ]),

  h3('2.3.3. Do Fumus Boni Iuris nas Demais Teses Revisionais'),

  j([
    bold('17. '),
    txt('O mérito da ação sustenta-se também nas seguintes ilegalidades contratuais:'),
  ]),

  j([
    bold('(i) Tarifa de Cadastro de R$ 23.481,91'),
    txt(' — equivalente a 18,25% do valor líquido liberado, desproporcional e potencialmente inaplicável (STJ, Tema 618); embutida no principal, gera cobrança de juros sobre encargo acessório;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('(ii) IOF financiado'),
    txt(' — o IOF de R$ 4.856,45 foi incorporado ao principal, sobre o qual incidem juros por 60 meses, configurando cobrança de juros sobre tributo — prática vedada por jurisprudência consolidada do STJ;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('(iii) Disparidade entre taxa nominal e CET'),
    txt(' — a taxa nominal é de 2,01% a.m. (26,95% a.a.), mas o CET é de 2,77% a.m. (38,75% a.a.), diferença de 11,80 pontos percentuais anuais que encobre encargos não discriminados (Resolução CMN nº 3.517/2007);'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('(iv) Cláusula de irresponsabilidade pelo inadimplemento do fornecedor'),
    txt(' — o item 1(e) da CCB isenta as Rés de qualquer responsabilidade pelo inadimplemento do fornecedor, violando a responsabilidade solidária na cadeia de consumo (CDC, art. 34).'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  h2('2.4. Da Tutela Inibitória — Independência em Relação ao Dano'),

  j([
    bold('18. '),
    txt('A presente tutela é de '),
    bold('natureza inibitória'),
    txt(', voltada à prevenção de ato ilícito ainda não consumado, o que afasta a necessidade de demonstração de dano efetivo ou de sua extensão. É o que dispõe expressamente o art. 497, parágrafo único, do CPC:'),
  ]),

  citacao('"Para a tutela inibitória e a tutela de remoção de ilícito, é irrelevante a demonstração da ocorrência de dano ou de sua extensão."'),

  j([
    bold('19. '),
    txt('Assim, basta a demonstração da '),
    bold('ameaça concreta e atual do ato ilícito'),
    txt(' — presente na própria literalidade da cláusula contratual impugnada, que confere às Rés a prerrogativa de desligar remotamente o sistema a qualquer momento, a seu exclusivo critério, sem processo judicial.'),
  ]),

  j([
    bold('20. '),
    txt('O perigo de frustração do resultado útil do processo (CPC, art. 300) decorre de fatores objetivos:'),
  ]),

  j([
    bold('a) Operacionalidade imediata do Ampera'),
    txt(': o dispositivo está instalado e ativo no sistema fotovoltaico do Autor — não há qualquer barreira técnica ou jurídica, fora esta ação, que impeça as Rés de o acionarem;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('b) Discricionariedade irrestrita das Rés'),
    txt(': a própria CCB (item 5) lista como hipóteses de vencimento antecipado situações amplíssimas, como '),
    italic('"mudança adversa na condição financeira, legal ou reputacional"'),
    txt(' do devedor — redação vaga que confere às Rés poder de acionar o desligamento mesmo sem inadimplência técnica;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('c) Dano de difícil reparação após consumado'),
    txt(': desligado o sistema, o Autor arcará com a conta integral da concessionária enquanto continua devendo as parcelas — duplo encargo de apuração complexa e restituição incerta;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('d) Inutilidade da tutela post factum'),
    txt(': a religação posterior não reparará o dano já acumulado — tornando inútil qualquer medida reparatória e frustrando o resultado prático do processo (CPC, art. 300, in fine).'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  h2('2.5. Da Adequação da Astreinte'),

  j([
    bold('21. '),
    txt('Para efetividade da tutela, requer-se a cominação de '),
    bold('multa diária (astreinte)'),
    txt(' em caso de descumprimento, nos termos do art. 536, § 1º, do CPC. O valor de R$ 1.000,00 por dia mostra-se adequado à capacidade econômica das Rés e à gravidade do ato vedado, podendo ser majorado pelo Juízo se verificado descumprimento reiterado.'),
  ]),
];

// ── SEÇÃO III ──
const secaoIII = [
  h1('III — DOS PEDIDOS'),

  par([bold('O Autor requer:')], { spacing: { after: pt(8) } }),

  h2('A) TUTELA DE URGÊNCIA INIBITÓRIA LIMINAR — INAUDITA ALTERA PARTE (CPC, arts. 300, 497, parágrafo único, e 536)'),

  j([
    txt('Que Vossa Excelência, '),
    bold('inaudita altera parte'),
    txt(', conceda '),
    bold('tutela de urgência de natureza inibitória'),
    txt(' — sendo '),
    bold('irrelevante a demonstração de dano efetivo ou de sua extensão'),
    txt(' (CPC, art. 497, parágrafo único) —, determinando:'),
  ]),

  j([
    bold('(i)'),
    txt(' que as Rés '),
    bold('BMP MONEY PLUS SOCIEDADE DE CRÉDITO DIRETO S.A.'),
    txt(' e '),
    bold('SOLFÁCIL ENERGIA SOLAR TECNOLOGIA & SERVIÇOS FINANCEIROS LTDA.'),
    txt(' se abstenham, de imediato e durante todo o trâmite desta ação, de proceder ao '),
    bold('desligamento remoto do sistema fotovoltaico'),
    txt(' instalado na Avenida dos Aimorés, nº 67, Bairro Beirol, Macapá-AP, CEP 68.902-140, seja por meio do dispositivo "Ampera", seja por qualquer outro meio tecnológico, eletrônico ou físico;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('(ii)'),
    txt(' que as Rés '),
    bold('mantenham o sistema fotovoltaico em pleno funcionamento'),
    txt(', sem qualquer intervenção que comprometa sua capacidade de geração de energia fotovoltaica, durante o trâmite desta ação;'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  j([
    bold('(iii)'),
    txt(' a imposição de '),
    bold('multa diária (astreinte) de R$ 1.000,00 (mil reais) por dia de descumprimento'),
    txt(', podendo ser majorada por Vossa Excelência se verificado descumprimento reiterado, nos termos do art. 537 do CPC, sem prejuízo da responsabilidade civil por perdas e danos.'),
  ], { indent: { left: convertInchesToTwip(0.4) } }),

  h2('B) NO MÉRITO'),

  j([txt('Ao final, requer que seja:')]),

  j([bold('1. '), txt('Declarada a '), bold('nulidade de pleno direito'), txt(' da cláusula contratual que autoriza o desligamento remoto do sistema fotovoltaico, com fundamento no art. 51, IV e VIII, do CDC;')]),
  j([bold('2. '), txt('Declarada a '), bold('nulidade da cláusula de eleição de foro'), txt(' em São Paulo/SP, por violação ao art. 101, I, do CDC;')]),
  j([bold('3. '), txt('Declarada a '), bold('nulidade da cláusula de irresponsabilidade'), txt(' das Rés pelo inadimplemento do fornecedor (item 1, alínea "e", da CCB), com reconhecimento da responsabilidade solidária na cadeia de fornecimento (CDC, art. 34);')]),
  j([bold('4. '), txt('Reconhecida a '), bold('ilegalidade da Tarifa de Cadastro'), txt(' (R$ 23.481,91), com sua exclusão do contrato e restituição ao Autor dos valores pagos a esse título, atualizados pelo IPCA desde cada pagamento, acrescidos de juros de 1% ao mês (Súmula 54 do STJ);')]),
  j([bold('5. '), txt('Reconhecida a '), bold('ilegalidade da cobrança de juros sobre IOF financiado'), txt(', com recálculo do saldo devedor e restituição dos valores pagos a maior;')]),
  j([bold('6. '), txt('Determinado o '), bold('recálculo do contrato'), txt(' com expurgo das cobranças ilegais e reapresentação de novo plano de amortização com as condições saneadas, com reflexo nas parcelas vincendas;')]),
  j([bold('7. '), txt('Condenadas as Rés ao pagamento de '), bold('danos morais'), txt(', em valor a ser arbitrado por Vossa Excelência, pelo constrangimento e pela ameaça permanente a que o Autor foi submetido pela existência da cláusula de desligamento remoto;')]),
  j([bold('8. '), txt('Condenadas as Rés ao pagamento das '), bold('custas processuais e honorários advocatícios'), txt(' nos termos do art. 85 do CPC.')]),

  h2('C) DA GRATUIDADE DE JUSTIÇA'),

  j([
    txt('Requer a concessão dos benefícios da '),
    bold('gratuidade de justiça'),
    txt(', nos termos do art. 98 do CPC e art. 5º, LXXIV, da Constituição Federal, com fundamento na declaração de hipossuficiência em anexo (Documento 03).'),
  ]),
];

// ── SEÇÃO IV ──
const secaoIV = [
  h1('IV — DAS PROVAS'),
  j([txt('O Autor protesta provar o alegado por todos os meios de prova em direito admitidos, especialmente:')]),
  j([txt('— Prova documental (CCB e seus anexos — Documentos 01 a 04);')], { indent: { left: convertInchesToTwip(0.3) } }),
  j([txt('— Prova pericial técnica, para verificação do dispositivo "Ampera" instalado e sua capacidade de desligamento remoto;')], { indent: { left: convertInchesToTwip(0.3) } }),
  j([txt('— Prova testemunhal;')], { indent: { left: convertInchesToTwip(0.3) } }),
  j([txt('— Depoimento pessoal dos representantes das Rés.')], { indent: { left: convertInchesToTwip(0.3) } }),
];

// ── SEÇÃO V ──
const secaoV = [
  h1('V — DO VALOR DA CAUSA'),
  j([
    txt('Atribui-se à causa o valor de '),
    bold('R$ 157.006,36 (cento e cinquenta e sete mil, seis reais e trinta e seis centavos)'),
    txt(', correspondente ao valor total do contrato objeto da presente ação, nos termos do art. 292, II, do CPC.'),
  ]),
];

// ── FECHAMENTO ──
const fechamento = [
  espaco(),
  par([txt('Nestes termos,')], { alignment: AlignmentType.JUSTIFIED }),
  par([txt('Pede deferimento.')], { alignment: AlignmentType.JUSTIFIED }),
  espaco(),
  par([txt('Macapá-AP, 22 de junho de 2026.')], { alignment: AlignmentType.JUSTIFIED, spacing: { after: pt(60) } }),
  par([txt('_________________________________________________')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([bold('CONCEIÇÃO MARIA DUARTE PORTILHO')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([txt('Advogada — OAB/AP 3.576')], { alignment: AlignmentType.CENTER, spacing: { after: pt(2) } }),
  par([txt('conceicaoduarte.adv@gmail.com')], { alignment: AlignmentType.CENTER }),
];

// ── DOCUMENTOS ──
const documentos = [
  espaco(),
  h1('DOCUMENTOS QUE ACOMPANHAM ESTA PETIÇÃO'),
  par([txt('☐  Documento 01 — Procuração ad judicia et extra')], { spacing: { after: pt(4) } }),
  par([txt('☐  Documento 02 — Cópia integral da CCB nº f488962a-692e-4656-8dff-6926fe7df5d5 (Clicksign)')], { spacing: { after: pt(4) } }),
  par([txt('☐  Documento 03 — Declaração de hipossuficiência / gratuidade de justiça')], { spacing: { after: pt(4) } }),
  par([txt('☐  Documento 04 — Comprovante de residência do Autor em Macapá-AP')], { spacing: { after: pt(4) } }),
];

// ─── DOCUMENTO FINAL ─────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONTE, size: TAM } },
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal',
        run: { bold: true, font: FONTE, size: TAM, color: '000000' },
        paragraph: { spacing: { before: pt(12), after: pt(6) } },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal',
        run: { bold: true, font: FONTE, size: TAM, color: '000000' },
        paragraph: { spacing: { before: pt(8), after: pt(4) } },
      },
      {
        id: 'Heading3', name: 'Heading 3', basedOn: 'Normal',
        run: { bold: true, font: FONTE, size: TAM, color: '000000' },
        paragraph: { spacing: { before: pt(6), after: pt(4) } },
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
      titulo,
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
const saida = 'peticoes/jose-portilho/tutela-urgencia-desligamento-remoto-solfacil-bmp-2026-06-22.docx';
writeFileSync(saida, buffer);
console.log('✔  Arquivo gerado:', saida);
