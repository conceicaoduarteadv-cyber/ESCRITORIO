import {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, convertInchesToTwip, UnderlineType
} from 'docx';
import { writeFileSync } from 'fs';

const F = 'Times New Roman';
const pt = (n) => n * 20;
const S = 24; // 12pt

const t = (text, o = {}) => new TextRun({ text, font: F, size: S, ...o });
const b = (text, o = {}) => t(text, { bold: true, ...o });
const p = (children, opts = {}) => new Paragraph({ children, ...opts });
const esp = (n = 8) => p([t('')], { spacing: { after: pt(n) } });

const j = AlignmentType.JUSTIFIED;
const c = AlignmentType.CENTER;

const titulo = p(
  [b('CONTRATO DE PRESTAÇÃO DE SERVIÇOS ADVOCATÍCIOS', { size: 28 })],
  { alignment: c, spacing: { after: pt(20) } }
);

const partes = [
  p([
    b('CONTRATANTE: '),
    t('SIMONE MARQUES MARTINS SEPEDA, brasileira, portadora do RG nº 040825 (POLITEC/AP), inscrita no CPF sob o nº 324.863.672-68, residente e domiciliada na Rua Carlos Gomes, nº 301, Bairro Jesus de Nazaré, Macapá/AP, CEP 68908-125, doravante denominada '),
    b('CONTRATANTE'),
    t(';'),
  ], { alignment: j, spacing: { after: pt(10) } }),

  p([
    b('CONTRATADA: '),
    t('CONCEIÇÃO MARIA DUARTE PORTILHO, advogada, inscrita na Ordem dos Advogados do Brasil, Seccional do Amapá, sob o nº OAB/AP 3576, com endereço profissional na Av. Aimorés, nº 67, Beirol, Macapá/AP, CEP 68902-140, e-mail conceicaoduarte.adv@gmail.com, doravante denominada '),
    b('ADVOGADA'),
    t(';'),
  ], { alignment: j, spacing: { after: pt(10) } }),

  p([t('As partes acima qualificadas têm entre si justo e contratado o presente Contrato de Prestação de Serviços Advocatícios, que se regerá pelas cláusulas e condições a seguir:')],
    { alignment: j, spacing: { after: pt(16) } }),
];

const cl = (num, titulo_cl) => p(
  [b(`CLÁUSULA ${num} — ${titulo_cl}`)],
  { alignment: j, spacing: { before: pt(14), after: pt(8) } }
);

const sub = (text) => p([b(text)], { alignment: j, spacing: { before: pt(8), after: pt(4) } });

const corpo = (text, indent = false, opts = {}) => p(
  [t(text)],
  { alignment: j, spacing: { after: pt(6) }, indent: indent ? { left: convertInchesToTwip(0.3) } : undefined, ...opts }
);

const misto = (children, indent = false, opts = {}) => p(
  children,
  { alignment: j, spacing: { after: pt(6) }, indent: indent ? { left: convertInchesToTwip(0.3) } : undefined, ...opts }
);

const doc = new Document({
  styles: {
    default: { document: { run: { font: F, size: S } } },
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
      titulo,
      ...partes,

      // Cláusula 1
      cl('PRIMEIRA', 'DO OBJETO'),
      corpo('A ADVOGADA obriga-se a prestar os seguintes serviços advocatícios em favor da CONTRATANTE:'),
      misto([b('I. '), t('Propositura e acompanhamento de '), b('Ação Revisional de Contratos Bancários com Pedido de Tutela de Urgência'), t(' em face de '), b('MIDWAY S.A. — Crédito, Financiamento e Investimento'), t(' (CNPJ nº 09.464.032/0001-12), objetivando a revisão das seguintes Cédulas de Crédito Bancário:')], false),
      corpo('— CCB nº 75805602, emitida em 18/02/2026, no valor total de R$ 23.251,80;', true),
      corpo('— CCB nº 75825597, emitida em 09/03/2026, no valor total de R$ 29.838,00;', true),
      corpo('— CCB nº 75879070, emitida em 27/04/2026, no valor total de R$ 3.100,00.', true),
      misto([b('II. '), t('Os serviços compreendem todos os atos processuais necessários ao andamento do processo até o trânsito em julgado da sentença ou acordo homologado judicialmente, incluindo: elaboração de petições, recursos, manifestações, sustentações orais, negociações extrajudiciais e acompanhamento da fase de cumprimento de sentença.')], false),
      corpo('Parágrafo único. O objeto deste contrato limita-se às CCBs identificadas nesta cláusula. Eventual extensão dos serviços a outros contratos ou ações será objeto de aditivo escrito.'),

      // Cláusula 2
      cl('SEGUNDA', 'DOS HONORÁRIOS ADVOCATÍCIOS'),
      misto([t('Os honorários pelos serviços ora contratados serão de '), b('30% (trinta por cento) sobre o proveito econômico obtido'), t(' pela CONTRATANTE ao final do processo, na modalidade '), b('honorários de êxito'), t('.')]),

      sub('§ 1º — Definição de proveito econômico'),
      corpo('Considera-se proveito econômico para fins de cálculo dos honorários, o valor correspondente a:'),
      corpo('a) a diferença entre o total originalmente cobrado pelas CCBs e o total apurado após a revisão judicial das taxas, seguros e demais cláusulas impugnadas, calculada sobre as parcelas vincendas;', true),
      corpo('b) os valores efetivamente restituídos ou creditados à CONTRATANTE a título de parcelas pagas a maior, seguro prestamista indevido e IOF cobrado sobre a base de juros;', true),
      corpo('c) em caso de acordo extrajudicial ou judicial homologado, o valor total da vantagem patrimonial obtida pela CONTRATANTE, incluindo descontos, remissão de dívida e cancelamentos;', true),
      corpo('d) em caso de extinção ou rescisão contratual determinada judicialmente (superendividamento — Lei nº 14.181/2021), o valor remanescente da dívida extinta ou renegociada.', true),

      sub('§ 2º — Apuração do proveito econômico'),
      corpo('O proveito econômico será apurado:'),
      corpo('I. por meio do laudo pericial produzido nos autos, quando houver;', true),
      corpo('II. por planilha elaborada de comum acordo entre as partes, confrontando o total original das CCBs com o total resultante da revisão judicial;', true),
      corpo('III. nos casos de acordo, pelo valor líquido constante do instrumento homologado pelo Juízo.', true),

      sub('§ 3º — Momento do pagamento'),
      corpo('Os honorários de êxito serão devidos e exigíveis:'),
      corpo('I. após o trânsito em julgado de sentença favorável;', true),
      corpo('II. após a homologação judicial de acordo;', true),
      corpo('III. em caso de acordo extrajudicial celebrado durante o processo, no momento da assinatura do instrumento;', true),
      corpo('IV. em caso de levantamento de valores, previamente ao levantamento pela CONTRATANTE, descontando-se a quota parte da ADVOGADA diretamente nos autos.', true),

      sub('§ 4º — Honorários sucumbenciais'),
      corpo('Os honorários sucumbenciais fixados pelo Juízo (CPC, art. 85) pertencem exclusivamente à ADVOGADA e não se compensam com os honorários contratuais de êxito previstos neste instrumento, nos termos do art. 24, § 4º, do Estatuto da OAB.'),

      sub('§ 5º — Ausência de honorários fixos'),
      corpo('As partes ajustam que não haverá pagamento de honorários fixos, mensais ou por ato, sendo a remuneração da ADVOGADA exclusivamente condicionada ao êxito.'),

      // Cláusula 3
      cl('TERCEIRA', 'DAS DESPESAS PROCESSUAIS'),
      corpo('As custas judiciais, emolumentos cartorários, honorários periciais e demais despesas processuais são de responsabilidade exclusiva da CONTRATANTE.'),
      corpo('§ 1º — A CONTRATANTE declara, neste ato, ser pessoa hipossuficiente e autorizará a ADVOGADA a requerer a gratuidade da justiça nos autos. Caso a gratuidade seja indeferida, caberá à CONTRATANTE arcar com as custas correspondentes no prazo de 10 (dez) dias úteis da intimação.'),
      corpo('§ 2º — Na hipótese de sucumbência (condenação ao pagamento de honorários da parte contrária), as custas de sucumbência serão pagas pela CONTRATANTE, podendo ser descontadas do proveito econômico apurado.'),

      // Cláusula 4
      cl('QUARTA', 'DAS OBRIGAÇÕES DA ADVOGADA'),
      corpo('São obrigações da ADVOGADA:'),
      corpo('I. conduzir os serviços com diligência, competência e ética, nos termos do Código de Ética e Disciplina da OAB;', true),
      corpo('II. manter a CONTRATANTE informada sobre o andamento do processo, sempre que houver decisão relevante;', true),
      corpo('III. guardar sigilo sobre todas as informações da CONTRATANTE;', true),
      corpo('IV. tratar os dados pessoais da CONTRATANTE de acordo com a Lei nº 13.709/2018 (LGPD), utilizando-os exclusivamente para a execução deste contrato.', true),
      corpo('Parágrafo único. A ADVOGADA não garante resultado específico, sendo vedada a promessa de êxito nos termos do art. 34, inciso IV, do Estatuto da OAB e do art. 41 do Código de Ética e Disciplina da OAB.'),

      // Cláusula 5
      cl('QUINTA', 'DAS OBRIGAÇÕES DA CONTRATANTE'),
      corpo('São obrigações da CONTRATANTE:'),
      corpo('I. fornecer à ADVOGADA todos os documentos, informações e esclarecimentos necessários ao patrocínio da causa, com veracidade e completude;', true),
      corpo('II. comunicar à ADVOGADA imediatamente qualquer contato direto da parte contrária, proposta de acordo ou notificação extrajudicial recebida;', true),
      corpo('III. não celebrar acordo, transação ou desistência sem a prévia anuência da ADVOGADA;', true),
      corpo('IV. manter seus dados de contato (endereço, telefone, e-mail) atualizados junto à ADVOGADA;', true),
      corpo('V. pagar os honorários de êxito na forma e prazo estabelecidos neste contrato.', true),

      // Cláusula 6
      cl('SEXTA', 'DA RESCISÃO'),
      corpo('§ 1º — Rescisão pela CONTRATANTE: A CONTRATANTE poderá rescindir este contrato a qualquer momento, devendo pagar à ADVOGADA honorários proporcionais ao trabalho já realizado, com base na Tabela de Honorários da OAB/AP vigente na data da rescisão, sem prejuízo do reembolso das despesas já desembolsadas.'),
      corpo('§ 2º — Rescisão pela ADVOGADA: A ADVOGADA poderá rescindir o contrato nas hipóteses previstas no art. 36 do Estatuto da OAB, com comunicação prévia à CONTRATANTE em tempo hábil para que esta constitua novo patrono, sem prejuízo do direito ao recebimento de honorários proporcionais.'),
      corpo('§ 3º — Êxito pós-rescisão: Caso o processo resulte em êxito após a rescisão deste contrato, a ADVOGADA fará jus a honorários proporcionais ao trabalho realizado durante sua atuação, a serem fixados de comum acordo ou, em caso de divergência, pela Câmara de Honorários da OAB/AP.'),

      // Cláusula 7
      cl('SÉTIMA', 'DO ACORDO SEM INTERVENÇÃO DA ADVOGADA'),
      corpo('Se a CONTRATANTE celebrar acordo diretamente com a parte adversa, sem a participação da ADVOGADA, serão devidos honorários de êxito calculados sobre o proveito econômico obtido no referido acordo, nos mesmos percentuais previstos na Cláusula Segunda.'),

      // Cláusula 8
      cl('OITAVA', 'DA PROTEÇÃO DE DADOS (LGPD)'),
      corpo('Os dados pessoais da CONTRATANTE coletados para execução deste contrato serão tratados com base na hipótese legal de execução de contrato (Lei nº 13.709/2018, art. 7º, inciso V), sendo utilizados exclusivamente para a prestação dos serviços advocatícios aqui descritos e para o cumprimento de obrigações legais da ADVOGADA.'),
      corpo('A CONTRATANTE poderá, a qualquer tempo, solicitar acesso, correção ou eliminação de seus dados pessoais, nos termos do art. 18 da LGPD.'),

      // Cláusula 9
      cl('NONA', 'DAS DISPOSIÇÕES GERAIS'),
      corpo('§ 1º — Este contrato é celebrado com base nos arts. 22 a 26 da Lei nº 8.906/94 (Estatuto da OAB) e no art. 593 e seguintes do Código Civil.'),
      corpo('§ 2º — A Tabela de Honorários da OAB/AP serve como referência mínima para os honorários proporcionais devidos na hipótese de rescisão.'),
      corpo('§ 3º — Eventuais alterações neste contrato somente terão validade se formalizadas por escrito e assinadas por ambas as partes.'),
      corpo('§ 4º — A tolerância de qualquer das partes quanto ao descumprimento de cláusula deste contrato não importará novação, renúncia ou alteração do pactuado.'),

      // Cláusula 10
      cl('DÉCIMA', 'DO FORO'),
      corpo('As partes elegem o foro da Comarca de Macapá/AP para dirimir quaisquer controvérsias oriundas deste contrato, renunciando a qualquer outro, por mais privilegiado que seja.'),

      esp(40),
      p([t('Macapá/AP, _____ de __________________ de 2026.')], { alignment: j, spacing: { after: pt(60) } }),

      p([t('___________________________________________')], { alignment: c, spacing: { after: pt(2) } }),
      p([b('CONTRATANTE')], { alignment: c, spacing: { after: pt(2) } }),
      p([t('SIMONE MARQUES MARTINS SEPEDA')], { alignment: c, spacing: { after: pt(2) } }),
      p([t('CPF nº 324.863.672-68')], { alignment: c, spacing: { after: pt(40) } }),

      p([t('___________________________________________')], { alignment: c, spacing: { after: pt(2) } }),
      p([b('ADVOGADA')], { alignment: c, spacing: { after: pt(2) } }),
      p([t('CONCEIÇÃO MARIA DUARTE PORTILHO')], { alignment: c, spacing: { after: pt(2) } }),
      p([t('OAB/AP nº 3576')], { alignment: c, spacing: { after: pt(40) } }),

      p([b('TESTEMUNHAS:')], { alignment: j, spacing: { before: pt(8), after: pt(28) } }),
      p([t('___________________________________________')], { alignment: j, spacing: { after: pt(2) } }),
      p([t('Nome:')], { alignment: j, spacing: { after: pt(2) } }),
      p([t('CPF:')], { alignment: j, spacing: { after: pt(28) } }),
      p([t('___________________________________________')], { alignment: j, spacing: { after: pt(2) } }),
      p([t('Nome:')], { alignment: j, spacing: { after: pt(2) } }),
      p([t('CPF:')], { alignment: j }),
    ],
  }],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync('contratos/honorarios-simone-sepeda-midway.docx', buffer);
console.log('Arquivo gerado: contratos/honorarios-simone-sepeda-midway.docx');
