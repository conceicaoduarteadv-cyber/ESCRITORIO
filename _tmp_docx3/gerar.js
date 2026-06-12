const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle, ShadingType, Table, TableRow, TableCell, WidthType,
  Header, Footer, PageNumber, UnderlineType
} = require("docx");
const fs = require("fs");

const corAzul = "1F3864";

function p(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 24, font: "Times New Roman", ...opts })],
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { after: 200, line: 360 },
    indent: opts.indent ? { firstLine: 720 } : undefined,
  });
}

function pCenter(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 24, font: "Times New Roman", ...opts })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 160 },
  });
}

function subtitulo(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 24, font: "Times New Roman" })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 200 },
  });
}

function secao(num, text) {
  return new Paragraph({
    children: [new TextRun({ text: `${num} — ${text}`, bold: true, size: 24, font: "Times New Roman", allCaps: true })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" } },
  });
}

function linha() {
  return new Paragraph({
    text: "",
    spacing: { after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" } },
  });
}

function espaco(n = 200) {
  return new Paragraph({ text: "", spacing: { after: n } });
}

function citacao(text) {
  return new Paragraph({
    children: [new TextRun({ text, italics: true, size: 22, font: "Times New Roman" })],
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 120, after: 120 },
    indent: { left: 720, right: 720 },
  });
}

function pedido(letra, texto) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${letra}) `, bold: true, size: 24, font: "Times New Roman" }),
      new TextRun({ text: texto, size: 24, font: "Times New Roman" }),
    ],
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 200 },
    indent: { left: 360 },
  });
}

function subitem(letra, texto) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${letra}) `, bold: true, size: 24, font: "Times New Roman" }),
      new TextRun({ text: texto, size: 24, font: "Times New Roman" }),
    ],
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 160 },
    indent: { left: 360 },
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Times New Roman", size: 24 } } },
  },
  sections: [
    // =================== PETIÇÃO ===================
    {
      properties: {
        page: { margin: { top: 1701, bottom: 1134, left: 1701, right: 1134 } },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [new TextRun({ text: "Conceição Duarte Advocacia  |  OAB/AP nº 3.576  |  Processo 0006649-25.2024.8.26.0071", size: 16, color: "808080" })],
            alignment: AlignmentType.RIGHT,
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: "conceicaoduarte.adv@gmail.com  |  (96) 9 8138-1828  |  Pág. ", size: 16, color: "808080" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "808080" }),
            ],
            alignment: AlignmentType.CENTER,
          })],
        }),
      },
      children: [
        // Endereçamento
        pCenter("EXCELENTÍSSIMO SENHOR DOUTOR JUIZ DE DIREITO DA", { bold: true }),
        pCenter("3ª VARA DA FAMÍLIA E SUCESSÕES DA COMARCA DE BAURU", { bold: true }),
        pCenter("ESTADO DE SÃO PAULO", { bold: true }),
        espaco(),
        pCenter("Processo nº 0006649-25.2024.8.26.0071"),
        espaco(),
        linha(),
        espaco(),

        // Qualificação
        new Paragraph({
          children: [
            new TextRun({ text: "RAFAEL DA SILVA DUARTE", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: ", brasileiro, solteiro, músico, inscrito no CPF nº ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "907.963.903-63", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: ", portador do RG nº 256542414-06, residente e domiciliado na Avenida Aimorés, nº 67, A, Beirol, Macapá/AP, CEP 68.902-140, por meio de sua advogada que esta subscreve (procuração em anexo), vem, respeitosamente, à presença de Vossa Excelência, nos autos da execução de alimentos em epígrafe, com fundamento no ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "art. 528, §6º, do CPC", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: ", na ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "Súmula 309 do STJ", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: " e nos arts. 2º e 3º da ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "Lei nº 12.764/2012", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: ", requerer:", size: 24, font: "Times New Roman" }),
          ],
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 240, line: 360 },
        }),

        // Título
        new Paragraph({
          children: [new TextRun({ text: "REVOGAÇÃO DO MANDADO DE PRISÃO CIVIL", bold: true, size: 28, font: "Times New Roman", allCaps: true })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 400, after: 120 },
        }),
        subtitulo("Nº 0006649-25.2024.8.26.0071.01.0005-27"),
        new Paragraph({
          children: [new TextRun({ text: "e, subsidiariamente, a suspensão imediata do cumprimento e a adoção de medidas coercitivas alternativas compatíveis com a condição de saúde do Executado", italics: true, size: 24, font: "Times New Roman" })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),
        linha(),

        // I — Fatos
        secao("I", "DOS FATOS"),
        p("Em 02 de junho de 2026, Vossa Excelência decretou a prisão civil do Executado pelo prazo de 1 (um) mês, em regime fechado, nos termos do art. 528, §3º, do CPC, ante o inadimplemento de obrigação alimentar no montante de R$ 19.043,54 (dezenove mil e quarenta e três reais e cinquenta e quatro centavos), apurado em 18/05/2026.", { indent: true }),
        p("Ciente da grave situação, o Executado por meio de sua advogada vem comprovar o pagamento das 3 (três) últimas prestações alimentícias vencidas — que constituem, por força da Súmula 309 do STJ, o único fundamento juridicamente válido para a prisão civil — totalizando R$ 3.802,19 (três mil oitocentos e dois reais e dezenove centavos), correspondente a 3 (três) prestações mensais de R$ 1.267,40 cada, conforme comprovante em anexo.", { indent: true }),
        p("Fato de extrema relevância que impõe a atenção de Vossa Excelência: o Executado é portador de Transtorno do Espectro Autista — TEA Grau 2, condição que exige suporte substancial para o pleno exercício das atividades cotidianas, tornando o cumprimento de prisão em regime fechado — com separação dos presos comuns, conforme determinado no próprio mandado — inadequado, desproporcional e atentatório à dignidade da pessoa humana.", { indent: true }),

        // II — Direito
        secao("II", "DO DIREITO"),

        new Paragraph({
          children: [new TextRun({ text: "2.1  O pagamento das 3 últimas prestações obsta a prisão civil — CPC art. 528, §6º + Súmula 309/STJ", bold: true, underline: { type: UnderlineType.SINGLE }, size: 24, font: "Times New Roman" })],
          spacing: { before: 200, after: 160 },
        }),
        p("O art. 528, §6º, do Código de Processo Civil é imperativo:"),
        citacao('"Paga a prestação alimentícia, o Juiz suspenderá o cumprimento da ordem de prisão."'),
        p("O dispositivo não abre margem à interpretação: comprovado o pagamento, a suspensão é obrigatória e imediata.", { indent: true }),
        p("A Súmula 309 do Superior Tribunal de Justiça delimita o âmbito da prisão civil:"),
        citacao('"O débito alimentar que autoriza a prisão civil do alimentante é o que compreende as três prestações anteriores ao ajuizamento da execução e as que se vencerem no curso do processo."'),
        p("O saldo remanescente — relativo a prestações anteriores ao tríplice período autorizado — não legitima a privação da liberdade, devendo ser cobrado pelos meios ordinários da execução por quantia certa (CPC, arts. 824 e ss.).", { indent: true }),
        new Paragraph({
          children: [
            new TextRun({ text: "Comprovado o pagamento de ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "R$ 3.802,19", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: " (três mil oitocentos e dois reais e dezenove centavos), correspondente às 3 últimas prestações de R$ 1.267,40 cada, impõe-se a revogação imediata do mandado de prisão.", size: 24, font: "Times New Roman" }),
          ],
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 200, line: 360 },
          indent: { firstLine: 720 },
        }),

        espaco(),
        new Paragraph({
          children: [new TextRun({ text: "2.2  O Transtorno do Espectro Autista Grau 2 torna o encarceramento cruel e desproporcional", bold: true, underline: { type: UnderlineType.SINGLE }, size: 24, font: "Times New Roman" })],
          spacing: { before: 200, after: 160 },
        }),
        p("O Executado é portador de TEA Grau 2 — nível classificado internacionalmente como aquele em que o indivíduo necessita de suporte substancial, com déficits severos nas habilidades de comunicação social e padrões restritos de comportamento que causam prejuízo significativo ao funcionamento em múltiplos contextos (DSM-5 / CID-11).", { indent: true }),
        subitem("a", "Lei nº 12.764/2012 (Política Nacional de Proteção dos Direitos da Pessoa com TEA): garante, em seu art. 3º, a vida digna, a integridade física e moral e a proteção contra qualquer forma de abuso ou exploração da pessoa com TEA."),
        subitem("b", "Lei nº 13.146/2015 — Estatuto da Pessoa com Deficiência (LBI): o art. 8º veda o tratamento cruel ou degradante à pessoa com deficiência. O encarceramento em regime fechado de pessoa com TEA Grau 2, sem suporte especializado, viola diretamente esse comando legal."),
        subitem("c", "Constituição Federal, art. 5º, XLIX: é assegurado ao preso o respeito à integridade física e moral. A unidade prisional ordinária não possui estrutura para atender às necessidades específicas de pessoa com TEA Grau 2."),
        subitem("d", "Natureza coercitiva da prisão civil: a prisão por alimentos visa compelir ao cumprimento da obrigação — não punir. Quando a medida se torna desproporcional, o art. 139, IV, do CPC autoriza medidas alternativas (protesto, suspensão de CNH, negativação) com o mesmo efeito coercitivo e sem risco à saúde do Executado."),

        espaco(),
        new Paragraph({
          children: [new TextRun({ text: "2.3  Da limitação do débito exequível pela prisão civil", bold: true, underline: { type: UnderlineType.SINGLE }, size: 24, font: "Times New Roman" })],
          spacing: { before: 200, after: 160 },
        }),
        p("O valor de R$ 19.043,54 anotado no mandado representa débito alimentar acumulado ao longo de vários meses. A Súmula 309 do STJ limita o fundamento da coerção pessoal às 3 últimas prestações. O saldo excedente deve ser redirecionado ao rito da execução por quantia certa, vedada sua utilização como fundamento para a privação da liberdade.", { indent: true }),

        // III — Pedidos
        secao("III", "DOS PEDIDOS"),
        p("Ante o exposto, requer o Executado a Vossa Excelência:"),
        espaco(),
        pedido("a", "PEDIDO PRINCIPAL: A revogação imediata do Mandado de Prisão Civil nº 0006649-25.2024.8.26.0071.01.0005-27, nos termos do art. 528, §6º, do CPC, ante a comprovação do pagamento das 3 (três) últimas prestações alimentícias no valor total de R$ 3.802,19, com a expedição do respectivo ALVARÁ DE SOLTURA em favor do Executado, cuja minuta segue na segunda folha desta peça;"),
        pedido("b", "PEDIDO SUBSIDIÁRIO 1: A substituição da prisão civil por medidas coercitivas alternativas previstas no art. 139, IV, do CPC — protesto do pronunciamento judicial, negativação nos órgãos de proteção ao crédito, suspensão de carteira de habilitação e/ou passaporte —, em atenção à condição de pessoa com TEA Grau 2 do Executado;"),
        pedido("c", "PEDIDO SUBSIDIÁRIO 2: A revisão do montante objeto da coerção pessoal, limitando a prisão civil às 3 últimas prestações (R$ 3.802,19), com redirecionamento do saldo restante ao rito de execução por quantia certa (CPC, arts. 824 e ss.);"),
        pedido("d", "EM QUALQUER HIPÓTESE: O reconhecimento de que o regime fechado é incompatível com a condição de saúde do Executado (TEA Grau 2), nos termos da Lei nº 12.764/2012 e do art. 8º da Lei nº 13.146/2015."),

        espaco(),
        p("Documentos anexos: (i) comprovante de pagamento de R$ 3.802,19; (ii) laudo médico de TEA Grau 2; (iii) procuração ad judicia."),
        espaco(),
        p("Termos em que pede e espera deferimento."),
        espaco(),
        pCenter("Macapá/AP, 11 de junho de 2026."),
        espaco(600),
        linha(),
        espaco(),
        pCenter("Conceição Maria Duarte Portilho", { bold: true }),
        pCenter("OAB/AP nº 3.576"),
        pCenter("Conceição Duarte Advocacia e Consultoria Jurídica"),
        pCenter("Av. Aimorés, 67 A, Beirol, Macapá/AP — CEP 68.902-140"),
        pCenter("conceicaoduarte.adv@gmail.com | (96) 9 8138-1828"),
      ],
    },

    // =================== ALVARÁ DE SOLTURA ===================
    {
      properties: {
        page: { margin: { top: 1701, bottom: 1134, left: 1701, right: 1134 } },
      },
      children: [
        espaco(),
        new Paragraph({
          children: [new TextRun({ text: "TRIBUNAL DE JUSTIÇA DO ESTADO DE SÃO PAULO", bold: true, size: 22, font: "Times New Roman", allCaps: true })],
          alignment: AlignmentType.CENTER, spacing: { after: 80 },
        }),
        new Paragraph({
          children: [new TextRun({ text: "3ª VARA DA FAMÍLIA E SUCESSÕES DA COMARCA DE BAURU", bold: true, size: 22, font: "Times New Roman", allCaps: true })],
          alignment: AlignmentType.CENTER, spacing: { after: 80 },
        }),
        new Paragraph({
          children: [new TextRun({ text: "Processo nº 0006649-25.2024.8.26.0071", size: 22, font: "Times New Roman" })],
          alignment: AlignmentType.CENTER, spacing: { after: 400 },
        }),
        linha(),
        espaco(),

        // Título do alvará
        new Paragraph({
          children: [new TextRun({ text: "ALVARÁ DE SOLTURA", bold: true, size: 40, font: "Times New Roman", allCaps: true })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 600 },
          border: {
            bottom: { style: BorderStyle.DOUBLE, size: 6, color: corAzul },
            top: { style: BorderStyle.DOUBLE, size: 6, color: corAzul },
          },
        }),
        espaco(),

        // Autoridade
        new Paragraph({
          children: [new TextRun({ text: "O(A) MERITÍSSIMO(A) SENHOR(A) DOUTOR(A) JUIZ(A) DE DIREITO DA 3ª VARA DA FAMÍLIA E SUCESSÕES DA COMARCA DE BAURU — ESTADO DE SÃO PAULO", bold: true, size: 24, font: "Times New Roman" })],
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 400, line: 360 },
        }),

        // Ordem
        new Paragraph({
          children: [
            new TextRun({ text: "MANDA ", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: "a qualquer Autoridade Policial ou agente de segurança pública que, tendo o presente em mãos, ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "SOLTE E PONHA IMEDIATAMENTE EM LIBERDADE", bold: true, size: 24, font: "Times New Roman", allCaps: true }),
            new TextRun({ text: " — salvo se por outro motivo estiver preso — a pessoa de:", size: 24, font: "Times New Roman" }),
          ],
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 400, line: 360 },
        }),

        // Caixa de identificação
        (() => {
          const cell = new TableCell({
            children: [
              new Paragraph({ children: [new TextRun({ text: "RAFAEL DA SILVA DUARTE", bold: true, size: 28, font: "Times New Roman" })], alignment: AlignmentType.CENTER, spacing: { before: 120, after: 80 } }),
              new Paragraph({ children: [new TextRun({ text: "CPF: 907.963.903-63   |   RG: 256542414-06", size: 22, font: "Times New Roman" })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
              new Paragraph({ children: [new TextRun({ text: "Nascido em: 24/09/1982   |   Músico   |   Solteiro", size: 22, font: "Times New Roman" })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
              new Paragraph({ children: [new TextRun({ text: "Av. dos Aimorés, nº 67, A, Beirol, Macapá/AP — CEP 68.902-140", size: 22, font: "Times New Roman" })], alignment: AlignmentType.CENTER, spacing: { after: 120 } }),
            ],
            shading: { type: ShadingType.CLEAR, fill: "EBF3FB" },
            margins: { top: 200, bottom: 200, left: 200, right: 200 },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: corAzul },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: corAzul },
              left: { style: BorderStyle.SINGLE, size: 4, color: corAzul },
              right: { style: BorderStyle.SINGLE, size: 4, color: corAzul },
            },
          });
          return new Table({ rows: [new TableRow({ children: [cell] })], width: { size: 100, type: WidthType.PERCENTAGE } });
        })(),
        espaco(400),

        // Contexto
        new Paragraph({
          children: [
            new TextRun({ text: "que se encontrava preso nos termos do ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "Mandado de Prisão Civil nº 0006649-25.2024.8.26.0071.01.0005-27", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: ", expedido em 02 de junho de 2026, por dívida alimentar no valor de R$ 19.043,54.", size: 24, font: "Times New Roman" }),
          ],
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 240, line: 360 },
        }),

        // Fundamento
        new Paragraph({
          children: [
            new TextRun({ text: "Fundamento: ", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: "Comprovação do pagamento das 3 (três) últimas prestações alimentícias vencidas, no valor total de ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "R$ 3.802,19 (três mil oitocentos e dois reais e dezenove centavos)", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: ", nos termos do ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "art. 528, §6º, do Código de Processo Civil", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: " e da ", size: 24, font: "Times New Roman" }),
            new TextRun({ text: "Súmula 309 do Superior Tribunal de Justiça", bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: ", dando-se por livre e revogado o mandado de prisão civil supramencionado.", size: 24, font: "Times New Roman" }),
          ],
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 800, line: 360 },
        }),

        // Local e data
        new Paragraph({
          children: [new TextRun({ text: "Bauru/SP, _____ de _________________ de 2026.", size: 24, font: "Times New Roman" })],
          alignment: AlignmentType.LEFT,
          spacing: { after: 1000 },
        }),

        // Assinatura
        linha(),
        espaco(),
        pCenter("MARIO RAMOS DOS SANTOS", { bold: true }),
        pCenter("Juiz de Direito — 3ª Vara da Família e Sucessões"),
        pCenter("Comarca de Bauru — Tribunal de Justiça do Estado de São Paulo"),
        espaco(800),

        new Paragraph({
          children: [new TextRun({ text: "Minuta elaborada pela advogada da parte para subscrição de V. Exa.  |  e-SAJ TJSP  |  Processo 0006649-25.2024.8.26.0071", size: 16, italics: true, color: "808080" })],
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" } },
          spacing: { before: 400 },
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(
    "c:/Users/CONCEIÇÃO DUARTE/teste/peticoes/revogacao-prisao-civil-rafael-duarte-2026-06-11.docx",
    buf
  );
  console.log("ok");
});
