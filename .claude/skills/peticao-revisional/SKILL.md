---
name: peticao-revisional
description: >
  Prepara petição inicial de ação revisional de contrato bancário.
  Coleta os dados do cliente e do contrato, monta a peça completa com
  qualificação, fatos, fundamentos jurídicos, jurisprudência do STJ e pedidos.
  Salva em .md e gera .docx para protocolar.
  Use quando disser "faz uma revisional", "petição revisional", "ação revisional",
  "preparar petição de revisão de contrato", "cliente quer revisar contrato bancário".
---

# /peticao-revisional — Petição Revisional de Contrato Bancário

## Dependências

- `_contexto/empresa.md` — dados do escritório (advogada, OAB, endereço)
- `_contexto/preferencias.md` — estilo de escrita das peças

---

## Workflow

### Passo 1 — Coletar dados do cliente e do contrato

Perguntar em sequência, um dado por vez:

1. **Nome completo do cliente**
2. **CPF e RG**
3. **Endereço completo**
4. **Nome do banco réu**
5. **Número do contrato** (se disponível)
6. **Tipo de contrato** (empréstimo pessoal, financiamento, cartão de crédito, consignado, cheque especial...)
7. **Valor original contratado**
8. **Data de assinatura do contrato**
9. **Qual é o problema principal?** (juros abusivos, capitalização indevida, tarifa não contratada, seguro obrigatório, cobrança indevida — pode ser mais de um)
10. **Tem histórico de pagamentos ou extrato?** (pedir que coloque na pasta `dados/` se tiver)
11. **Foro** (cidade onde o cliente é domiciliado)

### Passo 2 — Identificar os fundamentos aplicáveis

Com base nos problemas informados, selecionar os fundamentos:

**Aplicação do CDC:**
- STJ, Súmula 297: "O Código de Defesa do Consumidor é aplicável às instituições financeiras."
- CDC, art. 6º, V: revisão de cláusulas abusivas
- CDC, art. 51, IV: nulidade de cláusulas que estabeleçam obrigações iníquas

**Juros abusivos:**
- STJ, REsp 1.061.530/RS (Tema 27, repetitivo): possibilidade de revisão dos juros remuneratórios
- CC, art. 422: boa-fé objetiva nos contratos
- Resolução BACEN nº 4.559/2017: transparência nas operações de crédito

**Capitalização de juros (anatocismo):**
- STJ, Súmula 539: "É permitida a capitalização de juros com periodicidade inferior à anual em contratos celebrados com instituições integrantes do Sistema Financeiro Nacional a partir de 31/3/2000."
- STJ, Súmula 541: "A previsão no contrato bancário de taxa de juros anual superior ao duodécuplo da mensal é suficiente para permitir a cobrança da taxa efetiva anual contratada."
- Se não houver previsão expressa de capitalização: Decreto nº 22.626/33 (Lei da Usura), art. 4º

**Tarifas e seguros não contratados:**
- STJ, REsp 1.251.331/RS e REsp 1.255.573/MS (Tema 618, repetitivo):
  - Tarifa de abertura de crédito (TAC) e tarifa de emissão de carnê (TEC) — válidas se contratadas
  - Tarifa de cadastro — válida apenas na abertura da relação, não em renovações
  - Seguro de proteção financeira — abusivo se contratação for compulsória (venda casada — CDC art. 39, I)
- BACEN, Resolução nº 3.919/2010: lista de tarifas permitidas

**Devolução em dobro:**
- CDC, art. 42, parágrafo único: devolução em dobro de valores cobrados indevidamente, salvo engano justificável
- STJ, Súmula 322: inaplicável a devolução em dobro quando a cobrança foi feita de boa-fé

**Dano moral (se aplicável):**
- Negativação indevida: STJ, Súmula 385
- Cobrança vexatória: CDC, art. 71

### Passo 3 — Montar a petição

Estrutura completa:

```
EXCELENTÍSSIMO(A) SENHOR(A) DOUTOR(A) JUIZ(A) DE DIREITO DA VARA CÍVEL DA COMARCA DE [FORO]

[NOME DO CLIENTE], [nacionalidade], [estado civil], [profissão],
portador(a) do CPF nº [CPF] e RG nº [RG],
residente e domiciliado(a) em [endereço completo],
por meio de seu advogado que esta subscreve,
vem, respeitosamente, à presença de Vossa Excelência,
propor a presente

AÇÃO REVISIONAL DE CONTRATO BANCÁRIO
COM PEDIDO DE TUTELA DE URGÊNCIA (se aplicável)

em face de [NOME DO BANCO], instituição financeira [qualificação do banco],
pelos fatos e fundamentos jurídicos a seguir expostos.

I — DOS FATOS
[narrativa clara e objetiva da situação do cliente, com datas e valores]

II — DO DIREITO

2.1 Da aplicabilidade do Código de Defesa do Consumidor
[Súmula 297 STJ + CDC art. 6º V]

2.2 [fundamento específico 1 — ex: Dos juros abusivos]
[fundamentação com lei + jurisprudência]

2.3 [fundamento específico 2 — se houver]
[fundamentação com lei + jurisprudência]

[demais fundamentos conforme os problemas informados]

III — DA TUTELA DE URGÊNCIA (se aplicável)
[suspensão de descontos ou negativação, com base no CPC art. 300]

IV — DOS PEDIDOS

Ante o exposto, requer:

a) a citação do réu;
b) a procedência do pedido para:
   — revisão das cláusulas abusivas identificadas;
   — devolução dos valores cobrados indevidamente [em dobro se aplicável — CDC art. 42 §ún];
   — [outros pedidos conforme o caso];
c) a condenação do réu ao pagamento das custas processuais e honorários advocatícios (CPC, art. 85);
d) a produção de todos os meios de prova admitidos em direito, em especial documental e pericial contábil.

Dá-se à causa o valor de R$ [valor].

[cidade], [data].

Conceição Maria Duarte Portilho
OAB/[estado] nº [número OAB]
```

### Passo 4 — Completar dados do escritório

Preencher automaticamente com os dados de `_contexto/empresa.md`.
Se OAB e endereço do escritório não estiverem salvos, perguntar ao usuário e sugerir salvar no arquivo.

### Passo 5 — Salvar os arquivos

1. Salvar o texto em `peticoes/revisional-[nome-cliente]-[data].md`
2. Gerar versão Word com a skill nativa `/docx`:
   - Chamar `/docx` passando o conteúdo da petição
   - Salvar como `peticoes/revisional-[nome-cliente]-[data].docx`

Confirmar ao usuário:

> "Petição salva em `peticoes/`. Versão .md e .docx prontas. Quer revisar algum trecho antes de protocolar?"

---

## Regras

- Sempre verificar quais fundamentos se aplicam ao caso — não incluir tudo de uma vez
- Linguagem formal e técnica, como toda peça processual
- Jurisprudência com número do processo ou tema quando possível
- Se o cliente não tiver todos os dados, montar o que for possível e deixar marcadores [A PREENCHER] onde faltarem informações
- Não inventar valores ou datas — usar apenas o que o usuário informou
- Se houver extrato ou contrato na pasta `dados/`, ler antes de montar a peça
