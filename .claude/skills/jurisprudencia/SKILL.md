---
name: jurisprudencia
description: >
  Pesquisa jurisprudência relevante sobre um tema no STJ, STF e TJAP.
  Retorna acórdãos e súmulas com texto formatado para citar diretamente em petições.
  Salva os melhores resultados em jurisprudencia/direito-bancario/ para consulta futura.
  Use quando disser "pesquisa jurisprudência sobre X", "acha julgado sobre X",
  "preciso de jurisprudência pra X", "STJ sobre X", "o que o STJ diz sobre X".
---

# /jurisprudencia — Pesquisa de Jurisprudência

## Dependências

- `_contexto/estrategia.md` — área de foco atual (Direito Bancário)

---

## Workflow

### Passo 1 — Identificar o tema

Se o usuário não informou o tema claramente, perguntar:

> "Qual é o tema ou o tipo de abuso que você quer pesquisar?"

Exemplos de temas recorrentes em Direito Bancário:
- Juros abusivos / revisão de taxa
- Capitalização de juros (anatocismo)
- Tarifas bancárias (TAC, TEC, cadastro, seguro)
- Negativação indevida
- Superendividamento
- Portabilidade de crédito
- Cobrança indevida / devolução em dobro
- Dano moral bancário

### Passo 2 — Pesquisar no STJ

Usar WebSearch com as seguintes buscas:

1. `site:stj.jus.br [tema] direito bancário`
2. `STJ súmula [tema] banco`
3. `STJ repetitivo [tema] instituição financeira`

Priorizar:
- Recursos repetitivos (Temas) — vinculantes para todo o judiciário
- Súmulas do STJ
- Acórdãos recentes das Turmas de Direito Privado (3ª e 4ª Turma)

Súmulas do STJ mais usadas em Direito Bancário (referência rápida):
- **Súmula 297** — CDC aplicável às instituições financeiras
- **Súmula 293** — Revisão de contratos bancários é admissível
- **Súmula 322** — Devolução simples (não em dobro) quando cobrança de boa-fé
- **Súmula 381** — Vedado ao juiz conhecer de ofício da abusividade
- **Súmula 385** — Negativação prévia afasta dano moral por nova negativação
- **Súmula 539** — Capitalização inferior à anual permitida a partir de 31/03/2000
- **Súmula 541** — Taxa anual superior ao duodécuplo da mensal permite cobrança da taxa efetiva

### Passo 3 — Pesquisar no STF (se relevante)

Usar WebSearch: `site:stf.jus.br [tema]`

Priorizar repercussão geral e ADIs com efeito vinculante.

### Passo 4 — Pesquisar no TJAP

Usar WebSearch: `site:tjap.jus.br [tema] OR "tjap.jus.br" [tema] banco`

Se não encontrar no site oficial, tentar: `TJAP "[tema]" acórdão banco`

### Passo 5 — Apresentar os resultados

Para cada resultado encontrado, apresentar no formato pronto pra citar:

```
---
TRIBUNAL: [STJ / STF / TJAP]
TIPO: [Súmula / Acórdão / Tema Repetitivo]
REFERÊNCIA: [Súmula nº X / REsp nº X.XXX.XXX/UF / Tema nº X]
RELATOR: [Nome do Ministro/Desembargador]
DATA: [data do julgamento]

EMENTA / TESE:
"[texto da ementa ou tese fixada — entre aspas, pronto pra colar]"

APLICAÇÃO: [em uma linha, como usar essa jurisprudência na petição]
---
```

Apresentar no mínimo 3 e no máximo 6 resultados por pesquisa, priorizando os mais recentes e os de maior força vinculante.

### Passo 6 — Salvar (opcional)

Perguntar ao usuário:

> "Quer que eu salve esses resultados no seu banco de jurisprudência?"

Se sim, salvar em `jurisprudencia/direito-bancario/[tema-resumido].md` no mesmo formato da apresentação.

Se o arquivo já existir, adicionar os novos resultados ao final sem apagar os anteriores.

---

## Regras

- Sempre trazer o texto da ementa/tese — não só o número
- Indicar claramente se é vinculante (repetitivo, súmula vinculante) ou apenas orientativo
- Se não encontrar resultado concreto, informar: "Não encontrei jurisprudência específica pra esse tema. Quer que eu amplie a busca ou tente outra abordagem?"
- Não inventar números de processo ou ementas — só usar o que foi encontrado na busca
- Resultados do TJAP têm menor força vinculante que STJ/STF — indicar isso quando relevante
