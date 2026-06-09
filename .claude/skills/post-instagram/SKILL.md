---
name: post-instagram
description: >
  Cria post completo para o Instagram da Conceição Duarte Advocacia.
  A partir de um tema ou ideia, escreve o texto do post (hook + desenvolvimento + CTA)
  e sugere o layout da imagem para montar no Canva.
  Use quando disser "cria um post", "faz um post sobre X", "quero postar sobre X",
  "post pro Instagram", "conteúdo sobre X".
---

# /post-instagram — Post para o Instagram

## Dependências

- `_contexto/empresa.md` — contexto do escritório e área de atuação
- `_contexto/preferencias.md` — tom de voz e o que evitar
- `marca/design-guide.md` — identidade visual para sugestão de imagem

---

## Workflow

### Passo 1 — Entender o tema

Identificar o tema do post. Se o usuário não informou, perguntar:

> "Qual é o tema do post? Pode ser um assunto jurídico, uma dica, um caso comum, uma conquista..."

Exemplos de temas recorrentes em Direito Bancário:
- Cobrança indevida de tarifas
- Negativação indevida
- Juros abusivos em contratos
- Superendividamento
- Portabilidade de crédito
- Direitos do consumidor bancário

### Passo 2 — Ler o contexto

Ler `_contexto/empresa.md` e `_contexto/preferencias.md`.

Calibrar:
- Tom: claro, fundamentado, acessível — sem juridiquês desnecessário
- Público: pessoas físicas com problemas bancários, consumidores
- Posicionamento: Conceição como referência em Direito Bancário

### Passo 3 — Escrever o texto do post

Estrutura obrigatória:

**Hook (primeiras 2 linhas — antes do "ver mais"):**
- Frase que para o scroll
- Pode ser uma pergunta, uma afirmação surpreendente ou um dado
- Não começar com "Olá" ou apresentação

**Desenvolvimento (3-5 parágrafos curtos):**
- Explicar o problema ou direito em linguagem acessível
- Incluir base legal quando relevante (ex: "segundo o CDC, art. 42...")
- Um ponto por parágrafo
- Sem bullet points excessivos

**CTA (última linha):**
- Direcionado à ação: "Me chama no WhatsApp", "Salva esse post", "Comenta sua dúvida"
- Não usar "não esquece de curtir/seguir"

**Hashtags (ao final, separadas do texto):**
- 5 a 8 hashtags
- Mix: área (#direitobancario #direitodoconsumidor), local (#advogada #advocacia), tema específico

### Passo 4 — Sugerir layout da imagem para Canva

Com base em `marca/design-guide.md`, sugerir:

```
SUGESTÃO DE IMAGEM (Canva):
- Formato: 1080x1080px (feed quadrado)
- Fundo: #1A1517 (preto carvão)
- Texto principal: [frase do hook ou uma frase-chave do post]
- Cor do texto: rose gold (#E8B4B0)
- Fonte do título: Cinzel Medium
- Fonte do subtítulo: Montserrat Light
- Elemento: logo brasão CD no canto inferior direito
- Visual: minimalista — texto centralizado, sem imagem de fundo
```

Adaptar a sugestão ao tema do post. Se for um post sobre cobrança indevida, por exemplo, sugerir o texto da imagem como "O banco te cobrou indevido? Você tem direito à devolução em dobro."

### Passo 5 — Salvar

Salvar em `marketing-pessoal/instagram/post-[tema-resumido]-[data].md`

Formato do arquivo:

```markdown
# Post — [Tema]
**Data:** [data de criação]

## Texto do post

[texto completo]

[hashtags]

## Sugestão de imagem (Canva)

[sugestão completa]
```

---

## Regras

- Nunca usar linguagem genérica de criador de conteúdo ("ei pessoal", "conteúdo incrível")
- Sempre calibrar pelo tom de `_contexto/preferencias.md`
- Base legal quando o tema exigir — mas sem transformar o post em petição
- O post deve soar como Conceição fala, não como um robô jurídico
- Frases curtas. Parágrafos de no máximo 3 linhas
