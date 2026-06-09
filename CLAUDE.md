# Conceição Duarte Advocacia — Claude Code OS

## O que é esse workspace

Workspace do escritório Conceição Duarte Advocacia e Consultoria Jurídica. Aqui ficam os clientes, peças jurídicas, jurisprudência, captação e marketing pessoal.

**Estrutura de pastas:**
- `_contexto/` — memória do sistema (não apagar)
- `clientes/` — um subdiretório por cliente com briefing e procuração
- `peticoes/` — petições por cliente ou por tipo
- `pareceres/` — pareceres jurídicos
- `contratos/` — minutas e contratos elaborados
- `jurisprudencia/` — pesquisas organizadas por área (prioridade: `direito-bancario/`)
- `captacao-clientes/` — materiais e estratégias de captação
- `marketing-pessoal/` — conteúdo Instagram, posicionamento, autoridade
- `dados/` — arquivos para análise (PDFs, CSVs, docs recebidos)
- `marca/` — logo e guia de design
- `templates/skills/` — templates de skills prontos pra personalizar com /mapear
- `templates/ferramentas/catalogo.md` — APIs e ferramentas disponíveis pra usar em skills
- `tarefas.md` — lista de tarefas corrente

## Sobre o negócio

Conceição Maria Duarte Portilho é advogada solo, titular do escritório Conceição Duarte Advocacia e Consultoria Jurídica. Atende clientes externos com foco atual em Direito Bancário. Toca todas as frentes do escritório — jurídica e administrativa.

## O que mais fazemos aqui

- Petições (iniciais, contestações, recursos)
- Pareceres jurídicos
- Contratos e minutas
- Consultas jurídicas
- Pesquisa de jurisprudência (foco: Direito Bancário)
- Atendimento e captação de clientes
- Conteúdo para Instagram e marketing pessoal

## Clientes e contexto

Atende clientes externos (pessoas físicas e jurídicas). Escritório solo — Conceição cuida de tudo.

## Tom de voz

Escrever com clareza, fundamentação jurídica e jurisprudência. Evitar textos vagos, genéricos ou sem base legal. Linguagem precisa e acessível ao mesmo tempo — o cliente precisa entender, o juiz precisa respeitar.

## Ferramentas conectadas

- [ ] Gmail — `claude mcp add gmail -- npx -y @gongrzhe/server-gmail-autoauth-mcp`
- [ ] Google Calendar — `claude mcp add google-calendar -- npx -y @gongrzhe/server-google-calendar-autoauth-mcp`
- [ ] Canva — `claude mcp add canva -- npx -y @canva/canva-mcp-server`

*(Marcar conforme for instalando os MCPs)*

---

## Contexto do negócio

No início de toda conversa, ler os seguintes arquivos (se existirem e estiverem configurados):

1. `_contexto/empresa.md` — quem é o usuário, o que faz, como funciona o negócio
2. `_contexto/preferencias.md` — tom de voz, estilo de escrita, o que evitar
3. `_contexto/estrategia.md` — foco atual, prioridades, o que pode esperar

Usar essas informações como base pra qualquer resposta ou decisão. Ao sugerir prioridades, formatos ou abordagens, considerar o foco atual descrito em `estrategia.md`.

Para qualquer tarefa visual (carrossel, proposta, slide, landing page), consultar `marca/design-guide.md` como referência de estilo.

Não é necessário listar o que foi lido nem confirmar a leitura. Apenas usar o contexto naturalmente.

---

## Fluxo de trabalho

Antes de executar qualquer tarefa, verificar se existe uma skill relevante em `.claude/skills/` ou `.claude/commands/`.
Se encontrar, seguir as instruções da skill.
Se não encontrar, executar a tarefa normalmente.

Ao concluir uma tarefa que não tinha skill mas parece repetível (o usuário provavelmente vai pedir de novo no futuro), perguntar:

> "Isso pode virar uma skill pra próxima vez. Quer que eu crie?"

Não perguntar pra tarefas pontuais ou perguntas simples. Só quando o padrão de repetição for claro.

---

## Aprender com correções

Quando o usuário corrigir algo, melhorar uma resposta ou dar uma instrução que parece permanente (frases como "na verdade é assim", "não faça mais isso", "prefiro assim", "sempre que...", "evita...", "da próxima vez..."), perguntar:

> "Quer que eu salve isso pra não precisar repetir?"

Se sim, identificar onde faz mais sentido salvar:

- **Sobre o negócio** → `_contexto/empresa.md`
- **Sobre preferências e estilo** → `_contexto/preferencias.md`
- **Sobre prioridades e foco atual** → `_contexto/estrategia.md`
- **Regra de comportamento nessa pasta** → `CLAUDE.md`

Salvar com uma linha nova clara, sem reformatar o arquivo inteiro. Confirmar o que foi salvo mostrando a linha adicionada.

---

## Manter contexto atualizado

Ao terminar uma tarefa que mudou algo relevante no projeto (novo cliente, nova skill, mudança de foco, novo processo, ferramenta instalada, estrutura de pastas alterada), perguntar:

> "Isso mudou algo no teu contexto. Quer que eu atualize os arquivos de memória?"

Se sim, identificar o que precisa atualizar:

- **Novo cliente, serviço, ferramenta, equipe** → `_contexto/empresa.md`
- **Mudança de prioridade ou foco** → `_contexto/estrategia.md`
- **Correção de tom ou estilo** → `_contexto/preferencias.md`
- **Nova pasta, regra de organização, skill criada** → `CLAUDE.md`
- **Mudança visual (cores, fontes, logo)** → `marca/design-guide.md`

Mostrar o que vai mudar antes de salvar. Não reformatar o arquivo inteiro, só adicionar ou editar a linha relevante.

---

## Criação de skills

Quando o usuário pedir pra criar uma nova skill:

1. Verificar se existe um template relevante em `templates/skills/`. Se existir, usar como base e adaptar pro contexto do usuário
2. Perguntar: "Essa skill é específica pra esse projeto ou vai ser útil em qualquer projeto?"
   - Específica desse negócio → salvar em `.claude/skills/nome-da-skill/SKILL.md` (local)
   - Útil em qualquer projeto → salvar em `~/.claude/skills/nome-da-skill/SKILL.md` (global)
3. Ler `_contexto/empresa.md` e `_contexto/preferencias.md` pra calibrar o conteúdo da skill ao contexto do negócio
4. Se a skill precisar de arquivos de apoio (templates, referências, exemplos), criar dentro da pasta da skill
5. Seguir o fluxo da skill-creator nativa do Claude Code
