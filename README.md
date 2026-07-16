# Consultor Pedagógico — Colégio Divino Coração (ACSC)

Skill de apoio aos professores do **Colégio Divino Coração** (Alegrete/RS), mantido pela **Associação Congregação de Santa Catarina (ACSC)**.

## Estrutura do Projeto

```
├── CDC/                          # Skill principal da escola
│   ├── SKILL.md                  # Instruções da skill (249 linhas)
│   ├── evals/
│   │   └── evals.json            # 3 casos de teste
│   └── references/               # Documentos institucionais
│       ├── regimento-parcial.md  # Regimento Escolar ACSC/RS (98 artigos)
│       ├── regimento-parcial.pdf # PDF original do regimento
│       ├── sistema-de-avaliacao.md
│       ├── metodologia.md
│       └── objetivos-gerais-da-escola.md
├── skill-assistente-pedagogico.md # Skill focada no Ensino Fundamental (273 linhas)
├── LDB_texto_integral.md          # Lei de Diretrizes e Bases da Educação (texto compilado)
├── docx/                          # Criação e edição de documentos Word (.docx)
├── pdf/                           # Leitura, extração e manipulação de PDFs
├── pptx/                          # Criação e edição de apresentações (.pptx)
└── skill-creator/                 # Criação e avaliação de skills
```

## Skills Inclusas

| Skill | Público | Etapas | Função |
|-------|---------|--------|--------|
| `apoio-docente-cdc` (CDC/SKILL.md) | Todos os professores | EI, EF, EM, NEM, AEE | Planos de aula, rubricas, engajamento, conselho de classe |
| `assistente-pedagogico-ef-cdc` (skill-assistente-pedagogico.md) | Professores do EF | 1º ao 9º ano | Planos de trabalho, parecer descritivo, escala 0-10, progressão parcial |

## Referências

- **PPPP** — Projeto Político Pedagógico Pastoral ACSC
- **Regimento Escolar Padrão ACSC/RS** (versão 24/12/2021)
- **LDB 9.394/96** — Lei de Diretrizes e Bases da Educação Nacional
- **BNCC** — Base Nacional Comum Curricular
- **Carisma Congregacional** — Bem-Aventurada Regina Protmann
- **Filosofia Cristã Católica**

## Como Usar

1. [Baixar skill-assistente-pedagogico.md](https://raw.githubusercontent.com/emersonrizzatti/consultor-divino-coracao/main/skill-assistente-pedagogico.md) (clique com o botão direito e "Salvar link como...") ou [Baixar CDC/SKILL.md](https://raw.githubusercontent.com/emersonrizzatti/consultor-divino-coracao/main/CDC/SKILL.md) (versão completa para todas as etapas)
2. Faça perguntas como professor do Colégio Divino Coração sobre planos de aula, avaliações ou engajamento de turmas
3. A IA irá responder com base nas informações contidas no arquivo Skill.

## Ferramentas

- **docx/** — Cria, edita e formata documentos Word (.docx) com sumário, cabeçalhos, numeração de páginas e timed e alterações controladas.
- **pdf/** — Extrai texto e tabelas de PDFs, combina, divide, aplica OCR, preenche formulários e converte documentos para PDF.
- **pptx/** — Cria e edita apresentações de slides (.pptx) a partir de texto, templates ou decks existentes, incluindo notas do apresentador.
- **skill-creator/** — Cria novas skills do zero, edita skills existentes, executa testes automatizados (evals) e mede o desempenho das skills.
