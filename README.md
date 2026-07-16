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
├── skill-creator/                 # Framework para criação de skills
└── pdf/                           # Skill para processamento de PDFs
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

1. Copie a pasta `CDC/` para o diretório de skills do Claude
2. Ou use o arquivo `skill-assistente-pedagogico.md` como skill avulsa
3. Faça perguntas como professor do Colégio Divino Coração sobre planos de aula, avaliações ou engajamento de turmas
