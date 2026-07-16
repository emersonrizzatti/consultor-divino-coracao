const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Assistente Pedagógico CDC";
pres.title = "IA na Sala de Aula — CDC";

// === CDC IDENTITY COLORS ===
const C = {
  primary:  "81211F",
  accent:   "FFCC2A",
  dark:     "1A1A1A",
  text:     "333333",
  lightBg:  "F5F5F5",
  white:    "FFFFFF",
  midGray:  "999999",
  darkBg:   "222222",
  footText: "CCCCCC",
};

const F = {
  h: "Georgia",
  b: "Calibri",
};

const mkShadow = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.10 });

function titleSlide(title, subtitle, footer) {
  const s = pres.addSlide();
  s.background = { color: C.primary };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.575, w: 10, h: 0.05, fill: { color: C.accent } });
  s.addShape(pres.shapes.OVAL, { x: 7.5, y: -1.0, w: 4.0, h: 4.0, fill: { color: C.accent, transparency: 85 } });
  s.addShape(pres.shapes.OVAL, { x: 8.5, y: 3.5, w: 3.0, h: 3.0, fill: { color: C.accent, transparency: 85 } });
  s.addText(title, { x: 0.7, y: 1.0, w: 7.5, h: 3.0, fontSize: 32, fontFace: F.h, color: C.white, bold: true, align: "left", valign: "middle", margin: 0 });
  if (subtitle) s.addText(subtitle, { x: 0.7, y: 3.8, w: 7.5, h: 0.7, fontSize: 14, fontFace: F.b, color: C.accent, align: "left", margin: 0 });
  if (footer) s.addText(footer, { x: 0.7, y: 4.8, w: 7.5, h: 0.4, fontSize: 11, fontFace: F.b, color: C.white, align: "left", margin: 0 });
  return s;
}

function sectionSlide(title, subtitle) {
  const s = pres.addSlide();
  s.background = { color: C.primary };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.575, w: 10, h: 0.05, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.3, w: 10, h: 0.005, fill: { color: C.accent, transparency: 60 } });
  s.addText(title, { x: 0.7, y: 1.8, w: 8.6, h: 1.2, fontSize: 36, fontFace: F.h, color: C.white, bold: true, align: "left", valign: "middle", margin: 0 });
  if (subtitle) s.addText(subtitle, { x: 0.7, y: 3.0, w: 8.6, h: 0.8, fontSize: 16, fontFace: F.b, color: C.accent, align: "left", valign: "top", margin: 0 });
  return s;
}

function contentSlide(title) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  // Top bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.9, w: 10, h: 0.04, fill: { color: C.accent } });
  // Title
  s.addText(title, { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 22, fontFace: F.h, color: C.white, bold: true, align: "left", valign: "middle", margin: 0 });
  // Subtle bottom bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.525, w: 10, h: 0.1, fill: { color: C.lightBg } });
  return s;
}

function addTextBlock(slide, items, x, y, w, h, opts) {
  slide.addText(items, {
    x: x || 0.6, y: y || 1.2, w: w || 8.8, h: h || 4.0,
    fontSize: 14, fontFace: F.b, color: C.text, valign: "top", align: "left",
    margin: 0, paraSpaceAfter: 4, ...opts,
  });
}

function addCard(slide, x, y, w, h) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.white }, shadow: mkShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.05, h, fill: { color: C.accent } });
}

// =======================================================
// SLIDE 1 — TITLE
// =======================================================
titleSlide(
  "Inteligência Artificial\ndo jeito que todo\nprofessor entende",
  "Uma comparação com o mundo dos veículos para descomplicar a IA",
  "Assistente Pedagógico | Colégio Divino Coração — ACSC"
);

// =======================================================
// SLIDE 2 — RESUMO DA METÁFORA (visão geral)
// =======================================================
{
  const s = sectionSlide("Resumo da Metáfora", "Antes de começar, veja como vamos comparar cada conceito de IA com algo do universo automotivo.");
  // Table
  const rows = [
    [{ text: "Conceito de IA", options: { bold: true, color: C.white, fill: { color: C.primary }, fontSize: 13 } },
     { text: "Metáfora Automotiva", options: { bold: true, color: C.white, fill: { color: C.primary }, fontSize: 13 } }],
    [{ text: "Inteligência Artificial", options: { color: C.text, fill: { color: C.white }, fontSize: 12 } },
     { text: "Indústria automotiva (todo o setor)", options: { color: C.text, fill: { color: C.lightBg }, fontSize: 12 } }],
    [{ text: "Modelo de IA", options: { color: C.text, fill: { color: C.white }, fontSize: 12 } },
     { text: "Montadora (Toyota, Volvo, Honda…)", options: { color: C.text, fill: { color: C.lightBg }, fontSize: 12 } }],
    [{ text: "Ferramenta de IA Generativa", options: { color: C.text, fill: { color: C.white }, fontSize: 12 } },
     { text: "Carro específico (sedan, hatch, SUV)", options: { color: C.text, fill: { color: C.lightBg }, fontSize: 12 } }],
    [{ text: "Agente de IA", options: { color: C.text, fill: { color: C.white }, fontSize: 12 } },
     { text: "Carro com piloto automático", options: { color: C.text, fill: { color: C.lightBg }, fontSize: 12 } }],
    [{ text: "Skill", options: { color: C.text, fill: { color: C.white }, fontSize: 12 } },
     { text: "Manual do proprietário", options: { color: C.text, fill: { color: C.lightBg }, fontSize: 12 } }],
  ];
  s.addTable(rows, { x: 0.7, y: 0.8, w: 8.6, colW: [3.2, 5.4], border: { pt: 0.5, color: C.primary }, rowH: 0.55, margin: [3, 8, 3, 8] });
}

// =======================================================
// SLIDE 3 — O QUE É IA?
// =======================================================
{
  const s = contentSlide("O que é Inteligência Artificial?");
  addTextBlock(s, [
    { text: "Imagine a indústria automotiva como um todo:", options: { bold: true, breakLine: true, fontSize: 16, color: C.primary } },
    { text: "Ela reúne fábricas, engenheiros, projetos, motores, estradas — tudo que envolve o mundo dos carros.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "A Inteligência Artificial é a mesma coisa:", options: { bold: true, breakLine: true, fontSize: 16, color: C.primary } },
    { text: "É todo o campo de estudo que ensina máquinas a pensar, aprender, reconhecer padrões e tomar decisões.", options: { breakLine: true, paraSpaceAfter: 6 } },
    { text: "Dentro dela existem vários tipos: sistemas que recomendam filmes, que reconhecem sua voz, que geram textos ou imagens…", options: { breakLine: true, paraSpaceAfter: 6 } },
    { text: "Assim como a indústria automotiva produz desde caminhões até carros de corrida, a IA produz soluções diferentes para problemas diferentes.", options: {} },
  ]);
  // Visual accent
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.2, w: 4.5, h: 0.005, fill: { color: C.accent } });
  s.addText("IA ≠ pensamento humano. IA = reconhecimento de padrões em escala.", {
    x: 0.6, y: 4.3, w: 8.8, h: 0.5, fontSize: 12, fontFace: F.b, color: C.midGray, italic: true, align: "left", margin: 0,
  });
}

// =======================================================
// SLIDE 4 — MODELOS = MONTADORAS
// =======================================================
{
  const s = contentSlide("Modelos de IA = Montadoras de Veículos");
  addTextBlock(s, [
    { text: "Se a IA é a indústria, os modelos de IA são as montadoras:", options: { bold: true, breakLine: true, fontSize: 16, color: C.primary, paraSpaceAfter: 6 } },
    { text: "Cada montadora tem sua engenharia, seu jeito de construir carros, suas especialidades. O mesmo vale para os modelos de IA.", options: { breakLine: true, paraSpaceAfter: 6 } },
  ], 0.6, 1.1, 8.8, 1.2);

  const pairs = [
    ["GPT (OpenAI)", "Toyota", "Versátil, confiável, o mais popular."],
    ["Claude (Anthropic)", "Volvo", "Cuidadoso, seguro, foco em ética."],
    ["Gemini (Google)", "Honda", "Integrado ao ecossistema Google."],
    ["DeepSeek / Llama", "BYD / JAC", "Abertos, acessíveis, crescimento rápido."],
  ];
  pairs.forEach((p, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.6, y = 2.4 + row * 1.5;
    addCard(s, x, y, 4.2, 1.3);
    s.addText(p[0], { x: x + 0.15, y: y + 0.05, w: 3.9, h: 0.35, fontSize: 16, fontFace: F.h, color: C.primary, bold: true, align: "left", margin: 0 });
    s.addText("→ " + p[1], { x: x + 0.15, y: y + 0.4, w: 3.9, h: 0.25, fontSize: 13, fontFace: F.b, color: C.accent, italic: true, align: "left", margin: 0 });
    s.addText(p[2], { x: x + 0.15, y: y + 0.7, w: 3.9, h: 0.5, fontSize: 12, fontFace: F.b, color: C.midGray, align: "left", margin: 0 });
  });
}

// =======================================================
// SLIDE 5 — COMO OS MODELOS FUNCIONAM
// =======================================================
{
  const s = contentSlide("Como os Modelos de IA Funcionam?");
  // Left content
  addTextBlock(s, [
    { text: "Diferente do que muitos pensam, a IA não \"pensa\" ou \"entende\" como uma pessoa.", options: { bold: true, breakLine: true, fontSize: 16, color: C.primary, paraSpaceAfter: 8 } },
    { text: "Ela funciona com dois ingredientes:", options: { bold: true, breakLine: true, fontSize: 16, color: C.primary, paraSpaceAfter: 6 } },
  ], 0.6, 1.1, 5.2, 1.0);

  // Stat card
  [["1. Estatística", "O modelo foi treinado com bilhões de textos. Aprendeu quais palavras costumam aparecer juntas."],
   ["2. Previsão", "Quando você pergunta algo, ele calcula a palavra mais provável — palavra por palavra."],
  ].forEach((item, i) => {
    const y = 2.1 + i * 0.65;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 5.2, h: 0.55, fill: { color: C.lightBg }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.05, h: 0.55, fill: { color: C.accent } });
    s.addText([
      { text: item[0] + " ", options: { bold: true, color: C.primary } },
      { text: item[1], options: { color: C.text } },
    ], { x: 0.85, y: y + 0.02, w: 4.8, h: 0.5, fontSize: 12, fontFace: F.b, align: "left", valign: "middle", margin: 0 });
  });

  // Right highlight box
  s.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 1.1, w: 3.4, h: 4.0, fill: { color: C.primary }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 1.1, w: 3.4, h: 0.04, fill: { color: C.accent } });
  s.addText("Na prática, a IA é uma máquina de", { x: 6.4, y: 1.3, w: 3.0, h: 0.6, fontSize: 12, fontFace: F.b, color: C.white, align: "center", margin: 0 });
  s.addText("ADIVINHAR", { x: 6.4, y: 1.9, w: 3.0, h: 0.7, fontSize: 32, fontFace: F.h, color: C.accent, bold: true, align: "center", margin: 0 });
  s.addText("a próxima palavra", { x: 6.4, y: 2.5, w: 3.0, h: 0.5, fontSize: 14, fontFace: F.b, color: C.white, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 3.2, w: 2.2, h: 0.03, fill: { color: C.accent } });
  s.addText("Ela é muito boa nisso,\nmas continua sendo\num palpite estatístico.", { x: 6.4, y: 3.4, w: 3.0, h: 0.9, fontSize: 11, fontFace: F.b, color: C.accent, align: "center", margin: 0 });

  // Bottom comparison
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.4, w: 5.2, h: 0.7, fill: { color: C.lightBg }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.4, w: 0.05, h: 0.7, fill: { color: C.primary } });
  s.addText([
    { text: "Pense num corretor de texto do celular: ", options: { bold: true, color: C.primary } },
    { text: "ele sugere a próxima palavra baseado no que você já digitou. A IA generativa faz a mesma coisa — só que em escala gigante.", options: { color: C.text } },
  ], { x: 0.85, y: 4.45, w: 4.8, h: 0.6, fontSize: 11, fontFace: F.b, align: "left", valign: "middle", margin: 0 });
}

// =======================================================
// SLIDE 6 — FERRAMENTAS DE IA GENERATIVA
// =======================================================
{
  const s = contentSlide("Ferramentas de IA Generativa = Carros que você dirige");
  addTextBlock(s, [
    { text: "Se os modelos são montadoras, as ferramentas são os carros que chegam às concessionárias:", options: { bold: true, breakLine: true, fontSize: 16, color: C.primary, paraSpaceAfter: 6 } },
    { text: "ChatGPT → sedã popular. Fácil, escreve textos, gera ideias. Milhares usam no dia a dia.", options: { breakLine: true, paraSpaceAfter: 4 } },
    { text: "Gemini → hatch conectado. Integrado com YouTube, Gmail, Google Drive.", options: { breakLine: true, paraSpaceAfter: 4 } },
    { text: "Claude → sedã premium. Respostas estruturadas, ótimo para análise de documentos.", options: { breakLine: true, paraSpaceAfter: 4 } },
    { text: "Copilot (Microsoft) → carro corporativo integrado ao Office (Word, Excel, PowerPoint).", options: { breakLine: true, paraSpaceAfter: 4 } },
    { text: "Geradores de imagem (DALL-E, Midjourney) → carros customizados para um fim específico.", options: { paraSpaceAfter: 4 } },
  ], 0.6, 1.1, 8.8, 3.0);

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.2, w: 8.8, h: 0.005, fill: { color: C.accent } });
  s.addText("Cada ferramenta tem seu ponto forte. Escolha a que melhor se encaixa no que você precisa fazer.", {
    x: 0.6, y: 4.3, w: 8.8, h: 0.5, fontSize: 12, fontFace: F.b, color: C.midGray, italic: true, align: "left", margin: 0,
  });
}

// =======================================================
// SLIDE 7 — AGENTES DE IA
// =======================================================
{
  const s = contentSlide("Agentes de IA = Carros com Piloto Automático");
  addTextBlock(s, [
    { text: "Até aqui, as ferramentas dependem de você dar os comandos (o prompt).", options: { bold: true, breakLine: true, fontSize: 16, color: C.primary, paraSpaceAfter: 4 } },
    { text: "É como dirigir um carro manual: você acelera, vira, freia, escolhe a rota.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "Um agente de IA é como um carro com piloto automático avançado:", options: { bold: true, breakLine: true, fontSize: 16, color: C.primary, paraSpaceAfter: 6 } },
    { text: "Você diz o destino e o carro planeja a rota, decide quando acelerar ou frear, busca informações no GPS, recalcula se houver trânsito.", options: { breakLine: true, paraSpaceAfter: 6 } },
    { text: "Na prática: você dá um objetivo para o agente, e ele decide quais ferramentas usar, que passos seguir e quando parar.", options: { breakLine: true, paraSpaceAfter: 6 } },
    { text: "Exemplo: um agente pode pesquisar um tema na web, resumir, criar uma apresentação e salvar o arquivo — tudo sozinho.", options: {} },
  ]);
}

// =======================================================
// SLIDE 8 — CONTEXTO E REGRAS
// =======================================================
{
  const s = contentSlide("Contexto e Regras = Instruções para a IA \"adivinhar\" melhor");

  // Left card — problem
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.2, h: 3.4, fill: { color: C.lightBg }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.2, h: 0.4, fill: { color: C.primary } });
  s.addText("Sem instruções, a IA \"chuta\"", { x: 0.65, y: 1.25, w: 3.9, h: 0.3, fontSize: 14, fontFace: F.h, color: C.white, bold: true, align: "left", margin: 0 });
  s.addText([
    { text: "Se você perguntar \"Crie uma atividade\" sem contexto:", options: { bold: true, breakLine: true, fontSize: 13, paraSpaceAfter: 3 } },
    { text: "Pode vir para qualquer série", options: { bullet: true, breakLine: true, fontSize: 12 } },
    { text: "Pode ignorar o regimento da escola", options: { bullet: true, breakLine: true, fontSize: 12 } },
    { text: "Pode usar linguagem inadequada", options: { bullet: true, breakLine: true, fontSize: 12 } },
    { text: "Pode sugerir algo contra os valores da ACSC", options: { bullet: true, breakLine: true, fontSize: 12, paraSpaceAfter: 4 } },
    { text: "O palpite dela pode ser genérico ou errado.", options: { italic: true, fontSize: 12, color: C.midGray } },
  ], { x: 0.7, y: 1.75, w: 3.8, h: 2.7, fontSize: 13, fontFace: F.b, color: C.text, align: "left", valign: "top", margin: 0 });

  // Right card — solution
  s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.2, w: 4.4, h: 4.0, fill: { color: C.primary }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.2, w: 4.4, h: 0.04, fill: { color: C.accent } });
  s.addText("Como melhorar o palpite?", { x: 5.3, y: 1.3, w: 4.0, h: 0.4, fontSize: 15, fontFace: F.h, color: C.accent, bold: true, align: "left", margin: 0 });
  s.addText([
    { text: "Dê CONTEXTO", options: { bold: true, breakLine: true, fontSize: 13, color: C.accent, paraSpaceAfter: 1 } },
    { text: "\"Sou professor do 4º ano do CDC…\"", options: { breakLine: true, fontSize: 11, color: C.white, paraSpaceAfter: 5 } },
    { text: "Estabeleça REGRAS", options: { bold: true, breakLine: true, fontSize: 13, color: C.accent, paraSpaceAfter: 1 } },
    { text: "\"...respeitando o PPPP da ACSC e a BNCC\"", options: { breakLine: true, fontSize: 11, color: C.white, paraSpaceAfter: 5 } },
    { text: "Defina o FORMATO DE SAÍDA", options: { bold: true, breakLine: true, fontSize: 13, color: C.accent, paraSpaceAfter: 1 } },
    { text: "\"Estruture em tabela: objetivos, atividades…\"", options: { breakLine: true, fontSize: 11, color: C.white, paraSpaceAfter: 5 } },
    { text: "Exija EXEMPLOS", options: { bold: true, breakLine: true, fontSize: 13, color: C.accent, paraSpaceAfter: 1 } },
    { text: "\"Inclua um exemplo de parecer descritivo\"", options: { breakLine: true, fontSize: 11, color: C.white, paraSpaceAfter: 5 } },
    { text: "Quanto mais informação, mais preciso será o \"palpite\" da IA.", options: { bold: true, fontSize: 12, color: C.accent, italic: true } },
  ], { x: 5.3, y: 1.8, w: 4.0, h: 3.2, fontSize: 12, fontFace: F.b, color: C.white, align: "left", valign: "top", margin: 0 });

  // Bottom bar on left
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.7, w: 4.2, h: 0.55, fill: { color: C.primary } });
  s.addText("É aí que entram as SKILLS!", { x: 0.5, y: 4.7, w: 4.2, h: 0.55, fontSize: 15, fontFace: F.h, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0 });
}

// =======================================================
// SLIDE 9 — SKILLS = MANUAL DO PROPRIETÁRIO
// =======================================================
{
  const s = contentSlide("Skills = Manual do Proprietário");
  // Left
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 3.7, fill: { color: C.lightBg }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.04, fill: { color: C.accent } });
  s.addText("Imagine um carro novo:", { x: 0.7, y: 1.35, w: 3.9, h: 0.3, fontSize: 15, fontFace: F.h, color: C.primary, bold: true, align: "left", margin: 0 });
  s.addText([
    { text: "Você pode apenas sentar e dirigir (funciona!), mas o manual do proprietário ensina:", options: { breakLine: true, fontSize: 12, paraSpaceAfter: 4 } },
    { text: "Como usar cada botão e função", options: { bullet: true, breakLine: true, fontSize: 12 } },
    { text: "A maneira correta de fazer a manutenção", options: { bullet: true, breakLine: true, fontSize: 12 } },
    { text: "Procedimentos de segurança", options: { bullet: true, breakLine: true, fontSize: 12 } },
    { text: "Modos de condução específicos", options: { bullet: true, fontSize: 12 } },
  ], { x: 0.7, y: 1.8, w: 3.9, h: 2.9, fontSize: 12, fontFace: F.b, color: C.text, align: "left", valign: "top", margin: 0 });

  // Right
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 3.7, fill: { color: C.primary }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.04, fill: { color: C.accent } });
  s.addText("As Skills fazem o mesmo com a IA:", { x: 5.4, y: 1.35, w: 3.9, h: 0.3, fontSize: 15, fontFace: F.h, color: C.accent, bold: true, align: "left", margin: 0 });
  s.addText([
    { text: "São instruções que ensinam o modelo a:", options: { breakLine: true, fontSize: 12, paraSpaceAfter: 4, color: C.white } },
    { text: "Saber quem é você e seu contexto", options: { bullet: true, breakLine: true, fontSize: 12, color: C.white } },
    { text: "Seguir regras da sua escola", options: { bullet: true, breakLine: true, fontSize: 12, color: C.white } },
    { text: "Usar os documentos certos", options: { bullet: true, breakLine: true, fontSize: 12, color: C.white } },
    { text: "Falar como um especialista pedagógico", options: { bullet: true, breakLine: true, fontSize: 12, color: C.white } },
    { text: "Seguir processos passo a passo", options: { bullet: true, breakLine: true, fontSize: 12, paraSpaceAfter: 6, color: C.white } },
    { text: "Sem a skill o modelo \"até funciona\". Com ela, ele age como um profissional da sua área.", options: { bold: true, italic: true, fontSize: 12, color: C.accent } },
  ], { x: 5.4, y: 1.8, w: 3.9, h: 2.9, fontSize: 12, fontFace: F.b, color: C.text, align: "left", valign: "top", margin: 0 });
}

// =======================================================
// SLIDE 10 — MÃO NA MASSA
// =======================================================
{
  const s = pres.addSlide();
  s.background = { color: C.primary };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.575, w: 10, h: 0.05, fill: { color: C.accent } });
  s.addText("Mão na Massa", { x: 0.7, y: 0.3, w: 8.6, h: 0.6, fontSize: 26, fontFace: F.h, color: C.white, bold: true });

  const steps = [
    { n: "1", t: "Baixe o arquivo Skill", d: "Acesse o repositório e baixe skill-assistente-pedagogico.md" },
    { n: "2", t: "Abra o ChatGPT ou Gemini", d: "Entre no chat da plataforma que você mais usa" },
    { n: "3", t: "Carregue o arquivo", d: "Clique em anexar e selecione o arquivo baixado" },
    { n: "4", t: "Faça um pedido", d: "\"Crie um plano de aula de Ciências para o 6º ano\"" },
    { n: "5", t: "Refine", d: "\"Coloque uma atividade lúdica\" ou \"Adicione oração inicial\"" },
    { n: "6", t: "Aproveite", d: "Copie o resultado ou peça um arquivo .docx" },
  ];
  steps.forEach((st, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.15, y = 1.2 + row * 2.1;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.85, h: 1.8, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.85, h: 0.04, fill: { color: C.accent } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.2, w: 0.4, h: 0.4, fill: { color: C.accent } });
    s.addText(st.n, { x: x + 0.15, y: y + 0.2, w: 0.4, h: 0.4, fontSize: 14, fontFace: F.b, color: C.primary, bold: true, align: "center", valign: "middle" });
    s.addText(st.t, { x: x + 0.65, y: y + 0.2, w: 2.0, h: 0.35, fontSize: 13, fontFace: F.h, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    s.addText(st.d, { x: x + 0.15, y: y + 0.7, w: 2.55, h: 0.9, fontSize: 11, fontFace: F.b, color: C.text, align: "left", valign: "top", margin: 0 });
  });
}

// =======================================================
// SLIDE 11 — EXEMPLO PRÁTICO
// =======================================================
{
  const s = contentSlide("Exemplo Prático com a Skill CDC");
  // Prompt box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9.0, h: 1.8, fill: { color: C.lightBg }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 0.05, h: 1.8, fill: { color: C.accent } });
  s.addText("Prompt para testar (copie e cole após carregar a skill):", { x: 0.8, y: 1.25, w: 8.5, h: 0.25, fontSize: 11, fontFace: F.b, color: C.midGray, italic: true, margin: 0 });
  s.addText([
    { text: "\"Sou professor do 4º ano do Colégio Divino Coração.", options: { breakLine: true } },
    { text: "Preciso de um parecer descritivo trimestral para um aluno", options: { breakLine: true } },
    { text: "que está com dificuldades em leitura e interpretação de texto.", options: { breakLine: true } },
    { text: "Inclua sugestões para a família e alinhamento ao PPPP da ACSC.\"", options: {} },
  ], { x: 0.8, y: 1.55, w: 8.5, h: 1.2, fontSize: 12, fontFace: "Consolas", color: C.text, align: "left", valign: "top", margin: 0 });

  // Result box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.3, w: 9.0, h: 1.9, fill: { color: C.primary }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.3, w: 0.05, h: 1.9, fill: { color: C.accent } });
  s.addText("O que a IA vai responder:", { x: 0.8, y: 3.4, w: 8.5, h: 0.25, fontSize: 12, fontFace: F.b, color: C.accent, bold: true, margin: 0 });
  s.addText([
    { text: "Parecer descritivo completo nos padrões do Colégio", options: { bullet: true, breakLine: true, color: C.white } },
    { text: "Alinhado ao sistema trimestral e ao PPPP", options: { bullet: true, breakLine: true, color: C.white } },
    { text: "Linguagem adequada para o boletim escolar", options: { bullet: true, breakLine: true, color: C.white } },
    { text: "Sugestões pedagógicas para a família", options: { bullet: true, breakLine: true, color: C.white } },
    { text: "Estratégias de intervenção em sala de aula", options: { bullet: true, color: C.white } },
  ], { x: 0.8, y: 3.7, w: 8.5, h: 1.3, fontSize: 12, fontFace: F.b, color: C.text, align: "left", valign: "top", margin: 0 });
}

// =======================================================
// SLIDE 12 — OUTRAS FERRAMENTAS
// =======================================================
{
  const s = contentSlide("Outras Skills Disponíveis no Repositório");

  const tools = [
    ["docx/", "Cria e edita documentos Word (.docx) com formatação profissional"],
    ["xlsx/", "Cria e edita planilhas (.xlsx, .csv) com fórmulas, gráficos e formatação"],
    ["pdf/", "Extrai texto, combina, divide, aplica OCR e preenche formulários PDF"],
    ["pptx/", "Cria e edita apresentações de slides (.pptx) com design personalizado"],
    ["skill-creator/", "Cria, edita e testa skills com avaliação automatizada de desempenho"],
  ];

  const startY = 1.2;
  tools.forEach((t, i) => {
    const y = startY + i * 0.8;
    addCard(s, 1.0, y, 8.0, 0.6);
    s.addText(t[0], { x: 1.25, y: y + 0.05, w: 1.8, h: 0.5, fontSize: 16, fontFace: F.h, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    s.addText(t[1], { x: 3.1, y: y + 0.05, w: 5.7, h: 0.5, fontSize: 13, fontFace: F.b, color: C.text, align: "left", valign: "middle", margin: 0 });
  });

  s.addText("Use a skill pptx/ para criar ou editar apresentações como esta!", {
    x: 0.6, y: 5.2, w: 8.8, h: 0.3, fontSize: 11, fontFace: F.b, color: C.midGray, italic: true, align: "center", margin: 0,
  });
}

// =======================================================
// SLIDE 13 — FINAL
// =======================================================
{
  const s = pres.addSlide();
  s.background = { color: C.primary };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.575, w: 10, h: 0.05, fill: { color: C.accent } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 3.5, w: 3.5, h: 3.5, fill: { color: C.accent, transparency: 85 } });
  s.addShape(pres.shapes.OVAL, { x: 8.0, y: -1.0, w: 3.5, h: 3.5, fill: { color: C.accent, transparency: 85 } });

  s.addText("Vamos começar?", { x: 0.7, y: 1.2, w: 8.6, h: 1.0, fontSize: 36, fontFace: F.h, color: C.white, bold: true, align: "center" });
  s.addText("Baixe a skill, carregue no chat e faça seu primeiro pedido.", { x: 0.7, y: 2.2, w: 8.6, h: 0.5, fontSize: 16, fontFace: F.b, color: C.accent, align: "center" });
  s.addText("Em poucos minutos você terá um assistente\npedagógico personalizado para o Colégio Divino Coração.", { x: 0.7, y: 2.7, w: 8.6, h: 0.6, fontSize: 14, fontFace: F.b, color: C.white, align: "center" });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 3.6, w: 3.0, h: 0.03, fill: { color: C.accent } });
  s.addText("Assistente Pedagógico CDC", { x: 0.7, y: 3.8, w: 8.6, h: 0.4, fontSize: 12, fontFace: F.b, color: C.accent, align: "center" });
}

// =======================================================
// SAVE
// =======================================================
pres.writeFile({ fileName: "IA_na_Sala_de_Aula_CDC_2026.pptx" })
  .then(() => console.log("OK: IA_na_Sala_de_Aula_CDC_2026.pptx"))
  .catch(err => console.error("ERRO:", err));
