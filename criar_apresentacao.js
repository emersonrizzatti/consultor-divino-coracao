const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Assistente Pedagógico CDC";
pres.title = "IA na Sala de Aula";

// === Colors ===
const C = {
  dark:     "1A3A3A",
  terracota:"C75B39",
  sage:     "5B8C5A",
  cream:    "F7F3EE",
  white:    "FFFFFF",
  text:     "2D2D2D",
  lightText:"5A5A5A",
  gold:     "D4A843",
};

// === Helpers ===
const mkShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.12 });

function addSectionSlide(title, subtitle) {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.gold } });
  s.addText(title, { x: 0.8, y: 1.2, w: 8.4, h: 1.8, fontSize: 36, fontFace: "Georgia", color: C.white, bold: true, align: "left", valign: "middle" });
  if (subtitle) {
    s.addText(subtitle, { x: 0.8, y: 3.0, w: 8.4, h: 1.2, fontSize: 16, fontFace: "Calibri", color: C.cream, align: "left", valign: "top" });
  }
  return s;
}

function addContentSlide(title, items, opts = {}) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.dark } });
  s.addText(title, { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 24, fontFace: "Georgia", color: C.white, bold: true, align: "left", valign: "middle" });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.9, w: 10, h: 0.04, fill: { color: C.terracota } });

  const startY = opts.startY || 1.2;
  const colW = opts.colW || 8.8;
  const colX = opts.colX || 0.6;
  const col2 = opts.col2 || false;

  if (col2) {
    // Two-column layout
    const leftItems = items.slice(0, Math.ceil(items.length / 2));
    const rightItems = items.slice(Math.ceil(items.length / 2));
    s.addText(leftItems, { x: colX, y: startY, w: 4.1, h: 4, fontSize: 14, fontFace: "Calibri", color: C.text, valign: "top", align: "left", margin: 0, paraSpaceAfter: 6 });
    s.addText(rightItems, { x: colX + 4.5, y: startY, w: 4.1, h: 4, fontSize: 14, fontFace: "Calibri", color: C.text, valign: "top", align: "left", margin: 0, paraSpaceAfter: 6 });
  } else {
    s.addText(items, { x: colX, y: startY, w: colW, h: 4, fontSize: 14, fontFace: "Calibri", color: C.text, valign: "top", align: "left", margin: 0, paraSpaceAfter: 4 });
  }
  return s;
}

// =============================================
// SLIDE 1 — TITLE
// =============================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  // Accent bar at top
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.gold } });
  // Decorative circle
  s.addShape(pres.shapes.OVAL, { x: 7.8, y: -1.2, w: 3.5, h: 3.5, fill: { color: C.terracota, transparency: 30 } });
  s.addShape(pres.shapes.OVAL, { x: 8.5, y: -0.8, w: 2.5, h: 2.5, fill: { color: C.sage, transparency: 40 } });

  s.addText("Inteligência Artificial\ndo jeito que todo\nprofessor entende", {
    x: 0.8, y: 1.0, w: 7.5, h: 3.0, fontSize: 32, fontFace: "Georgia", color: C.white, bold: true, align: "left", valign: "middle", margin: 0,
  });
  s.addText("Uma comparação com o mundo dos veículos\npara descomplicar a IA", {
    x: 0.8, y: 3.8, w: 7.5, h: 0.6, fontSize: 14, fontFace: "Calibri", color: C.cream, align: "left", margin: 0,
  });
  s.addText("Assistente Pedagógico | Colégio Divino Coração", {
    x: 0.8, y: 4.8, w: 7.5, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.gold, align: "left", margin: 0,
  });
}

// =============================================
// SLIDE 2 — O QUE É IA? (indústria automotiva)
// =============================================
addContentSlide("O que é Inteligência Artificial?", [
  { text: "Imagine a indústria automotiva como um todo:", options: { bold: true, breakLine: true, fontSize: 16, color: C.terracota } },
  { text: "Ela reúne fábricas, engenheiros, projetos, motores, estradas — tudo que envolve o mundo dos carros.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 8 } },
  { text: "A Inteligência Artificial é a mesma coisa:", options: { bold: true, breakLine: true, fontSize: 16, color: C.terracota } },
  { text: "É todo o campo de estudo que ensina máquinas a pensar, aprender, reconhecer padrões e tomar decisões — como um grande setor que cria tecnologias inteligentes.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 8 } },
  { text: "Dentro dela existem vários tipos: sistemas que recomendam filmes, que reconhecem sua voz, que geram textos ou imagens...", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 8 } },
  { text: "Assim como a indústria automotiva produz desde caminhões até carros de corrida, a IA produz soluções diferentes para problemas diferentes.", options: { fontSize: 14 } },
], { startY: 1.2 });

// =============================================
// SLIDE 3 — COMO OS MODELOS FUNCIONAM (estatística / adivinhação)
// =============================================
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.dark } });
  s.addText("Como os Modelos de IA Funcionam?", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 24, fontFace: "Georgia", color: C.white, bold: true });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.9, w: 10, h: 0.04, fill: { color: C.terracota } });

  // Left column — explanation
  s.addText([
    { text: "Diferente do que muitos pensam, a IA não \"pensa\" ou \"entende\" como uma pessoa.", options: { bold: true, breakLine: true, fontSize: 16, color: C.terracota, paraSpaceAfter: 8 } },
    { text: "Ela funciona com dois ingredientes:", options: { bold: true, breakLine: true, fontSize: 16, color: C.terracota, paraSpaceAfter: 6 } },
  ], { x: 0.6, y: 1.2, w: 5.5, h: 1.2, fontSize: 14, fontFace: "Calibri", color: C.text, valign: "top", align: "left", margin: 0, paraSpaceAfter: 4 });

  // Stat card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.3, w: 5.5, h: 0.7, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.3, w: 0.06, h: 0.7, fill: { color: C.terracota } });
  s.addText([
    { text: "1. Estatística: ", options: { bold: true, color: C.dark } },
    { text: "O modelo foi treinado com bilhões de textos. Ele aprendeu quais palavras costumam aparecer juntas, em que ordem, em cada contexto.", options: { color: C.text } },
  ], { x: 0.85, y: 2.35, w: 5.1, h: 0.6, fontSize: 13, fontFace: "Calibri", align: "left", valign: "middle", margin: 0 });

  // Prediction card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.15, w: 5.5, h: 0.7, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.15, w: 0.06, h: 0.7, fill: { color: C.terracota } });
  s.addText([
    { text: "2. Previsão: ", options: { bold: true, color: C.dark } },
    { text: "Quando você pergunta algo, ele calcula: \"qual é a palavra mais provável de vir agora?\" Palavra por palavra, ele constrói a resposta.", options: { color: C.text } },
  ], { x: 0.85, y: 3.2, w: 5.1, h: 0.6, fontSize: 13, fontFace: "Calibri", align: "left", valign: "middle", margin: 0 });

  // Highlight box on the right
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 1.2, w: 3.1, h: 3.8, fill: { color: C.dark }, shadow: mkShadow() });
  s.addText("Na prática, a IA é uma máquina de", { x: 6.7, y: 1.4, w: 2.7, h: 0.6, fontSize: 13, fontFace: "Calibri", color: C.cream, align: "center", valign: "middle", margin: 0 });
  s.addText("ADIVINHAR", { x: 6.7, y: 2.0, w: 2.7, h: 0.7, fontSize: 32, fontFace: "Georgia", color: C.gold, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("a próxima\npalavra", { x: 6.7, y: 2.6, w: 2.7, h: 0.8, fontSize: 16, fontFace: "Calibri", color: C.cream, align: "center", valign: "middle", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.9, y: 3.5, w: 2.3, h: 0.04, fill: { color: C.gold } });
  s.addText("Ela é muito boa nisso,\nmas continua sendo\num palpite estatístico.", { x: 6.7, y: 3.6, w: 2.7, h: 1.0, fontSize: 12, fontFace: "Calibri", color: C.cream, align: "center", valign: "middle", margin: 0 });

  // Comparison bar at bottom
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.3, w: 5.5, h: 0.9, fill: { color: C.white }, shadow: mkShadow() });
  s.addText([
    { text: "Pense num corretor de texto do celular: ", options: { bold: true, color: C.terracota } },
    { text: "ele sugere a próxima palavra baseado no que você já digitou. A IA generativa faz a mesma coisa — só que em escala gigante e com muito mais contexto.", options: { color: C.text } },
  ], { x: 0.8, y: 4.4, w: 5.1, h: 0.7, fontSize: 12, fontFace: "Calibri", align: "left", valign: "middle", margin: 0 });
}

// =============================================
// SLIDE 4 — MODELOS DE IA (montadoras)
// =============================================
addContentSlide("Modelos de IA = Montadoras de Veículos", [
  { text: "Se a IA é a indústria automotiva, os modelos de IA são as montadoras:", options: { bold: true, breakLine: true, fontSize: 16, color: C.terracota, paraSpaceAfter: 8 } },
  { text: "Cada montadora (Ford, Toyota, Honda, Volkswagen) tem sua engenharia, seu jeito de construir carros, suas especialidades.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 6 } },
  { text: "Assim, cada modelo de IA tem sua própria arquitetura e forma de processar informações:", options: { breakLine: true, fontSize: 14 } },
], { startY: 1.2 });

// Comparison cards
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.dark } });
  s.addText("Modelos de IA = Montadoras", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 24, fontFace: "Georgia", color: C.white, bold: true });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.9, w: 10, h: 0.04, fill: { color: C.terracota } });

  const cards = [
    { model: "GPT (OpenAI)", car: "Toyota", desc: "Versátil, confiável, o mais popular. Serve para quase tudo." },
    { model: "Claude (Anthropic)", car: "Volvo", desc: "Cuidadoso, seguro, bem instruído. Foco em segurança e ética." },
    { model: "Gemini (Google)", car: "Honda", desc: "Integrado ao ecossistema Google. Bom em raciocínio e buscas." },
    { model: "DeepSeek / Llama", car: "BYD / JAC", desc: "Modelos abertos e acessíveis. Qualidade crescendo rápido." },
  ];

  cards.forEach((card, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const xPos = 0.5 + col * 4.7;
    const yPos = 1.2 + row * 2.1;

    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 4.3, h: 1.8, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 0.06, h: 1.8, fill: { color: C.terracota } });
    s.addText(card.model, { x: xPos + 0.2, y: yPos + 0.1, w: 3.9, h: 0.4, fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, align: "left", margin: 0 });
    s.addText("→ " + card.car, { x: xPos + 0.2, y: yPos + 0.5, w: 3.9, h: 0.3, fontSize: 13, fontFace: "Calibri", color: C.terracota, italic: true, align: "left", margin: 0 });
    s.addText(card.desc, { x: xPos + 0.2, y: yPos + 0.9, w: 3.9, h: 0.7, fontSize: 12, fontFace: "Calibri", color: C.lightText, align: "left", margin: 0 });
  });
}

// =============================================
// SLIDE 5 — FERRAMENTAS DE IA GENERATIVA (modelos de carro)
// =============================================
addContentSlide("Ferramentas de IA Generativa = Carros que você dirige", [
  { text: "Se os modelos são montadoras, as ferramentas são os carros que chegam às concessionárias:", options: { bold: true, breakLine: true, fontSize: 16, color: C.terracota, paraSpaceAfter: 8 } },
  { text: "ChatGPT → é como um sedan popular. Fácil de dirigir, vai buscar informação, escreve textos, gera ideias. Milhares de pessoas usam no dia a dia.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 4 } },
  { text: "Gemini (chat) → como um hatch conectado. Integrado com YouTube, Gmail, Google Drive. Útil para quem já vive no ecossistema Google.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 4 } },
  { text: "Claude (chat) → como um sedã premium. Mais contido, respostas bem estruturadas, ótimo para análise de documentos e textos longos.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 4 } },
  { text: "Copilot (Microsoft) → como um carro da empresa integrado ao Office. Ajuda no Word, Excel, PowerPoint.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 4 } },
  { text: "Geradores de imagem (DALL-E, Midjourney) → como carros customizados para um fim específico (reboque, corrida). Fazem uma coisa muito bem.", options: { fontSize: 14 } },
], { startY: 1.1 });

// =============================================
// SLIDE 6 — AGENTES DE IA (carro autônomo / motorista)
// =============================================
addContentSlide("Agentes de IA = Carros com Piloto Automático", [
  { text: "Até aqui, as ferramentas dependem de você dar os comandos (o prompt):", options: { bold: true, breakLine: true, fontSize: 16, color: C.terracota, paraSpaceAfter: 8 } },
  { text: "É como dirigir um carro manual: você acelera, vira, freia, escolhe a rota.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 8 } },
  { text: "Um agente de IA é como um carro com piloto automático avançado:", options: { bold: true, breakLine: true, fontSize: 16, color: C.terracota, paraSpaceAfter: 8 } },
  { text: "Você diz o destino e o carro planeja a rota, Decide quando acelerar ou frear, Busca informações no GPS, Recalcula se houver trânsito.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 4 } },
  { text: "Na prática: você dá um objetivo para o agente, e ele decide quais ferramentas usar, que passos seguir e quando parar.", options: { breakLine: true, fontSize: 14, paraSpaceAfter: 4 } },
  { text: "Exemplo: um agente pode \"passo a passo\" pesquisar um tema na web, resumir o conteúdo, criar uma apresentação e salvar o arquivo — tudo sozinho.", options: { fontSize: 14, paraSpaceAfter: 4 } },
], { startY: 1.1 });

// =============================================
// SLIDE 7 — CONTEXTO E REGRAS (dicas para dirigir melhor)
// =============================================
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.dark } });
  s.addText("Contexto e Regras = Instruções para a IA \"adivinhar\" melhor", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 22, fontFace: "Georgia", color: C.white, bold: true });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.9, w: 10, h: 0.04, fill: { color: C.terracota } });

  // Left: the problem
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 3.0, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.45, fill: { color: C.terracota } });
  s.addText("Sem instruções, a IA \"chuta\"", { x: 0.7, y: 1.25, w: 3.9, h: 0.35, fontSize: 14, fontFace: "Georgia", color: C.white, bold: true, align: "left", valign: "middle", margin: 0 });
  s.addText([
    { text: "Se você perguntar \"Crie uma atividade\" sem contexto:", options: { bold: true, breakLine: true, fontSize: 13, paraSpaceAfter: 4 } },
    { text: "• Pode vir para qualquer série", options: { bullet: true, breakLine: true, fontSize: 12 } },
    { text: "• Pode ignorar o regimento da escola", options: { bullet: true, breakLine: true, fontSize: 12 } },
    { text: "• Pode usar linguagem inadequada", options: { bullet: true, breakLine: true, fontSize: 12 } },
    { text: "• Pode sugerir algo contra os valores da ACSC", options: { bullet: true, breakLine: true, fontSize: 12, paraSpaceAfter: 6 } },
    { text: "O palpite dela pode ser genérico ou errado.", options: { italic: true, fontSize: 12, color: C.lightText } },
  ], { x: 0.7, y: 1.8, w: 3.9, h: 2.2, fontSize: 13, fontFace: "Calibri", color: C.text, align: "left", valign: "top", margin: 0 });

  // Right: the solution
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 3.7, fill: { color: C.dark }, shadow: mkShadow() });

  s.addText("Como melhorar o palpite?", { x: 5.4, y: 1.3, w: 3.9, h: 0.4, fontSize: 15, fontFace: "Georgia", color: C.gold, bold: true, align: "left", valign: "middle", margin: 0 });

  const tipItems = [
    { text: "Dê CONTEXTO", options: { bold: true, breakLine: true, fontSize: 13, color: C.gold, paraSpaceAfter: 2 } },
    { text: "\"Sou professor do 4º ano do Colégio Divino Coração...\"", options: { breakLine: true, fontSize: 11, color: C.cream, paraSpaceAfter: 6 } },
    { text: "Estabeleça REGRAS", options: { bold: true, breakLine: true, fontSize: 13, color: C.gold, paraSpaceAfter: 2 } },
    { text: "\"...respeitando o PPPP da ACSC e a BNCC\"", options: { breakLine: true, fontSize: 11, color: C.cream, paraSpaceAfter: 6 } },
    { text: "Defina o FORMATO DE SAÍDA", options: { bold: true, breakLine: true, fontSize: 13, color: C.gold, paraSpaceAfter: 2 } },
    { text: "\"Estruture em uma tabela com colunas: objetivos, atividades, recursos, avaliação\"", options: { breakLine: true, fontSize: 11, color: C.cream, paraSpaceAfter: 6 } },
    { text: "Exija EXEMPLOS", options: { bold: true, breakLine: true, fontSize: 13, color: C.gold, paraSpaceAfter: 2 } },
    { text: "\"Inclua um exemplo de parecer descritivo\"", options: { breakLine: true, fontSize: 11, color: C.cream, paraSpaceAfter: 6 } },
    { text: "Quanto mais informação você der, mais preciso será o \"palpite\" da IA.", options: { bold: true, breakLine: true, fontSize: 12, color: C.gold, italic: true } },
  ];

  s.addText(tipItems, { x: 5.4, y: 1.8, w: 3.9, h: 2.9, fontSize: 13, fontFace: "Calibri", color: C.text, align: "left", valign: "top", margin: 0 });

  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 4.3, h: 0.7, fill: { color: C.dark } });
  s.addText("É aí que entram as SKILLS!", { x: 0.5, y: 4.5, w: 4.3, h: 0.7, fontSize: 16, fontFace: "Georgia", color: C.gold, bold: true, align: "center", valign: "middle", margin: 0 });
}

// =============================================
// SLIDE 8 — SKILLS (manual do proprietário)
// =============================================
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.dark } });
  s.addText("Skills = Manual do Proprietário", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 24, fontFace: "Georgia", color: C.white, bold: true });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.9, w: 10, h: 0.04, fill: { color: C.terracota } });

  // Two-column layout with shape cards
  const leftItems = [
    { text: "Imagine que você comprou um carro novo:", options: { bold: true, breakLine: true, fontSize: 16, color: C.terracota, paraSpaceAfter: 6 } },
    { text: "Você pode apenas sentar e dirigir (funciona!), mas o manual do proprietário ensina:", options: { breakLine: true, fontSize: 13, paraSpaceAfter: 4 } },
    { text: "• Como usar cada botão e função", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "• A maneira correta de fazer a manutenção", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "• Procedimentos de segurança", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "• Modos de condução específicos (estrada, cidade, off-road)", options: { bullet: true, breakLine: true, fontSize: 13, paraSpaceAfter: 6 } },
  ];

  // Box on the right
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 3.8, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.06, fill: { color: C.sage } });

  const rightItems = [
    { text: "As Skills fazem o mesmo com a IA:", options: { bold: true, breakLine: true, fontSize: 16, color: C.sage, paraSpaceAfter: 6 } },
    { text: "São arquivos de instrução que ensinam o modelo de IA a:", options: { breakLine: true, fontSize: 13, paraSpaceAfter: 6 } },
    { text: "Saber quem é você e seu contexto", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Seguir regras da sua escola", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Usar os documentos certos", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Falar como um especialista pedagógico", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Seguir processos passo a passo", options: { bullet: true, breakLine: true, fontSize: 13, paraSpaceAfter: 6 } },
    { text: "Sem a skill o modelo \"até funciona\". Com ela, ele age como um profissional da sua área.", options: { bold: true, breakLine: true, fontSize: 13, italic: true } },
  ];

  s.addText(leftItems, { x: 0.6, y: 1.2, w: 4.3, h: 3.8, fontSize: 14, fontFace: "Calibri", color: C.text, valign: "top", align: "left", margin: 0, paraSpaceAfter: 4 });
  s.addText(rightItems, { x: 5.4, y: 1.4, w: 3.9, h: 3.4, fontSize: 14, fontFace: "Calibri", color: C.text, valign: "top", align: "left", margin: 0, paraSpaceAfter: 4 });
}

// =============================================
// SLIDE 9 — MÃO NA MASSA (passo a passo)
// =============================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.gold } });
  s.addText("Mão na Massa", { x: 0.8, y: 0.3, w: 8.4, h: 0.7, fontSize: 28, fontFace: "Georgia", color: C.white, bold: true });

  const steps = [
    { num: "1", title: "Baixe o arquivo Skill", desc: "Acesse o repositório e baixe skill-assistente-pedagogico.md (link no README)" },
    { num: "2", title: "Abra o ChatGPT ou Gemini", desc: "Entre no chat da plataforma que você mais usa" },
    { num: "3", title: "Carregue o arquivo", desc: "Clique no botão de anexar e selecione o arquivo baixado" },
    { num: "4", title: "Faça um pedido", desc: "Escreva algo como: \"Crie um plano de aula de Ciências para o 6º ano sobre cadeia alimentar\"" },
    { num: "5", title: "Refine", desc: "Peça ajustes: \"Coloque uma atividade lúdica\" ou \"Adicione um momento de oração inicial\"" },
    { num: "6", title: "Aproveite", desc: "Copie o resultado para seu plano de aula, salve ou peça um arquivo .docx" },
  ];

  steps.forEach((step, i) => {
    const row = Math.floor(i / 3);
    const col = i % 3;
    const xPos = 0.5 + col * 3.15;
    const yPos = 1.3 + row * 2.0;

    // Card
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 2.85, h: 1.7, fill: { color: C.white }, shadow: mkShadow() });
    // Number circle
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.15, y: yPos + 0.15, w: 0.45, h: 0.45, fill: { color: C.terracota } });
    s.addText(step.num, { x: xPos + 0.15, y: yPos + 0.15, w: 0.45, h: 0.45, fontSize: 14, fontFace: "Calibri", color: C.white, bold: true, align: "center", valign: "middle" });
    // Title
    s.addText(step.title, { x: xPos + 0.7, y: yPos + 0.15, w: 2.0, h: 0.4, fontSize: 14, fontFace: "Georgia", color: C.dark, bold: true, align: "left", valign: "middle", margin: 0 });
    // Description
    s.addText(step.desc, { x: xPos + 0.15, y: yPos + 0.65, w: 2.55, h: 0.9, fontSize: 11, fontFace: "Calibri", color: C.lightText, align: "left", valign: "top", margin: 0 });
  });
}

// =============================================
// SLIDE 10 — EXEMPLO PRÁTICO
// =============================================
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.dark } });
  s.addText("Exemplo Prático com a Skill CDC", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 24, fontFace: "Georgia", color: C.white, bold: true });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.9, w: 10, h: 0.04, fill: { color: C.terracota } });

  // Prompt example box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1.6, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 0.06, h: 1.6, fill: { color: C.terracota } });
  s.addText("Prompt para testar (copie e cole no chat após carregar a skill):", { x: 0.8, y: 1.3, w: 8.5, h: 0.3, fontSize: 12, fontFace: "Calibri", color: C.lightText, italic: true, margin: 0 });
  s.addText([
    { text: "\"Sou professor do 4º ano do Colégio Divino Coração.", options: { breakLine: true } },
    { text: "Preciso de um parecer descritivo trimestral para um aluno", options: { breakLine: true } },
    { text: "que está com dificuldades em leitura e interpretação de texto.", options: { breakLine: true } },
    { text: "Inclua sugestões para a família e alinhamento ao PPPP da ACSC.\"", options: {} },
  ], { x: 0.8, y: 1.6, w: 8.5, h: 1.0, fontSize: 13, fontFace: "Consolas", color: C.text, align: "left", valign: "top", margin: 0 });

  // Expected result box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.1, w: 9, h: 2.0, fill: { color: C.dark }, shadow: mkShadow() });
  s.addText("O que a IA vai responder:", { x: 0.8, y: 3.2, w: 8.5, h: 0.3, fontSize: 12, fontFace: "Calibri", color: C.gold, bold: true, margin: 0 });
  s.addText([
    { text: "Um parecer descritivo completo nos padrões do Colégio", options: { bullet: true, breakLine: true, color: C.cream } },
    { text: "Alinhado ao sistema trimestral e ao PPPP", options: { bullet: true, breakLine: true, color: C.cream } },
    { text: "Com linguagem adequada para o boletim escolar", options: { bullet: true, breakLine: true, color: C.cream } },
    { text: "Sugestões pedagógicas para a família", options: { bullet: true, breakLine: true, color: C.cream } },
    { text: "Estratégias de intervenção em sala de aula", options: { bullet: true, color: C.cream } },
  ], { x: 0.8, y: 3.5, w: 8.5, h: 1.4, fontSize: 13, fontFace: "Calibri", color: C.white, align: "left", valign: "top", margin: 0 });
}

// =============================================
// SLIDE 11 — RECAP
// =============================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.gold } });
  s.addText("Resumo da Metáfora", { x: 0.8, y: 0.3, w: 8.4, h: 0.7, fontSize: 28, fontFace: "Georgia", color: C.white, bold: true });

  const pairs = [
    ["Inteligência Artificial", "Indústria automotiva (todo o setor)"],
    ["Modelo de IA", "Montadora (Toyota, Volvo, Honda...)"],
    ["Ferramenta de IA Generativa", "Carro específico (sedan, hatch, SUV)"],
    ["Agente de IA", "Carro com piloto automático"],
    ["Skill", "Manual do proprietário"],
  ];

  const tableData = [
    [
      { text: "Conceito", options: { fill: { color: C.terracota }, color: C.white, bold: true, fontSize: 14 } },
      { text: "Metáfora", options: { fill: { color: C.terracota }, color: C.white, bold: true, fontSize: 14 } },
    ],
    ...pairs.map(([a, b]) => [
      { text: a, options: { fill: { color: C.white }, color: C.dark, bold: true, fontSize: 13 } },
      { text: b, options: { fill: { color: C.cream }, color: C.text, fontSize: 13 } },
    ]),
  ];

  s.addTable(tableData, {
    x: 0.8, y: 1.3, w: 8.4,
    colW: [3.5, 4.9],
    border: { pt: 0.5, color: C.dark },
    rowH: [0.5, 0.5, 0.5, 0.5, 0.5, 0.5],
    margin: [3, 8, 3, 8],
  });

  s.addText("Com a skill carregada, o ChatGPT ou Gemini \"vira\" um especialista no Colégio Divino Coração — pronto para ajudar você no dia a dia!", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.7, fontSize: 14, fontFace: "Calibri", color: C.gold, align: "center", valign: "middle", italic: true, margin: 0,
  });
}

// =============================================
// SLIDE 12 — FINAL
// =============================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.gold } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 3.5, w: 4, h: 4, fill: { color: C.terracota, transparency: 30 } });
  s.addShape(pres.shapes.OVAL, { x: 8.5, y: -1, w: 3, h: 3, fill: { color: C.sage, transparency: 40 } });

  s.addText("Vamos começar?", { x: 0.8, y: 1.2, w: 8.4, h: 1.0, fontSize: 36, fontFace: "Georgia", color: C.white, bold: true, align: "center" });
  s.addText("Baixe a skill, carregue no chat e faça seu primeiro pedido.", { x: 0.8, y: 2.2, w: 8.4, h: 0.5, fontSize: 16, fontFace: "Calibri", color: C.cream, align: "center" });
  s.addText("Em poucos minutos você terá um assistente\npedagógico personalizado para o Colégio Divino Coração.", { x: 0.8, y: 2.7, w: 8.4, h: 0.6, fontSize: 14, fontFace: "Calibri", color: C.cream, align: "center" });

  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 3.7, w: 3, h: 0.04, fill: { color: C.gold } });
  s.addText("Assistente Pedagógico CDC", { x: 0.8, y: 3.9, w: 8.4, h: 0.4, fontSize: 12, fontFace: "Calibri", color: C.gold, align: "center" });
}

// =============================================
// SAVE
// =============================================
pres.writeFile({ fileName: "IA_na_Sala_de_Aula_CDC.pptx" })
  .then(() => console.log("OK: IA_na_Sala_de_Aula_CDC.pptx"))
  .catch(err => console.error("ERRO:", err));
