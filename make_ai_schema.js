const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_WIDE';

const slide = prs.addSlide();
slide.background = { color: 'FFFFFF' };

const F = 'Calibri';

// Color palette — all gray
const C = {
  bg:       'F4F4F4',
  border:   'AAAAAA',
  dark:     '2B2B2B',
  mid:      '555555',
  light:    '888888',
  lighter:  'BBBBBB',
  white:    'FFFFFF',
  blockFill:'ECECEC',
  dangerFill:'F0F0F0',
  dangerBrd: '777777',
  lineDark:  '555555',
  lineLight: 'AAAAAA',
};

// Title
slide.addText('Схема работы ИИ-помощника', {
  x: 0.3, y: 0.12, w: 12.6, h: 0.42,
  fontSize: 18, bold: true, color: C.dark, fontFace: F, align: 'center',
});
slide.addText('backend/routes/ai.js  ·  frontend/src/components/ChatWidget.jsx', {
  x: 0.3, y: 0.55, w: 12.6, h: 0.25,
  fontSize: 9, color: C.light, fontFace: F, align: 'center',
});

// Column separator lines (dashed)
const sepY = 0.85, sepH = 5.8;
[2.35, 5.05, 8.3].forEach(x => {
  slide.addShape(prs.ShapeType.line, {
    x, y: sepY, w: 0, h: sepH,
    line: { color: C.lighter, width: 0.8, dashType: 'dash' },
  });
});

// Column labels
[
  { x: 0.1,  w: 2.2,  label: 'ФРОНТЕНД' },
  { x: 2.4,  w: 2.6,  label: 'БЭКЕНД' },
  { x: 5.1,  w: 3.15, label: 'GROQ API' },
  { x: 8.35, w: 4.9,  label: 'БАЗА ДАННЫХ' },
].forEach(col => {
  slide.addText(col.label, {
    x: col.x, y: sepY, w: col.w, h: 0.22,
    fontSize: 8, bold: true, color: C.lighter, fontFace: F, align: 'center',
  });
});

// ── Helper functions ──────────────────────────────────────────────

function box(x, y, w, h, lines, opts = {}) {
  slide.addShape(prs.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: opts.fill || C.blockFill },
    line: { color: opts.brd || C.border, width: opts.brdW || 1.2 },
    rectRadius: 0.08,
  });
  if (!lines) return;
  const arr = Array.isArray(lines) ? lines : [lines];
  const lineH = h / (arr.length + 0.6);
  arr.forEach((txt, i) => {
    slide.addText(txt.text || txt, {
      x: x + 0.05, y: y + lineH * (i + 0.3), w: w - 0.1, h: lineH,
      fontSize: txt.size || 9,
      bold: txt.bold || false,
      color: txt.color || C.dark,
      fontFace: F,
      align: 'center',
      valign: 'middle',
    });
  });
}

function diamond(x, y, w, h, lines) {
  // Approximate diamond with rotated rectangle or use freeform
  // pptxgenjs supports custom geometry via addShape with ShapeType
  // Use a simple rect with different styling to simulate — or use triangle approximation
  // Actually pptxgenjs has decagon but not diamond directly. Use ShapeType.diamond if available.
  slide.addShape('diamond', {
    x, y, w, h,
    fill: { color: C.blockFill },
    line: { color: C.border, width: 1.2 },
  });
  const arr = Array.isArray(lines) ? lines : [lines];
  const lh = h / (arr.length + 1);
  arr.forEach((txt, i) => {
    slide.addText(txt.text || txt, {
      x: x + w * 0.15, y: y + lh * (i + 0.5), w: w * 0.7, h: lh,
      fontSize: txt.size || 8.5,
      bold: txt.bold || false,
      color: C.dark,
      fontFace: F,
      align: 'center',
      valign: 'middle',
    });
  });
}

function arr(x1, y1, x2, y2, opts = {}) {
  slide.addShape(prs.ShapeType.line, {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    w: Math.abs(x2 - x1) || 0.01,
    h: Math.abs(y2 - y1) || 0.01,
    line: {
      color: opts.color || C.lineDark,
      width: opts.w || 1.2,
      dashType: opts.dash ? 'dash' : 'solid',
      endArrowType: opts.noHead ? 'none' : 'arrow',
      beginArrowType: 'none',
      // flip if going right-to-left or bottom-to-top
    },
    flipH: x2 < x1,
    flipV: y2 < y1,
  });
}

function label(x, y, w, h, txt, opts = {}) {
  slide.addText(txt, {
    x, y, w, h,
    fontSize: opts.size || 8,
    color: opts.color || C.mid,
    fontFace: F,
    align: opts.align || 'center',
    bold: opts.bold || false,
  });
}

// ── Layout constants ──────────────────────────────────────────────
// Rows Y positions
const R = {
  r1: 1.05,  // User / Jailbreak check / Groq #1
  r2: 2.0,   // Blocked / Synonyms+Stemming / DB
  r3: 3.05,  // Build Prompt
  r4: 4.0,   // Groq #2
  r5: 5.1,   // Content Check / Response
};
const BH = 0.72; // standard block height

// ── BLOCKS ───────────────────────────────────────────────────────

// 1. User
box(0.12, R.r1, 2.0, BH, [
  { text: 'Пользователь', bold: true, size: 10 },
  { text: 'ChatWidget.jsx', size: 8, color: C.light },
]);

// 2. Jailbreak check (diamond)
diamond(2.45, R.r1, 2.5, BH, [
  { text: 'Проверка', bold: true, size: 9 },
  { text: 'на джейлбрейк', size: 8 },
  { text: 'isJailbreakAttempt()', size: 7, color: C.light },
]);

// 3. Blocked
box(2.45, R.r2, 2.5, BH, [
  { text: 'ЗАБЛОКИРОВАНО', bold: true, size: 9 },
  { text: 'стандартный ответ', size: 8, color: C.light },
], { brd: C.mid });

// 4. Groq Call 1
box(5.15, R.r1, 3.0, BH, [
  { text: 'Groq #1  —  llama-3.1-8b', bold: true, size: 9 },
  { text: 'Извлечь поисковые термины', size: 8 },
  { text: 'max_tokens: 60  ·  temp: 0.3', size: 7.5, color: C.light },
  { text: 'ai.js : 173', size: 7, color: C.lighter },
]);

// 5. Groq Server (top)
box(8.35, R.r1, 4.5, BH, [
  { text: 'groq.com API', bold: true, size: 10 },
  { text: 'LPU-чипы  ·  ~0.1 сек', size: 8, color: C.light },
]);

// 6. Synonyms + Stemming
box(5.15, R.r2, 3.0, BH, [
  { text: 'Синонимы + стемминг', bold: true, size: 9 },
  { text: 'getSearchVariants() + SYNONYMS', size: 8 },
  { text: 'телефон → iphone, samsung...', size: 7.5, color: C.light },
  { text: 'ai.js : 50', size: 7, color: C.lighter },
]);

// 7. SQLite DB
box(8.35, R.r2, 4.5, BH, [
  { text: 'SQLite', bold: true, size: 10 },
  { text: 'products · categories · users', size: 8, color: C.light },
]);

// 8. Build Prompt
box(5.15, R.r3, 3.0, BH, [
  { text: 'Сборка системного промпта', bold: true, size: 9 },
  { text: 'SYSTEM_PROMPT + товары из БД', size: 8 },
  { text: '+ история диалога (8 сообщений)', size: 8 },
  { text: 'ai.js : 200', size: 7, color: C.lighter },
]);

// 9. Groq Call 2
box(5.15, R.r4, 3.0, BH, [
  { text: 'Groq #2  —  llama-3.3-70b', bold: true, size: 9 },
  { text: 'Генерация финального ответа', size: 8 },
  { text: 'max_tokens: 600  ·  temp: 0.7', size: 7.5, color: C.light },
  { text: 'ai.js : 205', size: 7, color: C.lighter },
]);

// 10. Groq Server (bottom)
box(8.35, R.r4, 4.5, BH, [
  { text: 'groq.com API', bold: true, size: 10 },
  { text: '70B параметров  ·  ~0.3 сек', size: 8, color: C.light },
]);

// 11. Content check (diamond)
diamond(2.45, R.r5, 2.5, BH, [
  { text: 'Проверка', bold: true, size: 9 },
  { text: 'ответа', size: 8.5 },
  { text: 'containsDangerous()', size: 7, color: C.light },
]);

// 12. Response to user
box(0.12, R.r5, 2.0, BH, [
  { text: 'Ответ пользователю', bold: true, size: 9 },
  { text: 'текст + карточки товаров', size: 8 },
  { text: 'ChatWidget.jsx : 46', size: 7, color: C.lighter },
]);

// ── ARROWS ────────────────────────────────────────────────────────

const midBH = BH / 2;

// User → Jailbreak
arr(2.12, R.r1 + midBH, 2.45, R.r1 + midBH);
label(2.13, R.r1 + midBH - 0.22, 0.32, 0.2, 'запрос', { size: 7.5 });

// Jailbreak → Blocked (YES, down)
arr(2.7, R.r1 + BH, 2.7, R.r2);
label(2.75, R.r1 + BH + 0.05, 0.5, 0.2, 'ДА', { size: 7.5, bold: true, align: 'left' });

// Jailbreak → Groq#1 (NO, right)
arr(4.95, R.r1 + midBH, 5.15, R.r1 + midBH);
label(4.75, R.r1 + midBH - 0.22, 0.4, 0.2, 'НЕТ', { size: 7.5, bold: true });

// Groq#1 → Groq Server (right)
arr(8.15, R.r1 + midBH, 8.35, R.r1 + midBH);
label(8.0, R.r1 + midBH - 0.22, 0.35, 0.2, 'запрос', { size: 7 });

// Groq Server → Groq#1 (return, dashed)
arr(8.35, R.r1 + BH * 0.65, 8.15, R.r1 + BH * 0.65, { dash: true, color: C.lineLight });
label(8.0, R.r1 + BH * 0.65, 0.35, 0.2, 'термины', { size: 7 });

// Groq#1 → Synonyms (down)
arr(6.65, R.r1 + BH, 6.65, R.r2);

// Synonyms → SQLite (right)
arr(8.15, R.r2 + midBH, 8.35, R.r2 + midBH);
label(7.85, R.r2 + midBH - 0.22, 0.65, 0.2, 'SQL LIKE', { size: 7 });

// SQLite → Synonyms (return, dashed)
arr(8.35, R.r2 + BH * 0.65, 8.15, R.r2 + BH * 0.65, { dash: true, color: C.lineLight });
label(7.85, R.r2 + BH * 0.65, 0.7, 0.2, 'до 8 товаров', { size: 7 });

// Synonyms → Build Prompt (down)
arr(6.65, R.r2 + BH, 6.65, R.r3);

// Build Prompt → Groq#2 (down)
arr(6.65, R.r3 + BH, 6.65, R.r4);

// Groq#2 → Groq Server bottom (right)
arr(8.15, R.r4 + midBH, 8.35, R.r4 + midBH);
label(8.0, R.r4 + midBH - 0.22, 0.35, 0.2, 'запрос', { size: 7 });

// Groq Server → Groq#2 (return, dashed)
arr(8.35, R.r4 + BH * 0.65, 8.15, R.r4 + BH * 0.65, { dash: true, color: C.lineLight });
label(7.55, R.r4 + BH * 0.65, 1.0, 0.2, 'текст + IDs', { size: 7 });

// Groq#2 → Content check
// go left from Groq#2 bottom center to content check
arr(5.15, R.r4 + BH * 0.75, 4.95, R.r4 + BH * 0.75);
// vertical down
slide.addShape(prs.ShapeType.line, {
  x: 4.94, y: R.r4 + BH * 0.75, w: 0.01, h: R.r5 + midBH - (R.r4 + BH * 0.75),
  line: { color: C.lineDark, width: 1.2, endArrowType: 'none' },
});
// horizontal to diamond center
arr(4.95, R.r5 + midBH, 4.95, R.r5 + midBH, { noHead: true });
slide.addShape(prs.ShapeType.line, {
  x: 4.95, y: R.r5 + midBH, w: 0.01, h: 0.01,
  line: { color: C.lineDark, width: 1.2, endArrowType: 'arrow' },
});
// simple: just draw vertical then horizontal
// Actually let's just do a direct arrow from bottom of Groq#2 to top of content check diamond
// The diamond center is at (2.45 + 2.5/2, R.r5 + BH/2) = (3.7, R.r5 + 0.36)
// Let's route: down from Groq#2 bottom, then left
const gx2cx = 5.15 + 3.0 / 2; // center x of groq#2 = 6.65
const g2bottom = R.r4 + BH;
const checkCenterX = 2.45 + 2.5 / 2; // 3.7
const checkTop = R.r5;

// Vertical segment down
slide.addShape(prs.ShapeType.line, {
  x: gx2cx, y: g2bottom, w: 0.01, h: checkTop - g2bottom,
  line: { color: C.lineDark, width: 1.2, endArrowType: 'none' },
});
// Horizontal segment left
slide.addShape(prs.ShapeType.line, {
  x: checkCenterX, y: checkTop, w: gx2cx - checkCenterX, h: 0.01,
  line: { color: C.lineDark, width: 1.2, endArrowType: 'none', beginArrowType: 'none' },
  flipH: true,
});
// Arrow down into diamond
arr(checkCenterX, checkTop, checkCenterX, R.r5);

// Content check → Response (safe, left)
arr(2.45, R.r5 + midBH, 2.12, R.r5 + midBH);
label(2.13, R.r5 + midBH - 0.22, 0.6, 0.2, 'безопасно', { size: 7.5 });

// Content check → blocked (NO path, down label)
slide.addShape(prs.ShapeType.line, {
  x: 3.7, y: R.r5 + BH, w: 0.01, h: 0.25,
  line: { color: C.lineLight, width: 1, endArrowType: 'none', dashType: 'dash' },
});
label(3.75, R.r5 + BH + 0.04, 1.0, 0.2, 'опасно → блок', { size: 7.5, align: 'left', color: C.light });

// History loop: left side vertical dashed
slide.addShape(prs.ShapeType.line, {
  x: 0.5, y: R.r5, w: 0.01, h: -(R.r5 - R.r1 - BH),
  line: { color: C.lineLight, width: 1, dashType: 'lgDash', endArrowType: 'none' },
  flipV: true,
});
slide.addShape(prs.ShapeType.line, {
  x: 0.5, y: R.r1 + BH, w: 4.65, h: 0.01,
  line: { color: C.lineLight, width: 1, dashType: 'lgDash', endArrowType: 'arrow' },
});
label(1.5, R.r2 + 0.1, 1.8, 0.22, 'история диалога (8 сообщений)', { size: 7, color: C.lighter });

// ── LEGEND ───────────────────────────────────────────────────────
const lx = 0.12, ly = 2.15, lw = 2.0, lh2 = 2.45;
box(lx, ly, lw, lh2, null, { fill: 'F8F8F8', brd: C.lighter });
slide.addText('Обозначения', {
  x: lx + 0.1, y: ly + 0.08, w: lw - 0.2, h: 0.28,
  fontSize: 9, bold: true, color: C.dark, fontFace: F,
});

const legends = [
  { dash: false, txt: 'прямой запрос' },
  { dash: true,  txt: 'ответ / возврат' },
];
legends.forEach((l, i) => {
  const y = ly + 0.42 + i * 0.38;
  slide.addShape(prs.ShapeType.line, {
    x: lx + 0.15, y: y + 0.1, w: 0.55, h: 0,
    line: { color: C.lineDark, width: 1.2, dashType: l.dash ? 'dash' : 'solid', endArrowType: 'arrow' },
  });
  slide.addText(l.txt, {
    x: lx + 0.78, y: y, w: lw - 0.9, h: 0.26,
    fontSize: 8, color: C.mid, fontFace: F,
  });
});

const shapes = [
  { shape: 'roundRect', txt: 'блок обработки' },
  { shape: 'diamond',   txt: 'условие / проверка' },
];
shapes.forEach((s, i) => {
  const y = ly + 1.25 + i * 0.52;
  if (s.shape === 'roundRect') {
    slide.addShape(prs.ShapeType.roundRect, {
      x: lx + 0.12, y, w: 0.55, h: 0.28,
      fill: { color: C.blockFill }, line: { color: C.border, width: 1 }, rectRadius: 0.05,
    });
  } else {
    slide.addShape('diamond', {
      x: lx + 0.12, y, w: 0.55, h: 0.28,
      fill: { color: C.blockFill }, line: { color: C.border, width: 1 },
    });
  }
  slide.addText(s.txt, {
    x: lx + 0.75, y: y + 0.02, w: lw - 0.88, h: 0.26,
    fontSize: 8, color: C.mid, fontFace: F,
  });
});

// ── WRITE ─────────────────────────────────────────────────────────
const out = 'C:\\Users\\Никита\\figa\\portfolio\\Флип_Схема_ИИ.pptx';
prs.writeFile({ fileName: out }).then(() => console.log('DONE:', out)).catch(console.error);
