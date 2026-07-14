const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_WIDE'; // 13.33 × 7.5 inches

const slide = prs.addSlide();
slide.background = { color: 'FFFFFF' };

const F = 'Calibri';
const SW = 13.33;

// ── Palette ───────────────────────────────────────────────────────────
const C = {
  dark:   '222222',
  mid:    '555555',
  gray:   '888888',
  light:  'AAAAAA',
  faint:  'CCCCCC',
  fillD:  'EBEBEB',  // default box fill
  fillL:  'E4E4E4',  // left branch
  fillR:  'EEEEEE',  // right branch
  fillM:  'DEDEDE',  // merge box
  brd:    '999999',
  brdH:   '666666',  // heavier border
};

// ── Layout constants ──────────────────────────────────────────────────
const BW  = 2.9;   // box width
const BH  = 0.60;  // box height
const DW  = 3.4;   // diamond width
const DH  = 0.74;  // diamond height
const GAP = 0.28;  // vertical gap between boxes (arrow space)

const CX   = SW / 2;            // 6.665  — slide centre
const LCX  = 2.85;              // left branch centre X
const RCX  = SW - LCX;         // 10.48  — right branch centre X

const MBX  = CX  - BW / 2;     // 5.215
const LBX  = LCX - BW / 2;     // 0.40
const RBX  = RCX - BW / 2;     // 8.03

// Rows
const R1   = 0.68;
const R2   = R1  + BH + GAP;   // 1.56
const DY   = R2  + BH + GAP;   // 2.44   diamond top
const R4   = DY  + DH + 0.46;  // 3.64   first branch row
const R5   = R4  + BH + GAP;   // 4.52
const R6   = R5  + BH + GAP;   // 5.40
const COLY = R6  + BH + 0.30;  // 6.28   collector line Y
const R7   = COLY + 0.18;      // 6.46   merge box top

// ── Helpers ───────────────────────────────────────────────────────────

// Draw a rounded box with up to 3 text lines
function box(x, y, w, h, lines, fill, brd) {
  slide.addShape(prs.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: fill || C.fillD },
    line: { color: brd  || C.brd,  width: 1.1 },
    rectRadius: 0.08,
  });
  const rows = Array.isArray(lines) ? lines : [lines];
  const rowH = (h - 0.10) / rows.length;
  rows.forEach((row, i) => {
    const r = typeof row === 'string' ? { text: row } : row;
    slide.addText(r.text, {
      x: x + 0.12, y: y + 0.05 + i * rowH, w: w - 0.24, h: rowH,
      fontSize: r.size  || 10,
      bold:     r.bold  || false,
      color:    r.color || C.dark,
      fontFace: F, align: 'center', valign: 'middle',
    });
  });
}

// Draw a diamond (decision shape) with text constrained to inner 50%
function diamond(cx, y, w, h, title) {
  slide.addShape('diamond', {
    x: cx - w / 2, y, w, h,
    fill: { color: C.fillD },
    line: { color: C.brdH, width: 1.3 },
  });
  slide.addText(title, {
    x: cx - w * 0.30, y: y + h * 0.28, w: w * 0.60, h: h * 0.44,
    fontSize: 10, bold: true, color: C.dark, fontFace: F,
    align: 'center', valign: 'middle',
  });
}

// Small label tag above or beside a box
function tag(x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontSize: opts.size  || 8,
    bold:     opts.bold  || false,
    color:    opts.color || C.gray,
    fontFace: F,
    align:    opts.align || 'center',
    valign:   'middle',
    italic:   opts.italic || false,
  });
}

// Vertical arrow (downward)
function vArrow(x, y1, y2, noHead) {
  slide.addShape(prs.ShapeType.line, {
    x, y: y1, w: 0.01, h: y2 - y1,
    line: {
      color: C.mid, width: 1.2,
      endArrowType: noHead ? 'none' : 'arrow',
    },
  });
}

// Horizontal line (left→right always, use flipH for right→left)
function hLine(x1, x2, y, withArrow) {
  const goLeft = x2 < x1;
  slide.addShape(prs.ShapeType.line, {
    x: Math.min(x1, x2), y, w: Math.abs(x2 - x1), h: 0.01,
    line: {
      color: C.mid, width: 1.2,
      endArrowType: withArrow ? 'arrow' : 'none',
    },
    flipH: goLeft,
  });
}

// ── TITLE ─────────────────────────────────────────────────────────────
slide.addText('User Flow — Маркетплейс Флип', {
  x: 0.5, y: 0.10, w: 12.33, h: 0.40,
  fontSize: 18, bold: true, color: C.dark, fontFace: F, align: 'center',
});
slide.addText('Путь пользователя: от входа на сайт до выполнения заказа', {
  x: 0.5, y: 0.50, w: 12.33, h: 0.20,
  fontSize: 9, color: C.gray, fontFace: F, align: 'center',
});

// ── MAIN TRUNK ────────────────────────────────────────────────────────

// R1 — Главная страница
box(MBX, R1, BW, BH, [
  { text: 'Главная страница', bold: true, size: 10 },
  { text: 'Просмотр каталога, поиск товаров', size: 8.5, color: C.mid },
]);

vArrow(CX, R1 + BH, R2);

// R2 — Регистрация / Вход
box(MBX, R2, BW, BH, [
  { text: 'Регистрация / Вход', bold: true, size: 10 },
  { text: 'Email + пароль, создание аккаунта', size: 8.5, color: C.mid },
]);

vArrow(CX, R2 + BH, DY);

// Diamond — Выбор типа продавца
diamond(CX, DY, DW, DH, 'Тип продавца?');

// ── BRANCH LINES FROM DIAMOND ─────────────────────────────────────────
const DLP_X = CX - DW / 2;   // diamond left  point X = 4.965
const DRP_X = CX + DW / 2;   // diamond right point X = 8.365
const DMY   = DY + DH / 2;   // diamond mid   Y

// Left label above the horizontal branch line
tag(LCX - 1.2, DMY - 0.24, 2.4, 0.22,
  'ЧАСТНЫЙ ПРОДАВЕЦ', { size: 8, bold: true, color: C.mid, align: 'center' });

// Right label
tag(RCX - 1.2, DMY - 0.24, 2.4, 0.22,
  'КОМПАНИЯ / МАГАЗИН', { size: 8, bold: true, color: C.mid, align: 'center' });

// Horizontal line: diamond left → LCX (going left → flipH)
hLine(DLP_X, LCX, DMY, false);

// Vertical arrow down to R4 left branch
vArrow(LCX, DMY, R4);

// Horizontal line: diamond right → RCX
hLine(DRP_X, RCX, DMY, false);

// Vertical arrow down to R4 right branch
vArrow(RCX, DMY, R4);

// ── DIVIDER LINE between branches (faint dashed) ─────────────────────
slide.addShape(prs.ShapeType.line, {
  x: CX, y: DMY, w: 0.01, h: COLY - DMY,
  line: { color: C.faint, width: 0.8, dashType: 'sysDash', endArrowType: 'none' },
});

// ── LEFT BRANCH — Частный продавец ───────────────────────────────────
const leftRows = [
  [
    { text: 'Профиль частника',        bold: true, size: 10 },
    { text: 'Имя, город, описание',    size: 8.5, color: C.mid },
  ],
  [
    { text: 'Добавить объявление',               bold: true, size: 10 },
    { text: 'Название, цена, фото, состояние (б/у)',  size: 8.5, color: C.mid },
  ],
  [
    { text: 'Объявление активно',                bold: true, size: 10 },
    { text: 'Видно в каталоге покупателям',      size: 8.5, color: C.mid },
  ],
];

[R4, R5, R6].forEach((y, i) => {
  box(LBX, y, BW, BH, leftRows[i], C.fillL);
  vArrow(LCX, y + BH, [R5, R6, COLY][i], i < 2 ? false : true);
});

// ── RIGHT BRANCH — Компания ───────────────────────────────────────────
const rightRows = [
  [
    { text: 'Профиль компании',                       bold: true, size: 10 },
    { text: 'Название магазина, описание, лого',       size: 8.5, color: C.mid },
  ],
  [
    { text: 'Dashboard продавца',                     bold: true, size: 10 },
    { text: 'Продажи, просмотры, активные товары',    size: 8.5, color: C.mid },
  ],
  [
    { text: 'Управление товарами',                    bold: true, size: 10 },
    { text: 'Добавить / редактировать / удалить',     size: 8.5, color: C.mid },
  ],
];

[R4, R5, R6].forEach((y, i) => {
  box(RBX, y, BW, BH, rightRows[i], C.fillR);
  vArrow(RCX, y + BH, [R5, R6, COLY][i], i < 2 ? false : true);
});

// ── COLLECTOR + MERGE ─────────────────────────────────────────────────
// Horizontal line from LCX to CX (no arrow)
hLine(LCX, CX, COLY, false);
// Horizontal line from RCX to CX (no arrow)
hLine(RCX, CX, COLY, false);
// Arrow down from CX to merge box
vArrow(CX, COLY, R7);

// Merge box
box(MBX, R7, BW, BH, [
  { text: 'Заказ выполнен',                          bold: true, size: 10 },
  { text: 'Статус обновлён · Покупатель оставил отзыв', size: 8.5, color: C.mid },
], C.fillM, C.brdH);

// ── LEGEND ────────────────────────────────────────────────────────────
// Tiny legend bottom-right
const LGX = 10.5, LGY = 3.8;
slide.addShape(prs.ShapeType.roundRect, {
  x: LGX, y: LGY, w: 2.58, h: 2.12,
  fill: { color: 'F8F8F8' }, line: { color: C.faint, width: 0.8 }, rectRadius: 0.07,
});
tag(LGX + 0.12, LGY + 0.10, 2.3, 0.22, 'Обозначения', { size: 8.5, bold: true, color: C.dark });

// Arrow line sample
slide.addShape(prs.ShapeType.line, {
  x: LGX + 0.18, y: LGY + 0.52, w: 0.52, h: 0.01,
  line: { color: C.mid, width: 1.2, endArrowType: 'arrow' },
});
tag(LGX + 0.80, LGY + 0.40, 1.65, 0.26, 'Переход / шаг', { size: 8, align: 'left' });

// Box sample
slide.addShape(prs.ShapeType.roundRect, {
  x: LGX + 0.18, y: LGY + 0.84, w: 0.52, h: 0.26,
  fill: { color: C.fillD }, line: { color: C.brd, width: 0.9 }, rectRadius: 0.04,
});
tag(LGX + 0.80, LGY + 0.78, 1.65, 0.35, 'Блок обработки', { size: 8, align: 'left' });

// Diamond sample
slide.addShape('diamond', {
  x: LGX + 0.18, y: LGY + 1.22, w: 0.52, h: 0.30,
  fill: { color: C.fillD }, line: { color: C.brdH, width: 1.0 },
});
tag(LGX + 0.80, LGY + 1.18, 1.65, 0.35, 'Условие / ветвление', { size: 8, align: 'left' });

// Left branch fill sample
slide.addShape(prs.ShapeType.roundRect, {
  x: LGX + 0.18, y: LGY + 1.64, w: 0.52, h: 0.26,
  fill: { color: C.fillL }, line: { color: C.brd, width: 0.9 }, rectRadius: 0.04,
});
tag(LGX + 0.80, LGY + 1.58, 1.65, 0.35, 'Ветка: частник', { size: 8, align: 'left' });

// Right branch fill sample
slide.addShape(prs.ShapeType.roundRect, {
  x: LGX + 0.18, y: LGY + 1.94, w: 0.52, h: 0.22,
  fill: { color: C.fillR }, line: { color: C.brd, width: 0.9 }, rectRadius: 0.04,
});
tag(LGX + 0.80, LGY + 1.88, 1.65, 0.35, 'Ветка: компания', { size: 8, align: 'left' });

// ── WRITE ─────────────────────────────────────────────────────────────
const out = 'C:\\Users\\Никита\\figa\\portfolio\\Флип_UserFlow.pptx';
prs.writeFile({ fileName: out })
  .then(() => console.log('DONE:', out))
  .catch(e => console.error('ERROR:', e.message));
