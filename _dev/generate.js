const fs = require('fs');

const raw = fs.readFileSync('./node_1864.json', 'utf8').replace(/^﻿/, '');
const data = JSON.parse(raw);
const rootDoc = data.nodes['1:864'].document;

const ROOT_BB = rootDoc.absoluteBoundingBox;
const ROOT_W = ROOT_BB.width;
const ROOT_H = ROOT_BB.height;

// ── Export overrides ─────────────────────────────────────────
// Render these nodes as a single exported PNG at NATURAL @1x size (no stretching).
// naturalW/H = exported PNG width/height ÷ 2 (exported at @2x).
// xOffset/yOffset = extra shift so the PNG content aligns with its actual visual origin.
//   1:880 – mask shape (Rectangle 27) is invisible; actual photo starts at abs x=489.526 → offset +114.526
//   1:885 – Vector 70 content starts at abs x=529 → offset +154
const EXPORT_AS_WHOLE_PNG = {
  '1:880': { file: '1-880.png', naturalW: 940, naturalH: 1258, xOffset: 0, yOffset: 0 },
  '1:885': { file: '1-885.png', naturalW: 609, naturalH: 910,  xOffset: 0, yOffset: 0 },
};

// Natural SVG dimensions (from actual SVG width/height attributes).
// Used instead of absoluteBoundingBox size to avoid distortion.
const SVG_NATURAL = {
  '1:876': [728,  1397],
  '1:878': [658,  1453],
  '1:896': [1770, 713 ],
  '1:897': [1764, 714 ],
  '1:898': [1721, 671 ],
  '1:899': [1681, 831 ],
  '1:900': [1681, 831 ],
};

// Render these nodes as SVG img (exported from Figma).
// Covers border rectangles with gradient strokes + BOOLEAN_OPERATION frames.
const EXPORT_AS_SVG = {
  '1:868':  { file: '1-868.svg'  }, // title text
  '1:1026': { file: '1-1026.svg' }, // BOOLEAN_OPERATION Subtract (section 20 left)
  '1:1029': { file: '1-1029.svg' }, // BOOLEAN_OPERATION Subtract (section 20 right)
  '1:1043': { file: '1-1043.svg' }, // border rect section 19
  '1:1044': { file: '1-1044.svg' }, // border rect section 19
  '1:1053': { file: '1-1053.svg' }, // border rect section 17
  '1:1058': { file: '1-1058.svg' }, // border rect section 16
  '1:1060': { file: '1-1060.svg' }, // border rect section 15
  '1:1069': { file: '1-1069.svg' }, // border rect section 13
  '1:1076': { file: '1-1076.svg' }, // border rect section 10
  '1:1081': { file: '1-1081.svg' }, // border rect section 9
  '1:1101': { file: '1-1101.svg' }, // border rect section 6
  '1:1110': { file: '1-1110.svg' }, // border rect section 4
  '1:1117': { file: '1-1117.svg' }, // border rect section 3
  '1:1123': { file: '1-1123.svg' }, // border rect section 2
  '1:1126': { file: '1-1126.svg' }, // border rect section 1
};

// Per-node font overrides (Figma API may return a fallback/wrong family name)
const FONT_OVERRIDES = {
  '1:891': { fontFamily: 'Stem', fontWeight: 300 }, // bio text — Stem Light
};

// ── Helpers ──────────────────────────────────────────────────
function colorToCss(color, opacity) {
  if (!color) return 'transparent';
  const r = Math.round(color.r * 255);
  const g = Math.round(color.g * 255);
  const b = Math.round(color.b * 255);
  const a = Math.min(1, (opacity !== undefined ? opacity : 1) * (color.a !== undefined ? color.a : 1));
  return `rgba(${r},${g},${b},${a.toFixed(3)})`;
}

function blendModeToCSS(bm) {
  const map = {
    NORMAL:'normal',MULTIPLY:'multiply',SCREEN:'screen',OVERLAY:'overlay',
    DARKEN:'darken',LIGHTEN:'lighten',COLOR_DODGE:'color-dodge',COLOR_BURN:'color-burn',
    HARD_LIGHT:'hard-light',SOFT_LIGHT:'soft-light',DIFFERENCE:'difference',
    EXCLUSION:'exclusion',HUE:'hue',SATURATION:'saturation',COLOR:'color',
    LUMINOSITY:'luminosity',LINEAR_BURN:'multiply',LINEAR_DODGE:'screen',PASS_THROUGH:'normal'
  };
  return map[bm] || 'normal';
}

function getFills(node) {
  if (!node.fills || !node.fills.length) return '';
  const visible = node.fills.filter(f => f.visible !== false);
  if (!visible.length) return '';
  const parts = [];
  for (const fill of [...visible].reverse()) {
    if (fill.type === 'SOLID') {
      parts.push(colorToCss(fill.color, fill.opacity));
    } else if (fill.type === 'IMAGE') {
      const fn = node.id.replace(/:/g, '-') + '.png';
      parts.push(`url('./assets/${fn}')`);
    } else if (fill.type === 'GRADIENT_LINEAR') {
      let angle = 180;
      if (fill.gradientHandlePositions && fill.gradientHandlePositions.length >= 2) {
        const [p0, p1] = fill.gradientHandlePositions;
        angle = Math.round(Math.atan2(p1.x - p0.x, -(p1.y - p0.y)) * 180 / Math.PI);
      }
      const stops = fill.gradientStops.map(s => `${colorToCss(s.color)} ${(s.position*100).toFixed(1)}%`).join(', ');
      parts.push(`linear-gradient(${angle}deg, ${stops})`);
    } else if (fill.type === 'GRADIENT_RADIAL') {
      const stops = fill.gradientStops.map(s => `${colorToCss(s.color)} ${(s.position*100).toFixed(1)}%`).join(', ');
      parts.push(`radial-gradient(ellipse at center, ${stops})`);
    }
  }
  return parts.join(', ');
}

function getEffects(node) {
  const result = { boxShadow: '', filter: '' };
  if (!node.effects || !node.effects.length) return result;
  const shadows = [];
  let filters = '';
  for (const e of node.effects) {
    if (e.visible === false) continue;
    if (e.type === 'DROP_SHADOW')
      shadows.push(`${e.offset.x}px ${e.offset.y}px ${e.radius}px ${e.spread||0}px ${colorToCss(e.color)}`);
    else if (e.type === 'INNER_SHADOW')
      shadows.push(`inset ${e.offset.x}px ${e.offset.y}px ${e.radius}px ${e.spread||0}px ${colorToCss(e.color)}`);
    else if (e.type === 'LAYER_BLUR')
      filters += `blur(${e.radius}px) `;
  }
  result.boxShadow = shadows.join(', ');
  result.filter = filters.trim();
  return result;
}

function getBorderRadius(node) {
  if (node.rectangleCornerRadii) {
    const [tl, tr, br, bl] = node.rectangleCornerRadii;
    if (tl === tr && tr === br && br === bl) return `${tl}px`;
    return `${tl}px ${tr}px ${br}px ${bl}px`;
  }
  if (node.cornerRadius) return `${node.cornerRadius}px`;
  return '';
}

function getStrokeShadow(node) {
  if (!node.strokes || !node.strokes.length) return '';
  const stroke = node.strokes.find(s => s.visible !== false);
  if (!stroke || !stroke.color) return '';
  const color = colorToCss(stroke.color, stroke.opacity);
  const w = node.strokeWeight || 1;
  const align = node.strokeAlign || 'INSIDE';
  if (align === 'INSIDE') return `inset 0 0 0 ${w}px ${color}`;
  if (align === 'OUTSIDE') return `0 0 0 ${w}px ${color}`;
  return `inset 0 0 0 ${w/2}px ${color}, 0 0 0 ${w/2}px ${color}`;
}

// ── Node renderer ─────────────────────────────────────────────
function renderNode(node, parentBB) {
  if (node.visible === false) return '';
  const bb = node.absoluteBoundingBox;
  if (!bb) return '';

  const x = bb.x - parentBB.x;
  const y = bb.y - parentBB.y;
  const w = bb.width;
  const h = bb.height;
  const type = node.type;
  const opacity = node.opacity !== undefined ? node.opacity : 1;
  const blendMode = blendModeToCSS(node.blendMode);
  const effects = getEffects(node);
  const strokeShadow = getStrokeShadow(node);
  const boxShadows = [effects.boxShadow, strokeShadow].filter(Boolean).join(', ');

  // ── SVG export override ──
  if (EXPORT_AS_SVG[node.id]) {
    const { file } = EXPORT_AS_SVG[node.id];
    let css = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;`;
    if (opacity !== 1) css += `opacity:${opacity};`;
    if (blendMode !== 'normal') css += `mix-blend-mode:${blendMode};`;
    return `<img src="./assets/${file}" style="${css}" alt="">`;
  }

  // ── Whole-PNG export override (natural @1x size, no stretch) ──
  if (EXPORT_AS_WHOLE_PNG[node.id]) {
    const { file, naturalW, naturalH, xOffset = 0, yOffset = 0 } = EXPORT_AS_WHOLE_PNG[node.id];
    const px = x + xOffset;
    const py = y + yOffset;
    let css = `position:absolute;left:${px}px;top:${py}px;width:${naturalW}px;height:${naturalH}px;`;
    if (opacity !== 1) css += `opacity:${opacity};`;
    if (blendMode !== 'normal') css += `mix-blend-mode:${blendMode};`;
    if (boxShadows) css += `box-shadow:${boxShadows};`;
    return `<img src="./assets/${file}" style="${css}" alt="">`;
  }

  // ── TEXT ──
  if (type === 'TEXT') {
    const style = node.style || {};
    const fills = (node.fills || []).filter(f => f.visible !== false);
    const color = fills.length && fills[0].color
      ? colorToCss(fills[0].color, fills[0].opacity) : '#fff';
    const fo = FONT_OVERRIDES[node.id] || {};
    const fontFamily = fo.fontFamily || style.fontFamily || 'sans-serif';
    const fontSize = style.fontSize || 16;
    const fontWeight = fo.fontWeight || style.fontWeight || 400;
    const lineHeight = style.lineHeightPx ? `${style.lineHeightPx}px`
      : (style.lineHeightUnit === 'PERCENT' ? `${style.lineHeightPercentFontSize||120}%` : 'normal');
    const letterSpacing = style.letterSpacing ? `${style.letterSpacing}px` : 'normal';
    const textAlign = (style.textAlignHorizontal || 'LEFT').toLowerCase();
    const textTransform = style.textCase === 'UPPER' ? 'uppercase'
      : style.textCase === 'LOWER' ? 'lowercase' : 'none';
    const fontStyle = style.italic ? 'italic' : 'normal';
    const textDecoration = style.textDecoration === 'UNDERLINE' ? 'underline'
      : style.textDecoration === 'STRIKETHROUGH' ? 'line-through' : 'none';

    let css = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;`;
    css += `color:${color};font-family:'${fontFamily}',sans-serif;font-size:${fontSize}px;font-weight:${fontWeight};`;
    css += `line-height:${lineHeight};letter-spacing:${letterSpacing};text-align:${textAlign};`;
    css += `text-transform:${textTransform};font-style:${fontStyle};text-decoration:${textDecoration};`;
    css += `opacity:${opacity};mix-blend-mode:${blendMode};overflow:hidden;white-space:pre-wrap;`;
    if (effects.filter) css += `filter:${effects.filter};`;

    const content = (node.characters || '')
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    return `<div style="${css}">${content}</div>`;
  }

  // ── VECTOR ──
  if (type === 'VECTOR') {
    const fn = node.id.replace(/:/g,'-') + '.svg';
    const svgExists = fs.existsSync(`./assets/${fn}`);
    // Use natural SVG dimensions if known — avoids distortion from stretching to bb size
    const nat = SVG_NATURAL[node.id];
    const svgW = nat ? nat[0] : w;
    const svgH = nat ? nat[1] : h;
    let css = `position:absolute;left:${x}px;top:${y}px;width:${svgW}px;height:${svgH}px;`;
    if (opacity !== 1) css += `opacity:${opacity};`;
    if (blendMode !== 'normal') css += `mix-blend-mode:${blendMode};`;
    if (effects.filter) css += `filter:${effects.filter};`;
    if (svgExists)
      return `<img src="./assets/${fn}" style="${css}" alt="">`;
    const bg = getFills(node);
    return `<div style="${css}${bg ? `background:${bg};` : ''}"></div>`;
  }

  // ── LINE ──
  if (type === 'LINE') {
    const stroke = node.strokes && node.strokes[0];
    const color = stroke ? colorToCss(stroke.color) : '#fff';
    const weight = node.strokeWeight || 1;
    let css = `position:absolute;left:${x}px;top:${y}px;width:${Math.max(w,1)}px;height:${weight}px;background:${color};`;
    if (opacity !== 1) css += `opacity:${opacity};`;
    if (blendMode !== 'normal') css += `mix-blend-mode:${blendMode};`;
    return `<div style="${css}"></div>`;
  }

  // ── BOOLEAN_OPERATION ──
  if (type === 'BOOLEAN_OPERATION') {
    const fn = node.id.replace(/:/g,'-') + '.png';
    let css = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;`;
    if (opacity !== 1) css += `opacity:${opacity};`;
    if (blendMode !== 'normal') css += `mix-blend-mode:${blendMode};`;
    if (boxShadows) css += `box-shadow:${boxShadows};`;
    if (fs.existsSync(`./assets/${fn}`))
      css += `background:url('./assets/${fn}') center/contain no-repeat;`;
    return `<div style="${css}"></div>`;
  }

  // ── FRAME / GROUP / RECTANGLE / ELLIPSE ──
  const bg = getFills(node);
  const hasImageFill = (node.fills || []).some(f => f.type === 'IMAGE');
  const borderRadius = getBorderRadius(node);

  let css = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;`;
  if (opacity !== 1) css += `opacity:${opacity};`;
  if (blendMode !== 'normal') css += `mix-blend-mode:${blendMode};`;
  if (boxShadows) css += `box-shadow:${boxShadows};`;
  if (effects.filter) css += `filter:${effects.filter};`;
  if (bg) {
    css += `background:${bg};`;
    if (hasImageFill) css += `background-size:cover;background-position:center;`;
  }
  if (type === 'ELLIPSE') css += `border-radius:50%;`;
  else if (borderRadius) css += `border-radius:${borderRadius};`;

  const clips = type === 'FRAME' ? (node.clipsContent !== false) : (node.clipsContent === true);
  if (clips) css += `overflow:hidden;`;

  let children = '';
  if (node.children) {
    for (const child of node.children) {
      children += renderNode(child, bb);
    }
  }
  return `<div style="${css}">${children}</div>`;
}

// ── Build HTML ────────────────────────────────────────────────
let body = '';
for (const child of (rootDoc.children || [])) {
  body += renderNode(child, ROOT_BB);
}

const rootBg = getFills(rootDoc) || 'rgb(18,19,22)';

const fonts = `
@font-face {
  font-family: 'Druk Text Wide';
  src: url('./assets/DrukTextWide-Super.ttf') format('truetype');
  font-weight: 100 900;
  font-style: normal;
}
@font-face {
  font-family: 'Disruptors Script';
  src: url('./assets/Disruptors-Script.otf') format('opentype');
  font-weight: 100 900;
  font-style: normal;
}
@font-face {
  font-family: 'Stem';
  src: url('./assets/Stem-Light.ttf') format('truetype');
  font-weight: 300;
  font-style: normal;
}
@font-face {
  font-family: 'Stem';
  src: url('./assets/Stem-Light.ttf') format('truetype');
  font-weight: 350 500;
  font-style: normal;
}
@font-face {
  font-family: 'Helvetica';
  src: url('./assets/helvetica_light.otf') format('opentype');
  font-weight: 300 400;
  font-style: normal;
}`;

const html = `<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Портфолио — Смирнов Никита</title>
<style>
${fonts}
* { margin: 0; padding: 0; box-sizing: border-box; }
body { background: #000; }
.canvas {
  position: relative;
  width: ${ROOT_W}px;
  height: ${ROOT_H}px;
  background: ${rootBg};
  overflow: hidden;
  margin: 0 auto;
}
</style>
</head>
<body>
<div class="canvas">
${body}
</div>
</body>
</html>`;

fs.writeFileSync('./index.html', html, 'utf8');
console.log('HTML generated:', html.length, 'chars');
console.log('Canvas:', ROOT_W + 'x' + ROOT_H + 'px');
