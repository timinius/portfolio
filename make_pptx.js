const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();

prs.layout = 'LAYOUT_WIDE';
prs.title = 'Флип — Маркетплейс';

const BG = '#FFFFFF';
const DARK = '#1A1A1A';
const BLUE = '#2563EB';
const GREEN = '#16A34A';
const PURPLE = '#7C3AED';
const ORANGE = '#EA580C';
const GRAY = '#6B7280';
const LIGHTBG = '#F8F9FA';
const LIGHTBLUE = '#EFF6FF';
const LIGHTGREEN = '#F0FDF4';
const LIGHTPURPLE = '#F5F3FF';

const fontMain = 'Calibri';

function addSlide(title, subtitle, fn) {
  const slide = prs.addSlide();
  slide.background = { color: BG };

  // Top accent line
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.06, fill: { color: DARK } });

  // Slide number in top right
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.4, y: 0.18, w: 10, h: 0.35,
      fontSize: 10, color: GRAY, fontFace: fontMain, bold: false,
    });
  }

  slide.addText(title, {
    x: 0.4, y: 0.5, w: 12, h: 0.7,
    fontSize: 30, bold: true, color: DARK, fontFace: fontMain,
  });

  fn(slide);
  return slide;
}

function box(slide, x, y, w, h, fill, borderColor) {
  slide.addShape(prs.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: fill || LIGHTBG },
    line: { color: borderColor || '#E5E7EB', width: 1 },
    rectRadius: 0.1,
  });
}

// ─── SLIDE 1: TITLE ───────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: DARK };
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: DARK } });

  slide.addText('Флип', {
    x: 0.6, y: 1.2, w: 12, h: 1.4,
    fontSize: 72, bold: true, color: '#FFFFFF', fontFace: fontMain,
  });
  slide.addText('Маркетплейс нового поколения', {
    x: 0.6, y: 2.7, w: 10, h: 0.6,
    fontSize: 22, color: '#9CA3AF', fontFace: fontMain,
  });
  slide.addText('B2B + B2C · React + Node.js + SQLite · Railway + Groq AI', {
    x: 0.6, y: 3.3, w: 11, h: 0.45,
    fontSize: 14, color: '#6B7280', fontFace: fontMain,
  });

  // Steps bar
  const steps = ['01. Анализ', '02. Дизайн', '03. Реализация', '04. Деплой', '05. Оценка'];
  const colors = [BLUE, GREEN, PURPLE, ORANGE, '#0891B2'];
  steps.forEach((s, i) => {
    slide.addShape(prs.ShapeType.roundRect, {
      x: 0.6 + i * 2.6, y: 4.2, w: 2.3, h: 0.55,
      fill: { color: colors[i] }, line: { color: colors[i] }, rectRadius: 0.08,
    });
    slide.addText(s, {
      x: 0.6 + i * 2.6, y: 4.2, w: 2.3, h: 0.55,
      fontSize: 11, bold: true, color: '#FFFFFF', fontFace: fontMain, align: 'center', valign: 'middle',
    });
  });
}

// ─── SLIDE 2: АНАЛИЗ КОНКУРЕНТОВ ──────────────────────────────────
addSlide('01. Анализ конкурентов', 'Исследование рынка · выбор ниши · прототип', (slide) => {
  // Left: comparison table
  const tableData = [
    [{ text: 'Платформа', options: { bold: true, fill: { color: '#F3F4F6' } } },
     { text: 'Модель', options: { bold: true, fill: { color: '#F3F4F6' } } },
     { text: 'ИИ', options: { bold: true, fill: { color: '#F3F4F6' } } },
     { text: 'B2B', options: { bold: true, fill: { color: '#F3F4F6' } } },
     { text: 'B2C', options: { bold: true, fill: { color: '#F3F4F6' } } }],
    ['Авито', 'C2C', '✗', '✗', '✓'],
    ['Ozon', 'B2C', '✗', '✓', '✓'],
    ['Wildberries', 'B2C', '✗', '✓', '✓'],
    ['Юла', 'C2C', '✗', '✗', '✓'],
    [{ text: 'Флип 🏆', options: { bold: true } },
     { text: 'B2B+B2C', options: { bold: true } },
     { text: '✓', options: { bold: true, color: GREEN } },
     { text: '✓', options: { bold: true, color: GREEN } },
     { text: '✓', options: { bold: true, color: GREEN } }],
  ];

  slide.addTable(tableData, {
    x: 0.4, y: 1.4, w: 6.2, h: 2.6,
    fontSize: 12, fontFace: fontMain, color: DARK,
    border: { pt: 1, color: '#E5E7EB' },
    rowH: 0.38,
  });

  // Right: B2B+B2C explanation
  box(slide, 7.0, 1.4, 5.8, 2.6, LIGHTBLUE, BLUE);
  slide.addText('Почему B2B + B2C?', {
    x: 7.2, y: 1.5, w: 5.4, h: 0.4,
    fontSize: 14, bold: true, color: BLUE, fontFace: fontMain,
  });

  const reasons = [
    '→ Авито/Юла — только C2C, нет бизнес-аккаунтов',
    '→ Ozon/WB — высокий барьер входа для малого бизнеса',
    '→ Флип принимает всех: магазины и частников',
    '→ Магазины дают стабильный поток товаров',
    '→ Частники обеспечивают разнообразие каталога',
    '→ Единый кабинет: режим «компания» и «частник»',
  ];
  reasons.forEach((r, i) => {
    slide.addText(r, {
      x: 7.2, y: 1.95 + i * 0.33, w: 5.4, h: 0.3,
      fontSize: 11, color: DARK, fontFace: fontMain,
    });
  });

  // Bottom: Gaps found
  box(slide, 0.4, 4.15, 12.4, 1.2, LIGHTGREEN, GREEN);
  slide.addText('Незакрытые ниши конкурентов → наши возможности', {
    x: 0.6, y: 4.22, w: 12, h: 0.35,
    fontSize: 13, bold: true, color: GREEN, fontFace: fontMain,
  });
  const gaps = [
    { label: 'Нет ИИ ни у кого', desc: 'Groq AI для подбора товаров' },
    { label: 'Сложный онбординг', desc: 'Простая регистрация за 1 мин' },
    { label: 'Разделение аудиторий', desc: 'Один маркетплейс для всех' },
    { label: 'Нет персонализации', desc: 'Рекомендации на базе LLM' },
  ];
  gaps.forEach((g, i) => {
    slide.addText(`✗ ${g.label}`, { x: 0.7 + i * 3.1, y: 4.62, w: 2.9, h: 0.22, fontSize: 10, bold: true, color: '#DC2626', fontFace: fontMain });
    slide.addText(`✓ ${g.desc}`, { x: 0.7 + i * 3.1, y: 4.85, w: 2.9, h: 0.22, fontSize: 10, color: GREEN, fontFace: fontMain });
  });
});

// ─── SLIDE 3: ПРОЕКТИРОВАНИЕ ───────────────────────────────────────
addSlide('02. Проектирование', 'База данных · Figma · Компонентная архитектура', (slide) => {
  // DB Schema
  box(slide, 0.4, 1.4, 4.1, 2.1, '#FFFBEB', '#F59E0B');
  slide.addText('Схема БД (SQLite)', { x: 0.6, y: 1.48, w: 3.8, h: 0.35, fontSize: 13, bold: true, color: '#92400E', fontFace: fontMain });
  const tables = [
    { name: 'users', fields: 'id, name, email, password_hash, account_type, rating' },
    { name: 'products', fields: 'id, seller_id, title, price, category, condition' },
    { name: 'orders', fields: 'id, buyer_id, seller_id, product_id, status, total' },
    { name: 'cart_items', fields: 'id, user_id, product_id, quantity' },
    { name: 'favorites', fields: 'id, user_id, product_id' },
  ];
  tables.forEach((t, i) => {
    slide.addText(t.name, { x: 0.6, y: 1.9 + i * 0.32, w: 1.3, h: 0.28, fontSize: 10, bold: true, color: BLUE, fontFace: 'Courier New' });
    slide.addText(t.fields, { x: 1.95, y: 1.9 + i * 0.32, w: 2.4, h: 0.28, fontSize: 9, color: GRAY, fontFace: 'Courier New' });
  });

  // Frontend arch
  box(slide, 4.7, 1.4, 4.1, 2.1, LIGHTPURPLE, PURPLE);
  slide.addText('Frontend (React + Vite)', { x: 4.9, y: 1.48, w: 3.8, h: 0.35, fontSize: 13, bold: true, color: PURPLE, fontFace: fontMain });
  const fe = [
    ['Pages', '18 страниц, React Router'],
    ['Components', 'Header, ProductCard, AiChat'],
    ['Contexts', 'AuthContext — JWT глобально'],
    ['API Client', 'Axios + interceptors'],
    ['Стили', 'CSS переменные, Onest font'],
  ];
  fe.forEach(([k, v], i) => {
    slide.addText(k, { x: 4.9, y: 1.9 + i * 0.32, w: 1.3, h: 0.28, fontSize: 10, bold: true, color: PURPLE, fontFace: fontMain });
    slide.addText(v, { x: 6.25, y: 1.9 + i * 0.32, w: 2.4, h: 0.28, fontSize: 10, color: DARK, fontFace: fontMain });
  });

  // Backend arch
  box(slide, 9.0, 1.4, 4.2, 2.1, LIGHTGREEN, GREEN);
  slide.addText('Backend (Node.js + Express)', { x: 9.2, y: 1.48, w: 3.9, h: 0.35, fontSize: 13, bold: true, color: GREEN, fontFace: fontMain });
  const be = [
    ['Routes', '/api/products, users, orders...'],
    ['Auth', 'JWT, bcrypt, middleware'],
    ['DB', 'node-sqlite3-wasm, initDB()'],
    ['AI', 'POST /api/ai/chat → Groq'],
    ['Static', 'serve React build в prod'],
  ];
  be.forEach(([k, v], i) => {
    slide.addText(k, { x: 9.2, y: 1.9 + i * 0.32, w: 1.2, h: 0.28, fontSize: 10, bold: true, color: GREEN, fontFace: fontMain });
    slide.addText(v, { x: 10.45, y: 1.9 + i * 0.32, w: 2.6, h: 0.28, fontSize: 10, color: DARK, fontFace: fontMain });
  });

  // Design tokens
  box(slide, 0.4, 3.65, 12.8, 1.7, '#F9FAFB', '#E5E7EB');
  slide.addText('Figma — 18 экранов + дизайн-токены', { x: 0.6, y: 3.73, w: 12, h: 0.35, fontSize: 13, bold: true, color: DARK, fontFace: fontMain });
  const pages = ['Home', 'Catalog', 'Product', 'Cart', 'Checkout', 'Orders', 'Profile', 'Login', 'Register', 'Dashboard', 'Seller', 'Favorites', 'MyListings', 'AddProduct', 'EditProduct', 'OrderDetail', 'NotFound', 'AI Chat'];
  const colors2 = [BLUE, BLUE, BLUE, GREEN, GREEN, GREEN, PURPLE, PURPLE, PURPLE, ORANGE, ORANGE, '#0891B2', '#0891B2', '#0891B2', '#0891B2', DARK, DARK, PURPLE];
  pages.forEach((p, i) => {
    slide.addShape(prs.ShapeType.roundRect, { x: 0.55 + (i % 9) * 1.42, y: 4.15 + Math.floor(i / 9) * 0.5, w: 1.3, h: 0.36, fill: { color: colors2[i] }, line: { color: colors2[i] }, rectRadius: 0.06 });
    slide.addText(p, { x: 0.55 + (i % 9) * 1.42, y: 4.15 + Math.floor(i / 9) * 0.5, w: 1.3, h: 0.36, fontSize: 9, bold: true, color: '#FFFFFF', fontFace: fontMain, align: 'center', valign: 'middle' });
  });
});

// ─── SLIDE 4: ИИ-ПОМОЩНИК — GROQ ──────────────────────────────────
addSlide('03. Реализация ИИ-помощника', 'Groq API · llama-3.3-70b · Контекст из БД', (slide) => {
  // Why Groq table
  box(slide, 0.4, 1.4, 6.0, 2.2, LIGHTPURPLE, PURPLE);
  slide.addText('Почему Groq?', { x: 0.6, y: 1.48, w: 5.6, h: 0.38, fontSize: 14, bold: true, color: PURPLE, fontFace: fontMain });

  const groqData = [
    [{ text: 'Провайдер', options: { bold: true, fill: { color: '#EDE9FE' } } },
     { text: 'Скорость', options: { bold: true, fill: { color: '#EDE9FE' } } },
     { text: 'Цена', options: { bold: true, fill: { color: '#EDE9FE' } } },
     { text: 'Free tier', options: { bold: true, fill: { color: '#EDE9FE' } } }],
    ['OpenAI GPT-4', '~3–5 сек', '$$$', '✗'],
    ['Anthropic Claude', '~3–6 сек', '$$$', '✗'],
    ['Gemini', '~2–4 сек', '$$', '✓'],
    [{ text: 'Groq 🏆', options: { bold: true } },
     { text: '~0.3 сек ⚡', options: { bold: true, color: GREEN } },
     { text: 'Free', options: { bold: true, color: GREEN } },
     { text: '✓', options: { bold: true, color: GREEN } }],
  ];
  slide.addTable(groqData, {
    x: 0.5, y: 1.92, w: 5.7, h: 1.55,
    fontSize: 11, fontFace: fontMain, color: DARK,
    border: { pt: 1, color: '#DDD6FE' },
    rowH: 0.3,
  });

  // Groq advantages
  box(slide, 6.6, 1.4, 6.6, 2.2, '#FFF7ED', ORANGE);
  slide.addText('Преимущества Groq', { x: 6.8, y: 1.48, w: 6.2, h: 0.38, fontSize: 14, bold: true, color: ORANGE, fontFace: fontMain });
  const adv = [
    '⚡ LPU-чипы — скорость вывода в 10× быстрее GPU',
    '🆓 llama-3.3-70b полностью бесплатно на старте',
    '🔌 OpenAI-совместимый API — легкий переход',
    '🌍 Модель llama обучена на 70 млрд параметрах',
    '🔒 Данные не используются для обучения',
  ];
  adv.forEach((a, i) => {
    slide.addText(a, { x: 6.8, y: 1.95 + i * 0.33, w: 6.2, h: 0.3, fontSize: 11, color: DARK, fontFace: fontMain });
  });

  // AI Architecture flow
  box(slide, 0.4, 3.75, 12.8, 1.6, LIGHTBLUE, BLUE);
  slide.addText('Архитектура запроса к ИИ', { x: 0.6, y: 3.83, w: 12, h: 0.35, fontSize: 13, bold: true, color: BLUE, fontFace: fontMain });

  const flowItems = [
    { label: '👤 Пользователь', sub: 'Вопрос в чате', color: BLUE },
    { label: '⚙️ Express API', sub: 'POST /api/ai/chat', color: DARK },
    { label: '🗄️ SQLite', sub: 'Топ-товары из БД', color: GREEN },
    { label: '📝 System Prompt', sub: 'Контекст платформы', color: ORANGE },
    { label: '🤖 Groq API', sub: 'llama-3.3-70b', color: PURPLE },
    { label: '💬 Ответ', sub: '~0.3 сек пользователю', color: GREEN },
  ];
  flowItems.forEach((item, i) => {
    const x = 0.55 + i * 2.15;
    slide.addShape(prs.ShapeType.roundRect, { x, y: 4.25, w: 1.95, h: 0.75, fill: { color: '#FFFFFF' }, line: { color: item.color, width: 1.5 }, rectRadius: 0.08 });
    slide.addText(item.label, { x, y: 4.32, w: 1.95, h: 0.28, fontSize: 10, bold: true, color: item.color, fontFace: fontMain, align: 'center' });
    slide.addText(item.sub, { x, y: 4.6, w: 1.95, h: 0.28, fontSize: 9, color: GRAY, fontFace: fontMain, align: 'center' });
    if (i < flowItems.length - 1) {
      slide.addText('→', { x: x + 1.95, y: 4.4, w: 0.2, h: 0.3, fontSize: 14, color: GRAY, fontFace: fontMain, align: 'center' });
    }
  });
});

// ─── SLIDE 5: USER FLOW ИИ-ПОМОЩНИКА ──────────────────────────────
addSlide('03. User Flow — ИИ-помощник', 'Путь пользователя от вопроса до покупки', (slide) => {
  const steps = [
    { icon: '🏠', title: 'Любая страница', desc: 'Пользователь на сайте', color: BLUE },
    { icon: '💬', title: 'Кнопка ИИ', desc: 'Плавающий виджет справа внизу', color: BLUE },
    { icon: '⌨️', title: 'Вопрос', desc: 'Напр.: «наушники до 3000₽»', color: PURPLE },
    { icon: '⚙️', title: 'Backend', desc: 'Запрос к БД + сборка промпта', color: ORANGE },
    { icon: '🤖', title: 'Groq AI', desc: 'Генерация ответа ~0.3 сек', color: PURPLE },
    { icon: '📦', title: 'Ответ', desc: 'Список товаров со ссылками', color: GREEN },
    { icon: '🛒', title: 'Покупка', desc: 'Переход в карточку → корзина', color: GREEN },
  ];

  steps.forEach((s, i) => {
    const y = 1.4 + i * 0.72;
    slide.addShape(prs.ShapeType.roundRect, { x: 0.4, y, w: 0.65, h: 0.55, fill: { color: s.color }, line: { color: s.color }, rectRadius: 0.1 });
    slide.addText(s.icon, { x: 0.4, y, w: 0.65, h: 0.55, fontSize: 18, align: 'center', valign: 'middle', fontFace: fontMain });

    slide.addShape(prs.ShapeType.roundRect, { x: 0.4, y: y + 0.58, w: 0.65, h: 0.1, fill: { color: i < steps.length - 1 ? '#D1D5DB' : 'transparent' }, line: { color: 'transparent' } });

    slide.addText(`${String(i + 1).padStart(2, '0')}. ${s.title}`, { x: 1.2, y: y + 0.05, w: 5, h: 0.25, fontSize: 13, bold: true, color: s.color, fontFace: fontMain });
    slide.addText(s.desc, { x: 1.2, y: y + 0.3, w: 5, h: 0.22, fontSize: 11, color: GRAY, fontFace: fontMain });

    if (i < steps.length - 1) {
      slide.addText('↓', { x: 0.53, y: y + 0.58, w: 0.4, h: 0.18, fontSize: 13, color: '#9CA3AF', fontFace: fontMain, align: 'center' });
    }
  });

  // Right: system prompt detail
  box(slide, 6.8, 1.4, 6.4, 3.6, '#F9FAFB', '#E5E7EB');
  slide.addText('Что входит в System Prompt', { x: 7.0, y: 1.5, w: 6.0, h: 0.38, fontSize: 14, bold: true, color: DARK, fontFace: fontMain });
  const prompts = [
    ['🏪 Платформа', 'Название «Флип», категории, правила общения'],
    ['📦 Товары', 'TOP-10 актуальных товаров из SQLite (title, price, id)'],
    ['🎯 Поведение', 'Отвечай кратко, давай ссылки /product/:id'],
    ['🔒 Ограничения', 'Не выходи за рамки маркетплейса'],
    ['💬 История', 'messages[] — весь диалог передаётся в Groq'],
  ];
  prompts.forEach(([k, v], i) => {
    slide.addText(k, { x: 7.0, y: 1.97 + i * 0.55, w: 1.8, h: 0.28, fontSize: 11, bold: true, color: PURPLE, fontFace: fontMain });
    slide.addText(v, { x: 8.85, y: 1.97 + i * 0.55, w: 4.1, h: 0.28, fontSize: 10.5, color: DARK, fontFace: fontMain });
  });

  box(slide, 6.8, 5.2, 6.4, 0.9, LIGHTPURPLE, PURPLE);
  slide.addText('Ключевой инсайт:', { x: 7.0, y: 5.28, w: 6.0, h: 0.28, fontSize: 12, bold: true, color: PURPLE, fontFace: fontMain });
  slide.addText('ИИ знает о каталоге благодаря системному промпту — он подбирает товары из реальной базы данных, а не выдумывает', {
    x: 7.0, y: 5.58, w: 6.0, h: 0.44, fontSize: 10.5, color: DARK, fontFace: fontMain,
  });
});

// ─── SLIDE 6: ДЕПЛОЙ ──────────────────────────────────────────────
addSlide('04. Деплой на Railway', 'GitHub → автодеплой → production за 2 минуты', (slide) => {
  // Why Railway
  box(slide, 0.4, 1.4, 6.0, 2.3, LIGHTGREEN, GREEN);
  slide.addText('Почему Railway?', { x: 0.6, y: 1.48, w: 5.6, h: 0.38, fontSize: 14, bold: true, color: GREEN, fontFace: fontMain });

  const rwData = [
    [{ text: 'Платформа', options: { bold: true, fill: { color: '#DCFCE7' } } },
     { text: 'SQLite', options: { bold: true, fill: { color: '#DCFCE7' } } },
     { text: 'Автодеплой', options: { bold: true, fill: { color: '#DCFCE7' } } },
     { text: 'Cold start', options: { bold: true, fill: { color: '#DCFCE7' } } }],
    ['Heroku', '✗ (эфемерная FS)', '✓', 'Нет'],
    ['Vercel', '✗ (serverless)', '✓', 'Нет'],
    ['Render free', '✓', '✓', '✗ Есть (30с)'],
    [{ text: 'Railway 🏆', options: { bold: true } },
     { text: '✓ persistent', options: { bold: true, color: GREEN } },
     { text: '✓', options: { bold: true, color: GREEN } },
     { text: '✓ нет', options: { bold: true, color: GREEN } }],
  ];
  slide.addTable(rwData, {
    x: 0.5, y: 1.92, w: 5.7, h: 1.6,
    fontSize: 11, fontFace: fontMain, color: DARK,
    border: { pt: 1, color: '#BBF7D0' },
    rowH: 0.3,
  });

  // CI/CD Pipeline
  box(slide, 6.6, 1.4, 6.6, 2.3, LIGHTBLUE, BLUE);
  slide.addText('CI/CD Pipeline', { x: 6.8, y: 1.48, w: 6.2, h: 0.38, fontSize: 14, bold: true, color: BLUE, fontFace: fontMain });
  const pipeline = [
    { icon: '⬡', label: 'git push → GitHub (main)', color: DARK },
    { icon: '↓', label: '', color: '#9CA3AF' },
    { icon: 'R', label: 'Railway — webhook → Build', color: PURPLE },
    { icon: '↓', label: '', color: '#9CA3AF' },
    { icon: '⚙️', label: 'npm install + npm run build', color: ORANGE },
    { icon: '↓', label: '', color: '#9CA3AF' },
    { icon: '✓', label: 'Deploy + initDB() — seed данные', color: GREEN },
  ];
  pipeline.forEach((p, i) => {
    if (p.label) {
      slide.addText(`${p.icon}  ${p.label}`, { x: 6.8, y: 1.92 + i * 0.28, w: 6.2, h: 0.26, fontSize: 11, bold: p.icon !== '↓', color: p.color, fontFace: fontMain });
    } else {
      slide.addText(p.icon, { x: 6.85, y: 1.92 + i * 0.28, w: 0.3, h: 0.26, fontSize: 12, color: p.color, fontFace: fontMain, align: 'center' });
    }
  });

  // Env vars
  box(slide, 0.4, 3.85, 5.8, 1.5, '#FFFBEB', '#F59E0B');
  slide.addText('ENV переменные (Railway Dashboard)', { x: 0.6, y: 3.93, w: 5.4, h: 0.35, fontSize: 13, bold: true, color: '#92400E', fontFace: fontMain });
  const envs = ['NODE_ENV=production', 'GROQ_API_KEY=gsk_...', 'JWT_SECRET=...', 'PORT=3001  # /data/flip.db'];
  envs.forEach((e, i) => {
    slide.addText(e, { x: 0.6, y: 4.35 + i * 0.24, w: 5.4, h: 0.22, fontSize: 11, color: DARK, fontFace: 'Courier New' });
  });

  // Key advantage
  box(slide, 6.6, 3.85, 6.6, 1.5, LIGHTPURPLE, PURPLE);
  slide.addText('Ключевое преимущество Railway', { x: 6.8, y: 3.93, w: 6.2, h: 0.35, fontSize: 13, bold: true, color: PURPLE, fontFace: fontMain });
  const rails = [
    '✓ Persistent filesystem — SQLite не сбрасывается',
    '✓ Push to GitHub = деплой без доп. конфигурации',
    '✓ Нет cold start (в отличие от Render free)',
    '✓ Идеально для MVP без PostgreSQL',
    '✓ Бесплатный tier на старте ($5 кредитов)',
  ];
  rails.forEach((r, i) => {
    slide.addText(r, { x: 6.8, y: 4.35 + i * 0.24, w: 6.2, h: 0.22, fontSize: 11, color: i < 2 ? GREEN : DARK, fontFace: fontMain, bold: i < 2 });
  });
});

// ─── SLIDE 7: ОЦЕНКА ──────────────────────────────────────────────
addSlide('05. Оценка и дальнейшее развитие', 'Итоги · Проблемы · Вектор', (slide) => {
  // Stats
  const stats = [
    { num: '18', label: 'страниц\nреализовано' },
    { num: '60+', label: 'демо-товаров\nв каталоге' },
    { num: '~0.3с', label: 'ответ\nИИ-помощника' },
    { num: '2', label: 'демо-аккаунта\n(B2B + B2C)' },
  ];
  stats.forEach((s, i) => {
    box(slide, 0.4 + i * 3.2, 1.38, 3.0, 1.0, '#F3F4F6', '#E5E7EB');
    slide.addText(s.num, { x: 0.4 + i * 3.2, y: 1.42, w: 3.0, h: 0.52, fontSize: 30, bold: true, color: DARK, fontFace: fontMain, align: 'center' });
    slide.addText(s.label, { x: 0.4 + i * 3.2, y: 1.95, w: 3.0, h: 0.38, fontSize: 10, color: GRAY, fontFace: fontMain, align: 'center' });
  });

  // What works
  box(slide, 0.4, 2.55, 6.2, 2.6, LIGHTGREEN, GREEN);
  slide.addText('✓ Что работает', { x: 0.6, y: 2.63, w: 5.8, h: 0.38, fontSize: 14, bold: true, color: GREEN, fontFace: fontMain });
  const works = [
    'Регистрация, вход, JWT-авторизация',
    'Каталог с фильтрацией (категория, цена, город)',
    'Корзина, избранное, оформление заказов',
    'Личный кабинет продавца (B2B и B2C режимы)',
    'Добавление / редактирование / удаление товаров',
    'ИИ-помощник с реальным контекстом из БД',
    'Автодеплой на Railway через GitHub',
  ];
  works.forEach((w, i) => {
    slide.addText(`✓  ${w}`, { x: 0.6, y: 3.1 + i * 0.3, w: 5.8, h: 0.28, fontSize: 11, color: DARK, fontFace: fontMain });
  });

  // Problems
  box(slide, 6.8, 2.55, 6.4, 1.1, '#FEF2F2', '#FCA5A5');
  slide.addText('✗ Выявленные ограничения', { x: 7.0, y: 2.63, w: 6.0, h: 0.38, fontSize: 14, bold: true, color: '#DC2626', fontFace: fontMain });
  const probs = ['SQLite не масштабируется горизонтально', 'Нет реальной оплаты (только эмуляция)', 'Нет push-уведомлений продавцу', 'Поиск — простой LIKE, без full-text'];
  probs.forEach((p, i) => {
    slide.addText(`✗  ${p}`, { x: 7.0, y: 3.1 + i * 0.27, w: 6.0, h: 0.25, fontSize: 11, color: '#DC2626', fontFace: fontMain });
  });

  // Roadmap
  box(slide, 6.8, 3.8, 6.4, 1.35, LIGHTBLUE, BLUE);
  slide.addText('→ Вектор развития', { x: 7.0, y: 3.88, w: 6.0, h: 0.35, fontSize: 14, bold: true, color: BLUE, fontFace: fontMain });
  const road = ['PostgreSQL вместо SQLite', 'Stripe / ЮKassa — реальная оплата', 'ИИ-рекомендации по истории просмотров', 'React Native — мобильное приложение'];
  road.forEach((r, i) => {
    slide.addText(`→  ${r}`, { x: 7.0, y: 4.3 + i * 0.27, w: 6.0, h: 0.25, fontSize: 11, color: BLUE, fontFace: fontMain });
  });
});

// ─── SLIDE 8: ФИНАЛЬНЫЙ ────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: DARK };
  slide.addText('Флип', { x: 1, y: 1.5, w: 12, h: 1.2, fontSize: 60, bold: true, color: '#FFFFFF', fontFace: fontMain });
  slide.addText('Маркетплейс для покупки и продажи — простой, умный, доступный', {
    x: 1, y: 2.9, w: 11, h: 0.6, fontSize: 20, color: '#9CA3AF', fontFace: fontMain,
  });

  const bullets = ['B2B + B2C в одной платформе', 'ИИ-помощник на Groq (0.3 сек)', 'Автодеплой Railway из GitHub', 'Готов к масштабированию'];
  bullets.forEach((b, i) => {
    slide.addShape(prs.ShapeType.roundRect, { x: 1 + i * 3.2, y: 3.9, w: 3.0, h: 0.9, fill: { color: [BLUE, PURPLE, GREEN, ORANGE][i] }, line: { color: [BLUE, PURPLE, GREEN, ORANGE][i] }, rectRadius: 0.12 });
    slide.addText(b, { x: 1 + i * 3.2, y: 3.9, w: 3.0, h: 0.9, fontSize: 12, bold: true, color: '#FFFFFF', fontFace: fontMain, align: 'center', valign: 'middle' });
  });

  slide.addText('github.com · Railway · nasmirnov21@gmail.com', {
    x: 1, y: 5.2, w: 11.5, h: 0.4, fontSize: 12, color: '#4B5563', fontFace: fontMain,
  });
}

const outPath = 'C:\\Users\\Никита\\figa\\portfolio\\Флип_Презентация.pptx';
prs.writeFile({ fileName: outPath }).then(() => {
  console.log('DONE: ' + outPath);
}).catch(e => console.error(e));
