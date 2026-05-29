const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 16:9

// ─── Color palette ───────────────────────────────────────────
const C = {
  bg1:    '07091c',
  bg2:    '0c1028',
  bgCard: '111829',
  cyan:   '00d4ff',
  orange: 'ff6335',
  yellow: 'ffd600',
  purple: 'c084fc',
  green:  '00e676',
  red:    'ff4444',
  t1:     'f0f4ff',
  t2:     'c0d2ee',
  t3:     '93a8d2',
  gold:   'f5c842',
  white:  'FFFFFF',
};

// ─── Helpers ──────────────────────────────────────────────────
function darkBg(slide) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: C.bg1 },
    line: { color: C.bg1 },
  });
}

function slideHeader(slide, num, label) {
  slide.addText(`${num}  ${label}`, {
    x: 0.4, y: 0.15, w: 11, h: 0.35,
    fontSize: 10, color: C.t3, bold: false,
    fontFace: 'Arial',
  });
  // thin separator line
  slide.addShape(pptx.ShapeType.line, {
    x: 0.4, y: 0.55, w: 12.2, h: 0,
    line: { color: C.t3, width: 0.5, dashType: 'dash' },
  });
}

function card(slide, opts) {
  // opts: x, y, w, h, color (border/accent), title, body (array of strings)
  const bdr = opts.color || C.cyan;
  slide.addShape(pptx.ShapeType.rect, {
    x: opts.x, y: opts.y, w: opts.w, h: opts.h,
    fill: { color: C.bgCard },
    line: { color: bdr, width: 1.5 },
    rectRadius: 0.05,
  });
  // left accent bar
  slide.addShape(pptx.ShapeType.rect, {
    x: opts.x, y: opts.y, w: 0.06, h: opts.h,
    fill: { color: bdr },
    line: { color: bdr },
  });
  if (opts.title) {
    slide.addText(opts.title, {
      x: opts.x + 0.15, y: opts.y + 0.12, w: opts.w - 0.25, h: 0.32,
      fontSize: 13, bold: true, color: bdr, fontFace: 'Arial',
    });
  }
  if (opts.body && opts.body.length) {
    const bodyStr = opts.body.map(t => ({ text: t + '\n', options: {} }));
    slide.addText(opts.body.join('\n'), {
      x: opts.x + 0.15, y: opts.y + (opts.title ? 0.48 : 0.15),
      w: opts.w - 0.25,
      h: opts.h - (opts.title ? 0.6 : 0.25),
      fontSize: 11, color: C.t2, fontFace: 'Arial',
      valign: 'top', wrap: true,
    });
  }
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 01 — TITLE
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);

  // Decorative grid dots (simplified)
  for (let r = 0; r < 5; r++) {
    for (let c = 0; c < 14; c++) {
      slide.addShape(pptx.ShapeType.ellipse, {
        x: 0.3 + c * 0.95, y: 0.3 + r * 1.35, w: 0.04, h: 0.04,
        fill: { color: '1a2540' }, line: { color: '1a2540' },
      });
    }
  }

  // Radial glow circle
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 2.5, y: 0.5, w: 8, h: 6,
    fill: { type: 'solid', color: '0d1a40' },
    line: { color: '0d1a40' },
  });

  slide.addText('What Makes Things Fun?', {
    x: 1.5, y: 1.8, w: 10, h: 0.5,
    fontSize: 14, color: C.cyan, bold: true, fontFace: 'Arial',
    align: 'center',
  });

  slide.addText('「面白い」とは何か', {
    x: 1.5, y: 2.4, w: 10, h: 1.2,
    fontSize: 52, bold: true, color: C.t1, fontFace: 'Arial',
    align: 'center',
  });

  // divider
  slide.addShape(pptx.ShapeType.rect, {
    x: 4.5, y: 3.8, w: 4, h: 0.04,
    fill: { color: C.cyan }, line: { color: C.cyan },
  });

  slide.addText('ゲームデザインを通じて「面白さ」の構造を解き明かす', {
    x: 1.5, y: 4.0, w: 10, h: 0.55,
    fontSize: 14, color: C.t2, fontFace: 'Arial', align: 'center',
  });

  slide.addText('▶  PRESS SPACE TO BEGIN', {
    x: 3.5, y: 5.5, w: 6, h: 0.4,
    fontSize: 10, color: C.t3, fontFace: 'Arial', align: 'center',
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 02 — 問い
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '02', '問い');

  // Category pills
  const categories = ['ゲーム', '映画', '漫画', '小説', '会話', '仕事', 'スポーツ', '人生'];
  const colors = [C.cyan, C.orange, C.yellow, C.purple, C.green, C.t3, C.orange, C.cyan];
  categories.forEach((cat, i) => {
    const x = (i % 4) * 1.8 + 0.6;
    const y = Math.floor(i / 4) * 0.55 + 0.7;
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y, w: 1.5, h: 0.38,
      fill: { color: C.bgCard },
      line: { color: colors[i], width: 1.2 },
      rectRadius: 0.1,
    });
    slide.addText(cat, {
      x, y, w: 1.5, h: 0.38,
      fontSize: 13, color: colors[i], bold: true,
      fontFace: 'Arial', align: 'center', valign: 'middle',
    });
  });

  slide.addText('そもそも、', {
    x: 1, y: 2.0, w: 11, h: 0.7,
    fontSize: 28, color: C.t2, fontFace: 'Arial', align: 'center',
  });
  slide.addText('「面白い」とは 何でしょうか？', {
    x: 0.5, y: 2.65, w: 12, h: 1.2,
    fontSize: 44, bold: true, color: C.t1, fontFace: 'Arial', align: 'center',
  });
  slide.addText('日常的に使う言葉ほど、説明するのは難しい。', {
    x: 1.5, y: 4.1, w: 10, h: 0.45,
    fontSize: 14, color: C.t3, fontFace: 'Arial', align: 'center',
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 03 — なぜこの問いか
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '03', 'なぜこの問いか');

  const items = [
    { icon: '🎮', text: '子どもの頃からゲームが好きで、「なぜこれは面白いのか」をずっと考えていた' },
    { icon: '🔧', text: '作る側になってわかった—— 「面白さは設計できる」という事実' },
    { icon: '🌐', text: 'その構造はゲームだけでなく、プロダクト・マーケ・教育・組織にも使える' },
  ];

  items.forEach((item, i) => {
    const y = 0.8 + i * 1.4;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 0.5, y: y + 0.05, w: 0.6, h: 0.6,
      fill: { color: C.bgCard }, line: { color: C.cyan, width: 1.5 },
    });
    slide.addText(item.icon, {
      x: 0.5, y: y + 0.05, w: 0.6, h: 0.6,
      fontSize: 16, align: 'center', valign: 'middle',
    });
    slide.addText(item.text, {
      x: 1.3, y: y, w: 8, h: 0.7,
      fontSize: 16, color: C.t2, fontFace: 'Arial', valign: 'middle', wrap: true,
    });
  });

  // Right side callout
  slide.addShape(pptx.ShapeType.rect, {
    x: 9.8, y: 0.7, w: 3.2, h: 4.8,
    fill: { color: '0a1535' }, line: { color: C.cyan, width: 1.5 },
  });
  slide.addText('THE CORE QUESTION', {
    x: 9.9, y: 0.9, w: 3.0, h: 0.35,
    fontSize: 9, color: C.cyan, bold: true, fontFace: 'Arial', align: 'center',
  });
  slide.addText('「面白い」は\n偶然ではなく\n設計できる', {
    x: 9.9, y: 1.4, w: 3.0, h: 3.8,
    fontSize: 22, bold: true, color: C.t1, fontFace: 'Arial',
    align: 'center', valign: 'middle',
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 04 — 岡本吉起さんの視点
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '04', '岡本吉起さんの視点');

  // Profile card (left)
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.4, y: 0.65, w: 3.5, h: 5.0,
    fill: { color: '0e1830' }, line: { color: C.gold, width: 1.5 },
  });
  slide.addText('LEGEND', {
    x: 0.5, y: 0.8, w: 3.3, h: 0.35,
    fontSize: 10, color: C.gold, bold: true, fontFace: 'Arial', align: 'center',
  });
  slide.addText('岡本吉起', {
    x: 0.5, y: 1.5, w: 3.3, h: 0.7,
    fontSize: 26, bold: true, color: C.t1, fontFace: 'Arial', align: 'center',
  });
  slide.addText('元カプコン 執行役員\nゲームクリエイター', {
    x: 0.5, y: 2.25, w: 3.3, h: 0.65,
    fontSize: 11, color: C.t3, fontFace: 'Arial', align: 'center',
  });
  // Works
  ['ストリートファイターII', 'バイオハザード', 'モンスターストライク'].forEach((w, i) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.6, y: 3.1 + i * 0.55, w: 3.1, h: 0.42,
      fill: { color: C.bgCard }, line: { color: C.gold, width: 0.8 },
      rectRadius: 0.05,
    });
    slide.addText(w, {
      x: 0.6, y: 3.1 + i * 0.55, w: 3.1, h: 0.42,
      fontSize: 11, color: C.gold, fontFace: 'Arial', align: 'center', valign: 'middle',
    });
  });

  // Quote
  slide.addText('「面白いゲームには、4段階の面白さがある」', {
    x: 4.2, y: 0.75, w: 8.8, h: 0.65,
    fontSize: 16, color: C.t1, fontFace: 'Arial', italic: true,
  });

  // 4 stages
  const stages = [
    { num: '1', label: '接触前', sub: '聞いて面白い', items: ['企画・コンセプト', 'タイトル・コピー', '口コミ・評判'], note: '触れる前に「気になる」を作る。' },
    { num: '2', label: '初見',   sub: '見て面白い',   items: ['キャラクター', 'グラフィック', '世界観・雰囲気'],  note: '第一印象で期待値が決まる。' },
    { num: '3', label: 'プレイ中', sub: '遊んで面白い', items: ['操作性・手触り', 'ゲームシステム', 'チュートリアル'], note: 'ここが最も重要。' },
    { num: '4', label: '継続', sub: '続けて面白い', items: ['成長実感', '変化・驚き', 'コミュニティ'], note: '長期エンゲージメント。' },
  ];

  const stageColors = [C.cyan, C.yellow, C.orange, C.purple];
  stages.forEach((s, i) => {
    const x = 4.2 + i * 2.3;
    const y = 1.55;
    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: 2.15, h: 3.8,
      fill: { color: C.bgCard }, line: { color: stageColors[i], width: 1.2 },
    });
    slide.addText(s.num, {
      x, y: y + 0.05, w: 2.15, h: 0.4,
      fontSize: 18, bold: true, color: stageColors[i], fontFace: 'Arial', align: 'center',
    });
    slide.addText(s.label, {
      x, y: y + 0.42, w: 2.15, h: 0.38,
      fontSize: 14, bold: true, color: C.t1, fontFace: 'Arial', align: 'center',
    });
    slide.addText(s.sub, {
      x, y: y + 0.82, w: 2.15, h: 0.32,
      fontSize: 10, color: stageColors[i], fontFace: 'Arial', align: 'center',
    });
    s.items.forEach((item, j) => {
      slide.addText(`• ${item}`, {
        x: x + 0.1, y: y + 1.22 + j * 0.42, w: 1.95, h: 0.38,
        fontSize: 10, color: C.t2, fontFace: 'Arial',
      });
    });
    slide.addText(s.note, {
      x: x + 0.05, y: y + 3.3, w: 2.05, h: 0.42,
      fontSize: 9, color: stageColors[i], fontFace: 'Arial', italic: true, wrap: true,
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 05 — 先人たちの視点
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '05', '先人たちの視点');

  const legends = [
    { p: 'P1', name: '桜井政博', works: '星のカービィ / スマブラシリーズ', keyword: 'リスクと\nリターン', desc: '面白さはリスクを取ってリターンを得る構造から生まれる', color: C.cyan },
    { p: 'P2', name: 'Sid Meier', works: 'Civilization / Pirates!', keyword: '面白い\n意思決定', desc: 'ゲームとは「一連の面白い意思決定」であるとして知られる考え方', color: C.yellow },
    { p: 'P3', name: 'Raph Koster', works: 'Ultima Online / A Theory of Fun', keyword: '楽しさは\n学習', desc: '脳が新しいパターンを学習するときに楽しさが生まれる、という整理', color: C.orange },
    { p: 'P4', name: '宮本茂', works: 'スーパーマリオ / ゼルダ / Wii Sports', keyword: '触って\nわかる体験', desc: '説明しなくても触った瞬間に「面白い」と伝わる体験を重視する姿勢', color: C.purple },
  ];

  legends.forEach((l, i) => {
    const x = 0.3 + i * 3.3;
    const y = 0.7;
    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: 3.1, h: 5.2,
      fill: { color: C.bgCard }, line: { color: l.color, width: 1.2 },
    });
    slide.addText(l.p, {
      x, y: y + 0.08, w: 3.1, h: 0.35,
      fontSize: 10, color: l.color, bold: true, fontFace: 'Arial', align: 'center',
    });
    slide.addText(l.name, {
      x, y: y + 0.45, w: 3.1, h: 0.5,
      fontSize: 16, bold: true, color: C.t1, fontFace: 'Arial', align: 'center',
    });
    slide.addText(l.works, {
      x: x + 0.1, y: y + 0.95, w: 2.9, h: 0.5,
      fontSize: 9, color: C.t3, fontFace: 'Arial', align: 'center', wrap: true,
    });
    // keyword box
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.2, y: y + 1.55, w: 2.7, h: 1.0,
      fill: { color: '07091c' }, line: { color: l.color, width: 1.5 },
      rectRadius: 0.08,
    });
    slide.addText(l.keyword, {
      x: x + 0.2, y: y + 1.55, w: 2.7, h: 1.0,
      fontSize: 18, bold: true, color: l.color, fontFace: 'Arial',
      align: 'center', valign: 'middle',
    });
    slide.addText(l.desc, {
      x: x + 0.1, y: y + 2.7, w: 2.9, h: 2.2,
      fontSize: 10, color: C.t2, fontFace: 'Arial', wrap: true, valign: 'top',
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 06 — 共通するものは何か
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '06', '共通するものは何か');

  slide.addText('COMMON CORE', {
    x: 0.5, y: 0.75, w: 12, h: 0.4,
    fontSize: 11, color: C.cyan, bold: true, fontFace: 'Arial', align: 'center',
  });

  slide.addText('共通しているのは\n「人の中で', {
    x: 0.5, y: 1.2, w: 12, h: 1.4,
    fontSize: 28, color: C.t2, fontFace: 'Arial', align: 'center',
  });

  slide.addText('感情が動いている', {
    x: 0.5, y: 2.5, w: 12, h: 0.9,
    fontSize: 44, bold: true, color: C.cyan, fontFace: 'Arial', align: 'center',
  });

  slide.addText('」こと', {
    x: 0.5, y: 3.35, w: 12, h: 0.6,
    fontSize: 28, color: C.t2, fontFace: 'Arial', align: 'center',
  });

  // Conclusion box
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.5, y: 4.1, w: 10, h: 1.6,
    fill: { color: '0a1535' }, line: { color: C.cyan, width: 1.5 },
  });
  slide.addText('「面白い」とは、', {
    x: 1.6, y: 4.2, w: 9.8, h: 0.45,
    fontSize: 16, color: C.t2, fontFace: 'Arial', align: 'center',
  });
  slide.addText('感情の運動 である。', {
    x: 1.6, y: 4.65, w: 9.8, h: 0.7,
    fontSize: 26, bold: true, color: C.cyan, fontFace: 'Arial', align: 'center',
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 07 — 設計の原則
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '07', '設計の原則');

  slide.addText('ゲームはプレイヤーに\n「楽しくなれ」と命令できない', {
    x: 0.5, y: 0.7, w: 12, h: 1.1,
    fontSize: 28, bold: true, color: C.t1, fontFace: 'Arial', align: 'center',
  });

  slide.addText('だからこそ、できるのは', {
    x: 0.5, y: 1.85, w: 12, h: 0.45,
    fontSize: 16, color: C.t3, fontFace: 'Arial', align: 'center',
  });
  slide.addText('「感情が動く状態」を設計する  ことだけ', {
    x: 0.5, y: 2.3, w: 12, h: 0.65,
    fontSize: 22, bold: true, color: C.cyan, fontFace: 'Arial', align: 'center',
  });

  // Two columns
  // Left: ゲーム側
  const leftItems = ['状態をつくる', '期待させる', '気づかせる', '挑ませる', '変化を感じさせる'];
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.4, y: 3.15, w: 5.2, h: 2.6,
    fill: { color: C.bgCard }, line: { color: C.cyan, width: 1.2 },
  });
  slide.addText('ゲーム側 — 設計できる', {
    x: 0.5, y: 3.25, w: 5.0, h: 0.38,
    fontSize: 12, color: C.cyan, bold: true, fontFace: 'Arial',
  });
  leftItems.forEach((t, i) => {
    slide.addText(`✓  ${t}`, {
      x: 0.6, y: 3.68 + i * 0.38, w: 4.8, h: 0.35,
      fontSize: 12, color: C.t2, fontFace: 'Arial',
    });
  });

  // Right: プレイヤー側
  const rightItems = ['期待が生まれる', '気づきが生まれる', '挑戦心が生まれる', '変化を知覚する'];
  slide.addShape(pptx.ShapeType.rect, {
    x: 7.4, y: 3.15, w: 5.2, h: 2.6,
    fill: { color: C.bgCard }, line: { color: C.orange, width: 1.2 },
  });
  slide.addText('プレイヤー側 — 設計できない', {
    x: 7.5, y: 3.25, w: 5.0, h: 0.38,
    fontSize: 12, color: C.orange, bold: true, fontFace: 'Arial',
  });
  rightItems.forEach((t, i) => {
    slide.addText(`→  ${t}`, {
      x: 7.6, y: 3.68 + i * 0.38, w: 4.8, h: 0.35,
      fontSize: 12, color: C.t2, fontFace: 'Arial',
    });
  });

  // Arrow between
  slide.addShape(pptx.ShapeType.rect, {
    x: 5.7, y: 4.3, w: 1.6, h: 0.4,
    fill: { color: C.bgCard }, line: { color: C.t3, width: 0.8 },
  });
  slide.addText('境界 ➔', {
    x: 5.7, y: 4.3, w: 1.6, h: 0.4,
    fontSize: 11, color: C.t3, fontFace: 'Arial', align: 'center', valign: 'middle',
  });

  slide.addText('その先で初めて 感情が動く', {
    x: 0.5, y: 5.85, w: 12, h: 0.4,
    fontSize: 13, color: C.yellow, fontFace: 'Arial', align: 'center', italic: true,
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 08 — 4要素 循環モデル
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '08', '感情の運動を起こす4要素【循環モデル】');

  // Center label
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 5.0, y: 2.0, w: 3.0, h: 3.0,
    fill: { color: '0c1028' }, line: { color: C.cyan, width: 2 },
  });
  slide.addText('感情の\n運動', {
    x: 5.0, y: 2.0, w: 3.0, h: 3.0,
    fontSize: 22, bold: true, color: C.cyan, fontFace: 'Arial',
    align: 'center', valign: 'middle',
  });

  // 4 elements
  const elements = [
    { label: '推測', sub: '次はどうなる？\n何をすべき？', color: C.cyan,   x: 1.2, y: 0.7  },
    { label: '観察', sub: '結果を確認する\n情報を得る',   color: C.yellow, x: 9.8, y: 0.7  },
    { label: '成長', sub: '上手くなる\n理解が深まる',   color: C.orange, x: 9.8, y: 3.8  },
    { label: '変化', sub: '世界が変わる\n自分が変わる',  color: C.purple, x: 1.2, y: 3.8  },
  ];

  elements.forEach(el => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: el.x, y: el.y, w: 2.5, h: 2.0,
      fill: { color: C.bgCard }, line: { color: el.color, width: 1.8 },
      rectRadius: 0.1,
    });
    slide.addText(el.label, {
      x: el.x, y: el.y + 0.15, w: 2.5, h: 0.65,
      fontSize: 26, bold: true, color: el.color, fontFace: 'Arial', align: 'center',
    });
    slide.addText(el.sub, {
      x: el.x + 0.1, y: el.y + 0.85, w: 2.3, h: 1.0,
      fontSize: 11, color: C.t2, fontFace: 'Arial', align: 'center', wrap: true,
    });
  });

  // Circular arrows (simplified with text labels)
  const arrows = [
    { text: '期待する ➔', x: 3.5, y: 1.0 },
    { text: '➔ 確認する', x: 7.2, y: 1.0 },
    { text: '気づく ➔', x: 9.3, y: 3.0 },
    { text: '◀ 実践する', x: 3.5, y: 5.1 },
    { text: '◀ 振り返る', x: 1.0, y: 3.0 },
  ];
  arrows.forEach(a => {
    slide.addText(a.text, {
      x: a.x, y: a.y, w: 2.0, h: 0.38,
      fontSize: 10, color: C.t3, fontFace: 'Arial', align: 'center',
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 09 — 対戦ゲームにおける感情の運動
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '09', '例：対戦ゲームにおける感情の運動');

  const steps = [
    { label: '勝ちたい',     type: '推測', color: C.cyan,   x: 0.3  },
    { label: '相手を\n観察する', type: '観察', color: C.yellow, x: 2.0  },
    { label: '負けて\n悔しい',   type: '成長', color: C.orange, x: 3.7  },
    { label: '原因に\n気づく',   type: '観察', color: C.yellow, x: 5.4  },
    { label: '練習する',    type: '成長', color: C.orange, x: 7.1  },
    { label: 'できるように\nなる', type: '成長', color: C.orange, x: 8.8  },
    { label: '読みが\n当たる',   type: '推測', color: C.cyan,   x: 10.5 },
    { label: '勝つ',       type: '変化', color: C.purple, x: 12.2 },
  ];

  // Emotion wave (simplified as colored line)
  const waveY = 2.8;
  steps.forEach((s, i) => {
    const x = s.x;
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: 0.7, w: 1.6, h: 1.8,
      fill: { color: C.bgCard }, line: { color: s.color, width: 1.2 },
      rectRadius: 0.06,
    });
    slide.addText(s.type, {
      x, y: 0.75, w: 1.6, h: 0.3,
      fontSize: 8, color: s.color, bold: true, fontFace: 'Arial', align: 'center',
    });
    slide.addText(s.label, {
      x, y: 1.05, w: 1.6, h: 1.3,
      fontSize: 12, color: C.t1, fontFace: 'Arial', align: 'center', valign: 'middle', wrap: true,
    });
    if (i < steps.length - 1) {
      slide.addText('→', {
        x: x + 1.62, y: 1.4, w: 0.36, h: 0.4,
        fontSize: 14, color: C.t3, fontFace: 'Arial', align: 'center',
      });
    }
  });

  // Emotion peak callout
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.3, y: 3.5, w: 12.5, h: 2.2,
    fill: { color: '0a1535' }, line: { color: C.purple, width: 1.5 },
  });
  slide.addText('感情の頂点', {
    x: 0.5, y: 3.65, w: 4, h: 0.4,
    fontSize: 12, color: C.purple, bold: true, fontFace: 'Arial',
  });
  slide.addText('勝利が嬉しいのは、勝ったからだけではない。', {
    x: 0.5, y: 4.1, w: 12, h: 0.5,
    fontSize: 16, bold: true, color: C.t1, fontFace: 'Arial',
  });
  slide.addText('そこに至るまでの  期待・悔しさ・気づき・成長  があるから嬉しい。', {
    x: 0.5, y: 4.65, w: 12, h: 0.55,
    fontSize: 14, color: C.t2, fontFace: 'Arial',
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 10 — 期待値コントロール
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '10', '期待値コントロール');

  // Alert box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.4, y: 0.65, w: 12.2, h: 0.85,
    fill: { color: '200a0a' }, line: { color: C.red, width: 1.5 },
  });
  slide.addText('⚠  ユーザーは勝手に期待して、勝手に失望する', {
    x: 0.5, y: 0.68, w: 12, h: 0.45,
    fontSize: 16, bold: true, color: C.red, fontFace: 'Arial',
  });
  slide.addText('期待値はプレイが始まる前から動いている。これがクソゲー化を生む最大の要因。', {
    x: 0.5, y: 1.1, w: 12, h: 0.32,
    fontSize: 11, color: C.t3, fontFace: 'Arial',
  });

  // Scale diagram
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: 1.65, w: 12, h: 3.0,
    fill: { color: C.bgCard }, line: { color: C.t3, width: 0.8 },
  });

  // Left side: 期待が重い
  slide.addText('期待', {
    x: 1.0, y: 1.9, w: 2.5, h: 0.5,
    fontSize: 20, bold: true, color: C.orange, fontFace: 'Arial',
  });
  slide.addText('期待が重い', {
    x: 1.0, y: 2.42, w: 2.5, h: 0.32,
    fontSize: 10, color: C.t3, fontFace: 'Arial',
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 0.8, y: 2.8, w: 2.0, h: 2.0,
    fill: { color: '200a0a' }, line: { color: C.orange, width: 1.5 },
  });
  slide.addText('✕\nクソゲー', {
    x: 0.8, y: 2.8, w: 2.0, h: 2.0,
    fontSize: 15, bold: true, color: C.orange, fontFace: 'Arial',
    align: 'center', valign: 'middle',
  });

  // Right side: 回収
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 8.5, y: 1.8, w: 2.5, h: 2.5,
    fill: { color: '0a1a0a' }, line: { color: C.green, width: 1.5 },
  });
  slide.addText('回収', {
    x: 8.5, y: 1.8, w: 2.5, h: 1.2,
    fontSize: 20, bold: true, color: C.green, fontFace: 'Arial', align: 'center', valign: 'middle',
  });
  slide.addText('✓\n面白い', {
    x: 8.5, y: 2.9, w: 2.5, h: 1.4,
    fontSize: 15, bold: true, color: C.green, fontFace: 'Arial', align: 'center', valign: 'middle',
  });

  // Bottom line
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.4, y: 4.85, w: 12.2, h: 0.75,
    fill: { color: '0a1535' }, line: { color: C.cyan, width: 1.2 },
  });
  slide.addText('見ているのは  報酬の大きさ  ではなく、  期待とのギャップ', {
    x: 0.5, y: 4.9, w: 12, h: 0.6,
    fontSize: 16, bold: true, color: C.t1, fontFace: 'Arial', align: 'center',
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 11 — 設計領域として整理する
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '11', '設計領域として整理する');

  slide.addText('面白さは 中心にひとつ。多くの 設計領域 が 4要素 を通って、そこへ集まる', {
    x: 0.4, y: 0.65, w: 12, h: 0.75,
    fontSize: 15, color: C.t2, fontFace: 'Arial',
  });

  // Center: 面白さ
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 5.0, y: 1.8, w: 3.0, h: 3.0,
    fill: { color: '0a1535' }, line: { color: C.cyan, width: 2.5 },
  });
  slide.addText('面白さ\n(感情の運動)', {
    x: 5.0, y: 1.8, w: 3.0, h: 3.0,
    fontSize: 16, bold: true, color: C.cyan, fontFace: 'Arial',
    align: 'center', valign: 'middle',
  });

  // Design domains around center
  const domains = [
    { label: 'ゲームメカニクス', color: C.cyan,   x: 0.3, y: 1.5 },
    { label: 'ストーリー',     color: C.yellow, x: 0.3, y: 3.2 },
    { label: 'UI/UX',        color: C.orange, x: 0.3, y: 4.9 },
    { label: 'グラフィック',   color: C.purple, x: 9.7, y: 1.5 },
    { label: 'サウンド',      color: C.green,  x: 9.7, y: 3.2 },
    { label: '難易度設計',    color: C.yellow, x: 9.7, y: 4.9 },
  ];

  domains.forEach(d => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: d.x, y: d.y, w: 2.8, h: 0.7,
      fill: { color: C.bgCard }, line: { color: d.color, width: 1.2 },
      rectRadius: 0.08,
    });
    slide.addText(d.label, {
      x: d.x, y: d.y, w: 2.8, h: 0.7,
      fontSize: 13, bold: true, color: d.color, fontFace: 'Arial',
      align: 'center', valign: 'middle',
    });
  });

  // 4 elements band
  const els = ['推測', '観察', '成長', '変化'];
  const elColors = [C.cyan, C.yellow, C.orange, C.purple];
  els.forEach((e, i) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 1.5 + i * 2.4, y: 5.55, w: 2.1, h: 0.55,
      fill: { color: C.bgCard }, line: { color: elColors[i], width: 1.2 },
      rectRadius: 0.06,
    });
    slide.addText(e, {
      x: 1.5 + i * 2.4, y: 5.55, w: 2.1, h: 0.55,
      fontSize: 14, bold: true, color: elColors[i], fontFace: 'Arial',
      align: 'center', valign: 'middle',
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 12 — この構造はゲームだけではない
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);
  slideHeader(slide, '12', 'この構造はゲームだけではない');

  const domains = [
    { icon: '📱', label: 'プロダクト',       sub: 'オンボーディング設計', desc: '使うほど手に馴染んでいく', color: C.cyan },
    { icon: '📣', label: 'マーケティング',    sub: '期待値とストーリー設計', desc: '続きが見たくなる', color: C.yellow },
    { icon: '📚', label: '教育',             sub: '学習体験の設計',   desc: '「わかった！」の連続', color: C.orange },
    { icon: '🏢', label: '組織マネジメント',  sub: '成長実感の設計',   desc: '自分の変化が見える', color: C.purple },
    { icon: '🤝', label: '接客・サービス',    sub: '体験の起伏設計',   desc: '期待を超えるもてなし', color: C.green },
    { icon: '🎉', label: 'イベント',          sub: '一体感の設計',    desc: 'その場でしか味わえない', color: C.gold },
  ];

  domains.forEach((d, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.3 + col * 4.3;
    const y = 0.7 + row * 2.2;
    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: 4.0, h: 2.0,
      fill: { color: C.bgCard }, line: { color: d.color, width: 1.2 },
    });
    slide.addText(d.icon, {
      x: x + 0.1, y: y + 0.1, w: 0.7, h: 0.7,
      fontSize: 22, fontFace: 'Arial',
    });
    slide.addText(d.label, {
      x: x + 0.85, y: y + 0.1, w: 3.0, h: 0.45,
      fontSize: 15, bold: true, color: d.color, fontFace: 'Arial',
    });
    slide.addText(d.sub, {
      x: x + 0.85, y: y + 0.55, w: 3.0, h: 0.32,
      fontSize: 10, color: C.t3, fontFace: 'Arial',
    });
    slide.addText(d.desc, {
      x: x + 0.1, y: y + 1.3, w: 3.8, h: 0.5,
      fontSize: 13, color: C.t2, fontFace: 'Arial', italic: true,
    });
  });

  slide.addText('「面白い」を設計する思想は、人が関わるあらゆる場面で使える', {
    x: 0.3, y: 5.1, w: 12.4, h: 0.55,
    fontSize: 14, bold: true, color: C.cyan, fontFace: 'Arial', align: 'center',
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 13 — MY CONCLUSION
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);

  // Dark overlay with glow
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 1, y: -1, w: 11, h: 9,
    fill: { color: '080f25' }, line: { color: '080f25' },
  });

  slide.addText('MY CONCLUSION', {
    x: 0.5, y: 0.6, w: 12, h: 0.5,
    fontSize: 12, color: C.cyan, bold: true, fontFace: 'Arial', align: 'center',
  });

  slide.addText('ゲームの面白さとは、', {
    x: 0.5, y: 1.3, w: 12, h: 0.7,
    fontSize: 26, color: C.t2, fontFace: 'Arial', align: 'center',
  });

  slide.addText('推測・観察・成長・変化', {
    x: 0.5, y: 2.1, w: 12, h: 0.9,
    fontSize: 38, bold: true, color: C.cyan, fontFace: 'Arial', align: 'center',
  });

  slide.addText('によって起こる、', {
    x: 0.5, y: 3.05, w: 12, h: 0.6,
    fontSize: 26, color: C.t2, fontFace: 'Arial', align: 'center',
  });

  slide.addText('感情の運動である。', {
    x: 0.5, y: 3.7, w: 12, h: 0.9,
    fontSize: 38, bold: true, color: C.t1, fontFace: 'Arial', align: 'center',
  });

  slide.addText('— 先人の知見と、作る側として積み上げてきた実践から体系化した結論。', {
    x: 1, y: 4.85, w: 11, h: 0.42,
    fontSize: 12, color: C.t3, fontFace: 'Arial', align: 'center', italic: true,
  });
  slide.addText('この構造はゲームだけでなく、プロダクト・教育・組織にも当てはまる。', {
    x: 1, y: 5.3, w: 11, h: 0.42,
    fontSize: 12, color: C.t3, fontFace: 'Arial', align: 'center', italic: true,
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 14 — FINAL
// ═══════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  darkBg(slide);

  slide.addShape(pptx.ShapeType.ellipse, {
    x: 0, y: -1, w: 13, h: 9,
    fill: { color: '060c20' }, line: { color: '060c20' },
  });

  slide.addText('GAME DESIGN PHILOSOPHY', {
    x: 0.5, y: 1.0, w: 12, h: 0.45,
    fontSize: 11, color: C.t3, bold: true, fontFace: 'Arial', align: 'center',
  });

  slide.addText('「面白い」を作るには、', {
    x: 0.5, y: 1.7, w: 12, h: 0.75,
    fontSize: 30, color: C.t1, fontFace: 'Arial', align: 'center',
  });

  slide.addText('感情の運動を設計 すればいい。', {
    x: 0.5, y: 2.5, w: 12, h: 1.0,
    fontSize: 42, bold: true, color: C.cyan, fontFace: 'Arial', align: 'center',
  });

  slide.addText('その構造はゲームに限らない。', {
    x: 0.5, y: 3.75, w: 12, h: 0.5,
    fontSize: 16, color: C.t2, fontFace: 'Arial', align: 'center',
  });

  slide.addText('プロダクト・教育・マーケティング・組織', {
    x: 0.5, y: 4.3, w: 12, h: 0.5,
    fontSize: 16, color: C.t3, fontFace: 'Arial', align: 'center',
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 4.5, y: 4.95, w: 4, h: 0.04,
    fill: { color: C.t3 }, line: { color: C.t3 },
  });

  slide.addText('「面白い」を意図的に生み出したいすべての場所で、この考え方は使える。', {
    x: 0.5, y: 5.15, w: 12, h: 0.55,
    fontSize: 14, color: C.t2, fontFace: 'Arial', align: 'center',
  });
}

// ─── Save ─────────────────────────────────────────────────────
pptx.writeFile({ fileName: '/Users/hiroki_torii/claudecode/game-design-presentation/game-design-presentation.pptx' })
  .then(() => console.log('✅ Saved: game-design-presentation.pptx'))
  .catch(e => { console.error(e); process.exit(1); });
