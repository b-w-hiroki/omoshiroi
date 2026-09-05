/* ============================================================
   export-pptx.js — index.html を画像化して PPTX に書き出す
   ------------------------------------------------------------
   使い方:  node export-pptx.js [出力ファイル名] [--steps]
   出力:    presentation-export.pptx（16:9、各スライド全面画像＋スピーカーノート）
   仕様:
     - 標準は 1 セクション = 1 スライド。サブステップ（→キーで進む演出）は
       すべて表示した最終状態で撮影する（06 のチップ全表示、12 の先人カード4人展開）
     - --steps を付けると 1 ステップ = 1 スライドに展開する
     - スピーカーノートは index.html 内の <!-- Speaker note: ... --> を転記
     - 画像は 2560×1440 (Retina 相当) の JPEG で描画
   前提:    playwright（Chromium）と pptxgenjs が入っていること
            npm install && npx playwright install chromium
   ============================================================ */
const fs   = require('fs');
const path = require('path');

function loadPlaywright() {
  try { return require('playwright'); }
  catch (_) { return require('/opt/node22/lib/node_modules/playwright'); }
}
const { chromium } = loadPlaywright();
const PptxGenJS = require('pptxgenjs');

const args = process.argv.slice(2);
const EXPAND_STEPS = args.includes('--steps');
const OUT = args.find(a => !a.startsWith('--')) || 'presentation-export.pptx';
const ROOT = __dirname;
const HTML = path.join(ROOT, 'index.html');
const TMP  = path.join(ROOT, '.pptx-frames');

// ── speaker notes from HTML comments (per <section>) ──
function extractNotes() {
  const html = fs.readFileSync(HTML, 'utf8');
  const notes = [];
  const re = /<section class="slide[^>]*>([\s\S]*?)<\/section>/g;
  let m;
  while ((m = re.exec(html))) {
    const n = /<!--\s*Speaker note:\s*([\s\S]*?)-->/.exec(m[1]);
    notes.push(n ? n[1].trim() : '');
  }
  return notes;
}

(async () => {
  const notes = extractNotes();
  fs.rmSync(TMP, { recursive: true, force: true });
  fs.mkdirSync(TMP);

  const launchOpts = {};
  if (fs.existsSync('/opt/pw-browsers/chromium')) launchOpts.executablePath = '/opt/pw-browsers/chromium';
  const browser = await chromium.launch(launchOpts);
  const page = await browser.newPage({ viewport: { width: 1280, height: 720 }, deviceScaleFactor: 2 });
  await page.goto('file://' + HTML);
  let css = '#nav-hint,#slide-counter{display:none!important} .title-deco{animation:none!important} *{animation:none!important;transition:none!important}';
  if (!EXPAND_STEPS) {
    // 1 セクション = 1 枚: サブステップの要素をすべて表示した状態で撮る
    css += ' .legend-card{flex-grow:1!important;opacity:1!important} .legend-steps{display:none!important}'
         + ' .lc-active .legend-body{max-height:none!important;padding:12px 16px 14px!important}';
  }
  await page.addStyleTag({ content: css });
  await page.waitForTimeout(300);

  // number of sub-steps per slide (same logic as script.js)
  const stepCounts = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.slide')).map(s => {
      const n = s.querySelectorAll('[data-step]').length;
      if (!n) return 1;
      return s.dataset.stepMode === 'cumulative' ? n + 1 : n;
    })
  );

  const frames = []; // { file, slideIdx, step, steps }
  for (let i = 0; i < stepCounts.length; i++) {
    for (let k = 0; k < stepCounts[i]; k++) {
      if (!(i === 0 && k === 0)) await page.keyboard.press('ArrowRight');
      if (!EXPAND_STEPS && k < stepCounts[i] - 1) continue; // 最終ステップだけ撮る
      if (!EXPAND_STEPS) {
        // 最終状態: すべての data-step 要素を表示
        await page.evaluate(() => document.querySelectorAll('.slide.active [data-step]')
          .forEach(el => el.classList.add('lc-active')));
      }
      await page.waitForTimeout(120);
      const file = path.join(TMP, `s${String(i + 1).padStart(2, '0')}-${k + 1}.jpg`);
      await page.screenshot({ path: file, type: 'jpeg', quality: 90 });
      frames.push({ file, slideIdx: i, step: k, steps: stepCounts[i] });
    }
  }
  await browser.close();

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9'; // 10in × 5.625in
  pptx.title = '「面白い」とは何か';
  for (const f of frames) {
    const slide = pptx.addSlide();
    slide.background = { color: '07091c' };
    slide.addImage({ path: f.file, x: 0, y: 0, w: 10, h: 5.625 });
    let note = notes[f.slideIdx] || '';
    if (EXPAND_STEPS && f.steps > 1) note = `[ステップ ${f.step + 1}/${f.steps}] ` + note;
    if (note) slide.addNotes(note);
  }
  await pptx.writeFile({ fileName: path.join(ROOT, OUT) });
  fs.rmSync(TMP, { recursive: true, force: true });
  console.log(`wrote ${OUT}: ${frames.length} slides (${stepCounts.length} sections)`);
})().catch(e => { console.error(e); process.exit(1); });
