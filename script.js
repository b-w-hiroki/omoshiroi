/* ============================================================
   GAME DESIGN PRESENTATION — script.js  (v2: sub-navigation)
   ============================================================ */

(function () {
  'use strict';

  const slides    = Array.from(document.querySelectorAll('.slide'));
  const totalEl   = document.getElementById('total-slides');
  const currentEl = document.getElementById('current-slide');

  let current = 0;
  const stepState = {}; // { slideIndex: currentStep }

  /* ── Helpers ────────────────────────────────────────────── */

  /** Cards/items with data-step inside a given slide */
  function getStepItems(idx) {
    return Array.from(slides[idx].querySelectorAll('[data-step]'));
  }

  /**
   * Activate a specific step within a slide.
   * Adds .lc-active to the matching item, removes from others.
   * Updates progress dots.
   */
  function applyStep(idx, step) {
    const items = getStepItems(idx);
    if (!items.length) return;

    items.forEach((el, i) => el.classList.toggle('lc-active', i === step));

    const dots = slides[idx].querySelectorAll('.legend-step-dot');
    dots.forEach((dot, i) => dot.classList.toggle('active', i === step));

    stepState[idx] = step;
  }

  /* ── Core navigation ────────────────────────────────────── */

  /**
   * Jump directly to a slide.
   * @param {number} idx   — target slide index
   * @param {number} dir   — direction: +1 (forward) or -1 (backward)
   */
  function goTo(idx, dir) {
    if (idx < 0 || idx >= slides.length) return;

    slides[current].classList.remove('active');
    current = idx;
    slides[current].classList.add('active');
    currentEl.textContent = current + 1;

    // Initialise sub-steps for slides that have them
    const items = getStepItems(current);
    if (items.length) {
      // Coming from left → start at step 0; from right → start at last step
      applyStep(current, dir >= 0 ? 0 : items.length - 1);
    }
  }

  /**
   * Try to advance: move to next sub-step first, then next slide.
   */
  function tryNext() {
    const items = getStepItems(current);
    if (items.length) {
      const step = stepState[current] ?? 0;
      if (step < items.length - 1) {
        applyStep(current, step + 1);
        return;
      }
    }
    goTo(current + 1, 1);
  }

  /**
   * Try to go back: retreat sub-step first, then previous slide.
   */
  function tryPrev() {
    const items = getStepItems(current);
    if (items.length) {
      const step = stepState[current] ?? 0;
      if (step > 0) {
        applyStep(current, step - 1);
        return;
      }
    }
    goTo(current - 1, -1);
  }

  /* ── Keyboard ───────────────────────────────────────────── */
  document.addEventListener('keydown', function (e) {
    switch (e.key) {
      case 'ArrowRight':
      case 'ArrowDown':
      case ' ':
        e.preventDefault();
        tryNext();
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
        e.preventDefault();
        tryPrev();
        break;
      case 'Home':
        e.preventDefault();
        goTo(0, 1);
        break;
      case 'End':
        e.preventDefault();
        goTo(slides.length - 1, -1);
        break;
    }
  });

  /* ── Touch / swipe ──────────────────────────────────────── */
  let touchStartX = 0;
  document.addEventListener('touchstart', function (e) {
    touchStartX = e.touches[0].clientX;
  }, { passive: true });

  document.addEventListener('touchend', function (e) {
    const dx = e.changedTouches[0].clientX - touchStartX;
    if (Math.abs(dx) > 50) {
      dx < 0 ? tryNext() : tryPrev();
    }
  }, { passive: true });

  /* ── Scale to fit viewport (16:9 固定) ─────────────────── */
  const BASE_W = 1280;
  const BASE_H = 720;
  const presentation = document.getElementById('presentation');

  function scalePresentation() {
    const scale = Math.min(
      window.innerWidth  / BASE_W,
      window.innerHeight / BASE_H
    );
    presentation.style.transform =
      `translate(-50%, -50%) scale(${scale})`;
  }

  window.addEventListener('resize', scalePresentation);
  scalePresentation();

  /* ── Init ───────────────────────────────────────────────── */
  totalEl.textContent = slides.length;
  goTo(0, 1);
})();
