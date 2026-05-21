/* Lyros Workbook Finder — page interaction */
(function () {
  'use strict';

  // ── Library catalogue used for the hero matcher ──────────────────────────
  // Each entry has a few keywords + a workbook title/desc/tag set.
  const LIBRARY = [
    {
      kws: ['monthly', 'management', 'report', 'pack', 'mgmt', 'commentary', 'mom', 'month'],
      title: 'Monthly Management Report Pack',
      desc:  'Revenue, gross margin, wages, cash movement and month-on-month commentary. Designed for Xero exports.',
      tags:  ['XLSX', 'Xero-ready', 'Management reporting', 'SME finance', 'Free download'],
    },
    {
      kws: ['cash', 'flow', 'cashflow', 'forecast', '13-week', 'payroll', 'gst', 'supplier'],
      title: 'Cash Flow Timing Workbook',
      desc:  'Rolling 13-week cash forecast with payroll risk, GST timing and supplier payment scheduling.',
      tags:  ['XLSX', 'Xero-ready', 'Cash flow', 'Treasury', 'Free download'],
    },
    {
      kws: ['board', 'pack', 'directors', 'governance', 'revenue', 'margin', 'expenses', 'cash movement'],
      title: 'Board Reporting Pack',
      desc:  'One-page summary, KPIs, financials, cash and a CFO commentary template. Board-ready.',
      tags:  ['XLSX', 'Board', 'Reporting', 'Governance', 'Free download'],
    },
    {
      kws: ['budget', 'actual', 'variance', 'forecast', 'fp&a', 'plan', 'vs'],
      title: 'Budget vs Actual Workbook',
      desc:  'Monthly variance, YTD bridge, departmental rollup with commentary cells.',
      tags:  ['XLSX', 'Xero-ready', 'FP&A', 'Budgeting', 'Free download'],
    },
    {
      kws: ['supplier', 'spend', 'review', 'creditor', 'ap', 'vendor', 'procurement'],
      title: 'Supplier Spend Review',
      desc:  'Top suppliers, ABN dedup, payment terms, concentration risk and renewals.',
      tags:  ['XLSX', 'AP', 'Cost control', 'Suppliers', 'Free download'],
    },
    {
      kws: ['payroll', 'wages', 'staff', 'headcount', 'risk', 'super', 'leave'],
      title: 'Payroll & Wages Analysis',
      desc:  'Headcount, ratios, leave liability and payroll-to-revenue trend.',
      tags:  ['XLSX', 'Payroll', 'People', 'STP', 'Free download'],
    },
    {
      kws: ['gst', 'bas', 'tax', 'compliance', 'instalment', 'paygi', 'lodgement'],
      title: 'GST & BAS Timing Tracker',
      desc:  'BAS cycle, GST cash impact and instalment ladder for quarterly lodgement.',
      tags:  ['XLSX', 'Tax', 'Compliance', 'Beta', 'Free download'],
    },
    {
      kws: ['margin', 'gross', 'bridge', 'volume', 'price', 'mix', 'product'],
      title: 'Revenue & Margin Bridge',
      desc:  'Volume, price and mix bridge. Explains what moved last month.',
      tags:  ['XLSX', 'Revenue', 'Margin', 'Beta', 'Free download'],
    },
    {
      kws: ['working capital', 'debtor', 'creditor', 'days', 'inventory', 'dso', 'dpo'],
      title: 'Working Capital Tracker',
      desc:  'Debtor days, creditor days, inventory days and working capital cycle.',
      tags:  ['XLSX', 'Cash', 'Operations', 'Coming soon'],
    },
    {
      kws: ['ndis', 'support', 'disability', 'provider', 'services'],
      title: 'Services / NDIS Monthly Pack',
      desc:  'Monthly management reporting tailored to a services or NDIS provider: revenue, payroll mix, cash and compliance notes.',
      tags:  ['XLSX', 'Xero-ready', 'NDIS', 'Services', 'Free download'],
    },
  ];

  function rank(query) {
    const q = query.toLowerCase();
    let best = null, bestScore = 0;
    for (const item of LIBRARY) {
      let s = 0;
      for (const kw of item.kws) {
        if (q.includes(kw)) s += kw.length > 4 ? 2 : 1;
      }
      if (s > bestScore) { bestScore = s; best = item; }
    }
    return { item: best, score: bestScore };
  }

  // ── DOM refs ────────────────────────────────────────────────────────────
  const prompt = document.getElementById('promptBox');
  const input  = document.getElementById('promptInput');
  const send   = document.getElementById('sendBtn');

  const stateInput     = prompt.querySelector('[data-state="input"]');
  const stateSearching = prompt.querySelector('[data-state="searching"]');
  const stateMatch     = prompt.querySelector('[data-state="match"]');
  const stateMissing   = prompt.querySelector('[data-state="missing"]');

  const steps = [...prompt.querySelectorAll('.search-step')];
  const matchTitle = document.getElementById('matchTitle');
  const matchDesc  = document.getElementById('matchDesc');
  const matchTags  = document.getElementById('matchTags');

  // Animation speed is settable via Tweaks panel.
  function getSpeed() { return Number(document.documentElement.dataset.demoSpeed || 1); }
  const beat = (ms) => ms / getSpeed();

  let running = false;

  // ── States ───────────────────────────────────────────────────────────────
  function show(...names) {
    [stateInput, stateSearching, stateMatch, stateMissing].forEach((el) => { el.hidden = true; });
    names.forEach((n) => {
      const el = prompt.querySelector(`[data-state="${n}"]`);
      if (el) el.hidden = false;
    });
  }

  function reset() {
    show('input');
    send.disabled = false;
    send.innerHTML = `Find a workbook <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 5l7 7-7 7"/></svg>`;
    steps.forEach((s) => { s.classList.remove('in', 'active', 'done'); });
    input.focus();
  }

  function setSearching() {
    show('searching');
    steps.forEach((s) => { s.classList.remove('in', 'active', 'done'); });
    send.disabled = true;
    send.innerHTML = `<span style="display:inline-block;width:12px;height:12px;border-radius:50%;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;animation:spin 0.7s linear infinite;"></span> Searching…`;
  }

  function activateStep(i, finalState = 'done') {
    steps.forEach((s, idx) => {
      if (idx < i) s.classList.add('in', 'done');
      else if (idx === i) {
        s.classList.add('in', finalState === 'active' ? 'active' : 'active');
      }
    });
  }

  async function runSearch(query) {
    if (running) return;
    running = true;

    if (!query.trim()) { running = false; return; }

    setSearching();

    // Step 1: Understanding
    await wait(120);
    steps[0].classList.add('in', 'active');
    await wait(beat(750));
    steps[0].classList.remove('active'); steps[0].classList.add('done');

    // Step 2: Searching library
    steps[1].classList.add('in', 'active');
    await wait(beat(950));
    steps[1].classList.remove('active'); steps[1].classList.add('done');

    // Step 3: Closest match
    steps[2].classList.add('in', 'active');
    await wait(beat(700));
    steps[2].classList.remove('active'); steps[2].classList.add('done');

    await wait(beat(220));

    // Result
    const r = rank(query);
    if (r.item && r.score > 0) {
      matchTitle.textContent = r.item.title;
      matchDesc.textContent  = r.item.desc;
      matchTags.innerHTML = '';
      r.item.tags.forEach((t, idx) => {
        const el = document.createElement('span');
        el.className = 'tag' + (idx === 0 || /free/i.test(t) ? ' green' : '');
        el.textContent = t;
        matchTags.appendChild(el);
      });
      show('match');
    } else {
      show('missing');
    }

    send.disabled = false;
    send.innerHTML = `Find a workbook <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 5l7 7-7 7"/></svg>`;
    running = false;
  }

  function wait(ms) { return new Promise((r) => setTimeout(r, ms)); }

  // ── Wire up ─────────────────────────────────────────────────────────────
  function trySearch() {
    const q = input.value.trim();
    if (!q) { input.focus(); return; }
    runSearch(q);
  }

  send.addEventListener('click', trySearch);
  input.addEventListener('keydown', (e) => {
    if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') { e.preventDefault(); trySearch(); }
  });

  document.getElementById('chips').addEventListener('click', (e) => {
    const btn = e.target.closest('.chip');
    if (!btn) return;
    const text = btn.dataset.chip;
    // Compose a natural sentence around the chip
    const sentence = `I need a ${text.toLowerCase()} workbook for an SME running on Xero.`;
    input.value = sentence;
    show('input');
    runSearch(sentence);
  });

  prompt.addEventListener('click', (e) => {
    const r = e.target.closest('[data-reset]');
    if (r) { reset(); }
  });

  // ── Catalogue filter ────────────────────────────────────────────────────
  const filters = document.getElementById('catFilters');
  const catalogue = document.getElementById('catalogue');
  if (filters && catalogue) {
    filters.addEventListener('click', (e) => {
      const b = e.target.closest('.cat-filter');
      if (!b) return;
      filters.querySelectorAll('.cat-filter').forEach((x) => x.classList.remove('active'));
      b.classList.add('active');
      const f = b.dataset.filter;
      catalogue.querySelectorAll('.cat-card').forEach((card) => {
        card.style.display = (f === 'all' || card.dataset.category === f) ? '' : 'none';
      });
    });
  }
})();
