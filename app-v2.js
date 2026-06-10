/* Lyros Workbook Finder v2 — page interaction
   Hero-only funnel: describe → match → download / book a call / enquire. */
(function () {
  'use strict';

  // ── Library catalogue — mirrors the real Lyros workbook library (25 built) ──
  const LIBRARY = [
    {
      kws: ['chart of accounts', 'coa setup', 'account codes', 'chart setup', 'new chart', 'account structure'],
      title: 'Chart of Accounts Setup',
      desc:  'Setup register for the chart of accounts your bookkeeper will load into your software.',
      filename: '1010 - Chart of Accounts Setup.xlsx',
    },
    {
      kws: ['supplier mapping', 'coa mapping', 'invoice coding', 'default account', 'approval rule', 'supplier rules', 'map supplier', 'gst treatment'],
      title: 'Supplier to COA Mapping',
      desc:  'Maps each supplier to a default account, GST treatment, and approval rule.',
      filename: '1020 - Supplier to COA Mapping.xlsx',
    },
    {
      kws: ['month-end', 'month end', 'close checklist', 'closing', 'sign-off', 'close process', 'eom'],
      title: 'Month-End Close Checklist',
      desc:  'Step-by-step close checklist with sign-off, ageing, and reconciliation gates.',
      filename: '2010 - Month-End Close Checklist.xlsx',
    },
    {
      kws: ['file review', 'accounting file', 'hygiene', 'quarterly review', 'posting', 'data integrity', 'health check'],
      title: 'Accounting File Review Checklist',
      desc:  'Quarterly review of account balances, posting hygiene, and lookup mappings.',
      filename: '2020 - Accounting File Review Checklist.xlsx',
    },
    {
      kws: ['management report', 'management reporting', 'monthly management', 'monthly report', 'mgmt', 'management pack', 'leadership', 'monthly p&l'],
      title: 'Management Reporting Pack',
      desc:  'Monthly P&L pack with headline, monthly trend, quarter comparison, wages, and working capital.',
      filename: '4000 - Management Reporting Pack.xlsx',
    },
    {
      kws: ['board', 'board pack', 'directors', 'governance', 'investor', 'board report'],
      title: 'Board Reporting Pack',
      desc:  'One-page board read-out plus detail, cash position, and commentary.',
      filename: '4001 - Board Reporting Pack.xlsx',
    },
    {
      kws: ['budget vs actual', 'budget', 'vs actual', 'variance to budget', 'actuals', 'reforecast'],
      title: 'Budget vs Actual',
      desc:  'Side-by-side actuals and budget with monthly and YTD variance.',
      filename: '4010 - Budget vs Actual.xlsx',
    },
    {
      kws: ['variance bridge', 'ytd bridge', 'year-to-date', 'variance drivers', 'monthly variance', 'walk the bridge'],
      title: 'Monthly Variance and YTD Bridge',
      desc:  'Walk-the-bridge from prior-month and YTD variance drivers to current actuals.',
      filename: '4020 - Monthly Variance and YTD Bridge.xlsx',
    },
    {
      kws: ['department', 'cost centre', 'cost center', 'rollup', 'tracking categories', 'divisional', 'per department'],
      title: 'Departmental Variance Rollup',
      desc:  'Department-level P&L rollup with variance to budget per cost centre.',
      filename: '4030 - Departmental Variance Rollup.xlsx',
    },
    {
      kws: ['scenario', 'upside', 'downside', 'driver-based', 'what if', 'sensitivity', 'forecast scenarios', 'scenario planning'],
      title: 'Budgeting Scenario Flex',
      desc:  'Driver-based scenario flex (base, upside, downside) for forward 12 months.',
      filename: '4040 - Budgeting Scenario Flex.xlsx',
    },
    {
      kws: ['consolidation', 'consolidated', 'multiple entities', 'multi-entity', 'intercompany', 'elimination', 'trial balance', 'group accounts'],
      title: 'Consolidation from Trial Balances',
      desc:  'Per-entity trial balances with eliminations and combined consolidated P&L and BS.',
      filename: '4050 - Consolidation from Trial Balances.xlsx',
    },
    {
      kws: ['cash flow', 'cashflow', 'cash', '13 week', '13-week', 'thirteen week', 'rolling cash', 'cash forecast'],
      title: '13 Week Rolling Cash Flow',
      desc:  'Forward 13-week cash flow with inflows, outflows, and opening and closing balances.',
      filename: '5100 - 13 Week Rolling Cash Flow.xlsx',
    },
    {
      kws: ['cash position', 'bank account', 'bank balances', 'daily cash', 'by bank', 'multiple banks'],
      title: 'Cash Position by Bank',
      desc:  'Daily cash position split by bank account with reconciliation to the GL.',
      filename: '5110 - Cash Position by Bank.xlsx',
    },
    {
      kws: ['gst timing', 'bas timing', 'gst cash', 'bas payment', 'bas cash', 'gst', 'tax set aside'],
      title: 'GST and BAS Cash Flow Timing',
      desc:  'Cash impact of upcoming BAS payments and refunds across the next four quarters.',
      filename: '5120 - GST and BAS Cash Flow Timing.xlsx',
    },
    {
      kws: ['working capital', 'cash conversion', 'debtor days', 'creditor days', 'inventory days', 'dso', 'dpo', 'dio'],
      title: 'Working Capital Cycle',
      desc:  'DSO, DPO, DIO, and cash conversion cycle trend with rolling 12-month chart.',
      filename: '5150 - Working Capital Cycle.xlsx',
    },
    {
      kws: ['chase list', 'chase', 'collections', 'follow-up', 'overdue invoices', 'debtors list', 'credit control'],
      title: 'Aged Receivables Chase List',
      desc:  'Customer-level ageing buckets with follow-up tier and priority for the credit team.',
      filename: '5200 - Aged Receivables Chase List.xlsx',
    },
    {
      kws: ['aged receivables', 'ageing', 'aging', 'receivables', 'debtor', 'debtor concentration', 'ar analysis'],
      title: 'Aged Receivables Analysis',
      desc:  'Heatmap of customer ageing exposure, concentration, and top-debtor drill-down.',
      filename: '5210 - Aged Receivables Analysis.xlsx',
    },
    {
      kws: ['depreciation', 'asset register', 'fixed asset', 'wdv', 'written down', 'disposals', 'roll forward'],
      title: 'Depreciation Roll Forward',
      desc:  'Asset register with WDV opening, additions, disposals, depreciation, and closing balances.',
      filename: '5300 - Depreciation Roll Forward.xlsx',
    },
    {
      kws: ['lodgement', 'lodgment', 'bas tracker', 'bas due', 'bas register', 'bas', 'lodged'],
      title: 'GST and BAS Lodgement Tracker',
      desc:  'Quarterly lodgement and payment register with due dates, amounts, and status.',
      filename: '6100 - GST and BAS Lodgement Tracker.xlsx',
    },
    {
      kws: ['fbt', 'fringe benefits', 'entertainment', 'company vehicle', 'novated', 'employee benefits'],
      title: 'FBT Exposure Quick Check',
      desc:  'Common FBT triggers screened against transactional data with exposure estimate.',
      filename: '6200 - FBT Exposure Quick Check.xlsx',
    },
    {
      kws: ['leave liability', 'leave balance', 'accrued wages', 'annual leave', 'leave provision', 'long service'],
      title: 'Wages and Leave Liability',
      desc:  'Period-end leave balance and accrued wages liability per employee.',
      filename: '6300 - Wages and Leave Liability.xlsx',
    },
    {
      kws: ['top customers', 'customer revenue', 'customer concentration', 'biggest customers', 'revenue by customer', 'key accounts'],
      title: 'Top Customers by Revenue',
      desc:  'Ranked customer revenue with year-on-year and concentration share.',
      filename: '7000 - Top Customers by Revenue.xlsx',
    },
    {
      kws: ['margin', 'margin bridge', 'revenue bridge', 'margin analysis', 'gross margin', 'price volume mix', 'revenue decomposition'],
      title: 'Revenue and Margin Bridge',
      desc:  'Account-level revenue decomposition between prior and current quarter.',
      filename: '7010 - Revenue and Margin Bridge.xlsx',
    },
    {
      kws: ['supplier spend', 'top supplier', 'supplier review', 'vendor spend', 'procurement', 'creditor spend', 'supplier negotiation'],
      title: 'Top Supplier Review',
      desc:  'Ranked supplier spend with ABN-based dedup and year-on-year change.',
      filename: '8000 - Top Supplier Review.xlsx',
    },
    {
      kws: ['payroll', 'payroll to revenue', 'payroll ratio', 'wage cost', 'wages ratio', 'headcount cost', 'staff cost'],
      title: 'Payroll to Revenue Ratio',
      desc:  'Payroll cost as a share of revenue over time, by department.',
      filename: '8100 - Payroll to Revenue Ratio.xlsx',
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
  const downloadBtn = document.getElementById('downloadBtn');

  const enquireLink = document.getElementById('enquireLink');
  const requestLink = document.getElementById('requestLink');

  const EMAIL = 'chris@lyros.com.au';

  function setMailto(link, subject, body) {
    if (!link) return;
    link.href = 'mailto:' + EMAIL +
      '?subject=' + encodeURIComponent(subject) +
      '&body=' + encodeURIComponent(body);
  }

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

  const SEND_LABEL = `Find a workbook <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 5l7 7-7 7"/></svg>`;
  const RETRY_LABEL = `Try again <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 5l7 7-7 7"/></svg>`;
  let hasSearched = false;

  function setSearching() {
    show('input', 'searching');
    steps.forEach((s) => { s.classList.remove('in', 'active', 'done'); });
    send.disabled = true;
    send.innerHTML = `<span style="display:inline-block;width:12px;height:12px;border-radius:50%;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;animation:spin 0.7s linear infinite;"></span> Searching…`;
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
      // Serve the real workbook file from the library.
      if (downloadBtn) {
        downloadBtn.href = 'library/' + encodeURIComponent(r.item.filename);
        downloadBtn.setAttribute('download', r.item.filename);
      }
      setMailto(
        enquireLink,
        'Workbook enquiry — ' + r.item.title,
        'Hi Chris,\n\nI used the workbook finder and matched with "' + r.item.title + '".\n\nWhat I asked for:\n"' + query.trim() + '"\n\nI\'d like to talk about getting this populated with my accounting data.\n\nThanks,'
      );
      show('input', 'match');
    } else {
      setMailto(
        requestLink,
        'Workbook request — not in the library yet',
        'Hi Chris,\n\nI used the workbook finder but there was no match. Here\'s what I need:\n\n"' + query.trim() + '"\n\nPlease let me know if you can build it.\n\nThanks,'
      );
      show('input', 'missing');
    }

    hasSearched = true;
    send.disabled = false;
    send.innerHTML = RETRY_LABEL;
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
    const text = btn.dataset.chip.toLowerCase();
    const article = /^[aeiou]/.test(text) ? 'an' : 'a';
    const sentence = `I need ${article} ${text} workbook for an SME running on Xero.`;
    input.value = sentence;
    runSearch(sentence);
  });
})();
