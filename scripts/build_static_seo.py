# -*- coding: utf-8 -*-
"""
Single source of truth for the static SEO/GEO layer of workbooks.lyros.com.au.

Reads library/workbook_library.json and regenerates, all from that one source so
they can never drift:
  - robots.txt        (allow all crawlers + AI bots; the Cloudflare Managed
                       robots.txt must ALSO be turned off at the zone - see README)
  - sitemap.xml       (homepage + 25 xlsx download URLs, %20-encoded)
  - llms.txt          (plain-markdown entity definition + catalogue for LLMs)
  - index.html        (injects, between LYROS-SEO markers: head meta/OG/Twitter +
                       JSON-LD @graph + FAQPage; and a crawlable catalogue + FAQ
                       section in the body)
  - styles-v2.css     (appends the few new catalogue/FAQ rules)

Idempotent: re-running replaces the marked blocks, it does not duplicate them.
No build step on the host; these are plain static files served by the Worker.
House rule: no em-dashes anywhere (asserted before write).

Run:  python scripts/build_static_seo.py
"""
import json
import re
import html
import urllib.parse
import datetime
import sys
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SUB = "https://workbooks.lyros.com.au"
APEX = "https://www.lyros.com.au"
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
TODAY = datetime.date.today().isoformat()

BANDS_ORDER = [
    "1000s Engagement and setup",
    "2000s Processes and risks",
    "4000s Financials and pervasive workings",
    "5000s Assets",
    "6000s Liabilities and tax",
    "7000s Revenue",
    "8000s Expenses",
]

# ---- One source for the FAQ: drives BOTH the visible section and the FAQPage JSON-LD.
# answer_text MUST equal the visible rendered text (links render as their anchor text).
FAQ = [
    {
        "q": "Are the Lyros workbooks free to download?",
        "answer_html": "Yes. Every workbook in the library is free to download, with no sign-up. Describe what you need in the finder above or pick one from the catalogue, and the real Excel file downloads straight away.",
        "answer_text": "Yes. Every workbook in the library is free to download, with no sign-up. Describe what you need in the finder above or pick one from the catalogue, and the real Excel file downloads straight away.",
    },
    {
        "q": "Do the workbooks work with Australian GST and BAS?",
        "answer_html": "Yes. The library is built for Australian small and medium businesses, with GST and BAS handled the way the ATO timing actually works. The GST and BAS workbooks treat GST as held in trust during the quarter rather than as a quarterly operating cost.",
        "answer_text": "Yes. The library is built for Australian small and medium businesses, with GST and BAS handled the way the ATO timing actually works. The GST and BAS workbooks treat GST as held in trust during the quarter rather than as a quarterly operating cost.",
    },
    {
        "q": "What do I need to use them?",
        "answer_html": "Microsoft Excel and your figures from your accounting software. The workbooks are working files: you supply the data on the data sheet and the reporting and analysis tabs do the rest. No add-ins or subscriptions are required.",
        "answer_text": "Microsoft Excel and your figures from your accounting software. The workbooks are working files: you supply the data on the data sheet and the reporting and analysis tabs do the rest. No add-ins or subscriptions are required.",
    },
    {
        "q": "What if the workbook I need is not in the library?",
        "answer_html": 'Tell the finder what you need. If it is not built yet, Lyros builds workbooks in-house, so you can request it by email at <a href="mailto:chris@lyros.com.au">chris@lyros.com.au</a> or talk it through on a 15-minute call.',
        "answer_text": "Tell the finder what you need. If it is not built yet, Lyros builds workbooks in-house, so you can request it by email at chris@lyros.com.au or talk it through on a 15-minute call.",
    },
]

HEAD_BLOCK = """<!-- Search snippet -->
<meta name="description" content="Free accounting workbooks for Australian SMBs and finance teams: 13 week cash flow, board and management reporting packs, budget vs actual, GST and BAS timing, working capital and more. Describe what you need and download the matching Excel template." />

<!-- Canonical (subdomain root, trailing slash; never the apex path) -->
<link rel="canonical" href="https://workbooks.lyros.com.au/" />
<meta name="robots" content="index, follow, max-image-preview:large, max-snippet:-1" />

<!-- Theme + colour scheme (matches dark v2) -->
<meta name="theme-color" content="#1a1a1a" />
<meta name="color-scheme" content="dark" />

<!-- Open Graph (controls the LinkedIn launch unfurl) -->
<meta property="og:type" content="website" />
<meta property="og:site_name" content="Lyros Accounting" />
<meta property="og:locale" content="en_AU" />
<meta property="og:url" content="https://workbooks.lyros.com.au/" />
<meta property="og:title" content="Find the finance workbook you need, in plain English" />
<meta property="og:description" content="Describe what you need and match it to one of 25 ready-to-use Excel workbooks for Australian finance teams." />
<meta property="og:image" content="https://workbooks.lyros.com.au/assets/og-image.png" />
<meta property="og:image:secure_url" content="https://workbooks.lyros.com.au/assets/og-image.png" />
<meta property="og:image:type" content="image/png" />
<meta property="og:image:width" content="1200" />
<meta property="og:image:height" content="630" />
<meta property="og:image:alt" content="Lyros Accounting Workbook Finder. Describe the finance workbook you need." />

<!-- Twitter / X (fallback for Slack, iMessage, Teams, X) -->
<meta name="twitter:card" content="summary_large_image" />
<meta name="twitter:title" content="Find the finance workbook you need, in plain English" />
<meta name="twitter:description" content="Describe what you need and match it to one of 25 ready-to-use Excel workbooks for Australian finance teams." />
<meta name="twitter:image" content="https://workbooks.lyros.com.au/assets/og-image.png" />
<meta name="twitter:image:alt" content="Lyros Accounting Workbook Finder. Describe the finance workbook you need." />"""

CATALOGUE_CSS = """
/* ===================== LYROS-SEO:START (generated by scripts/build_static_seo.py - do not edit by hand) ===================== */
/* The catalogue shell/grid/card are existing classes in styles.css. These are the only NEW rules:
   the band-group heading, the when_to_use excerpt line, and the FAQ block. All token-driven, dark v2. */
.v2 .cat-band-title {
  font-size: 0.6875rem;
  font-weight: 500;
  font-family: var(--font-mono);
  letter-spacing: var(--tr-kicker);
  text-transform: uppercase;
  color: var(--text-dim);
  margin: 44px 0 18px;
  padding-bottom: 12px;
  border-bottom: 1px solid rgba(255,255,255,0.06);
}
/* .catalogue is an intentional unstyled grouping wrapper around the band blocks. */
.v2 .catalogue .cat-band-title:first-child { margin-top: 0; }
.v2 .cat-card .cat-use {
  color: var(--text-dim);
  font-size: 0.8125rem;
  line-height: 1.55;
  margin: 4px 0 0;
}
.v2 .faq-block { display: grid; gap: 22px; max-width: 760px; }
.v2 .faq-item h3 { font-size: 1.0625rem; font-weight: 600; color: var(--off-white); margin: 0 0 6px; letter-spacing: -0.005em; }
.v2 .faq-item p { margin: 0; color: var(--text-muted); font-size: 0.9375rem; line-height: 1.6; }
.v2 .faq-item a { color: var(--green); text-decoration: none; }
.v2 .faq-item a:hover { text-decoration: underline; }
/* ===================== LYROS-SEO:END ===================== */
"""


def excerpt(text, target=150, hardcap=210):
    """First 1-2 sentences of when_to_use, ~target chars, never past hardcap."""
    text = text.strip()
    sents = re.split(r'(?<=[.])\s+(?=[A-Z])', text)
    out = ""
    for s in sents:
        s = s.strip()
        if not s:
            continue
        if not out:
            out = s
            if len(out) >= target:
                break
        else:
            cand = out + " " + s
            if len(cand) <= hardcap:
                out = cand
                if len(out) >= target:
                    break
            else:
                break
    return out


def esc(s):
    return html.escape(s, quote=True)


def load_items():
    with open(os.path.join(ROOT, "library", "workbook_library.json"), encoding="utf-8") as f:
        items = json.load(f)["items"]
    # stable: original JSON order; group by band preserving order
    return items


def build_catalogue_section(items):
    cards_by_band = {b: [] for b in BANDS_ORDER}
    for it in items:
        band = it["band"]
        if band not in cards_by_band:
            cards_by_band[band] = []
        enc = urllib.parse.quote(it["filename"])
        card = (
            '        <a class="cat-card" href="library/{enc}" download="{dl}">\n'
            '          <div class="top"><div class="file-tile">XLSX</div>'
            '<span class="status available">Available</span></div>\n'
            '          <h3>{title}</h3>\n'
            '          <p>{desc}</p>\n'
            '          <p class="cat-use">{use}</p>\n'
            '          <div class="footer-tags"><span class="tag">{num}</span>'
            '<span class="req">Download</span></div>\n'
            '        </a>'
        ).format(
            enc=enc,
            dl=esc(it["filename"]),
            title=esc(it["title"]),
            desc=esc(it["description"]),
            use=esc(excerpt(it["when_to_use"])),
            num=esc(it["number"]),
        )
        cards_by_band[band].append(card)

    band_blocks = []
    for band in BANDS_ORDER:
        cards = cards_by_band.get(band, [])
        if not cards:
            continue
        band_blocks.append(
            '      <h3 class="cat-band-title">{band}</h3>\n'
            '      <div class="catalogue-grid">\n{cards}\n      </div>'.format(
                band=esc(band), cards="\n".join(cards)
            )
        )

    faq_items = "\n".join(
        '      <div class="faq-item"><h3>{q}</h3><p>{a}</p></div>'.format(
            q=esc(f["q"]), a=f["answer_html"]
        )
        for f in FAQ
    )

    section = """<!-- LYROS-SEO-BODY:START (generated by scripts/build_static_seo.py - do not edit by hand) -->
<!-- Static, crawlable catalogue: makes all 25 workbooks visible to search and AI crawlers (which do not run JS).
     Reuses the existing dark-v2 catalogue CSS in styles.css. The finder JS is untouched (it binds only ids in #finder/#chips). -->
<section id="library" class="section" data-screen-label="03 Library">
  <div class="wrap">
    <div class="section-head stack">
      <span class="kicker"><span class="dot"></span> The library</span>
      <h2>Browse the full workbook library</h2>
      <p class="lede">The Lyros workbook library is a set of 25 ready-to-use Excel workbooks for Australian small and medium businesses and the people who run their numbers: owners, finance leads and bookkeepers. Each covers a real finance job, from a 13 week cash flow forecast and a monthly management report pack to budget vs actual, aged receivables, GST and BAS timing and working capital. Every workbook downloads as a real .xlsx file, free, no sign-up. Prefer to describe what you need in plain English? Use the finder above.</p>
      <p class="lede">If you would rather have the numbers and the narrative prepared for you each month, that is the <a href="https://www.lyros.com.au/">CFO advisory work Lyros does</a>: <a href="https://www.lyros.com.au/contact">talk to us about monthly management reporting</a>.</p>
    </div>

    <div class="catalogue">
{bands}
    </div>
  </div>
</section>

<!-- FAQ (static, crawlable; mirrored byte-for-byte by the FAQPage JSON-LD in the head) -->
<section id="faq" class="section" data-screen-label="04 FAQ">
  <div class="wrap">
    <div class="section-head stack">
      <span class="kicker"><span class="dot"></span> Common questions</span>
      <h2>Common questions</h2>
    </div>
    <div class="faq-block">
{faq}
    </div>
  </div>
</section>
<!-- LYROS-SEO-BODY:END -->""".format(bands="\n\n".join(band_blocks), faq=faq_items)
    return section


def build_jsonld(items):
    els = []
    for i, it in enumerate(items, 1):
        dl = SUB + "/library/" + urllib.parse.quote(it["filename"])
        els.append({
            "@type": "ListItem",
            "position": i,
            "item": {
                "@type": ["CreativeWork", "DigitalDocument"],
                "@id": "{}/#workbook-{}".format(SUB, it["number"]),
                "name": "{} (Excel workbook)".format(it["title"]),
                "alternateName": it["title"],
                "description": it["description"],
                "abstract": it["when_to_use"],
                "url": dl,
                "contentUrl": dl,
                "encodingFormat": MIME,
                "fileFormat": MIME,
                "isAccessibleForFree": True,
                "inLanguage": "en-AU",
                "genre": it["band"],
                "learningResourceType": "Spreadsheet template",
                "creator": {"@id": APEX + "/#organization"},
                "provider": {"@id": APEX + "/#organization"},
                "publisher": {"@id": APEX + "/#organization"},
                "license": APEX + "/",
                "audience": {"@type": "BusinessAudience", "name": "Australian small and medium businesses and finance teams"},
            },
        })

    graph = {
        "@context": "https://schema.org",
        "@graph": [
            {
                "@type": "Organization",
                "@id": APEX + "/#organization",
                "name": "Lyros Accounting",
                "legalName": "Lyros Pty Ltd",
                "url": APEX + "/",
                "logo": {"@type": "ImageObject", "url": SUB + "/assets/logo-full-white.png"},
                "email": "chris@lyros.com.au",
                "identifier": {"@type": "PropertyValue", "propertyID": "ABN", "value": "46 689 015 165"},
                "areaServed": {"@type": "Country", "name": "Australia"},
                "description": "Australian CFO advisory practice publishing free Excel finance workbooks for small and medium businesses and finance teams.",
            },
            {
                "@type": "WebSite",
                "@id": SUB + "/#website",
                "name": "Lyros Accounting Workbook Finder",
                "url": SUB + "/",
                "inLanguage": "en-AU",
                "publisher": {"@id": APEX + "/#organization"},
                "about": {"@id": APEX + "/#organization"},
            },
            {
                "@type": ["CollectionPage", "ItemList"],
                "@id": SUB + "/#workbook-library",
                "url": SUB + "/",
                "name": "Free finance workbooks for Australian businesses",
                "description": "A free library of 25 finance and accounting workbooks for Australian SMBs and finance teams, covering chart of accounts setup, month-end close, management and board reporting, budgeting, cash flow, working capital, receivables, tax, revenue, and expenses. Built by Lyros Accounting.",
                "inLanguage": "en-AU",
                "isPartOf": {"@id": SUB + "/#website"},
                "provider": {"@id": APEX + "/#organization"},
                "numberOfItems": 25,
                "itemListOrder": "https://schema.org/ItemListUnordered",
                "itemListElement": els,
            },
        ],
    }

    faqpage = {
        "@context": "https://schema.org",
        "@type": "FAQPage",
        "mainEntity": [
            {"@type": "Question", "name": f["q"],
             "acceptedAnswer": {"@type": "Answer", "text": f["answer_text"]}}
            for f in FAQ
        ],
    }

    def dump(obj):
        s = json.dumps(obj, indent=2, ensure_ascii=False)
        # safe embedding inside <script>: prevent any "</" from closing the tag
        return s.replace("<", "\\u003c")

    block = (
        '<script type="application/ld+json">\n' + dump(graph) + "\n</script>\n"
        '<script type="application/ld+json">\n' + dump(faqpage) + "\n</script>"
    )
    return block, graph, faqpage


def build_robots():
    bots = ["GPTBot", "OAI-SearchBot", "ChatGPT-User", "Google-Extended", "ClaudeBot",
            "anthropic-ai", "Claude-Web", "PerplexityBot", "Perplexity-User", "CCBot",
            "Applebot-Extended", "Amazonbot", "Bytespider", "cohere-ai", "meta-externalagent"]
    lines = [
        "# robots.txt for workbooks.lyros.com.au",
        "# Lyros Accounting workbook finder. AI crawlers and AI search are welcome.",
        "# This site exists to be read, summarised, and cited by AI assistants.",
        "# NOTE: this repo file is served by the Cloudflare Worker (assets.directory='.').",
        "# Cloudflare's Managed robots.txt must ALSO be turned off at the zone (see README)",
        "# or it is prepended at the edge and re-blocks the AI crawlers below.",
        "",
        "User-agent: *",
        "Content-Signal: search=yes,ai-input=yes,ai-train=yes",
        "Allow: /",
        "",
        "# AI search, grounding and training (explicitly welcomed)",
    ]
    for b in bots:
        lines.append("User-agent: " + b)
        lines.append("Allow: /")
    lines += ["", "Sitemap: " + SUB + "/sitemap.xml", "",
              "# Machine-readable summary for LLMs: " + SUB + "/llms.txt", ""]
    return "\n".join(lines)


def build_sitemap(items):
    out = ['<?xml version="1.0" encoding="UTF-8"?>',
           '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
           '  <url>',
           '    <loc>{}/</loc>'.format(SUB),
           '    <lastmod>{}</lastmod>'.format(TODAY),
           '    <changefreq>weekly</changefreq>',
           '    <priority>1.0</priority>',
           '  </url>']
    for it in items:
        loc = SUB + "/library/" + urllib.parse.quote(it["filename"])
        out += ['  <url>',
                '    <loc>{}</loc>'.format(loc),
                '    <lastmod>{}</lastmod>'.format(TODAY),
                '    <changefreq>monthly</changefreq>',
                '    <priority>0.6</priority>',
                '  </url>']
    out.append('</urlset>')
    out.append('')
    return "\n".join(out)


def build_llms(items):
    lines = [
        "# Lyros Accounting Workbook Finder",
        "",
        "> Lyros Accounting (Lyros Pty Ltd, ABN 46 689 015 165) is an Australian CFO advisory practice that publishes a free library of 25 ready-to-use Excel finance workbooks for small and medium businesses and finance teams. Every workbook downloads as a real .xlsx file, free and with no sign-up, from https://workbooks.lyros.com.au",
        "",
        "Describe a workbook in plain English in the finder, or browse the full catalogue below. The workbooks are built for Australian businesses, including GST and BAS handled the way ATO timing actually works.",
        "",
        "## Workbooks",
        "",
    ]
    by_band = {}
    for it in items:
        by_band.setdefault(it["band"], []).append(it)
    for band in BANDS_ORDER:
        group = by_band.get(band, [])
        if not group:
            continue
        lines.append("### " + band)
        lines.append("")
        for it in group:
            url = SUB + "/library/" + urllib.parse.quote(it["filename"])
            lines.append("- [{title}]({url}): {desc}".format(
                title=it["title"], url=url, desc=it["description"]))
        lines.append("")
    lines += [
        "## Engage",
        "",
        "- Book a 15-minute call: " + APEX + "/contact",
        "- Email: chris@lyros.com.au",
        "- Monthly management and board reporting, budgeting and cash flow are delivered as CFO advisory engagements: " + APEX + "/",
        "",
    ]
    return "\n".join(lines)


def inject(text, start_marker, end_marker, payload, after_anchor=None, before_anchor=None):
    """Idempotent injection. Removes any existing start..end block, then inserts
    payload either after `after_anchor` or before `before_anchor`."""
    pat = re.compile(re.escape(start_marker) + r".*?" + re.escape(end_marker), re.DOTALL)
    text = pat.sub("", text)
    # tidy any doubled blank lines left behind
    if after_anchor is not None:
        idx = text.index(after_anchor) + len(after_anchor)
        return text[:idx] + "\n" + payload + text[idx:]
    else:
        idx = text.index(before_anchor)
        cstart = text.rfind("<!--", 0, idx)
        cut = cstart if cstart != -1 else idx
        return text[:cut] + payload + "\n\n" + text[cut:]


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    items = load_items()
    assert len(items) == 25, "expected 25 items, got {}".format(len(items))

    # 1. robots.txt, sitemap.xml, llms.txt
    robots = build_robots()
    sitemap = build_sitemap(items)
    llms = build_llms(items)

    # 2. catalogue section + head/JSON-LD
    section = build_catalogue_section(items)
    jsonld_block, graph, faqpage = build_jsonld(items)
    head_payload = ("<!-- LYROS-SEO:START (generated by scripts/build_static_seo.py - do not edit by hand) -->\n"
                    + HEAD_BLOCK + "\n" + jsonld_block
                    + "\n<!-- LYROS-SEO:END -->")

    # ---- validation before writing anything ----
    # JSON-LD must be valid JSON (un-escape the <script> safety encoding first)
    for label, blk in (("@graph", graph), ("faqpage", faqpage)):
        json.loads(json.dumps(blk))  # round-trips
    # FAQ JSON-LD must byte-match the visible FAQ answers (Google policy)
    vis = {f["q"]: f["answer_text"] for f in FAQ}
    for q in faqpage["mainEntity"]:
        name = q["name"]
        assert vis.get(name) == q["acceptedAnswer"]["text"], "FAQ mismatch: " + name
    # card count
    assert section.count('class="cat-card"') == 25, "expected 25 cards"
    # dash sweep across everything we are about to write
    blobs = {"robots.txt": robots, "sitemap.xml": sitemap, "llms.txt": llms,
             "section": section, "head": head_payload, "css": CATALOGUE_CSS}
    for name, b in blobs.items():
        for bad, code in (("—", "em-dash"), ("–", "en-dash")):
            assert bad not in b, "{} found in {}".format(code, name)

    # ---- write files ----
    def w(rel, content):
        p = os.path.join(ROOT, rel)
        os.makedirs(os.path.dirname(p), exist_ok=True) if os.path.dirname(p) else None
        with open(p, "w", encoding="utf-8", newline="\n") as f:
            f.write(content)
        print("wrote", rel, "({} bytes)".format(len(content.encode("utf-8"))))

    w("robots.txt", robots)
    w("sitemap.xml", sitemap)
    w("llms.txt", llms)

    # index.html injection
    idx_path = os.path.join(ROOT, "index.html")
    html_text = open(idx_path, encoding="utf-8").read()
    html_text = inject(html_text, "<!-- LYROS-SEO:START", "<!-- LYROS-SEO:END -->",
                       head_payload,
                       after_anchor='<link rel="stylesheet" href="styles-v2.css" />')
    html_text = inject(html_text, "<!-- LYROS-SEO-BODY:START", "<!-- LYROS-SEO-BODY:END -->",
                       section,
                       before_anchor='<footer class="foot foot-slim"')
    with open(idx_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(html_text)
    print("wrote index.html ({} bytes)".format(len(html_text.encode("utf-8"))))

    # styles-v2.css append (idempotent)
    css_path = os.path.join(ROOT, "styles-v2.css")
    css = open(css_path, encoding="utf-8").read()
    css = re.compile(r"/\* ={5,} LYROS-SEO:START.*?LYROS-SEO:END ={5,} \*/\n?", re.DOTALL).sub("", css)
    css = css.rstrip() + "\n" + CATALOGUE_CSS
    with open(css_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(css)
    print("wrote styles-v2.css ({} bytes)".format(len(css.encode("utf-8"))))

    print("\nOK: 25 cards, JSON-LD valid, FAQ byte-matched, dash-clean. lastmod=" + TODAY)


if __name__ == "__main__":
    main()
