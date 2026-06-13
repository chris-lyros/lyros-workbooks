"""
Generates assets/og-image.png (1200x630) for workbooks.lyros.com.au social unfurls.

No design tool required. Renders the Lyros dark "v2" aesthetic with Pillow:
  - #1a1a1a page background
  - bottom-anchored #3a9e6e radial glow (matches the on-page hero motif)
  - the real logo-full-white.png composited top-left
  - an enterprise headline + subline in Segoe UI

Run (from repo root) with the venv Python that has Pillow:
  python scripts/make-og-image.py

Re-run any time the share copy changes. Output is a committed static file
(assets/og-image.png); the Cloudflare Worker serves it as-is with no build step.
"""
from PIL import Image, ImageDraw, ImageFont
import math

W, H = 1200, 630
BG = (26, 26, 26)          # #1a1a1a grey-black
GREEN = (58, 158, 110)     # #3a9e6e brand accent
WHITE = (244, 244, 244)    # #f4f4f4 near-white text
MUTED = (165, 165, 165)    # muted subline

FONTS = "C:/Windows/Fonts"
f_head = ImageFont.truetype(f"{FONTS}/seguisb.ttf", 66)   # Segoe UI Semibold
f_em   = ImageFont.truetype(f"{FONTS}/seguibl.ttf", 66)   # Segoe UI Black (accent word)
f_sub  = ImageFont.truetype(f"{FONTS}/segoeui.ttf", 32)
f_tag  = ImageFont.truetype(f"{FONTS}/seguisb.ttf", 24)

img = Image.new("RGB", (W, H), BG)

# Bottom-anchored radial green glow, mirrors the on-page hero gradient.
glow = Image.new("RGB", (W, H), BG)
gpx = glow.load()
cx, cy = W * 0.5, H * 1.02      # glow centre just below the canvas, like the page
rmax = W * 0.62
for y in range(H):
    for x in range(0, W, 1):
        d = math.hypot(x - cx, y - cy) / rmax
        if d >= 1.0:
            continue
        # ease-out falloff, peak ~28% green over the base background
        a = (1.0 - d) ** 2.2 * 0.30
        gpx[x, y] = (
            int(BG[0] + (GREEN[0] - BG[0]) * a),
            int(BG[1] + (GREEN[1] - BG[1]) * a),
            int(BG[2] + (GREEN[2] - BG[2]) * a),
        )
img = glow
draw = ImageDraw.Draw(img)

PAD = 80

# Logo top-left, scaled to a tidy height (keeps 1576x465 aspect).
logo = Image.open("assets/logo-full-white.png").convert("RGBA")
target_h = 70
target_w = int(logo.width * (target_h / logo.height))
logo = logo.resize((target_w, target_h), Image.LANCZOS)
img.paste(logo, (PAD, PAD), logo)

# Small kicker tag under the logo.
draw.text((PAD, PAD + target_h + 30), "WORKBOOK FINDER", font=f_tag, fill=GREEN)

# Headline block (two lines), the accent word in green/black weight.
y = 300
draw.text((PAD, y), "Describe the finance", font=f_head, fill=WHITE)
y += 84
# "workbook " white, "you need." in green accent weight on the same line
seg1 = "workbook "
draw.text((PAD, y), seg1, font=f_head, fill=WHITE)
w1 = draw.textlength(seg1, font=f_head)
draw.text((PAD + w1, y), "you need.", font=f_em, fill=GREEN)

# Subline.
y += 100
draw.text((PAD, y), "25 ready-to-use Excel workbooks for Australian finance teams.",
          font=f_sub, fill=MUTED)

# Thin green base rule for an enterprise finish.
draw.rectangle([(0, H - 8), (W, H)], fill=GREEN)

img.save("assets/og-image.png", "PNG", optimize=True)
print("Wrote assets/og-image.png", img.size)
