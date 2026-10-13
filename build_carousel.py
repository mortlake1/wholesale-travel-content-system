#!/usr/bin/env python3
"""Day-9 carousel artwork: 7 slides at 1080x1350 (SVG + PNG) plus an inline preview page."""
import os, subprocess

W, H = 1080, 1350
M = 92
INK = "#14283A"
MUTED = "#5A6B75"
TEAL = "#0E7C7B"
GOLD = "#C8922A"
NAVY = "#0F2438"
PAPER = "#F7F3EC"
LIGHTMUT = "#9FB6C2"
RED = "#A93226"
CROSS = "#E58E7A"
DOT_OFF_PAPER = "#D8D2C6"
DOT_OFF_NAVY = "#3E5566"

os.makedirs("carousel_day9", exist_ok=True)


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def svg_open(bg):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
            f'viewBox="0 0 {W} {H}"><rect width="{W}" height="{H}" fill="{bg}"/>')


def kicker(y, text, colour):
    return (f'<rect x="{M}" y="{y-34}" width="64" height="8" fill="{GOLD}"/>'
            f'<text x="{M}" y="{y}" font-family="DejaVu Sans" font-size="26" '
            f'font-weight="bold" fill="{colour}">{esc(text)}</text>')


def headline(y, lines, colour, size=58):
    out = []
    for i, ln in enumerate(lines):
        out.append(f'<text x="{M}" y="{y + i*int(size*1.28)}" font-family="DejaVu Sans" '
                   f'font-size="{size}" font-weight="bold" fill="{colour}">{esc(ln)}</text>')
    return "".join(out), y + len(lines) * int(size * 1.28)


def bullets(y, items, colour, marker="dot", mcolour=TEAL, size=34, lh=50):
    out = []
    for item in items:
        if marker == "dot":
            out.append(f'<circle cx="{M+8}" cy="{y-11}" r="7" fill="{mcolour}"/>')
        else:
            out.append(f'<path d="M {M} {y-26} l 20 20 M {M+20} {y-26} l -20 20" '
                       f'stroke="{mcolour}" stroke-width="6" stroke-linecap="round" fill="none"/>')
        for j, ln in enumerate(item):
            out.append(f'<text x="{M+40}" y="{y + j*lh}" font-family="DejaVu Sans" '
                       f'font-size="{size}" fill="{colour}">{esc(ln)}</text>')
        y += len(item) * lh + 34
    return "".join(out), y


def footer(idx, dark):
    mut = LIGHTMUT if dark else MUTED
    off = DOT_OFF_NAVY if dark else DOT_OFF_PAPER
    dots = []
    x0 = W - M - (7 * 34 - 12)
    for i in range(7):
        c = GOLD if i == idx else off
        dots.append(f'<circle cx="{x0 + i*34}" cy="{H-92}" r="9" fill="{c}"/>')
    tag = (f'<text x="{M}" y="{H-150}" font-family="DejaVu Sans" font-size="26" fill="{mut}">'
           f'wholesale travel, explained plainly</text>')
    handle = (f'<text x="{M}" y="{H-84}" font-family="DejaVu Sans" font-size="26" '
              f'font-weight="bold" fill="{mut}">@yourhandle</text>')
    return tag + handle + "".join(dots)


def slide(idx, dark, body_fn):
    bg = NAVY if dark else PAPER
    parts = [svg_open(bg)]
    parts.append(body_fn(dark))
    parts.append(footer(idx, dark))
    parts.append("</svg>")
    return "".join(parts)


# ---------------------------------------------------------------- slides
def s1(dark):
    o = [kicker(150, "THE PARTNER SIDE, EXPLAINED PLAINLY", GOLD)]
    h, y = headline(300, ["What this role is.", "What it is not."], "#FFFFFF", size=76)
    o.append(h)
    o.append(f'<rect x="{M}" y="{y+16}" width="180" height="10" fill="{GOLD}"/>')
    sub = ["Wholesale travel distribution,", "the honest version.", "",
           "No hype. No numbers I can't prove."]
    yy = y + 92
    for ln in sub:
        if ln:
            o.append(f'<text x="{M}" y="{yy}" font-family="DejaVu Sans" font-size="36" '
                     f'fill="{LIGHTMUT}">{esc(ln)}</text>')
        yy += 54
    o.append(f'<text x="{M}" y="{H-260}" font-family="DejaVu Sans" font-size="32" '
             f'font-weight="bold" fill="{GOLD}">Swipe  →   ·   Save this for later</text>')
    return "".join(o)


def s2(dark):
    o = [kicker(150, "SLIDE 2  ·  WHAT IT IS", TEAL)]
    h, y = headline(268, ["You introduce travellers.", "The company pays", "the commission."], INK)
    o.append(h)
    b, _ = bullets(y + 40, [
        ["You show people contracted hotel and", "resort rates they can't see on the", "public booking sites."],
        ["When one of them books a real trip,", "you're paid a commission on", "that booking."],
        ["That's the mechanism. There's a", "booking confirmation behind", "every rand."],
    ], "#33454F")
    o.append(b)
    return "".join(o)


def s3(dark):
    o = [kicker(150, "SLIDE 3  ·  WHAT IT IS NOT", GOLD)]
    h, y = headline(268, ["Three things this", "will never be."], "#FFFFFF")
    o.append(h)
    b, _ = bullets(y + 46, [
        ["A guaranteed income. Nobody can", "promise you an earnings figure", "— including me."],
        ["Getting paid to recruit. Commissions", "come from travel bookings, not", "from signing people up."],
        ["Something that runs without you.", "Your results depend entirely on", "the customers you build."],
    ], "#E8EEF2", marker="cross", mcolour=CROSS)
    o.append(b)
    return "".join(o)


def s4(dark):
    o = [kicker(150, "SLIDE 4  ·  WHAT IT COSTS", TEAL)]
    h, y = headline(268, ["The fee, out loud."], INK, size=64)
    o.append(h)
    y += 46
    rows = [("Joining fee", "R{joining}, once"),
            ("Ongoing", "R{monthly} per month"),
            ("Included", "portal, contracted rates,"),
            ("", "training and support")]
    o.append(f'<rect x="{M}" y="{y}" width="{W-2*M}" height="330" fill="#FFFFFF" stroke="{GOLD}" '
             f'stroke-width="3" stroke-dasharray="12 9" rx="14"/>')
    yy = y + 74
    for label, val in rows:
        if label:
            o.append(f'<text x="{M+36}" y="{yy}" font-family="DejaVu Sans" font-size="32" '
                     f'font-weight="bold" fill="{INK}">{esc(label)}</text>')
        o.append(f'<text x="{W-M-36}" y="{yy}" font-family="DejaVu Sans" font-size="32" '
                 f'fill="{MUTED}" text-anchor="end">{esc(val)}</text>')
        yy += 66
    cy = y + 330 + 60
    o.append(f'<rect x="{M}" y="{cy}" width="10" height="118" fill="{TEAL}"/>')
    o.append(f'<text x="{M+40}" y="{cy+42}" font-family="DejaVu Sans" font-size="34" '
             f'font-weight="bold" fill="{INK}">No obligation to book anything.</text>')
    o.append(f'<text x="{M+40}" y="{cy+96}" font-family="DejaVu Sans" font-size="34" '
             f'font-weight="bold" fill="{INK}">No obligation to invite anyone.</text>')
    o.append(f'<text x="{M}" y="{H-246}" font-family="DejaVu Sans" font-size="26" fill="{MUTED}">'
             f'Replace these placeholders with your company’s</text>')
    o.append(f'<text x="{M}" y="{H-208}" font-family="DejaVu Sans" font-size="26" fill="{MUTED}">'
             f'current fee schedule before this goes live.</text>')
    return "".join(o)


def s5(dark):
    o = [kicker(150, "SLIDE 5  ·  THE HONEST FIT", TEAL)]
    h, y = headline(268, ["Who it suits.", "Who it doesn't."], INK)
    o.append(h)
    y += 30
    o.append(f'<text x="{M}" y="{y}" font-family="DejaVu Sans" font-size="36" font-weight="bold" '
             f'fill="{TEAL}">Suits you if:</text>')
    y += 56
    b, y = bullets(y, [
        ["you already love planning travel"],
        ["you're happy to talk to people about it"],
        ["you want a side build, not a salary"],
        ["you can handle “no” without quitting"],
    ], "#33454F", size=33, lh=40)
    o.append(b)
    y += 8
    o.append(f'<text x="{M}" y="{y}" font-family="DejaVu Sans" font-size="36" font-weight="bold" '
             f'fill="{RED}">Wrong fit if:</text>')
    y += 56
    b2, _ = bullets(y, [
        ["you need money this month"],
        ["you want it to run itself"],
        ["you won't tell anyone about it"],
        ["you only want the savings — join as a", "customer instead. Genuinely fine."],
    ], "#33454F", size=33, lh=40)
    o.append(b2)
    return "".join(o)


def s6(dark):
    o = [kicker(150, "SLIDE 6  ·  THE DISCLOSURE", TEAL)]
    h, y = headline(268, ["Read this before anyone", "signs anything."], INK)
    o.append(h)
    y += 36
    body = ["Commission is earned only when real", "travellers book real holidays. What",
            "anyone earns depends entirely on the", "customers they build.", "",
            "So treat any specific income figure", "— from me or anyone else — as one",
            "person's story, not a promise."]
    for ln in body:
        if ln:
            o.append(f'<text x="{M}" y="{y}" font-family="DejaVu Sans" font-size="34" '
                     f'fill="#33454F">{esc(ln)}</text>')
        y += 52
    y += 40
    o.append(f'<rect x="{M}" y="{y}" width="{W-2*M}" height="190" fill="#FFFFFF" stroke="{TEAL}" '
             f'stroke-width="3" rx="14"/>')
    o.append(f'<text x="{M+36}" y="{y+66}" font-family="DejaVu Sans" font-size="31" '
             f'font-weight="bold" fill="{INK}">Company: {{registered name}}</text>')
    o.append(f'<text x="{M+36}" y="{y+112}" font-family="DejaVu Sans" font-size="31" '
             f'fill="{MUTED}">Reg. no {{number}} — ask me for the certificate</text>')
    o.append(f'<text x="{M+36}" y="{y+156}" font-family="DejaVu Sans" font-size="31" '
             f'fill="{MUTED}">and the compensation policy. One message, both docs.</text>')
    return "".join(o)


def s7(dark):
    o = [kicker(150, "SLIDE 7  ·  NEXT STEP", GOLD)]
    h, y = headline(300, ["Want the full picture?"], "#FFFFFF", size=72)
    o.append(h)
    o.append(f'<rect x="{M}" y="{y+10}" width="180" height="10" fill="{GOLD}"/>')
    yy = y + 110
    body = ["DM me BLUEPRINT and I'll send the", "20-minute overview plus the fee",
            "schedule.", "", "Then ask me the awkward questions.", "I'd rather you did that now."]
    for ln in body:
        if ln:
            o.append(f'<text x="{M}" y="{yy}" font-family="DejaVu Sans" font-size="38" '
                     f'fill="{LIGHTMUT}">{esc(ln)}</text>')
        yy += 58
    o.append(f'<text x="{M}" y="{H-252}" font-family="DejaVu Sans" font-size="26" fill="{LIGHTMUT}">'
             f'Rates and availability change daily.</text>')
    o.append(f'<text x="{M}" y="{H-214}" font-family="DejaVu Sans" font-size="26" fill="{LIGHTMUT}">'
             f'Nothing here is an earnings claim.</text>')
    return "".join(o)


SLIDES = [(s1, True), (s2, False), (s3, True), (s4, False), (s5, False), (s6, False), (s7, True)]

svgs = []
for i, (fn, dark) in enumerate(SLIDES):
    markup = slide(i, dark, fn)
    svgs.append(markup)
    with open(f"carousel_day9/slide{i+1}.svg", "w") as f:
        f.write(markup)

for i in range(1, 8):
    subprocess.run(["convert", "-background", "none", f"carousel_day9/slide{i}.svg",
                    f"carousel_day9/slide{i}.png"], check=True)

# ------------------------------------------------------------- preview page
grid = "\n".join(
    f'<figure><figcaption>Slide {i+1}</figcaption>{svg}</figure>' for i, svg in enumerate(svgs))
html = f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<title>Day-9 Carousel — What This Role Is. What It Is Not.</title>
<style>
 body {{ margin:0; background:#0B141B; color:#E8EEF2; font-family:-apple-system,'Segoe UI',Roboto,sans-serif; }}
 header {{ padding:44px 6vw 8px; }}
 h1 {{ font-size:30px; margin:0 0 8px; }}
 p.sub {{ color:#9FB6C2; max-width:70ch; line-height:1.55; margin:0 0 6px; }}
 .warn {{ background:#3A2E14; border:1px solid #C8922A; color:#F2D9A0; padding:12px 16px;
          border-radius:10px; max-width:70ch; margin:18px 0 0; line-height:1.5; font-size:14px; }}
 main {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(300px,1fr));
         gap:26px; padding:34px 6vw 70px; }}
 figure {{ margin:0; }}
 figcaption {{ font-size:13px; letter-spacing:.08em; text-transform:uppercase;
               color:#C8922A; margin-bottom:10px; font-weight:700; }}
 figure svg {{ width:100%; height:auto; border-radius:12px; display:block;
               box-shadow:0 12px 40px rgba(0,0,0,.5); }}
 footer {{ padding:0 6vw 60px; color:#9FB6C2; font-size:14px; line-height:1.6; }}
 code {{ background:#16222C; padding:2px 6px; border-radius:6px; }}
</style></head><body>
<header>
 <h1>Day-9 carousel — “What this role is. What it is not.”</h1>
 <p class="sub">Seven slides at 1080×1350, ready for Instagram and Facebook. PNG exports sit beside
 the SVG sources in <code>carousel_day9/</code>. Slides 3, 5 and 6 are the compliance backbone —
 they are the slides that make the other four believable. Do not cut them for brevity.</p>
 <div class="warn"><strong>Before posting:</strong> replace the placeholders —
 <code>@yourhandle</code> (every footer), <code>R{{joining}}</code> / <code>R{{monthly}}</code>
 (slide 4), <code>{{registered name}}</code> / <code>{{number}}</code> (slide 6). Every one of them is
 visible in the artwork below so you can't miss it.</div>
</header>
<main>
{grid}
</main>
<footer>
 Exported files: <code>slide1.svg … slide7.svg</code> (editable source) and
 <code>slide1.png … slide7.png</code> (1080×1350, upload-ready).<br>
 Caption for this carousel lives in <code>Week1_Captions_and_Hashtags.md</code> under “Day 9 bonus”.
 Post Sunday or Tuesday lunchtime; carousels earn their saves in the first two hours.
</footer>
</body></html>"""

with open("carousel_day9/carousel_preview.html", "w") as f:
    f.write(html)

print("slides + pngs + preview written")
