#!/usr/bin/env python3
"""Builds the 30-day SA travel-distribution content calendar (.xlsx)."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date, timedelta

# ---------------------------------------------------------------- palette
NAVY = "0F2438"
TEAL = "0E7C7B"
GOLD = "C8922A"
LIGHT = "F2F6F8"
W1 = "E8F1F5"   # week 1 tint
W2 = "EAF3EC"   # week 2 tint
W3 = "FBF1E3"   # week 3 tint
W4 = "F3ECF7"   # week 4 tint
WEEK_TINT = {1: W1, 2: W2, 3: W3, 4: W4, 5: W4}
RED = "C0392B"
AMBER = "B9770E"
GREEN = "1E7A46"

WHITE_BOLD = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BODY = Font(name="Calibri", size=11)
BODY_B = Font(name="Calibri", size=11, bold=True)
SMALL = Font(name="Calibri", size=9, color="5A6B75")
TITLE = Font(name="Calibri", size=20, bold=True, color=NAVY)
SUB = Font(name="Calibri", size=12, color="3F5563")
H2 = Font(name="Calibri", size=13, bold=True, color=NAVY)
WRAP = Alignment(wrap_text=True, vertical="top")
WRAP_C = Alignment(wrap_text=True, vertical="center", horizontal="center")
THIN = Side(style="thin", color="C9D4DA")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()


def style_header(ws, row, ncols, fill=NAVY):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = WHITE_BOLD
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = 34


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def body_row(ws, row, ncols, tint=None):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = BODY
        cell.alignment = WRAP
        cell.border = BORDER
        if tint:
            cell.fill = PatternFill("solid", fgColor=tint)


# ================================================================ 1. README
ws = wb.active
ws.title = "Start Here"
ws.sheet_view.showGridLines = False
set_widths(ws, [3, 26, 100])

ws["B2"] = "30-Day Wholesale Travel Content System"
ws["B2"].font = TITLE
ws["B3"] = "South African market  •  Savings-proof model (no income claims)  •  DM-first funnel"
ws["B3"].font = SUB

rows = [
    ("THE STRATEGY IN ONE LINE",
     "Sell the holiday first, the business second. Every piece of content must be true even if "
     "nobody ever joins your team — because the travel savings are what you can actually prove."),

    ("WHY THIS CHANGED",
     "Your original drafts argued legality out loud ('is this a Ponzi scheme?'), made income "
     "promises ('paid monthly for years', 'compounding residual wealth') and used an "
     "unsubstantiated 'up to 70% off'. You have real booking screenshots — so the calendar below "
     "puts every claim on proof you actually hold, and moves the business conversation into DMs "
     "where it converts better and carries far less risk."),

    ("THE TWO LANES",
     "Lane A = TRAVELLER. Wants a cheaper holiday. Public content, high volume, easy proof. "
     "This is 80% of the calendar.\n"
     "Lane B = PARTNER. Wants to build something. Invitation-only, DM and live formats, "
     "never earnings promises. This is 20% of the calendar.\n"
     "Lane B leads almost always arrive through Lane A content. Do not run Lane B ads cold."),

    ("HOW TO USE THE CALENDAR",
     "Tab '30-Day Calendar' = one row per day: format, hook, content beats, CTA, the proof asset "
     "you need, and a risk note. Filter by Week, Format or Status.\n"
     "Tab 'DM Scripts' = the BLUEPRINT flows, the pyramid-scheme objection, and the follow-up "
     "sequence.\n"
     "Tab 'Hooks Bank' = 24 drop-in first lines.\n"
     "Tab 'Claim Rules' = banned phrases and what to say instead. Read this before you post "
     "anything you wrote yourself.\n"
     "Tab 'Proof Checklist' = the screenshots and documents to collect once, and reuse for months.\n"
     "Tab 'Rewrites' = your three original pieces, rebuilt to fit this system."),

    ("NON-NEGOTIABLES",
     "1. No dollar/rand earnings figures, no 'passive', no 'residual income', no 'financial "
     "freedom' anywhere in public content.\n"
     "2. No 'comment a keyword' CTAs — Meta treats that as engagement bait and it exposes your "
     "leads publicly. Use 'DM me <WORD>' or a link.\n"
     "3. Every savings claim shows the screenshot: property, dates, room type, public rate, your "
     "rate, total saved. 'Up to X%' only appears next to a real example.\n"
     "4. Business-side content states plainly that commissions are paid on real travel bookings, "
     "and that results vary.\n"
     "5. Get written permission before using any customer name, photo or voice note."),

    ("SA LEGAL CONTEXT (why the wording matters)",
     "Section 43(2) of the Consumer Protection Act 68 of 2008 makes it an offence to directly OR "
     "INDIRECTLY promote, join or participate in a pyramid scheme, and s43(4) tests whether "
     "participants are compensated primarily from recruitment rather than from the sale of goods "
     "or services. Contravention can be pursued as criminal fraud, and the National Consumer "
     "Commission has investigated social-media-driven schemes — with participants, not just "
     "promoters, caught up in them. The practical takeaway for your copy: describe how the money "
     "is actually earned (real travel purchases), never describe how much people earn, and keep "
     "the recruitment message out of paid/public promotion. Confirm your specific company's "
     "structure and its own compliance policy with them in writing — this is marketing guidance, "
     "not legal advice."),

    ("POSTING CADENCE",
     "Reels: 3–4 per week. Carousel: 1 per week. Long-form Facebook post: every Sunday 19:30 SAST "
     "(plus two midweek posts on Thursdays). Stories: daily, 2–4 frames. Lives: Days 11, 13 and 27. "
     "Day 22 is the one full rest day; the other lighter days (3, 5, 12, 19) are 15–25 minute "
     "formats, which is what makes 30 consecutive days sustainable. Do not upgrade them into full "
     "productions — that is how this system dies in week two.\n"
     "Best SA windows: 06:00–07:30 (commute scroll), 12:00–13:30 (lunch), 19:00–21:30 (evening). "
     "Sunday evening is the strongest slot for long-form."),

    ("MEASUREMENT",
     "Ignore likes. Track four numbers weekly: DM conversations started, Lane A bookings "
     "requested, Lane B overview videos watched, and saves/shares per post. Saves and shares on "
     "the savings content are the leading indicator of a working hook."),
]

r = 5
for head, text in rows:
    ws.cell(row=r, column=2, value=head).font = Font(name="Calibri", size=11, bold=True, color=TEAL)
    ws.cell(row=r, column=2).alignment = WRAP
    c = ws.cell(row=r, column=3, value=text)
    c.font = BODY
    c.alignment = WRAP
    lines = max(text.count("\n") + 1, len(text) // 95 + 1)
    ws.row_dimensions[r].height = max(30, lines * 15.5)
    r += 1
    ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=LIGHT)
    r += 1

ws.freeze_panes = "A5"

# ================================================================ 2. CALENDAR
CAL_HEADERS = ["Day", "Date", "Week / Theme", "Format", "Hook (first line on screen or first sentence)",
               "Content Beats", "CTA (DM-first)", "Proof Asset Needed", "Risk / Compliance Note",
               "Est. Time to Make", "Status"]

START = date(2026, 10, 19)  # Monday

WEEK_THEME = {
    1: "Week 1 — Credibility: earn attention with travel savings. No business pitch.",
    2: "Week 2 — Curiosity: introduce the partner lane. Invitation, never recruitment pressure.",
    3: "Week 3 — Trust: seasonal urgency + objections answered honestly, in DMs.",
    4: "Week 4 — Conversion: two clear lanes, two clear next steps.",
    5: "Week 4 (cont.) — Conversion: two clear lanes, two clear next steps.",
}

NEW_CAL = [
    # ---- WEEK 1 — Credibility (Mon 19 Oct → Sun 25 Oct) ----
    (1, "Reel (30s) + 3 Story frames",
     "\"I found the same room on two websites. The price difference paid for our flights.\"",
     "Frame 1: split screen, same property, same dates, two rates. Frame 2: why the gap exists — contracted "
     "wholesale inventory vs public retail rate. Frame 3: what you do with the difference (flights, an extra "
     "night, the kids' activities). Close on the arithmetic, not on a promise.",
     "\"DM me the word RATES and I'll price your next trip against what you'd pay publicly.\"",
     "Screenshot pair: one property, identical dates/room, public rate vs your member rate.",
     "Green — claim is fully substantiated by the screenshot. Blur any member ID numbers.",
     "45 min"),

    (2, "Carousel (6 slides)",
     "\"How wholesale travel pricing actually works (no jargon, no hype)\"",
     "S1 cover. S2 the retail markup chain: hotel → wholesaler → booking site → you. S3 what a membership buys: "
     "access to contracted rates. S4 what it does NOT buy: guarantees, free travel, or magic. S5 a real worked "
     "example with numbers. S6 CTA.",
     "\"Save this for your next booking. DM me RATES for your own comparison.\"",
     "One worked-example slide using real numbers from a screenshot you hold.",
     "Green. Slide 4 is deliberately the trust-builder — keep it, do not soften it.",
     "60 min"),

    (3, "Story poll + Reel (20s)",
     "\"Be honest: when did you last check what a room actually cost before you booked it?\"",
     "Story poll (I always check / I book and hope). Reel: three questions to ask before confirming any booking — "
     "total cost not nightly rate, cancellation window, and whether the rate is contracted or dynamic.",
     "\"Reply to the poll and I'll send you my pre-booking checklist.\"",
     "None — educational only.",
     "Green. Polls are fine here: nothing is gated behind them, so this isn't engagement bait.",
     "20 min"),

    (4, "Reel (45s) — destination spotlight",
     "\"Mauritius in December for less than you'd pay in November. Here's the real number.\"",
     "Pick ONE destination. Show the property, the dates, the rate comparison, then the total trip maths for a "
     "family of four. Be specific about what is and isn't included.",
     "\"DM me MAURITIUS and I'll check live availability for your dates.\"",
     "Screenshot: Mauritius property, December dates, both rates, room type.",
     "Green if every figure on screen matches the screenshot. Never round the saving up.",
     "45 min"),

    (5, "Behind-the-scenes Story set (4 frames)",
     "\"What a booking actually looks like on my screen\"",
     "Frame 1: the search. Frame 2: the rate. Frame 3: the confirmation. Frame 4: the customer's message. Real, "
     "unpolished, timestamps visible.",
     "\"Want me to run your dates through this? DM me RATES.\"",
     "Screen recording or screenshots of a genuine booking (with permission).",
     "Green. Blur names, ID numbers, card details and booking references.",
     "15 min"),

    (6, "Customer story Reel (40s)",
     "\"Nadia saved R6,800 on a week in Zanzibar. She didn't do anything except book through me.\"",
     "Interview-style or voice-over: what she wanted, what she'd budgeted, what she paid, what she did with the "
     "difference. Let the customer say the saving — not you.",
     "\"Want the same check run on your next trip? DM me RATES.\"",
     "Written permission from the customer + the booking screenshot.",
     "Green with permission. First name only unless they've agreed otherwise. The saving figure must match the "
     "screenshot exactly.",
     "60 min"),

    (7, "Facebook long-form post — SUNDAY 19:30 SAST",
     "\"We almost booked December the expensive way. Here's the R difference.\"",
     "Tell the real story: what you wanted, what the public price was, what you found, what you did with the "
     "saving. Then a short, plain explanation of why the gap exists. End with a question, not a pitch. No "
     "business mention at all.",
     "\"Tell me where you're taking the family in December — I'll price three options for you.\"",
     "Real booking confirmation or screenshot, with dates and property visible.",
     "Green. Keep the whole post customer-facing. The moment an earnings line appears it stops being "
     "substantiated.",
     "30 min"),

    # ---- WEEK 2 — Curiosity (Mon 26 Oct → Sun 01 Nov) ----
    (8, "Reel (35s) — partner lane, soft",
     "\"Three things people ask me after I show them what a family saves in a year.\"",
     "Answer honestly: (1) does it cost to join, (2) is this a job or a side thing, (3) can I just use it myself "
     "and never invite anyone. Answer 3 first and answer it warmly — it is your strongest trust signal.",
     "\"DM me BLUEPRINT for the 20-minute overview. No pressure, and no follow-up sequence unless you ask.\"",
     "None — this is a values post, not a claims post.",
     "Amber — business lane. Say nothing about what people earn. If asked in the comments, reply \"I'll DM you "
     "the details\" and take it private.",
     "30 min"),

    (9, "Carousel (7 slides)",
     "\"What this role is. What it is not.\"",
     "S1 cover. S2 What it is: you introduce travellers to contracted rates and the company pays you a "
     "commission on bookings that actually happen. S3 What it is not: no earnings guarantee, no getting paid to "
     "recruit, results depend entirely on your own customers. S4 What it costs. S5 Who it suits / who it does "
     "not. S6 The disclosure line. S7 CTA.",
     "\"DM me BLUEPRINT if you want the full breakdown.\"",
     "Your company's current fee schedule and its published partner/compliance terms.",
     "Amber — the single most important compliance asset you will make this month. Slides 3, 5 and 6 are the ones "
     "that keep you safe. Do not cut them for brevity.",
     "90 min"),

    (10, "Reel (25s) — myth bust",
     "\"'You have to recruit your whole family.' No. Here's what actually happens.\"",
     "Show the customer-only path: someone joins purely to book their own trips and never invites a soul. Confirm "
     "this is allowed and normal in your company's terms before you film it.",
     "\"DM me RATES if you only want the savings.\"",
     "Company terms confirming there is no recruitment obligation.",
     "Amber — only post this if it is factually true for your company. If there IS a recruitment or volume "
     "requirement, do not run this reel at all.",
     "25 min"),

    (11, "LIVE — 30 min, 20:00 SAST (Thursday)",
     "\"Watch the 20-minute overview with me — I'll pause it and answer everything\"",
     "Screen-share the company overview video, pause at the four points people always ask about, answer live. "
     "Announce in Stories on Days 9–11.",
     "\"DM me BLUEPRINT for the link if you missed this.\"",
     "The overview video plus your own notes on the four pause points.",
     "Amber. You are responsible for what you say over the top of company material — pause it rather than "
     "repeating any claim you cannot substantiate yourself.",
     "45 min prep"),

    (12, "Voice-note / testimonial Reel (30s)",
     "\"This is what a customer said to me at 6am this morning.\"",
     "Play the voice note over a photo from the trip. No commentary from you, no caption selling anything. Just "
     "the customer.",
     "\"DM me RATES.\"",
     "Permission to share the voice note + the booking it refers to.",
     "Green with permission. Voice notes outperform polished testimonials because they can't be faked in the "
     "viewer's mind.",
     "15 min"),

    (13, "LIVE Q&A — 30 min, 11:00 SAST (Saturday)",
     "\"Ask me anything about wholesale travel — and about the partner side if you want it\"",
     "Structure: 10 min on how rates work with a live search on screen, 10 min open Q&A, 5 min on the partner "
     "side (honest, no numbers), 5 min wrap. Saturday morning catches people planning trips at home.",
     "\"If you want the overview after this, DM me BLUEPRINT and I'll send it.\"",
     "One live search ready to run, plus your fee schedule open in a browser tab.",
     "Amber. On a live you cannot edit — write your three answers about earnings in advance and practise them. "
     "The correct answer is always: it depends on your own customers, and I won't quote you a number.",
     "45 min prep"),

    (14, "Facebook long-form post — SUNDAY 19:30 SAST",
     "\"I used to think travel was a cost centre. I was wrong about why.\"",
     "Make the product-utility argument from your original Slide 4 — it was your best idea. People already buy "
     "travel at full price, already love it, already plan it years ahead. That is why retention is real. Do NOT "
     "mention income figures or recruitment. Also cut three 20-second clips from Day 13's live for Week 3 "
     "fillers.",
     "\"If you'd like to see how the partner side works, DM me BLUEPRINT.\"",
     "None required — reasoning only.",
     "Green. This is the compliant version of your residual-income argument: it sells product utility, not "
     "earnings.",
     "30 min"),

    # ---- WEEK 3 — Trust + December urgency (Mon 02 Nov → Sun 08 Nov) ----
    (15, "Reel (35s) — seasonal urgency",
     "\"If you haven't booked December yet, these are the dates still worth chasing.\"",
     "Three windows with real availability. SA school holidays run roughly 5 Dec to 13 Jan, so peak pricing is "
     "brutal — show the shoulder dates either side that still work for families with flexible leave.",
     "\"DM me DECEMBER with your dates and I'll come back to you within 24 hours.\"",
     "Live availability screenshots for the three date windows.",
     "Green. Availability moves fast — date-stamp the reel and take it down once the dates pass.",
     "40 min"),

    (16, "Carousel (5 slides)",
     "\"Book December now, or pay the January price for the same week.\"",
     "S1 cover. S2 how airline and hotel pricing moves as rooms fill. S3 a real example of the same property "
     "priced in October vs what it will be in December. S4 what to lock first (flights, then beds). S5 CTA.",
     "\"DM me DECEMBER and I'll price your week today.\"",
     "Two dated screenshots of the same property, captured weeks apart.",
     "Green. Keep the capture date visible in the screenshots — that is what makes the claim verifiable rather "
     "than theoretical.",
     "50 min"),

    (17, "Reel (60s) — the objection, answered properly",
     "\"The question I get asked every single week. Here's my actual answer.\"",
     "Don't get defensive. Three parts: (1) the company is paid when a real traveller books a real trip — show "
     "the booking; (2) nothing about your income requires anyone to recruit anyone; (3) here are the registered "
     "company details, check them yourself. Then the honest bit: some companies in this space ARE schemes, and "
     "here's the one question that tells them apart.",
     "\"If you want to see the structure yourself, DM me BLUEPRINT. Judge it on the facts.\"",
     "Company registration details, the compensation policy document, and one real booking.",
     "Amber — handle carefully. Never put 'Ponzi' or 'pyramid scheme' in the cover text or caption; say it once, "
     "on camera, then move on. Letting the algorithm index your post against scam vocabulary is how reach dies.",
     "75 min"),

    (18, "Facebook long-form post (midweek) — THU 20:00 SAST",
     "\"One family. Four trips a year. Here's the exact maths.\"",
     "A worked annual case study: a March long weekend in the Drakensberg, Easter in Cape Town, the June school "
     "holidays in Mozambique, December in Mauritius. Show each saving, then the year total. Compare it honestly "
     "against the membership cost and let the reader do the sums.",
     "\"DM me RATES and I'll build the same four-trip plan around your family's dates.\"",
     "Four real booking comparisons, or one real one plus clearly-labelled estimates.",
     "Amber. If any of the four trips is an estimate rather than a real booking, label it 'indicative' on screen. "
     "Mixing one real figure with three invented ones is how a truthful post becomes a misleading one.",
     "60 min"),

    (19, "Reel (25s) — cost transparency",
     "\"What the membership costs. Out loud. On camera.\"",
     "Say the number. Then show one booking that covers it. Transparency here is your competitive advantage — "
     "most people in this industry hide the fee.",
     "\"DM me BLUEPRINT for the fee schedule and the overview video.\"",
     "Current fee schedule + one booking comparison.",
     "Green. Naming the price filters out bad-fit DMs and dramatically improves the quality of the conversations "
     "you do have.",
     "25 min"),

    (20, "How-it-works Reel (45s)",
     "\"Watch me book a real holiday in 90 seconds.\"",
     "Screen recording, sped up: real property, real dates, real confirmation at the end. This is the most "
     "convincing asset in the whole system because it is unarguable.",
     "\"Your turn. DM me RATES with your dates.\"",
     "Full screen recording of one booking (permission + blurring).",
     "Green. Blur every personal and payment field. Keep the final confirmation visible.",
     "45 min"),

    (21, "Facebook long-form post — SUNDAY 19:30 SAST",
     "\"Who this is a bad fit for\"",
     "Long-form version of the disqualification carousel: you want money next week; you won't talk to anyone "
     "about it; you're looking for something that runs without you; you only want it for yourself (and that's "
     "genuinely fine — join as a customer instead). Then who it IS for.",
     "\"DM me BLUEPRINT if you're still reading.\"",
     "None.",
     "Green — likely your best-converting partner asset. Disqualifying openly signals honesty and pre-filters for "
     "people who actually convert.",
     "45 min"),

    # ---- WEEK 4 — Conversion (Mon 09 Nov → Tue 17 Nov) ----
    (22, "Rest / engage + admin",
     "No post. Clear the DM backlog and log every conversation.",
     "Name, what they asked for, lane (A or B), next action date. This sheet is worth more than any post you'll "
     "make this month. Also DM everyone who commented BLUEPRINT during the lives within 24 hours — speed is the "
     "whole game.",
     "n/a", "n/a",
     "Green. Rest days protect consistency more than daily posting does.",
     "45 min"),

    (23, "Reel (30s) — real scarcity",
     "\"I'm doing ten December travel plans this week. That's the honest limit.\"",
     "Explain why it's limited (each plan takes real research) rather than manufacturing urgency. Then show what "
     "the person gets: three options, real rates, booked or not.",
     "\"DM me DECEMBER. First ten, then I close it until January.\"",
     "None.",
     "Green as long as the limit is real. Fake scarcity in an industry already suspected of manipulation is the "
     "fastest way to confirm someone's suspicion.",
     "25 min"),

    (24, "Carousel (6 slides)",
     "\"Where we're going in 2027 — and what it costs to get there\"",
     "S1 cover. S2–S4 three destinations with real rate comparisons for early-2027 dates. S5 the booking window "
     "that gets the best price. S6 CTA.",
     "\"DM me 2027 and I'll price it.\"",
     "Three destination rate comparisons.",
     "Green. Early-bird content in October and November is the highest-intent travel content of the year — "
     "people are planning and nobody has quoted them yet.",
     "50 min"),

    (25, "Facebook long-form post — partner lane, THU 20:00 SAST",
     "\"I'm not going to tell you what you can earn. Here's what I will tell you.\"",
     "Turn your compliance constraint into your differentiator. Everyone else in this space quotes you a number; "
     "you won't, because you don't know their work ethic or their network. What you CAN tell them: how the "
     "commission works, what it costs, and what your own first six months looked like — as your story, labelled "
     "as yours.",
     "\"DM me BLUEPRINT. Ask me the hard questions; I'd rather you did that now.\"",
     "Your own real figures, clearly framed as one individual's experience.",
     "Amber. If you quote your own numbers you MUST label them as individual and atypical. The safest version "
     "tells the story with no figures at all and lets them ask in the DM.",
     "45 min"),

    (26, "Reel (40s) — proof stack",
     "\"Three trips I priced this week. Same properties, same dates, two prices each.\"",
     "Rapid-fire screenshot comparison: three real properties, the public rate and the member rate, the saving "
     "tallied at the end. No talking-head framing needed — voice-over and screen.",
     "\"DM me RATES and I'll be the fourth.\"",
     "Three fresh rate-comparison screenshots from this week.",
     "Green. Fresh screenshots with visible capture dates are the strongest proof format you have.",
     "35 min"),

    (27, "LIVE — 30 min, 11:00 SAST (Saturday)",
     "\"Book your December trip live with me\"",
     "Take three viewers' dates on the live, search them on screen, quote the member rate against the public "
     "rate in real time. Unscripted proof is the most persuasive thing in this entire system.",
     "\"DM me DECEMBER and I'll do yours next.\"",
     "Portal access, a stable connection, and your fee schedule open in a tab.",
     "Amber. Get each participant's on-camera or on-chat permission before pricing their trip, and blur personal "
     "details on any clip you repurpose.",
     "45 min prep"),

    (28, "Carousel (5 slides) — month recap",
     "\"This month in numbers: 14 families, R91,400 saved\"",
     "S1 cover with the two totals. S2–S4 three of the trips, each with its real saving and a screenshot. S5 "
     "what's open next month and how to get in touch. Sunday evening is when SA families scroll and plan, so "
     "the recap doubles as your December catch-up net.",
     "\"DM me RATES and you're next month's number.\"",
     "Your own booking log for the month + three booking screenshots.",
     "Green — only if the totals are literally true and countable. If your month produced four bookings, post "
     "four bookings. Small and real beats big and vague every single time.",
     "40 min"),

    (29, "Personal reflection Reel (45s)",
     "\"A year ago I booked my own holiday the expensive way.\"",
     "Straight to camera, no screenshots, no pitch. Tell the story of what you learned. End with the two doors: "
     "save on your travel, or build with travel — and that either one is a completely fine answer.",
     "\"DM me RATES or DM me BLUEPRINT. Whichever one is actually you.\"",
     "None.",
     "Green. This is the post people share, and it humanises everything else you published this month.",
     "30 min"),

    (30, "Facebook long-form + Story set — month reset",
     "\"November is open. Here's exactly what I can do for you.\"",
     "Two clear lanes, two clear next steps, one honest line about your capacity. Story set: behind the scenes of "
     "planning the month, a poll on where people are going, and the DM prompt. Include your month's real recap "
     "numbers if you have them.",
     "\"DM me RATES (holiday) or BLUEPRINT (the overview). I read every one.\"",
     "Your own booking log for the month.",
     "Green. Capacity honesty converts better than unlimited availability, because it implies real demand. Only "
     "quote recap totals you can literally count.",
     "25 min"),
]

CALENDAR = NEW_CAL

ws = wb.create_sheet("30-Day Calendar")
ws.sheet_view.showGridLines = False
set_widths(ws, [6, 13, 40, 24, 52, 74, 42, 40, 56, 13, 13])

ws["A1"] = "30-Day Content Calendar — Wholesale Travel (South Africa)"
ws["A1"].font = TITLE
ws.merge_cells("A1:K1")
ws["A2"] = ("Savings-proof model: every claim is backed by a screenshot you hold. Business-lane days are flagged "
            "Amber — read the risk note before posting. Start date is editable: change B5 and the column follows.")
ws["A2"].font = SUB
ws.merge_cells("A2:K2")
ws.row_dimensions[2].height = 30
ws["A2"].alignment = WRAP

hdr = 4
for i, h in enumerate(CAL_HEADERS, start=1):
    ws.cell(row=hdr, column=i, value=h)
style_header(ws, hdr, len(CAL_HEADERS))

r = hdr + 1
for (day, fmt, hook, beats, cta, proof, risk, ttime) in CALENDAR:
    week = (day - 1) // 7 + 1
    tint = WEEK_TINT[week]
    d = START + timedelta(days=day - 1)
    vals = [day, d.strftime("%a %d %b"), WEEK_THEME[week], fmt, hook, beats, cta, proof, risk, ttime, "Planned"]
    for i, v in enumerate(vals, start=1):
        ws.cell(row=r, column=i, value=v)
    body_row(ws, r, len(CAL_HEADERS), tint)
    ws.cell(row=r, column=1).font = BODY_B
    ws.cell(row=r, column=1).alignment = WRAP_C
    ws.cell(row=r, column=2).alignment = WRAP_C
    ws.cell(row=r, column=5).font = BODY_B
    ws.cell(row=r, column=10).alignment = WRAP_C
    ws.cell(row=r, column=11).alignment = WRAP_C
    # colour-code the risk cell
    rcell = ws.cell(row=r, column=9)
    if risk.startswith("Green"):
        rcell.font = Font(name="Calibri", size=10, color=GREEN)
    elif risk.startswith("Amber"):
        rcell.font = Font(name="Calibri", size=10, bold=True, color=AMBER)
    else:
        rcell.font = Font(name="Calibri", size=10, bold=True, color=RED)
    longest = max(len(beats) / 70, len(risk) / 52, len(hook) / 48, len(cta) / 40, len(proof) / 38)
    ws.row_dimensions[r].height = max(60, longest * 14.5)
    r += 1

last = r - 1
dv = DataValidation(type="list", formula1='"Planned,Filming,Editing,Scheduled,Posted,Skipped"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"K{hdr+1}:K{last}")
ws.auto_filter.ref = f"A{hdr}:K{last}"
ws.freeze_panes = f"E{hdr+1}"

# ================================================================ 3. DM SCRIPTS
ws = wb.create_sheet("DM Scripts")
ws.sheet_view.showGridLines = False
set_widths(ws, [3, 30, 112])

ws["B2"] = "DM Scripts — the BLUEPRINT funnel and everything around it"
ws["B2"].font = TITLE
ws.merge_cells("B2:C2")
ws["B3"] = ("Send within 15 minutes where you can. Reply to their Story or post first so you land in the main inbox "
            "rather than requests. Never paste all four messages at once — one message, then wait.")
ws["B3"].font = SUB
ws["B3"].alignment = WRAP
ws.merge_cells("B3:C3")
ws.row_dimensions[3].height = 32

DM_BLOCKS = [
    ("LANE A — 'RATES' (travel savings)", [
        ("Step 1 — Acknowledge + ask one question",
         "Hey {Name} 👋 Thanks for the message — happy to run your dates.\n\n"
         "Quick one so I quote the right thing: where are you thinking of going, roughly when, and "
         "how many people (kids' ages matter for the room setup)?"),
        ("Step 2 — Deliver value before asking for anything",
         "Got it. Give me a few hours and I'll come back with three options at member rates, plus "
         "what the same rooms are showing publicly today, so you can see the gap yourself.\n\n"
         "I'll send screenshots, not a summary — you should be able to check every number."),
        ("Step 3 — The comparison message",
         "Here's what I found for {destination}, {dates}, {n} people:\n\n"
         "1. {Property A} — public rate R{X}, member rate R{Y}. Saving R{Z}.\n"
         "2. {Property B} — public rate R{X}, member rate R{Y}. Saving R{Z}.\n"
         "3. {Property C} — public rate R{X}, member rate R{Y}. Saving R{Z}.\n\n"
         "Screenshots attached so you can verify each one. Rates move with availability, so these "
         "hold for about 48 hours.\n\n"
         "No obligation at all — if you'd rather book it yourself elsewhere, you've at least got "
         "the benchmark."),
        ("Step 4 — Membership question, answered plainly",
         "Fair question. It's R{fee} {monthly/annually} to have the access, and there's no "
         "obligation to book anything or invite anyone.\n\n"
         "On the trip above, the saving alone is R{Z} — so this one booking covers it {X} times "
         "over. If you travel once a year like this, it pays for itself. If you don't travel, it "
         "won't, and I'd tell you to skip it."),
        ("Step 5 — Close, or park it",
         "Want me to hold {option}? I'd need {details} to confirm.\n\n"
         "Or if the dates aren't firm yet, no stress — tell me when you're ready and I'll re-price "
         "it. I'm not going to chase you about it."),
    ]),
    ("LANE B — 'BLUEPRINT' (business-side)", [
        ("Step 1 — Qualify honestly",
         "Hey {Name} 👋 Here's the overview video: {link} — it's 20 minutes and it explains the "
         "structure properly, better than I can in a DM.\n\n"
         "Before you watch it, one honest question so I don't waste your time: are you mainly "
         "after cheaper travel for yourself, or are you genuinely looking at building something on "
         "the side? Both are good answers — they just need different info from me."),
        ("Step 2 — Set expectations, including the negative",
         "Worth saying up front, because most people in this industry won't:\n\n"
         "• It costs R{fee} to join and there are ongoing costs if you choose to run it as a "
         "business.\n"
         "• Commissions are paid when real travellers book real trips. Nothing is paid for signing "
         "people up.\n"
         "• I can't quote you an earnings figure, and you should be suspicious of anyone who does. "
         "It depends entirely on how many customers you build.\n"
         "• Most people who join just use it for their own travel. That's completely fine."),
        ("Step 3 — Invite the objection",
         "Watch it and then ask me the awkward questions. \"How do you actually get paid?\" \"What "
         "happens if I stop?\" \"Is this legal in SA?\" — I'd rather you asked me now than sat on "
         "it wondering.\n\n"
         "I'll give you straight answers, including the ones that might talk you out of it."),
        ("Step 4 — Follow-up (only if they engaged)",
         "Hey {Name}, no pressure at all — just checking whether the video answered what you were "
         "after?\n\n"
         "If it's a not-now, tell me and I'll leave it with you completely. If there's one bit "
         "that didn't make sense, I'm happy to walk through it on a quick voice note."),
        ("Step 5 — The 'no' response (send this one graciously, always)",
         "Totally fine — thanks for being straight with me. 👍\n\n"
         "The travel side still works for you as a customer if you ever want a holiday priced, so "
         "my DMs are open either way. Enjoy the rest of your week."),
    ]),
    ("OBJECTION — 'Is this a pyramid / MLM scheme?'", [
        ("The answer (voice note is better than text here)",
         "Honestly? It's the right question to ask, and I'd ask it too.\n\n"
         "Here's how I'd answer it if I were you: a scheme pays you from the money new people put "
         "in, and it stops working when recruitment stops. This one pays a commission when a "
         "traveller books a real holiday — there's a booking confirmation behind every rand. If "
         "nobody ever joined again, the company would still have customers going on holiday.\n\n"
         "Two things you should check for yourself rather than take my word for: the company's "
         "registration details, and whether your compensation depends on people you personally "
         "sign up. Mine doesn't. Ask me for the documents and I'll send them.\n\n"
         "And for what it's worth — there are genuine schemes in this space. If someone quotes you "
         "an income figure in the first message, run."),
    ]),
    ("OBJECTION — 'How much will I earn?'", [
        ("The only safe and honest answer",
         "I'm not going to give you a number, and I'd walk away from anyone who does.\n\n"
         "Here's why: your earnings come from commissions on trips your own customers book. That's "
         "a function of how many travellers you serve and how often they travel — nothing I can "
         "predict for you, and nothing the company can promise you either.\n\n"
         "What I can tell you is what it costs to start, exactly how the commission is calculated, "
         "and what my own first few months looked like as one person's experience. Want those?"),
    ]),
    ("OBJECTION — 'I can find cheaper on Booking.com'", [
        ("Benchmark it, don't argue",
         "You might! Send me the property and dates and I'll price it against my rate — if you find "
         "it cheaper publicly, book it there and don't give me a cent.\n\n"
         "Where the gap usually shows up is family rooms and multi-night stays, and on packages "
         "where the flight and bed are contracted together. But check it, don't take my word."),
    ]),
    ("RE-ENGAGEMENT — 21 days after a quiet lead", [
        ("Send this once, then stop",
         "Hey {Name} — no pitch, this is genuinely useful: {destination} for {month} has just "
         "dropped to R{rate} for {dates}, which is about R{saving} under what it was showing last "
         "month. Screenshot attached.\n\n"
         "If that trip isn't for you, ignore this and I won't message you again about it. If it "
         "is, tell me how many people and I'll check availability."),
    ]),
]

r = 5
for block_title, items in DM_BLOCKS:
    ws.cell(row=r, column=2, value=block_title).font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=TEAL)
    ws.cell(row=r, column=2).alignment = WRAP
    ws.cell(row=r, column=3).fill = PatternFill("solid", fgColor=TEAL)
    ws.row_dimensions[r].height = 24
    r += 1
    for label, script in items:
        ws.cell(row=r, column=2, value=label).font = BODY_B
        ws.cell(row=r, column=2).alignment = WRAP
        ws.cell(row=r, column=2).border = BORDER
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=LIGHT)
        c = ws.cell(row=r, column=3, value=script)
        c.font = BODY
        c.alignment = WRAP
        c.border = BORDER
        ws.row_dimensions[r].height = max(38, (script.count("\n") + 1) * 16 + len(script) // 105 * 15)
        r += 1
    r += 1

ws.freeze_panes = "A5"

# ================================================================ 4. HOOKS BANK
ws = wb.create_sheet("Hooks Bank")
ws.sheet_view.showGridLines = False
set_widths(ws, [5, 20, 78, 22, 46])

ws["A1"] = "Hooks Bank — 24 first lines that survive the scroll"
ws["A1"].font = TITLE
ws.merge_cells("A1:E1")
ws["A2"] = ("Rule: the hook must be true even if the viewer never buys anything. Number hooks only work when the "
            "number is real — screenshot it or don't say it.")
ws["A2"].font = SUB
ws.merge_cells("A2:E2")

HOOKS = [
    ("Savings / proof", "\"The same room. Two websites. R7,400 apart.\"", "Reel", "Highest save rate in this system"),
    ("Savings / proof", "\"I found out what a hotel actually charges vs what you pay. It's not a small gap.\"", "Reel", "Curiosity without a claim"),
    ("Savings / proof", "\"We booked four nights in Zanzibar for less than the flights.\"", "Reel", "Only if the screenshot says so"),
    ("Savings / proof", "\"This is what R12,000 gets a family of four in December. In Mauritius.\"", "Carousel", "Specific beats impressive"),
    ("Savings / proof", "\"Screenshot for screenshot: my rate vs the public one. Judge it yourself.\"", "Reel", "Radical transparency hook"),
    ("Savings / proof", "\"Fourteen families booked this month. Total saved: R91,400.\"", "Reel", "Use your real number"),
    ("Education", "\"You're being charged a markup you've never been told about. Here's the chain.\"", "Carousel", "Explains wholesale simply"),
    ("Education", "\"Three questions to ask before you confirm any hotel booking.\"", "Reel", "Pure value, zero pitch"),
    ("Education", "\"Why the same room costs more in December than November — and the dates that dodge it.\"", "Carousel", "Seasonal, high intent"),
    ("Education", "\"Dynamic pricing, explained in 30 seconds, with a real example.\"", "Reel", "Screenshare format"),
    ("Education", "\"What 'contracted rates' means and why it matters to your wallet.\"", "Reel", "Jargon-buster"),
    ("Objection / trust", "\"The question I get asked every single week. I'm going to answer it properly.\"", "Reel", "Safer cover than 'pyramid scheme'"),
    ("Objection / trust", "\"I'm not going to tell you what you can earn. Here's what I will tell you.\"", "Reel", "Turns compliance into a differentiator"),
    ("Objection / trust", "\"Who this is a terrible fit for.\"", "Carousel", "Disqualification converts"),
    ("Objection / trust", "\"What the membership costs. Out loud. On camera.\"", "Reel", "Naming the fee filters quality"),
    ("Objection / trust", "\"Some companies in this space really are schemes. Here's how to tell them apart.\"", "Reel", "Positions you as the honest one"),
    ("Story / human", "\"A year ago I booked my own holiday the expensive way.\"", "Reel", "Straight to camera, no cuts"),
    ("Story / human", "\"Nadia messaged me at 6am. Listen to this.\"", "Reel", "Voice-note social proof"),
    ("Story / human", "\"I used to think travel was a cost centre. I was wrong about why.\"", "Facebook post", "Product-utility argument"),
    ("Story / human", "\"My favourite part of this isn't the money. It's the photos people send me.\"", "Story set", "Warmth, no claim"),
    ("Seasonal / SA", "\"If you haven't booked December, these dates are still worth chasing.\"", "Reel", "Time-sensitive, date-stamp it"),
    ("Seasonal / SA", "\"June school holidays in Mozambique: what it actually costs a family of four.\"", "Carousel", "Book 4–6 months ahead"),
    ("Seasonal / SA", "\"Easter long weekend. Three drives from Joburg. Real prices.\"", "Reel", "Drakensberg / Magaliesburg / Vaal"),
    ("Seasonal / SA", "\"2027 is cheaper to book right now than it will be in March.\"", "Carousel", "Early-bird, high intent"),
]

hdr = 4
for i, h in enumerate(["#", "Category", "Hook", "Best Format", "Why It Works"], start=1):
    ws.cell(row=hdr, column=i, value=h)
style_header(ws, hdr, 5)
r = hdr + 1
for i, (cat, hook, fmt, why) in enumerate(HOOKS, start=1):
    ws.cell(row=r, column=1, value=i).alignment = WRAP_C
    ws.cell(row=r, column=2, value=cat)
    ws.cell(row=r, column=3, value=hook).font = BODY_B
    ws.cell(row=r, column=4, value=fmt)
    ws.cell(row=r, column=5, value=why)
    body_row(ws, r, 5, LIGHT if i % 2 == 0 else None)
    ws.cell(row=r, column=1).alignment = WRAP_C
    ws.cell(row=r, column=3).font = BODY_B
    ws.row_dimensions[r].height = 32
    r += 1
ws.freeze_panes = f"A{hdr+1}"
ws.auto_filter.ref = f"A{hdr}:E{r-1}"

# ================================================================ 5. CLAIM RULES
ws = wb.create_sheet("Claim Rules")
ws.sheet_view.showGridLines = False
set_widths(ws, [5, 34, 40, 62])

ws["A1"] = "Claim Rules — never say / say this instead"
ws["A1"].font = TITLE
ws.merge_cells("A1:D1")
ws["A2"] = ("You have savings proof but no typical-earnings data. Under that constraint, product and savings claims "
            "are fully available to you; earnings claims are not. Everything below follows from that one fact. "
            "Marketing guidance, not legal advice — confirm against your company's own compliance policy.")
ws["A2"].font = SUB
ws["A2"].alignment = WRAP
ws.merge_cells("A2:D2")
ws.row_dimensions[2].height = 34

RULES = [
    ("NEVER", "\"Passive income\"", "\"You earn a commission when travellers you introduce book trips.\" "
     "— describes the mechanism, promises no outcome."),
    ("NEVER", "\"Residual income\" / \"residual wealth\" / \"compounding income\"",
     "\"Customers who travel every year book through me every year.\" — same idea, stated as an "
     "observable fact about travel behaviour."),
    ("NEVER", "\"Financial freedom\" / \"quit your job\" / \"replace your salary\"",
     "\"A second income stream you build around something people already buy.\" — no magnitude "
     "implied."),
    ("NEVER", "\"Earn R20,000 a month\" or any earnings figure for other people",
     "Nothing. Delete it. If asked, use the 'I'm not going to give you a number' DM script."),
    ("NEVER", "\"Guaranteed\" / \"proven\" / \"no risk\" / \"can't lose\"",
     "\"Here's what it costs, here's what it doesn't guarantee.\" — the risk statement IS the "
     "trust-builder."),
    ("NEVER", "\"Get paid to recruit\" / \"build a team\" / \"downline\" / \"upline\"",
     "\"Serve more travellers.\" Recruitment vocabulary is what regulators and platforms pattern-match on."),
    ("NEVER", "\"Up to 70% off\" floating on its own",
     "\"Up to {X}% off — this Mauritius booking was {X}%. Screenshot attached.\" The number is only "
     "safe when it is anchored to a real, shown example."),
    ("NEVER", "Luxury lifestyle B-roll implying it came from this income",
     "Travel footage of trips that were genuinely taken, with the real saving figure on screen."),
    ("NEVER", "\"Comment BLUEPRINT below\" (gated content behind a comment)",
     "\"DM me BLUEPRINT.\" Meta treats keyword-comment gating as engagement bait and downranks it; "
     "it also publishes your lead list to competitors."),
    ("NEVER", "\"Ponzi\" or \"pyramid scheme\" in cover text or captions",
     "\"The question I get asked every week.\" Say the words once on camera, answer them, move on. "
     "Do not let the algorithm index your post against scam vocabulary."),
    ("ALWAYS", "Show the screenshot for any savings claim",
     "Property, dates, room type, public rate, member rate, total saved — all visible, all matching "
     "what you say."),
    ("ALWAYS", "State the membership cost before they ask",
     "Volunteering the price filters out bad-fit leads and is the fastest trust signal available to you."),
    ("ALWAYS", "Label personal results as personal",
     "\"This is what my first six months looked like\" — never \"this is what you'll get\"."),
    ("ALWAYS", "Include a results-vary line on business-side content",
     "\"Commission is earned on real travel bookings. What you earn depends entirely on the "
     "customers you build.\""),
    ("ALWAYS", "Get permission in writing before using a customer",
     "Name, photo, voice note, booking details — and blur ID numbers, card details and booking "
     "references every time."),
    ("ALWAYS", "Date-stamp seasonal availability content",
     "Rates and availability move. Put the date on the video and take it down when the dates pass."),
    ("ALWAYS", "Keep the business invitation in DMs, not paid ads",
     "Lane A content can be boosted. Lane B content should stay organic and invitation-led."),
    ("ALWAYS", "Keep your company's registration and compensation policy to hand",
     "When someone asks to verify — and the good ones will — sending the document within an hour "
     "wins the conversation."),
]

hdr = 4
for i, h in enumerate(["", "Rule", "Phrase / Practice", "Why, and What To Do Instead"], start=1):
    ws.cell(row=hdr, column=i, value=h)
style_header(ws, hdr, 4)
r = hdr + 1
for kind, phrase, replacement in RULES:
    ws.cell(row=r, column=1, value="🚫" if kind == "NEVER" else "✅")
    ws.cell(row=r, column=2, value=kind)
    ws.cell(row=r, column=3, value=phrase)
    ws.cell(row=r, column=4, value=replacement)
    fill = "FDECEA" if kind == "NEVER" else "EAF5EE"
    body_row(ws, r, 4, fill)
    ws.cell(row=r, column=1).alignment = WRAP_C
    ws.cell(row=r, column=1).font = Font(name="Calibri", size=13)
    ws.cell(row=r, column=2).font = Font(name="Calibri", size=10, bold=True,
                                         color=RED if kind == "NEVER" else GREEN)
    ws.cell(row=r, column=2).alignment = WRAP_C
    ws.cell(row=r, column=3).font = BODY_B
    ws.row_dimensions[r].height = max(40, len(replacement) / 58 * 15 + 22)
    r += 1
ws.freeze_panes = f"A{hdr+1}"

# ================================================================ 6. PROOF CHECKLIST
ws = wb.create_sheet("Proof Checklist")
ws.sheet_view.showGridLines = False
set_widths(ws, [5, 44, 58, 26, 15, 15])

ws["A1"] = "Proof Checklist — collect once, reuse for months"
ws["A1"].font = TITLE
ws.merge_cells("A1:F1")
ws["A2"] = ("Every item here converts a claim you cannot make into a claim you can. Build this folder in week one "
            "and the next eleven months of content write themselves.")
ws["A2"].font = SUB
ws.merge_cells("A2:F2")

PROOF = [
    ("Rate comparison screenshots — 10 minimum",
     "Same property, same dates, same room type: public rate next to member rate. Capture the URL "
     "bar or the app header so the source is visible.", "Days 1, 2, 4, 16, 24, 26", "HIGH"),
    ("One full booking screen recording",
     "Search → rate selection → confirmation. Blur every personal and payment field; keep the final "
     "confirmation legible.", "Days 5, 20", "HIGH"),
    ("Two dated screenshots of the same property weeks apart",
     "This is the only way to substantiate 'prices rise as rooms fill'. Include the capture date "
     "on screen.", "Day 16", "HIGH"),
    ("Four-trip annual case study numbers",
     "March (Drakensberg), Easter (Cape Town), June (Mozambique), December (Mauritius). Real where "
     "you have them, clearly labelled 'indicative' where you don't.", "Day 18", "HIGH"),
    ("Company registration details",
     "Registration number, registered address, and the compensation/commission policy document. "
     "Have these as a single PDF you can send in one message.", "Days 9, 17, 19", "HIGH"),
    ("Current fee schedule",
     "Joining fee, monthly fee, what's included, cancellation terms. Say it out loud on camera.",
     "Days 9, 19", "HIGH"),
    ("Written confirmation: no recruitment obligation",
     "From your company, in writing, that a member may join purely as a customer with no "
     "requirement to recruit. Do not run Day 11 without this.", "Day 11", "CRITICAL"),
    ("Customer permissions",
     "A short written 'yes, you can use my first name, my quote and this photo'. Keep them in one "
     "folder with the post they were used in.", "Days 6, 12, 18", "HIGH"),
    ("Three customer voice notes",
     "Unedited, 15–30 seconds, about the trip rather than about the business.", "Day 12", "MEDIUM"),
    ("Your own monthly booking log",
     "A simple sheet: date, customer, destination, saving. This produces your recap numbers and "
     "stops you ever being tempted to exaggerate.", "Days 22, 28, 30", "HIGH"),
    ("The company's 20-minute overview video",
     "Local link you can send in a DM, plus your notes on the four points to pause and clarify.",
     "Days 11, 27", "MEDIUM"),
    ("Screenshots of what you will NOT use",
     "Any company material that quotes typical earnings, shows luxury cars or makes lifestyle "
     "promises. Note it, then don't repost it.", "Ongoing", "CRITICAL"),
]

hdr = 4
for i, h in enumerate(["", "Asset", "What It Must Contain", "Used On", "Priority", "Done"], start=1):
    ws.cell(row=hdr, column=i, value=h)
style_header(ws, hdr, 6)
r = hdr + 1
for i, (asset, detail, used, prio) in enumerate(PROOF, start=1):
    ws.cell(row=r, column=1, value=i)
    ws.cell(row=r, column=2, value=asset).font = BODY_B
    ws.cell(row=r, column=3, value=detail)
    ws.cell(row=r, column=4, value=used)
    ws.cell(row=r, column=5, value=prio)
    ws.cell(row=r, column=6, value="☐")
    body_row(ws, r, 6, LIGHT if i % 2 == 0 else None)
    ws.cell(row=r, column=1).alignment = WRAP_C
    ws.cell(row=r, column=2).font = BODY_B
    ws.cell(row=r, column=4).alignment = WRAP_C
    pcell = ws.cell(row=r, column=5)
    pcell.alignment = WRAP_C
    pcell.font = Font(name="Calibri", size=10, bold=True,
                      color=RED if prio == "CRITICAL" else (AMBER if prio == "HIGH" else "3F5563"))
    ws.cell(row=r, column=6).alignment = WRAP_C
    ws.cell(row=r, column=6).font = Font(name="Calibri", size=14)
    ws.row_dimensions[r].height = max(44, len(detail) / 55 * 15 + 20)
    r += 1
dvp = DataValidation(type="list", formula1='"☐,✅,N/A"', allow_blank=True)
ws.add_data_validation(dvp)
dvp.add(f"F{hdr+1}:F{r-1}")
ws.freeze_panes = f"A{hdr+1}"

# ================================================================ 7. REWRITES
ws = wb.create_sheet("Rewrites")
ws.sheet_view.showGridLines = False
set_widths(ws, [3, 24, 74, 58])

ws["B2"] = "Your Three Original Pieces — Rewritten"
ws["B2"].font = TITLE
ws.merge_cells("B2:D2")
ws["B3"] = ("Left column = what you wrote. Right column = the version that fits this system. Changes are driven by "
            "three things: the 60-second timing problem, the unsubstantiated earnings claims, and the "
            "comment-keyword CTA.")
ws["B3"].font = SUB
ws["B3"].alignment = WRAP
ws.merge_cells("B3:D3")
ws.row_dimensions[3].height = 32

REWRITES = [
    # (section, LEFT = rewritten/ready, RIGHT = original/diagnosis)
    ("1. 60-SECOND REEL — rewritten script (148 words)",
     "[0:00\u20130:07] HOOK\n\"Same hotel. Same dates. Two different prices \u2014 and the gap paid for our flights.\"\n\n"
     "[0:07\u20130:20] PROOF  [screenshot on screen]\n\"This is a four-night stay in Mauritius in December. Public rate: "
     "R21,900. The rate I booked it at: R13,400. That's R8,500 back into the trip.\"\n\n"
     "[0:20\u20130:33] WHY\n\"Hotels contract rooms in bulk to wholesalers months ahead. Those rates never reach the public "
     "booking sites. I have access to them \u2014 that's the whole trick, and there's no trick beyond it.\"\n\n"
     "[0:33\u20130:47] THE POINT\n\"A job pays you once for an hour. A customer who travels every year books through you "
     "every year. Nobody needs another monthly subscription \u2014 but everybody needs a holiday.\"\n\n"
     "[0:47\u20131:00] CTA\n\"DM me the word RATES and I'll price your next trip against what you'd pay publicly. "
     "Screenshot for screenshot. If you find it cheaper, book it there.\"\n\n"
     "FILMING NOTE: cut on the number. Show the screenshot at 0:12 and hold it two full seconds \u2014 that is the "
     "moment people save the video, and saves are the metric that moves travel content. Swap R21,900 / R13,400 for "
     "your own real figures before you film.",
     "DIAGNOSIS \u2014 your original script is 230 words. At a normal on-camera pace (~150 wpm) that runs roughly 92 "
     "seconds, so the [0:00\u20131:00] timings cannot hold it: it will either be rushed or overrun and lose the "
     "completion-rate signal. Target for a true 60s is 145\u2013150 words; this rewrite is 148.\n\n"
     "The original hook also asks \"Is every online business secretly a pyramid or Ponzi scheme?\" \u2014 which primes "
     "the scam frame in viewers who never had it, and invites comments that read as hostile to the platform.\n\n"
     "YOUR ORIGINAL, VERBATIM:\n[0:00-0:08] \"Is every online business secretly a pyramid or Ponzi scheme? Let's "
     "talk about the legal line between a financial trap and true residual wealth.\"\n[0:08-0:22] \"A Ponzi scheme "
     "takes money from new recruits to pay off old ones because there is NO real product being sold... A legitimate "
     "network marketing model simply cuts out retail markups and pays everyday people to distribute real "
     "products.\"\n[0:22-0:38] \"Linear income pays you once for an hour of work. Residual income pays you "
     "repeatedly for work you did ONCE. That's true financial leverage.\"\n[0:38-0:50] \"For residual income to "
     "last, the product must be something people actually LOVE using. Nobody wants more monthly clutter\u2014but "
     "almost EVERYBODY wants to travel at wholesale prices.\"\n[0:50-1:00] \"When your travel savings far outweigh "
     "your membership cost, it's no longer a pitch\u2014it's a lifestyle. Comment 'BLUEPRINT' below...\"\n\n"
     "KEPT FROM YOUR DRAFT: the linear-vs-repeat contrast and \"nobody wants more monthly clutter.\" Both are "
     "genuinely strong and both survive intact."),

    ("2. CAROUSEL — rewritten (7 slides)",
     "S1 COVER\nHow Wholesale Travel Pricing Actually Works\n(No hype. Real screenshots.)\n\u2192 Save this for your "
     "next booking\n\n"
     "S2 THE CHAIN\nA hotel sells rooms in bulk to a wholesaler at a contracted rate. The wholesaler sells to "
     "booking sites, which add their markup. You pay the last price in that chain.\n\n"
     "S3 THE EXAMPLE\n[Screenshot] Mauritius, 4 nights, December. Public rate R21,900. Member rate R13,400. "
     "Saved: R8,500.\n\n"
     "S4 WHY THE GAP EXISTS\nContracted inventory is bought months ahead and is never released to public sites. "
     "Access, not magic.\n\n"
     "S5 WHAT IT IS NOT\nIt is not free travel. It is not guaranteed savings on every property. And it is not a "
     "way to get rich \u2014 it is a way to pay less for something you were already going to buy.\n\n"
     "S6 WHAT IT COSTS\nR{fee} {monthly/annually}. No obligation to book anything, and no obligation to invite "
     "anyone.\n\n"
     "S7 CTA\nDM me RATES and I'll price your next trip against the public rate \u2014 screenshot for screenshot.\n\n"
     "CAPTION FOOTNOTE: Rates and availability change daily. The example shown is a real booking. Nothing here is "
     "a promise of savings on your specific trip, and nothing here is an earnings claim.\n\n"
     "DESIGN: 1080\u00d71350. Maximum seven words per headline, one idea per slide, save prompt on Slide 1 only. "
     "Disclaimer goes in the caption AND on Slide 7, large enough to read on a phone.",
     "DIAGNOSIS \u2014 Slide 1: '\\rightarrow' leaked into your sub-text as an escape character. \"Read Before You "
     "Judge\" also reads as defensive before anyone has judged anything.\n\nSlide 3: \"Compounding\" breaks "
     "awkwardly on a mobile canvas, and \"Paid Monthly for Years\" is an earnings claim you cannot substantiate.\n\n"
     "Slide 5: \"up to 70% off\" floats with no example attached, plus a comment-gated CTA.\n\nMissing entirely: any "
     "results-vary line, and any identification of the company. The structural addition that matters most is Slide "
     "5 (\"What it is NOT\") \u2014 it is the slide that makes the other six believable, and the one most people in "
     "this industry delete.\n\n"
     "YOUR ORIGINAL, VERBATIM:\nSlide 1 (Cover) \u2014 \"Ponzi Scheme, Scam, or True Residual Wealth? (Read Before You "
     "Judge)\" / sub-text: \"Swipe to see the legal line between financial traps and real leverage ->\"\n"
     "Slide 2 \u2014 \"Ponzi vs. Direct Sales\" \u2022 Ponzi Scheme: No real product. Returns come strictly from "
     "headcounts. Collapses when recruiting stops. \u2022 Direct Sales: Real product/service. Commissions come "
     "strictly from real consumer orders and usage.\n"
     "Slide 3 \u2014 \"Linear Income vs. Compounding Residuals\" \u2022 Linear: 1 Hour Worked = 1 Payment. \u2022 "
     "Residual: Work Once = Paid Monthly for Years.\n"
     "Slide 4 \u2014 \"Why Product Value Changes Everything\": if customers buy and use the product because they love "
     "the value\u2014not just to earn a check\u2014it becomes a legitimate, sustainable business.\n"
     "Slide 5 \u2014 \"Why Wholesale Travel? We replaced physical clutter with wholesale 5-star hotel bookings, "
     "resorts, and annual rewards. Comment 'BLUEPRINT' below to get our private 20-minute overview video!\"\n\n"
     "KEPT FROM YOUR DRAFT: Slide 4's product-utility argument is the best idea in the whole pack. It now carries "
     "the Facebook post on Day 14 and the reel on Day 10."),

    ("3. FACEBOOK POST — rewritten",
     "We almost booked December the expensive way.\n\n"
     "Same resort we've wanted for three years. Public rate: R21,900 for four nights. I ran it through the "
     "wholesale portal I have access to: R13,400. Same room, same dates, same resort.\n\n"
     "R8,500 is not a small number. It's the flights. It's two extra nights. It's the reason we're going at all "
     "instead of talking about going.\n\n"
     "Here's why the gap exists, because it isn't a trick: hotels contract rooms in bulk to wholesalers months in "
     "advance. Those rates are never released to the public booking sites, because releasing them would undercut "
     "the hotel's own retail pricing. Most people simply never see the other side of that wall.\n\n"
     "And that's what made me think about the bigger question.\n\n"
     "Almost everything we're told to buy monthly, we eventually resent paying for. Another subscription. Another "
     "box in the cupboard. Another thing we bought once and forgot.\n\n"
     "Travel is the opposite. Nobody has ever resented a holiday. Families plan them years ahead, save for them, "
     "and then do it again the following year \u2014 which means a customer who travels is a customer who comes back.\n\n"
     "That's the part worth paying attention to. Not that people can save. That they save on something they were "
     "always going to buy anyway.\n\n"
     "So: where are you taking the family in December?\n\n"
     "Tell me in the comments and I'll price three options for you \u2014 screenshot for screenshot, member rate "
     "against public rate. If you find it cheaper elsewhere, book it there.\n\n"
     "And if you'd rather see how the partner side works, DM me BLUEPRINT for a 20-minute overview. I'll answer "
     "the awkward questions too.\n\n"
     "[Rates and availability change daily. The booking above is real and I can show you the confirmation. Nothing "
     "here is a promise about your own trip, and nothing here is an earnings claim \u2014 commission is earned only "
     "when real travellers book real holidays.]\n\n"
     "POST IT: Sunday, 19:30 SAST. First three lines decide whether anyone reaches the argument, so don't bury the "
     "R8,500.",
     "DIAGNOSIS \u2014 the strongest section of everything you wrote is the product-utility argument, and it sat 60% "
     "of the way down behind two traps and a legal defence.\n\n"
     "Issues: dollar signs for a South African audience; \"up to 70% off\" unsubstantiated; \"compounding residual "
     "vehicle for 2026\" is an earnings claim; the Ponzi paragraph spends 90 words defending an accusation most "
     "readers hadn't made; comment-gated CTA.\n\n"
     "Length is fine \u2014 SA Facebook audiences do read long-form on Sunday evenings \u2014 but the rewrite opens on a "
     "story instead of a taxonomy and moves the business invitation into a single closing line that routes to DM. "
     "Note what's missing on purpose: the word 'Ponzi', the word 'passive', and any number about what someone "
     "could earn.\n\n"
     "YOUR ORIGINAL, VERBATIM (opening):\n\"Most people are searching for financial freedom, but they get stuck in "
     "one of two traps: 1. The 40-Year Job Treadmill: Trading hours for dollars where income stops the second you "
     "take a break. 2. Financial Scams / Ponzi Schemes: Promises of 'get-rich-quick' cash with zero real products "
     "or commercial value behind them.\n\nSo what actually separates a legitimate online asset from a scheme? It "
     "comes down to TRUE PRODUCT UTILITY.\n\nA Ponzi scheme relies 100% on collecting membership fees from new "
     "recruits to pay old ones... A real business distribution model pays brand partners strictly when real "
     "consumers buy real, high-value products.\n\nWhat is a product that people ALREADY love and pay full price "
     "for every single day? Answer: TRAVEL.\n\n...a private wholesale booking engine that slashes up to 70% off "
     "public retail hotel and resort rates. When the travel savings alone exceed the cost of membership, customers "
     "stay for years simply because they love vacationing for wholesale. That is how real, compounding residual "
     "income is built.\n\nAre you operating on linear hourly work, or are you building a compounding residual "
     "vehicle for 2026? Drop the word 'BLUEPRINT' in the comments below or send me a DM...\""),
]

r = 5
for title, right, left in REWRITES:
    ws.cell(row=r, column=2, value=title).font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    for c in (2, 3, 4):
        ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=NAVY)
    ws.cell(row=r, column=2).alignment = WRAP
    ws.row_dimensions[r].height = 24
    r += 1
    ws.cell(row=r, column=2, value="").fill = PatternFill("solid", fgColor=LIGHT)
    ws.cell(row=r, column=3, value="REWRITTEN \u2014 READY TO USE").font = Font(name="Calibri", size=10, bold=True, color=RED)
    ws.cell(row=r, column=4, value="YOUR ORIGINAL / DIAGNOSIS").font = Font(name="Calibri", size=10, bold=True, color=GREEN)
    for c in (2, 3, 4):
        ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=r, column=c).border = BORDER
    r += 1
    ws.cell(row=r, column=3, value=right).font = BODY
    ws.cell(row=r, column=3).alignment = WRAP
    ws.cell(row=r, column=3).border = BORDER
    ws.cell(row=r, column=4, value=left).font = BODY
    ws.cell(row=r, column=4).alignment = WRAP
    ws.cell(row=r, column=4).border = BORDER
    ws.cell(row=r, column=2).border = BORDER
    lines = max(left.count("\n") + len(left) // 60, right.count("\n") + len(right) // 60)
    ws.row_dimensions[r].height = max(60, lines * 15)
    r += 2

ws.freeze_panes = "A5"

# ================================================================ 8. TRACKER
ws = wb.create_sheet("Weekly Tracker")
ws.sheet_view.showGridLines = False
set_widths(ws, [5, 14, 42, 16, 16, 18, 20, 18, 18, 34])

ws["A1"] = "Weekly Tracker — four numbers that matter"
ws["A1"].font = TITLE
ws.merge_cells("A1:J1")
ws["A2"] = ("Ignore likes and followers. These four are the ones that predict income: DM conversations started, "
            "Lane A bookings requested, overview videos watched (Lane B), and saves+shares on your best post.")
ws["A2"].font = SUB
ws.merge_cells("A2:J2")

hdr = 4
heads = ["", "Week Of", "Best Performing Post (hook)", "Posts Shipped", "DM Conversations Started",
         "Lane A: Trips Priced", "Lane A: Bookings Made", "Lane B: Overviews Sent",
         "Saves + Shares (best post)", "One Thing To Change Next Week"]
for i, h in enumerate(heads, start=1):
    ws.cell(row=hdr, column=i, value=h)
style_header(ws, hdr, len(heads))
r = hdr + 1
for w in range(1, 5):
    wk_start = START + timedelta(days=(w - 1) * 7)
    ws.cell(row=r, column=1, value=f"W{w}")
    ws.cell(row=r, column=2, value=wk_start.strftime("%a %d %b"))
    for c in range(1, len(heads) + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = BODY
        cell.alignment = WRAP_C if c not in (3, 10) else WRAP
        cell.border = BORDER
        cell.fill = PatternFill("solid", fgColor=WEEK_TINT[w])
    ws.row_dimensions[r].height = 34
    r += 1
ws.cell(row=r + 1, column=1, value="BENCHmarks").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
notes = ("Week 1 target: 10 DM conversations, 5 trips priced, 1 booking. If you land under half that, the hook is "
         "the problem, not the product — swap in a number-led hook from the Hooks Bank. Week 2 opens the partner "
         "lane: target your first 3 BLUEPRINT sends and treat the 'what it is NOT' carousel as the week's real "
         "deliverable. Week 3 is peak conversion because December urgency is genuine. Week 4 should produce your "
         "best numbers of the month — if it doesn't, you leaned too hard on Lane B and not enough on proof. "
         "Two rest/admin days are built in (22 and 29); do not fill them. Burnout kills more 30-day systems than "
         "bad copy does.")
c = ws.cell(row=r + 2, column=1, value=notes)
ws.merge_cells(start_row=r + 2, start_column=1, end_row=r + 2, end_column=10)
c.font = BODY
c.alignment = WRAP
c.fill = PatternFill("solid", fgColor=LIGHT)
ws.row_dimensions[r + 2].height = 62

wb.save("30-Day_Wholesale_Travel_Calendar.xlsx")
print("saved")
