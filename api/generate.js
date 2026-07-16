// api/generate.js — POST deck JSON, get back a .pptx file
// Deployed on Vercel. Make.com POSTs to https://<your-project>.vercel.app/api/generate

const pptxgen = require("pptxgenjs");

// ---- simple shared secret so randoms can't spam your endpoint ----
const SECRET = process.env.DECK_SECRET || "change-me";

// ---- palette (edit to taste) ----
const C = {
  bg: "0F172A",      // slate-900
  card: "1E293B",    // slate-800
  text: "F1F5F9",    // slate-100
  muted: "94A3B8",   // slate-400
  accent: "38BDF8",  // sky-400
  accent2: "A78BFA", // violet-400
};

function bgSlide(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  return s;
}

// small helper: a heading + optional kicker
function heading(slide, text, kicker) {
  if (kicker) {
    slide.addText(kicker.toUpperCase(), {
      x: 0.6, y: 0.5, w: 12, h: 0.4, fontSize: 14, color: C.accent,
      charSpacing: 2, bold: true, margin: 0,
    });
  }
  slide.addText(text || "", {
    x: 0.6, y: kicker ? 0.95 : 0.6, w: 12, h: 1.1, fontSize: 34,
    color: C.text, bold: true, margin: 0,
  });
}

// bullet list helper
function bullets(slide, items, x, y, w) {
  const arr = (items || []).filter(Boolean).map((t, i, a) => ({
    text: t, options: { bullet: true, color: C.text, fontSize: 18,
      breakLine: i !== a.length - 1, paraSpaceAfter: 10 },
  }));
  if (arr.length) slide.addText(arr, { x, y, w: w || 11.8, h: 4, margin: 0, valign: "top" });
}

module.exports = async (req, res) => {
  if (req.method !== "POST") return res.status(405).json({ error: "POST only" });
  if ((req.headers["x-deck-secret"] || "") !== SECRET)
    return res.status(401).json({ error: "bad secret" });

  let d = req.body;
  if (typeof d === "string") { try { d = JSON.parse(d); } catch { return res.status(400).json({ error: "bad JSON" }); } }
  const s = (d && d.slides) || {};

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

  // 1. Title
  {
    const sl = bgSlide(pres);
    sl.addText(s.title?.product_name || d.deck_title || "Untitled", {
      x: 0.6, y: 2.6, w: 12, h: 1.4, fontSize: 54, bold: true, color: C.text, margin: 0,
    });
    sl.addText(s.title?.tagline || "", {
      x: 0.6, y: 4.0, w: 12, h: 1, fontSize: 22, color: C.accent, margin: 0,
    });
  }

  // 2. Problem
  {
    const sl = bgSlide(pres);
    heading(sl, s.problem?.headline || "The Problem", "Problem");
    bullets(sl, s.problem?.pain_points, 0.6, 2.1);
  }

  // 3. Solution
  {
    const sl = bgSlide(pres);
    heading(sl, s.solution?.headline || "The Solution", "Solution");
    sl.addText(s.solution?.description || "", {
      x: 0.6, y: 2.0, w: 11.8, h: 1.2, fontSize: 18, color: C.muted, margin: 0, valign: "top",
    });
    bullets(sl, s.solution?.how_it_works, 0.6, 3.3);
  }

  // 4. Market / trend evidence
  {
    const sl = bgSlide(pres);
    heading(sl, "Market & Momentum", "Why now");
    bullets(sl, [
      s.market?.trend_evidence,
      s.market?.search_volume ? `Search signal: ${s.market.search_volume}` : null,
      s.market?.tam_note,
    ], 0.6, 2.1);
  }

  // 5. Product & technicals
  {
    const sl = bgSlide(pres);
    heading(sl, "Product & Technicals", "What we built");
    sl.addText("Features", { x: 0.6, y: 2.0, w: 5.8, h: 0.4, fontSize: 16, bold: true, color: C.accent, margin: 0 });
    bullets(sl, s.product_tech?.features, 0.6, 2.5, 5.8);
    sl.addText("Under the hood", { x: 6.9, y: 2.0, w: 5.8, h: 0.4, fontSize: 16, bold: true, color: C.accent2, margin: 0 });
    bullets(sl, s.product_tech?.technicals, 6.9, 2.5, 5.8);
  }

  // 6. Tech stack
  {
    const sl = bgSlide(pres);
    heading(sl, "Tech Stack", "How it's built");
    const layers = (s.tech_stack?.layers || []).slice(0, 4);
    layers.forEach((L, i) => {
      const y = 2.1 + i * 1.15;
      sl.addShape(pres.ShapeType.roundRect, { x: 0.6, y, w: 12, h: 0.95, fill: { color: C.card }, rectRadius: 0.08, line: { type: "none" } });
      sl.addText(L.name || "", { x: 0.85, y, w: 3.2, h: 0.95, fontSize: 16, bold: true, color: C.text, valign: "middle", margin: 0 });
      sl.addText((L.tools || []).join("  •  "), { x: 4.1, y, w: 8.3, h: 0.95, fontSize: 15, color: C.muted, valign: "middle", margin: 0 });
    });
  }

  // 7. Users / personas
  {
    const sl = bgSlide(pres);
    heading(sl, "Who It's For", "Users");
    const p = (s.users?.personas || []).slice(0, 3);
    const cw = 3.9, gap = 0.25;
    p.forEach((per, i) => {
      const x = 0.6 + i * (cw + gap);
      sl.addShape(pres.ShapeType.roundRect, { x, y: 2.1, w: cw, h: 3.6, fill: { color: C.card }, rectRadius: 0.08, line: { type: "none" } });
      sl.addText(per.who || "", { x: x + 0.25, y: 2.35, w: cw - 0.5, h: 0.6, fontSize: 18, bold: true, color: C.accent, margin: 0 });
      sl.addText([
        { text: "Problem", options: { fontSize: 12, bold: true, color: C.muted, breakLine: true, paraSpaceAfter: 4 } },
        { text: per.problem || "", options: { fontSize: 14, color: C.text, breakLine: true, paraSpaceAfter: 12 } },
        { text: "Gain", options: { fontSize: 12, bold: true, color: C.muted, breakLine: true, paraSpaceAfter: 4 } },
        { text: per.gain || "", options: { fontSize: 14, color: C.text } },
      ], { x: x + 0.25, y: 3.05, w: cw - 0.5, h: 2.5, margin: 0, valign: "top" });
    });
  }

  // 8. Competition / gap
  {
    const sl = bgSlide(pres);
    heading(sl, "The Gap", "Competition");
    sl.addText(s.competition?.gap || "", { x: 0.6, y: 2.0, w: 11.8, h: 1, fontSize: 18, color: C.text, margin: 0, valign: "top" });
    bullets(sl, s.competition?.alternatives, 0.6, 3.1);
    if (s.competition?.edge)
      sl.addText(`Our edge: ${s.competition.edge}`, { x: 0.6, y: 5.4, w: 11.8, h: 0.8, fontSize: 17, italic: true, color: C.accent, margin: 0 });
  }

  // 9. Go-to-market
  {
    const sl = bgSlide(pres);
    heading(sl, "Go To Market", "Plan");
    bullets(sl, [
      s.gtm?.channels ? `Channels: ${(s.gtm.channels || []).join(", ")}` : null,
      s.gtm?.pricing ? `Pricing: ${s.gtm.pricing}` : null,
      s.gtm?.first_move ? `First move: ${s.gtm.first_move}` : null,
    ], 0.6, 2.1);
  }

  // 10. The ask
  {
    const sl = bgSlide(pres);
    sl.addText(s.ask?.cta || "Let's talk", { x: 0.6, y: 2.6, w: 12, h: 1.2, fontSize: 40, bold: true, color: C.text, margin: 0 });
    bullets(sl, s.ask?.next_steps, 0.6, 4.0);
  }

  const buf = await pres.write("nodebuffer");
  const name = (d.deck_title || "pitch-deck").replace(/[^a-z0-9]+/gi, "-").toLowerCase();
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  res.setHeader("Content-Disposition", `attachment; filename="${name}.pptx"`);
  return res.status(200).send(buf);
};
