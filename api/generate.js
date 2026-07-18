// api/generate.js — POST deck JSON, get back a .pptx file
// Make.com POSTs to https://<your-project>.vercel.app/api/generate

const pptxgen = require("pptxgenjs");

const SECRET = process.env.DECK_SECRET || "change-me";

// ---- Palette: Midnight Executive, dark-dominant premium ----
const C = {
  ink:    "0B1020", // near-black navy (dominant)
  panel:  "161C33", // raised surface
  panel2: "1F2740", // secondary surface
  light:  "F4F6FB", // light slide bg
  lightPanel: "E6EAF5",
  text:   "F4F6FB",
  textDk: "0B1020",
  muted:  "9AA5C4",
  mutedDk:"5A6485",
  accent: "5EE6C5", // mint — the sharp accent
  accent2:"8FA6FF", // periwinkle support
};

const F = { head: "Cambria", body: "Calibri" };

function slideDark(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  return s;
}
function slideLight(pres) {
  const s = pres.addSlide();
  s.background = { color: C.light };
  return s;
}

// Slide number chip — the repeated visual motif (circle, not a stripe)
function chip(slide, n, onDark) {
  slide.addShape("ellipse", {
    x: 12.15, y: 0.45, w: 0.55, h: 0.55,
    fill: { color: onDark ? C.panel2 : C.lightPanel }, line: { type: "none" },
  });
  slide.addText(String(n), {
    x: 12.15, y: 0.45, w: 0.55, h: 0.55, align: "center", valign: "middle",
    fontSize: 13, bold: true, fontFace: F.body,
    color: onDark ? C.accent : C.mutedDk, margin: 0,
  });
}

function title(slide, text, onDark, y) {
  slide.addText(text || "", {
    x: 0.7, y: y == null ? 0.6 : y, w: 10.9, h: 1.0, fontSize: 36, bold: true,
    fontFace: F.head, color: onDark ? C.text : C.textDk, margin: 0, valign: "middle",
  });
}

function bullets(slide, items, opt) {
  const o = opt || {};
  const arr = (items || []).filter(Boolean).map((t, i, a) => ({
    text: t,
    options: {
      bullet: true, color: o.color || C.text, fontSize: o.fontSize || 15,
      fontFace: F.body, breakLine: i !== a.length - 1, paraSpaceAfter: 12,
    },
  }));
  if (arr.length) slide.addText(arr, {
    x: o.x || 0.7, y: o.y || 2.0, w: o.w || 11.9, h: o.h || 4.2,
    margin: 0, valign: "top",
  });
}

// numbered step row (process motif)
function stepRow(slide, i, txt, x, y, w, onDark) {
  slide.addShape("ellipse", {
    x, y, w: 0.42, h: 0.42,
    fill: { color: onDark ? C.accent : C.textDk }, line: { type: "none" },
  });
  slide.addText(String(i), {
    x, y, w: 0.42, h: 0.42, align: "center", valign: "middle",
    fontSize: 12, bold: true, fontFace: F.body,
    color: onDark ? C.ink : C.light, margin: 0,
  });
  slide.addText(txt || "", {
    x: x + 0.62, y: y - 0.04, w: w - 0.62, h: 0.5, fontSize: 15, fontFace: F.body,
    color: onDark ? C.text : C.textDk, margin: 0, valign: "middle",
  });
}

// horizontal score meter
function meter(slide, label, value, x, y, w) {
  const v = Math.max(0, Math.min(100, Number(value) || 0));
  slide.addText(label, {
    x, y, w: w * 0.55, h: 0.3, fontSize: 12, bold: true, fontFace: F.body,
    color: C.muted, margin: 0, valign: "middle",
  });
  slide.addText(String(v), {
    x: x + w * 0.55, y, w: w * 0.45, h: 0.3, fontSize: 12, bold: true,
    fontFace: F.body, color: C.accent, align: "right", margin: 0, valign: "middle",
  });
  slide.addShape("roundRect", {
    x, y: y + 0.34, w, h: 0.14, rectRadius: 0.07,
    fill: { color: C.panel2 }, line: { type: "none" },
  });
  if (v > 0) slide.addShape("roundRect", {
    x, y: y + 0.34, w: Math.max(0.14, w * (v / 100)), h: 0.14, rectRadius: 0.07,
    fill: { color: C.accent }, line: { type: "none" },
  });
}

module.exports = async (req, res) => {
  if (req.method !== "POST") return res.status(405).json({ error: "POST only" });
 // if ((req.headers["x-deck-secret"] || "") !== SECRET)
   // return res.status(401).json({ error: "bad secret" });

  let d = req.body;
  if (typeof d === "string") { try { d = JSON.parse(d); } catch { return res.status(400).json({ error: "bad JSON" }); } }
  const s = (d && d.slides) || {};
  const sc = d.scores || {};

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";

  // ---------- 1. Title (dark, big statement) ----------
  {
    const sl = slideDark(pres);
    sl.addShape("ellipse", {
      x: 9.9, y: -1.6, w: 5.6, h: 5.6,
      fill: { color: C.panel, transparency: 35 }, line: { type: "none" },
    });
    sl.addShape("ellipse", {
      x: 11.2, y: 4.4, w: 3.2, h: 3.2,
      fill: { color: C.panel2, transparency: 45 }, line: { type: "none" },
    });
    sl.addText((s.title?.product_name || d.deck_title || "Untitled"), {
      x: 0.8, y: 2.35, w: 9.4, h: 1.7, fontSize: 54, bold: true,
      fontFace: F.head, color: C.text, margin: 0, valign: "middle",
    });
    sl.addText(s.title?.tagline || "", {
      x: 0.8, y: 4.05, w: 8.8, h: 0.9, fontSize: 20, italic: true,
      fontFace: F.body, color: C.accent, margin: 0, valign: "top",
    });
    sl.addText("PITCH DECK", {
      x: 0.8, y: 1.75, w: 5, h: 0.4, fontSize: 12, bold: true, charSpacing: 3,
      fontFace: F.body, color: C.muted, margin: 0,
    });
  }

  // ---------- 2. Problem (light, pain cards) ----------
  {
    const sl = slideLight(pres);
    chip(sl, 2, false);
    title(sl, s.problem?.headline || "The Problem", false);
    const pts = (s.problem?.pain_points || []).filter(Boolean).slice(0, 3);
    const cw = 3.75, gap = 0.32;
    pts.forEach((p, i) => {
      const x = 0.7 + i * (cw + gap);
      sl.addShape("roundRect", {
        x, y: 2.15, w: cw, h: 3.3, rectRadius: 0.06,
        fill: { color: C.lightPanel }, line: { type: "none" },
      });
      sl.addText(String(i + 1).padStart(2, "0"), {
        x: x + 0.3, y: 2.42, w: 1.2, h: 0.6, fontSize: 30, bold: true,
        fontFace: F.head, color: C.mutedDk, margin: 0,
      });
      sl.addText(p, {
        x: x + 0.3, y: 3.1, w: cw - 0.6, h: 2.1, fontSize: 15,
        fontFace: F.body, color: C.textDk, margin: 0, valign: "top",
      });
    });
  }

  // ---------- 3. Solution (dark, two-column: statement + steps) ----------
  {
    const sl = slideDark(pres);
    chip(sl, 3, true);
    title(sl, s.solution?.headline || "The Solution", true);
    sl.addShape("roundRect", {
      x: 0.7, y: 2.0, w: 5.5, h: 3.6, rectRadius: 0.06,
      fill: { color: C.panel }, line: { type: "none" },
    });
    sl.addText(s.solution?.description || "", {
      x: 1.0, y: 2.3, w: 4.9, h: 3.0, fontSize: 16, fontFace: F.body,
      color: C.text, margin: 0, valign: "top",
    });
    const how = (s.solution?.how_it_works || []).filter(Boolean).slice(0, 3);
    sl.addText("HOW IT WORKS", {
      x: 6.7, y: 2.0, w: 5, h: 0.35, fontSize: 12, bold: true, charSpacing: 2,
      fontFace: F.body, color: C.muted, margin: 0,
    });
    how.forEach((h, i) => stepRow(sl, i + 1, h, 6.7, 2.55 + i * 1.0, 5.6, true));
  }

  // ---------- 4. Market (light, big stat callout + evidence) ----------
  {
    const sl = slideLight(pres);
    chip(sl, 4, false);
    title(sl, "Market & Momentum", false);
    sl.addShape("roundRect", {
      x: 0.7, y: 2.15, w: 4.6, h: 3.3, rectRadius: 0.06,
      fill: { color: C.ink }, line: { type: "none" },
    });
    sl.addText(s.market?.search_volume || "Emerging", {
      x: 1.0, y: 2.75, w: 4.0, h: 1.3, fontSize: 40, bold: true,
      fontFace: F.head, color: C.accent, margin: 0, valign: "middle",
    });
    sl.addText("DEMAND SIGNAL", {
      x: 1.0, y: 2.4, w: 4.0, h: 0.35, fontSize: 11, bold: true, charSpacing: 2,
      fontFace: F.body, color: C.muted, margin: 0,
    });
    sl.addText(s.market?.tam_note || "", {
      x: 1.0, y: 4.05, w: 4.0, h: 1.1, fontSize: 13, fontFace: F.body,
      color: C.muted, margin: 0, valign: "top",
    });
    sl.addText("WHY NOW", {
      x: 5.75, y: 2.25, w: 6, h: 0.35, fontSize: 12, bold: true, charSpacing: 2,
      fontFace: F.body, color: C.mutedDk, margin: 0,
    });
    sl.addText(s.market?.trend_evidence || "", {
      x: 5.75, y: 2.7, w: 6.4, h: 1.6, fontSize: 16, fontFace: F.body,
      color: C.textDk, margin: 0, valign: "top",
    });
    // score meters if provided
    if (sc.virality || sc.commercial_potential || sc.competition) {
      meter(sl, "Virality", sc.virality, 5.75, 4.35, 3.0);
      meter(sl, "Commercial", sc.commercial_potential, 5.75, 5.0, 3.0);
      meter(sl, "Open field", sc.competition, 9.15, 4.35, 3.0);
    }
  }

  // ---------- 5. Product & technicals (dark, split cards) ----------
  {
    const sl = slideDark(pres);
    chip(sl, 5, true);
    title(sl, "Product & Technicals", true);
    const cols = [
      { label: "FEATURES", items: s.product_tech?.features, x: 0.7 },
      { label: "UNDER THE HOOD", items: s.product_tech?.technicals, x: 6.75 },
    ];
    cols.forEach((col) => {
      sl.addShape("roundRect", {
        x: col.x, y: 2.0, w: 5.55, h: 3.6, rectRadius: 0.06,
        fill: { color: C.panel }, line: { type: "none" },
      });
      sl.addText(col.label, {
        x: col.x + 0.35, y: 2.25, w: 4.8, h: 0.35, fontSize: 12, bold: true,
        charSpacing: 2, fontFace: F.body, color: C.accent, margin: 0,
      });
      bullets(sl, (col.items || []).slice(0, 4), {
        x: col.x + 0.35, y: 2.75, w: 4.85, h: 2.6, fontSize: 14,
      });
    });
  }

  // ---------- 6. Tech stack (light, layered rows) ----------
  {
    const sl = slideLight(pres);
    chip(sl, 6, false);
    title(sl, "Tech Stack", false);
    const layers = (s.tech_stack?.layers || []).slice(0, 4);
    layers.forEach((L, i) => {
      const y = 2.15 + i * 0.92;
      sl.addShape("roundRect", {
        x: 0.7, y, w: 11.9, h: 0.75, rectRadius: 0.06,
        fill: { color: C.lightPanel }, line: { type: "none" },
      });
      sl.addShape("ellipse", {
        x: 0.95, y: y + 0.19, w: 0.37, h: 0.37,
        fill: { color: C.ink }, line: { type: "none" },
      });
      sl.addText(String(i + 1), {
        x: 0.95, y: y + 0.19, w: 0.37, h: 0.37, align: "center", valign: "middle",
        fontSize: 11, bold: true, fontFace: F.body, color: C.accent, margin: 0,
      });
      sl.addText(L.name || "", {
        x: 1.55, y, w: 3.0, h: 0.75, fontSize: 15, bold: true, fontFace: F.body,
        color: C.textDk, valign: "middle", margin: 0,
      });
      sl.addText((L.tools || []).join("   ·   "), {
        x: 4.7, y, w: 7.6, h: 0.75, fontSize: 14, fontFace: F.body,
        color: C.mutedDk, valign: "middle", margin: 0,
      });
    });
  }

  // ---------- 7. Users (dark, persona cards) ----------
  {
    const sl = slideDark(pres);
    chip(sl, 7, true);
    title(sl, "Who It's For", true);
    const p = (s.users?.personas || []).slice(0, 3);
    const cw = 3.75, gap = 0.32;
    p.forEach((per, i) => {
      const x = 0.7 + i * (cw + gap);
      sl.addShape("roundRect", {
        x, y: 2.0, w: cw, h: 3.55, rectRadius: 0.06,
        fill: { color: C.panel }, line: { type: "none" },
      });
      sl.addShape("ellipse", {
        x: x + 0.3, y: 2.3, w: 0.5, h: 0.5,
        fill: { color: C.accent }, line: { type: "none" },
      });
      sl.addText(String.fromCharCode(65 + i), {
        x: x + 0.3, y: 2.3, w: 0.5, h: 0.5, align: "center", valign: "middle",
        fontSize: 15, bold: true, fontFace: F.body, color: C.ink, margin: 0,
      });
      sl.addText(per.who || "", {
        x: x + 0.3, y: 2.95, w: cw - 0.6, h: 0.55, fontSize: 17, bold: true,
        fontFace: F.head, color: C.text, margin: 0, valign: "top",
      });
      sl.addText([
        { text: "PAIN", options: { fontSize: 10, bold: true, charSpacing: 1, color: C.muted, breakLine: true, paraSpaceAfter: 3 } },
        { text: per.problem || "", options: { fontSize: 13, color: C.text, breakLine: true, paraSpaceAfter: 10 } },
        { text: "GAIN", options: { fontSize: 10, bold: true, charSpacing: 1, color: C.muted, breakLine: true, paraSpaceAfter: 3 } },
        { text: per.gain || "", options: { fontSize: 13, color: C.accent } },
      ], { x: x + 0.3, y: 3.55, w: cw - 0.6, h: 1.8, margin: 0, valign: "top", fontFace: F.body });
    });
  }

  // ---------- 8. Competition (light, gap statement + alternatives) ----------
  {
    const sl = slideLight(pres);
    chip(sl, 8, false);
    title(sl, "The Gap", false);
    sl.addText(s.competition?.gap || "", {
      x: 0.7, y: 1.95, w: 11.9, h: 1.0, fontSize: 19, italic: true,
      fontFace: F.head, color: C.textDk, margin: 0, valign: "top",
    });
    const alts = (s.competition?.alternatives || []).filter(Boolean).slice(0, 3);
    const cw = 3.75, gap = 0.32;
    alts.forEach((a, i) => {
      const x = 0.7 + i * (cw + gap);
      sl.addShape("roundRect", {
        x, y: 3.2, w: cw, h: 1.35, rectRadius: 0.06,
        fill: { color: C.lightPanel }, line: { type: "none" },
      });
      sl.addText(a, {
        x: x + 0.28, y: 3.2, w: cw - 0.56, h: 1.35, fontSize: 14,
        fontFace: F.body, color: C.mutedDk, margin: 0, valign: "middle",
      });
    });
    if (s.competition?.edge) {
      sl.addShape("roundRect", {
        x: 0.7, y: 4.8, w: 11.9, h: 1.0, rectRadius: 0.06,
        fill: { color: C.ink }, line: { type: "none" },
      });
      sl.addText([
        { text: "OUR EDGE   ", options: { fontSize: 11, bold: true, charSpacing: 2, color: C.muted } },
        { text: s.competition.edge, options: { fontSize: 15, color: C.text } },
      ], { x: 1.0, y: 4.8, w: 11.3, h: 1.0, margin: 0, valign: "middle", fontFace: F.body });
    }
  }

  // ---------- 9. GTM (dark, three stat blocks) ----------
  {
    const sl = slideDark(pres);
    chip(sl, 9, true);
    title(sl, "Go To Market", true);
    const blocks = [
      { k: "CHANNELS", v: (s.gtm?.channels || []).join("\n") },
      { k: "PRICING", v: s.gtm?.pricing || "" },
      { k: "FIRST MOVE", v: s.gtm?.first_move || "" },
    ];
    const cw = 3.75, gap = 0.32;
    blocks.forEach((b, i) => {
      const x = 0.7 + i * (cw + gap);
      sl.addShape("roundRect", {
        x, y: 2.1, w: cw, h: 3.3, rectRadius: 0.06,
        fill: { color: i === 1 ? C.panel2 : C.panel }, line: { type: "none" },
      });
      sl.addText(b.k, {
        x: x + 0.3, y: 2.4, w: cw - 0.6, h: 0.35, fontSize: 11, bold: true,
        charSpacing: 2, fontFace: F.body, color: C.accent, margin: 0,
      });
      sl.addText(b.v, {
        x: x + 0.3, y: 2.9, w: cw - 0.6, h: 2.2, fontSize: 15,
        fontFace: F.body, color: C.text, margin: 0, valign: "top",
      });
    });
  }

  // ---------- 10. Ask (dark, closing statement) ----------
  {
    const sl = slideDark(pres);
    sl.addShape("ellipse", {
      x: -1.8, y: 3.6, w: 5.4, h: 5.4,
      fill: { color: C.panel, transparency: 40 }, line: { type: "none" },
    });
    sl.addText(s.ask?.cta || "Let's talk", {
      x: 0.9, y: 2.2, w: 11.4, h: 1.5, fontSize: 44, bold: true,
      fontFace: F.head, color: C.text, margin: 0, valign: "middle",
    });
    const steps = (s.ask?.next_steps || []).filter(Boolean).slice(0, 3);
    steps.forEach((t, i) => stepRow(sl, i + 1, t, 0.9, 3.95 + i * 0.72, 8.4, true));
  }

  const buf = await pres.write("nodebuffer");
  const name = (d.deck_title || "pitch-deck").replace(/[^a-z0-9]+/gi, "-").toLowerCase();
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  res.setHeader("Content-Disposition", `attachment; filename="${name}.pptx"`);
  return res.status(200).send(buf);
};
