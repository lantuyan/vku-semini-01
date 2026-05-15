// Build 2 slide decks for the iOS seminar at VKU.
// Output: seminar-ios-vku-ca1.pptx, seminar-ios-vku-ca2.pptx
// Tone: green (forest), large readable text for a 200-person hall.

const pptxgen = require("pptxgenjs");
const path = require("path");

// ===== Color palette (Forest & Moss inspired) =====
const C = {
  bg:        "F7FBF6", // off-white green tint
  bgDark:    "1B3A1F", // deep forest
  primary:   "2E7D32", // green 800
  primaryDk: "1B5E20", // green 900
  accent:    "97BC62", // moss
  ink:       "1A1A1A", // near black
  inkSoft:   "555555",
  white:     "FFFFFF",
};

const FONT = "Helvetica";
const ASSET_DIR = path.join(__dirname, "assets");

function asset(name) {
  return path.join(ASSET_DIR, name);
}

// ===================================================
// Helpers
// ===================================================
function addBaseBg(slide, dark = false) {
  slide.background = { color: dark ? C.bgDark : C.bg };
  // left accent stripe
  slide.addShape("rect", {
    x: 0, y: 0, w: 0.18, h: 7.5,
    fill: { color: dark ? C.accent : C.primary }, line: { color: dark ? C.accent : C.primary, width: 0 }
  });
  // top-right small block (visual motif — repeat across deck)
  slide.addShape("rect", {
    x: 12.6, y: 0, w: 0.7, h: 0.18,
    fill: { color: dark ? C.accent : C.primary }, line: { color: dark ? C.accent : C.primary, width: 0 }
  });
}

function addFooter(slide, deckLabel, dark = false) {
  slide.addText("Nguyễn Quế Lân · VNGalaxy", {
    x: 8.5, y: 7.1, w: 4.5, h: 0.3,
    fontFace: FONT, fontSize: 11, color: dark ? C.accent : C.inkSoft, align: "right",
  });
}

function addSource(slide, text, x = 0.65, y = 6.78, w = 8.2, dark = false) {
  slide.addText(text, {
    x, y, w, h: 0.25,
    fontFace: FONT, fontSize: 8.5,
    color: dark ? "DDE8D8" : C.inkSoft,
  });
}

function addPhotoPanel(slide, imagePath, x, y, w, h, caption) {
  slide.addImage({ path: imagePath, x, y, w, h, sizing: { type: "cover", w, h } });
  slide.addShape("rect", {
    x, y: y + h - 0.45, w, h: 0.45,
    fill: { color: C.bgDark, transparency: 12 },
    line: { color: C.bgDark, transparency: 100, width: 0 },
  });
  slide.addText(caption, {
    x: x + 0.18, y: y + h - 0.34, w: w - 0.36, h: 0.2,
    fontFace: FONT, fontSize: 8.5, color: C.white,
  });
}

function addSlideTitle(slide, title, subtitle = "") {
  slide.addText(title, {
    x: 0.5, y: 0.35, w: 12.5, h: 0.65,
    fontFace: FONT, fontSize: 30, bold: true, color: C.primaryDk,
  });
  slide.addShape("rect", {
    x: 0.5, y: 1.08, w: 1.5, h: 0.06,
    fill: { color: C.primary }, line: { color: C.primary, width: 0 },
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.5, y: 1.2, w: 11.9, h: 0.32,
      fontFace: FONT, fontSize: 13.5, color: C.inkSoft,
    });
  }
}

function addPill(slide, text, x, y, w, color = C.primary, fill = C.white) {
  slide.addShape("roundRect", {
    x, y, w, h: 0.44,
    rectRadius: 0.08,
    fill: { color: fill },
    line: { color, width: 1.1 },
  });
  slide.addText(text, {
    x: x + 0.08, y: y + 0.11, w: w - 0.16, h: 0.16,
    fontFace: FONT, fontSize: 9.5, bold: true, color, align: "center",
  });
}

function addMetricBar(slide, label, value, max, x, y, w, color, note = "") {
  slide.addText(label, { x, y, w: 2.1, h: 0.25, fontFace: FONT, fontSize: 12, bold: true, color: C.ink });
  slide.addShape("rect", { x: x + 2.25, y: y + 0.04, w, h: 0.2, fill: { color: "DDE8D8" }, line: { color: "DDE8D8", width: 0 } });
  slide.addShape("rect", { x: x + 2.25, y: y + 0.04, w: w * value / max, h: 0.2, fill: { color }, line: { color, width: 0 } });
  slide.addText(note || String(value), { x: x + 2.25 + w + 0.15, y: y - 0.01, w: 0.7, h: 0.2, fontFace: FONT, fontSize: 10, bold: true, color });
}

function titleSlide(pres, line1, line2, subtitle, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, true);
  addPhotoPanel(s, asset("devices-desk.jpg"), 8.35, 0.55, 4.15, 5.35, "Ảnh minh họa: Héctor Martínez / Wikimedia Commons CC0");
  s.addShape("rect", {
    x: 8.35, y: 0.55, w: 4.15, h: 5.35,
    fill: { color: C.bgDark, transparency: 55 },
    line: { color: C.accent, width: 0.8 },
  });
  s.addText(line1, {
    x: 0.8, y: 1.15, w: 7.35, h: 1.05,
    fontFace: FONT, fontSize: 56, bold: true, color: C.white,
  });
  s.addText(line2, {
    x: 0.8, y: 2.25, w: 7.35, h: 1.55,
    fontFace: FONT, fontSize: 56, bold: true, color: C.accent,
  });
  s.addText(subtitle, {
    x: 0.8, y: 4.08, w: 7.2, h: 0.66,
    fontFace: FONT, fontSize: 21, color: C.white,
  });
  [
    ["Swift", 8.65, 5.2],
    ["Xcode", 9.72, 5.2],
    ["Outsource", 10.85, 5.2],
  ].forEach(([label, x, y]) => addPill(s, label, x, y, 0.95, C.accent, C.bgDark));
  s.addText("Nguyễn Quế Lân — Technical Leader, VNGalaxy", {
    x: 0.8, y: 5.6, w: 12, h: 0.5,
    fontFace: FONT, fontSize: 20, color: C.white, bold: true,
  });
  s.addText("Sinh viên năm 2 — VKU · Sáng thứ Bảy · 2 ca × 90 phút", {
    x: 0.8, y: 6.1, w: 12, h: 0.5,
    fontFace: FONT, fontSize: 16, color: C.accent,
  });
  addFooter(s, deckLabel, true);
  return s;
}

function sectionDivider(pres, kicker, title, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, true);
  s.addText(kicker, {
    x: 0.8, y: 2.6, w: 12, h: 0.5,
    fontFace: FONT, fontSize: 18, color: C.accent, bold: true, charSpacing: 4,
  });
  s.addText(title, {
    x: 0.8, y: 3.1, w: 12, h: 2.0,
    fontFace: FONT, fontSize: 52, bold: true, color: C.white,
  });
  addFooter(s, deckLabel, true);
  return s;
}

function contentSlide(pres, title, bullets, deckLabel, opts = {}) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  s.addText(title, {
    x: 0.5, y: 0.35, w: 12.5, h: 0.8,
    fontFace: FONT, fontSize: opts.titleSize || 32, bold: true, color: C.primaryDk,
  });
  // accent underline
  s.addShape("rect", { x: 0.5, y: 1.1, w: 1.5, h: 0.06, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });

  // Bullets
  const items = bullets.map((b, i) => {
    if (typeof b === "string") {
      return { text: b, options: { bullet: { code: "25A0" }, breakLine: i !== bullets.length - 1, color: C.ink, fontSize: opts.bulletSize || 22 } };
    }
    return {
      text: b.text,
      options: {
        bullet: b.sub ? { indent: 24 } : { code: "25A0" },
        indentLevel: b.sub ? 1 : 0,
        breakLine: i !== bullets.length - 1,
        color: b.dim ? C.inkSoft : C.ink,
        fontSize: b.size || (opts.bulletSize || 22),
        bold: !!b.bold,
      }
    };
  });
  s.addText(items, {
    x: 0.7, y: 1.4, w: 12.0, h: 5.4,
    fontFace: FONT, valign: "top", paraSpaceAfter: 6,
  });
  addFooter(s, deckLabel, false);
  return s;
}

function twoColSlide(pres, title, leftTitle, left, rightTitle, right, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  s.addText(title, {
    x: 0.5, y: 0.35, w: 12.5, h: 0.8,
    fontFace: FONT, fontSize: 32, bold: true, color: C.primaryDk,
  });
  s.addShape("rect", { x: 0.5, y: 1.1, w: 1.5, h: 0.06, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });

  // Left card
  s.addShape("rect", {
    x: 0.5, y: 1.4, w: 6.1, h: 5.4,
    fill: { color: C.white }, line: { color: C.accent, width: 1 },
  });
  s.addText(leftTitle, { x: 0.7, y: 1.55, w: 5.8, h: 0.6, fontFace: FONT, fontSize: 22, bold: true, color: C.primaryDk });
  s.addText(left.map((t, i) => ({ text: t, options: { bullet: { code: "25A0" }, color: C.ink, fontSize: 18, breakLine: i !== left.length - 1 } })),
    { x: 0.85, y: 2.15, w: 5.65, h: 4.5, fontFace: FONT, valign: "top", paraSpaceAfter: 6 });

  // Right card
  s.addShape("rect", {
    x: 6.8, y: 1.4, w: 6.1, h: 5.4,
    fill: { color: C.white }, line: { color: C.accent, width: 1 },
  });
  s.addText(rightTitle, { x: 7.0, y: 1.55, w: 5.8, h: 0.6, fontFace: FONT, fontSize: 22, bold: true, color: C.primaryDk });
  s.addText(right.map((t, i) => ({ text: t, options: { bullet: { code: "25A0" }, color: C.ink, fontSize: 18, breakLine: i !== right.length - 1 } })),
    { x: 7.15, y: 2.15, w: 5.65, h: 4.5, fontFace: FONT, valign: "top", paraSpaceAfter: 6 });

  addFooter(s, deckLabel, false);
  return s;
}

function mobileTalentComparisonSlide(pres, deckLabel, emphasis) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  s.addText("Mobile talent: iOS ít người học hơn mức nhu cầu tuyển", {
    x: 0.5, y: 0.35, w: 12.5, h: 0.75,
    fontFace: FONT, fontSize: 30, bold: true, color: C.primaryDk,
  });
  s.addShape("rect", { x: 0.5, y: 1.08, w: 1.5, h: 0.06, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });

  s.addText("Nếu Android cần 10 người...", {
    x: 0.65, y: 1.35, w: 5.75, h: 0.35,
    fontFace: FONT, fontSize: 21, bold: true, color: C.primaryDk,
  });
  s.addText("...nhưng sinh viên học iOS lại rất ít", {
    x: 6.9, y: 1.35, w: 5.75, h: 0.35,
    fontFace: FONT, fontSize: 21, bold: true, color: C.primaryDk,
  });

  s.addShape("rect", { x: 0.6, y: 1.85, w: 5.85, h: 3.05, fill: { color: C.white }, line: { color: C.accent, width: 1 } });
  s.addText("Nhu cầu tuyển dụng", {
    x: 0.95, y: 2.05, w: 5.15, h: 0.32,
    fontFace: FONT, fontSize: 17, bold: true, color: C.ink,
  });

  const demandRatioRows = [
    { label: "Android", count: 10, raw: "22,6%", color: C.primary },
    { label: "iOS", count: 7, raw: "16,4%", color: C.accent },
  ];
  demandRatioRows.forEach((r, idx) => {
    const y = 2.65 + idx * 0.78;
    s.addText(r.label, { x: 0.95, y, w: 1.55, h: 0.3, fontFace: FONT, fontSize: 16, bold: true, color: C.ink });
    s.addText(`${r.count}`, { x: 2.6, y: y - 0.08, w: 0.65, h: 0.42, fontFace: FONT, fontSize: 27, bold: true, color: r.color, align: "right" });
    s.addShape("rect", { x: 3.45, y: y + 0.02, w: 2.1, h: 0.24, fill: { color: "DDE8D8" }, line: { color: "DDE8D8", width: 0 } });
    s.addShape("rect", { x: 3.45, y: y + 0.02, w: 2.1 * (r.count / 10), h: 0.24, fill: { color: r.color }, line: { color: r.color, width: 0 } });
    s.addText(r.raw, { x: 5.55, y: y - 0.02, w: 0.62, h: 0.25, fontFace: FONT, fontSize: 10, bold: true, color: C.inkSoft, align: "right" });
  });
  s.addText("Quy đổi từ mục tiêu tuyển dụng 2024: Android 22,6%, iOS 16,4%. Nếu Android = 10 thì iOS ≈ 7.", {
    x: 0.95, y: 4.08, w: 5.05, h: 0.42,
    fontFace: FONT, fontSize: 12.2, color: C.inkSoft,
  });

  s.addShape("rect", { x: 6.8, y: 1.85, w: 5.85, h: 3.05, fill: { color: C.white }, line: { color: C.accent, width: 1 } });
  s.addText("Pipeline sinh viên", {
    x: 7.15, y: 2.05, w: 5.1, h: 0.32,
    fontFace: FONT, fontSize: 17, bold: true, color: C.ink,
  });

  const studentRatioRows = [
    { label: "Android", count: 10, color: C.primary },
    { label: "iOS", count: 1, color: C.accent },
  ];
  studentRatioRows.forEach((r, idx) => {
    const y = 2.65 + idx * 0.78;
    s.addText(r.label, { x: 7.15, y, w: 1.55, h: 0.3, fontFace: FONT, fontSize: 16, bold: true, color: C.ink });
    s.addText(`${r.count}`, { x: 8.8, y: y - 0.08, w: 0.65, h: 0.42, fontFace: FONT, fontSize: 27, bold: true, color: r.color, align: "right" });
    s.addShape("rect", { x: 9.65, y: y + 0.02, w: 2.1, h: 0.24, fill: { color: "DDE8D8" }, line: { color: "DDE8D8", width: 0 } });
    s.addShape("rect", { x: 9.65, y: y + 0.02, w: 2.1 * (r.count / 10), h: 0.24, fill: { color: r.color }, line: { color: r.color, width: 0 } });
  });
  s.addText("Minh họa từ quan sát lớp học/khảo sát nhanh: nhiều bạn chọn Android trước; iOS bị rào cản MacBook, iPhone, App Store.", {
    x: 7.15, y: 4.08, w: 5.05, h: 0.42,
    fontFace: FONT, fontSize: 12.2, color: C.inkSoft,
  });

  s.addShape("rect", { x: 0.6, y: 5.2, w: 12.05, h: 1.2, fill: { color: C.bgDark }, line: { color: C.bgDark, width: 0 } });
  s.addText("Thông điệp chính", {
    x: 0.9, y: 5.36, w: 2.05, h: 0.28,
    fontFace: FONT, fontSize: 13, bold: true, color: C.accent,
  });
  s.addText(emphasis, {
    x: 0.9, y: 5.68, w: 11.45, h: 0.42,
    fontFace: FONT, fontSize: 17.5, bold: true, color: C.white,
  });
  s.addText("Nguồn tuyển dụng: TopDev Vietnam IT & Tech Talent Landscape 2024-2025, tr.34. Tỷ lệ sinh viên là minh họa/proxy, nên xác nhận bằng khảo sát nhanh tại lớp.", {
    x: 0.65, y: 6.65, w: 7.8, h: 0.28,
    fontFace: FONT, fontSize: 10.5, color: C.inkSoft,
  });

  addFooter(s, deckLabel, false);
  return s;
}

function timelineSlide(pres, title, rows, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  s.addText(title, { x: 0.5, y: 0.35, w: 12.5, h: 0.8, fontFace: FONT, fontSize: 32, bold: true, color: C.primaryDk });
  s.addShape("rect", { x: 0.5, y: 1.1, w: 1.5, h: 0.06, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });

  // Header
  const headerOpts = {
    fontFace: FONT, fontSize: 16, bold: true, color: C.white,
    fill: { color: C.primary }, align: "left", valign: "middle", margin: 6,
  };
  const cellOpts = { fontFace: FONT, fontSize: 16, color: C.ink, align: "left", valign: "middle", margin: 6,
    border: [{ pt: 0.5, color: "CCCCCC" }, { pt: 0.5, color: "CCCCCC" }, { pt: 0.5, color: "CCCCCC" }, { pt: 0.5, color: "CCCCCC" }] };

  const headerRow = [
    { text: "Thời gian", options: headerOpts },
    { text: "Phút", options: headerOpts },
    { text: "Nội dung", options: headerOpts },
  ];
  const bodyRows = rows.map(r => r.map((c, i) => ({ text: c, options: { ...cellOpts, bold: i === 0 } })));

  s.addTable([headerRow, ...bodyRows], {
    x: 0.5, y: 1.4, w: 12.3,
    colW: [2.0, 1.0, 9.3],
    rowH: 0.55,
    fontFace: FONT,
  });
  addFooter(s, deckLabel, false);
  return s;
}

function bigQuoteSlide(pres, quote, attribution, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, true);
  s.addText(`“${quote}”`, {
    x: 1.0, y: 2.0, w: 11.3, h: 3.0,
    fontFace: FONT, fontSize: 40, italic: true, color: C.white, valign: "middle",
  });
  s.addText(`— ${attribution}`, {
    x: 1.0, y: 5.2, w: 11.3, h: 0.6,
    fontFace: FONT, fontSize: 22, color: C.accent, bold: true,
  });
  addFooter(s, deckLabel, true);
  return s;
}

function platformEcosystemSlide(pres, deckLabel, flavor) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Hệ sinh thái Apple Platforms", "Một ngôn ngữ và một toolchain đi qua nhiều thiết bị.");

  addPhotoPanel(s, asset("devices-desk.jpg"), 8.55, 1.35, 3.95, 4.65, "Ảnh minh họa: Héctor Martínez / Wikimedia Commons CC0");
  s.addShape("rect", { x: 0.75, y: 1.85, w: 7.1, h: 0.12, fill: { color: "DDE8D8" }, line: { color: "DDE8D8", width: 0 } });
  const platforms = [
    ["iOS", "iPhone", 0.95, C.primary],
    ["iPadOS", "iPad", 2.28, C.accent],
    ["watchOS", "Watch", 3.78, C.primary],
    ["macOS", "Mac", 5.22, C.accent],
    ["visionOS", "Spatial", 6.48, C.primary],
  ];
  platforms.forEach(([name, device, x, color], idx) => {
    s.addShape("ellipse", { x, y: 1.53, w: 0.7, h: 0.7, fill: { color }, line: { color, width: 0 } });
    s.addText(String(idx + 1), { x, y: 1.72, w: 0.7, h: 0.2, fontFace: FONT, fontSize: 11, bold: true, color: C.white, align: "center" });
    s.addText(name, { x: x - 0.24, y: 2.38, w: 1.18, h: 0.25, fontFace: FONT, fontSize: 13.5, bold: true, color: C.primaryDk, align: "center" });
    s.addText(device, { x: x - 0.24, y: 2.72, w: 1.18, h: 0.22, fontFace: FONT, fontSize: 9.5, color: C.inkSoft, align: "center" });
  });
  s.addShape("rect", { x: 0.8, y: 3.45, w: 7.3, h: 1.3, fill: { color: C.white }, line: { color: C.accent, width: 1 } });
  s.addText("Ý cần giải thích cho sinh viên", { x: 1.05, y: 3.65, w: 3.2, h: 0.25, fontFace: FONT, fontSize: 13, bold: true, color: C.primaryDk });
  s.addText([
    "Học iOS không chỉ là học một chiếc điện thoại.",
    "Khi nắm Swift + Xcode + SDK, bạn có thể chuyển sang iPad, Mac, Watch hoặc visionOS dễ hơn nhiều so với học lại từ đầu.",
    flavor === "enterprise" ? "Với outsource, đây là lợi thế khi client cần mở rộng nhiều nền tảng Apple." : "Với indie, đây là lợi thế khi mở rộng sản phẩm ra nhiều thiết bị.",
  ].join("\n"), { x: 1.05, y: 3.98, w: 6.75, h: 0.62, fontFace: FONT, fontSize: 13.6, color: C.ink, breakLine: false, fit: "shrink" });
  addSource(s, "Nguồn ảnh: Wikimedia Commons; nội dung kỹ thuật: Apple Developer platform documentation.");
  addFooter(s, deckLabel, false);
  return s;
}

function appArchitectureVisualSlide(pres, deckLabel, flavor) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Cấu trúc một app iOS hiện đại", "Đọc project theo luồng: UI -> State/Logic -> Data -> Platform.");

  const layers = [
    ["UI", "SwiftUI / UIKit", "Màn hình, component, navigation", C.primary],
    ["Logic", "ViewModel / Use case", "State, validation, async/await", C.accent],
    ["Data", "REST/GraphQL + local DB", "Codable, cache, Core Data/SwiftData", C.primary],
    ["Platform", "Apple SDK", "Push, Maps, Health, AV, StoreKit", C.accent],
  ];
  layers.forEach((l, i) => {
    const y = 1.55 + i * 1.05;
    s.addShape("rect", { x: 0.85 + i * 0.18, y, w: 7.0 - i * 0.28, h: 0.72, fill: { color: i % 2 ? "EEF5EA" : C.white }, line: { color: l[3], width: 1 } });
    s.addText(l[0], { x: 1.1 + i * 0.18, y: y + 0.2, w: 0.9, h: 0.24, fontFace: FONT, fontSize: 14, bold: true, color: l[3] });
    s.addText(l[1], { x: 2.0 + i * 0.18, y: y + 0.15, w: 2.15, h: 0.24, fontFace: FONT, fontSize: 13, bold: true, color: C.ink });
    s.addText(l[2], { x: 4.15 + i * 0.18, y: y + 0.15, w: 3.25 - i * 0.18, h: 0.28, fontFace: FONT, fontSize: 11.5, color: C.inkSoft });
  });
  s.addShape("rect", { x: 8.55, y: 1.45, w: 3.95, h: 4.55, fill: { color: C.bgDark }, line: { color: C.bgDark, width: 0 } });
  s.addText("Ví dụ project outsource", { x: 8.88, y: 1.78, w: 3.25, h: 0.28, fontFace: FONT, fontSize: 15, bold: true, color: C.accent });
  const modules = flavor === "enterprise"
    ? ["Auth", "Account", "Transfer", "KYC", "Notification", "Analytics"]
    : ["Onboarding", "Paywall", "Core feature", "Settings", "Analytics", "Support"];
  modules.forEach((m, i) => {
    const x = 8.9 + (i % 2) * 1.55;
    const y = 2.35 + Math.floor(i / 2) * 0.72;
    addPill(s, m, x, y, 1.35, i % 2 ? C.accent : C.primary, C.bgDark);
  });
  s.addText("Câu hỏi hay để hỏi lớp: nếu bug ở màn hình login, em sẽ bắt đầu đọc từ layer nào?", {
    x: 8.88, y: 5.05, w: 3.25, h: 0.55,
    fontFace: FONT, fontSize: 13.5, bold: true, color: C.white,
  });
  addFooter(s, deckLabel, false);
  return s;
}

function languageHistoryVisualSlide(pres, deckLabel, flavor) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Objective-C -> Swift: đừng học như hai thế giới tách rời", "Swift là chính, Objective-C giúp đọc legacy và hiểu nền tảng Apple lâu đời.");
  s.addImage({ path: asset("swift-logo.png"), x: 9.25, y: 1.12, w: 2.35, h: 0.73 });

  const milestones = [
    ["1984", "Objective-C", "C + Smalltalk\nmessage syntax"],
    ["2008", "UIKit", "iPhone SDK\nproduction UI"],
    ["2014", "Swift", "type-safe\nmodern syntax"],
    ["2019", "SwiftUI", "declarative\nstate-driven"],
    ["2024", "Swift 6", "strict concurrency\nsafer async"],
  ];
  s.addShape("rect", { x: 0.95, y: 3.25, w: 10.8, h: 0.07, fill: { color: "CBDCC6" }, line: { color: "CBDCC6", width: 0 } });
  milestones.forEach((m, i) => {
    const x = 0.95 + i * 2.7;
    const color = i < 2 ? C.inkSoft : (i % 2 ? C.accent : C.primary);
    s.addShape("ellipse", { x: x - 0.18, y: 3.07, w: 0.42, h: 0.42, fill: { color }, line: { color, width: 0 } });
    s.addText(m[0], { x: x - 0.45, y: 2.34, w: 0.9, h: 0.26, fontFace: FONT, fontSize: 14, bold: true, color, align: "center" });
    s.addText(m[1], { x: x - 0.78, y: 2.66, w: 1.55, h: 0.28, fontFace: FONT, fontSize: 13, bold: true, color: C.ink, align: "center" });
    s.addText(m[2], { x: x - 0.85, y: 3.65, w: 1.7, h: 0.55, fontFace: FONT, fontSize: 10.5, color: C.inkSoft, align: "center" });
  });
  s.addShape("rect", { x: 1.0, y: 5.2, w: 11.55, h: 0.82, fill: { color: C.bgDark }, line: { color: C.bgDark, width: 0 } });
  s.addText(flavor === "enterprise"
    ? "Thông điệp phỏng vấn: viết Swift tốt, nhưng phải đọc được Obj-C/bridging khi gặp codebase cũ."
    : "Thông điệp indie: chọn Swift/SwiftUI để ship nhanh, biết Obj-C như kiến thức nền.",
    { x: 1.25, y: 5.48, w: 10.95, h: 0.25, fontFace: FONT, fontSize: 16, bold: true, color: C.white });
  addSource(s, "Logo Swift: Apple Inc via Wikimedia Commons, Apache 2.0. Mốc thời gian: Apple Developer / Swift project history.");
  addFooter(s, deckLabel, false);
  return s;
}

function frameworkComparisonVisualSlide(pres, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "UIKit vs SwiftUI: khác nhau ở cách nghĩ, không chỉ cú pháp", "Dùng ví dụ state để giải thích vì sao SwiftUI dễ bắt đầu hơn nhưng UIKit vẫn mạnh.");

  const cols = [
    ["UIKit", "Imperative", "Controller nói UI phải làm gì", "UIViewController\nviewDidLoad()\nbutton.addTarget(...)\ntableView.reloadData()", C.primary],
    ["SwiftUI", "Declarative", "State đổi -> View tự render lại", "struct ContentView: View\n@State var items = []\nList(items) { ... }", C.accent],
  ];
  cols.forEach((c, i) => {
    const x = 0.75 + i * 6.25;
    s.addShape("rect", { x, y: 1.55, w: 5.72, h: 4.75, fill: { color: C.white }, line: { color: c[4], width: 1.2 } });
    s.addShape("rect", { x, y: 1.55, w: 5.72, h: 0.12, fill: { color: c[4] }, line: { color: c[4], width: 0 } });
    s.addText(c[0], { x: x + 0.3, y: 1.9, w: 2.1, h: 0.35, fontFace: FONT, fontSize: 21, bold: true, color: C.primaryDk });
    s.addText(c[1], { x: x + 3.35, y: 1.95, w: 1.9, h: 0.25, fontFace: FONT, fontSize: 12.5, bold: true, color: c[4], align: "right" });
    s.addText(c[2], { x: x + 0.3, y: 2.45, w: 5.05, h: 0.38, fontFace: FONT, fontSize: 14.5, bold: true, color: C.ink });
    s.addShape("rect", { x: x + 0.35, y: 3.12, w: 5.0, h: 1.45, fill: { color: "102018" }, line: { color: "102018", width: 0 } });
    s.addText(c[3], { x: x + 0.58, y: 3.38, w: 4.55, h: 0.9, fontFace: "Courier New", fontSize: 12, color: C.accent });
    s.addText(i === 0 ? "Production lâu đời, control sâu, nhiều code legacy." : "Nhanh cho màn hình mới, preview tốt, gọn cho junior.",
      { x: x + 0.32, y: 4.95, w: 5.05, h: 0.5, fontFace: FONT, fontSize: 13.5, color: C.inkSoft });
  });
  addFooter(s, deckLabel, false);
  return s;
}

function frameworkDecisionSlide(pres, deckLabel, flavor) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Chọn UIKit hay SwiftUI trong thực tế 2026?", "Quyết định theo codebase, target iOS version và mức custom UI.");

  const decisions = [
    ["Màn mới, target iOS mới", "SwiftUI trước", C.accent],
    ["Codebase cũ / screen phức tạp", "UIKit vẫn sống khỏe", C.primary],
    ["Team outsource enterprise", "Hybrid là mặc định", C.primaryDk],
  ];
  decisions.forEach((d, i) => {
    const y = 1.65 + i * 1.15;
    s.addShape("rect", { x: 0.9, y, w: 4.0, h: 0.76, fill: { color: C.white }, line: { color: d[2], width: 1 } });
    s.addText(d[0], { x: 1.15, y: y + 0.18, w: 3.35, h: 0.22, fontFace: FONT, fontSize: 12.5, bold: true, color: C.ink });
    s.addShape("rect", { x: 5.15, y: y + 0.34, w: 0.8, h: 0.05, fill: { color: d[2] }, line: { color: d[2], width: 0 } });
    s.addShape("rect", { x: 5.82, y: y + 0.25, w: 0.18, h: 0.22, fill: { color: d[2] }, line: { color: d[2], width: 0 } });
    s.addShape("rect", { x: 6.2, y, w: 3.25, h: 0.76, fill: { color: d[2] }, line: { color: d[2], width: 0 } });
    s.addText(d[1], { x: 6.42, y: y + 0.2, w: 2.8, h: 0.2, fontFace: FONT, fontSize: 13.5, bold: true, color: C.white, align: "center" });
  });
  s.addShape("rect", { x: 9.85, y: 1.6, w: 2.35, h: 3.1, fill: { color: "EEF5EA" }, line: { color: C.accent, width: 1 } });
  s.addText("Bài học cho intern", { x: 10.1, y: 1.9, w: 1.85, h: 0.24, fontFace: FONT, fontSize: 13, bold: true, color: C.primaryDk, align: "center" });
  s.addText(flavor === "enterprise" ? "Học cả hai.\nInterview outsource hay hỏi: UIKit lifecycle, Auto Layout, SwiftUI state." : "SwiftUI để ship nhanh.\nUIKit đủ để đọc docs và xử lý edge case.",
    { x: 10.12, y: 2.55, w: 1.8, h: 1.4, fontFace: FONT, fontSize: 14, bold: true, color: C.ink, align: "center", valign: "middle" });
  addFooter(s, deckLabel, false);
  return s;
}

function toolchainVisualSlide(pres, deckLabel, flavor) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Bộ công cụ tối thiểu: Xcode là trung tâm", "Cho sinh viên nhìn toolchain như một pipeline thay vì danh sách tên.");
  addPhotoPanel(s, asset("xcode-screenshot.png"), 0.75, 1.45, 6.1, 3.7, "Screenshot Xcode: HtetPyae / Wikimedia Commons CC BY-SA 4.0");

  const tools = [
    ["Xcode", "Code + build + simulator"],
    ["SPM", "Quản lý package"],
    ["Instruments", "Memory / CPU / energy"],
    ["TestFlight", "Beta trước release"],
    [flavor === "indie" ? "RevenueCat" : "Xcode Cloud", flavor === "indie" ? "IAP/subscription" : "CI/CD chính chủ"],
  ];
  tools.forEach((t, i) => {
    const x = 7.35 + (i % 2) * 2.5;
    const y = 1.55 + Math.floor(i / 2) * 1.23;
    s.addShape("rect", { x, y, w: 2.15, h: 0.82, fill: { color: i % 2 ? "EEF5EA" : C.white }, line: { color: i % 2 ? C.accent : C.primary, width: 1 } });
    s.addText(t[0], { x: x + 0.18, y: y + 0.16, w: 1.75, h: 0.22, fontFace: FONT, fontSize: 13.5, bold: true, color: C.primaryDk });
    s.addText(t[1], { x: x + 0.18, y: y + 0.45, w: 1.78, h: 0.18, fontFace: FONT, fontSize: 9.4, color: C.inkSoft });
  });
  s.addShape("rect", { x: 1.0, y: 5.55, w: 11.4, h: 0.62, fill: { color: C.bgDark }, line: { color: C.bgDark, width: 0 } });
  s.addText("Cách giải thích: dev iOS không chỉ viết Swift; còn phải build, test, đo performance và phát hành qua App Store ecosystem.", {
    x: 1.28, y: 5.78, w: 10.9, h: 0.18, fontFace: FONT, fontSize: 13.8, bold: true, color: C.white,
  });
  addFooter(s, deckLabel, false);
  return s;
}

function tourCodeVisualSlide(pres, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Tour code mẫu: đọc project như đọc bản đồ", "Không cần gõ live; chỉ cần chỉ đúng vai trò từng thư mục.");
  const rows = [
    ["App/", "entry point, DI, app lifecycle"],
    ["Features/", "mỗi màn hình / flow là một module"],
    ["Core/", "networking, persistence, design system"],
    ["Resources/", "asset catalog, localizable strings"],
    ["Tests/", "unit test, UI test, fixtures"],
  ];
  rows.forEach((r, i) => {
    const y = 1.55 + i * 0.72;
    s.addText(r[0], { x: 0.85, y, w: 2.0, h: 0.28, fontFace: "Courier New", fontSize: 16, bold: true, color: i % 2 ? C.accent : C.primary });
    s.addShape("rect", { x: 2.95, y: y + 0.12, w: 0.85, h: 0.04, fill: { color: "CBDCC6" }, line: { color: "CBDCC6", width: 0 } });
    s.addText(r[1], { x: 4.0, y: y + 0.02, w: 3.8, h: 0.22, fontFace: FONT, fontSize: 13.2, color: C.ink });
  });
  s.addShape("rect", { x: 8.25, y: 1.45, w: 4.1, h: 3.95, fill: { color: "102018" }, line: { color: C.accent, width: 1 } });
  s.addText("struct ChecklistView: View {\n  @State var items: [Item]\n\n  var body: some View {\n    List(items) { item in\n      Row(item)\n    }\n    .task { await viewModel.load() }\n  }\n}", {
    x: 8.55, y: 1.85, w: 3.55, h: 2.65,
    fontFace: "Courier New", fontSize: 12.2, color: C.accent,
  });
  s.addText("Điểm nhấn demo: View -> ViewModel -> Service -> State", {
    x: 8.55, y: 4.85, w: 3.55, h: 0.25,
    fontFace: FONT, fontSize: 12.5, bold: true, color: C.white,
  });
  addFooter(s, deckLabel, false);
  return s;
}

function outsourceDefinitionVisualSlide(pres, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Outsource là gì?", "Hãy giải thích bằng luồng trách nhiệm, không chỉ bằng định nghĩa.");
  const nodes = [
    ["Client\nUS / JP / EU", 0.9, 2.1, C.primary],
    ["Team VN\nPM · Dev · QA", 4.55, 2.1, C.accent],
    ["App / Feature\nTestFlight -> Store", 8.2, 2.1, C.primary],
  ];
  nodes.forEach((n) => {
    s.addShape("rect", { x: n[1], y: n[2], w: 2.45, h: 1.3, fill: { color: C.white }, line: { color: n[3], width: 1.2 } });
    s.addText(n[0], { x: n[1] + 0.18, y: n[2] + 0.38, w: 2.05, h: 0.42, fontFace: FONT, fontSize: 16, bold: true, color: C.primaryDk, align: "center" });
  });
  [3.45, 7.1].forEach((x) => {
    s.addShape("rect", { x, y: 2.72, w: 0.8, h: 0.05, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });
    s.addShape("rect", { x: x + 0.66, y: 2.62, w: 0.2, h: 0.24, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });
  });
  s.addText("Client trả tiền cho năng lực delivery: hiểu requirement, estimate, code, test, release, maintain.", {
    x: 1.15, y: 4.45, w: 10.75, h: 0.35, fontFace: FONT, fontSize: 18, bold: true, color: C.ink, align: "center",
  });
  s.addText("Không sở hữu sản phẩm cuối, nhưng học được quy trình enterprise và làm việc với khách quốc tế.", {
    x: 1.45, y: 5.0, w: 10.15, h: 0.28, fontFace: FONT, fontSize: 14.5, color: C.inkSoft, align: "center",
  });
  addFooter(s, deckLabel, false);
  return s;
}

function daNangMarketVisualSlide(pres, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Đà Nẵng: outsource Mobile có đất diễn nếu bạn có iOS thật", "Dùng bức tranh địa phương để sinh viên thấy cơ hội gần mình.");
  addPhotoPanel(s, asset("dragon-bridge.jpg"), 7.35, 1.35, 5.05, 4.75, "Cầu Rồng Đà Nẵng: Tran Anh Khoa / Wikimedia Commons CC BY-SA 2.0");
  const points = [
    ["Top 3", "trung tâm CNTT lớn của Việt Nam", C.primary],
    ["JP · US · EU", "thị trường outsource phổ biến", C.accent],
    ["iOS", "cung nhân lực mỏng hơn Android", C.primary],
  ];
  points.forEach((p, i) => {
    const y = 1.55 + i * 1.22;
    s.addShape("rect", { x: 0.8, y, w: 5.9, h: 0.86, fill: { color: C.white }, line: { color: p[2], width: 1 } });
    s.addText(p[0], { x: 1.1, y: y + 0.2, w: 1.45, h: 0.28, fontFace: FONT, fontSize: 20, bold: true, color: p[2], align: "center" });
    s.addText(p[1], { x: 2.8, y: y + 0.23, w: 3.35, h: 0.22, fontFace: FONT, fontSize: 13.5, bold: true, color: C.ink });
  });
  addMetricBar(s, "Nhu cầu tuyển Android", 10, 10, 0.95, 5.12, 2.3, C.primary, "10");
  addMetricBar(s, "Nhu cầu tuyển iOS", 7, 10, 0.95, 5.55, 2.3, C.accent, "7");
  addSource(s, "Số liệu nhu cầu: quy đổi từ TopDev Vietnam IT & Tech Talent Landscape 2024-2025, tr.34. Ảnh: Wikimedia Commons.");
  addFooter(s, deckLabel, false);
  return s;
}

function dayInLifeVisualSlide(pres, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Một ngày của iOS dev outsource", "Sinh viên cần hình dung công việc lặp lại theo nhịp sprint và client timezone.");
  const items = [
    ["08:30", "Pull code\nJira\nDaily", C.primary],
    ["10:00", "Code feature\nUnit test", C.accent],
    ["14:00", "PR review\nQA bug", C.primary],
    ["16:30", "Demo\nRetro / planning", C.accent],
    ["21:00?", "Sync US\nkhi cần", C.primaryDk],
  ];
  s.addShape("rect", { x: 1.0, y: 3.1, w: 10.9, h: 0.08, fill: { color: "CBDCC6" }, line: { color: "CBDCC6", width: 0 } });
  items.forEach((it, i) => {
    const x = 1.0 + i * 2.72;
    s.addShape("ellipse", { x: x - 0.25, y: 2.85, w: 0.58, h: 0.58, fill: { color: it[2] }, line: { color: it[2], width: 0 } });
    s.addText(it[0], { x: x - 0.42, y: 2.1, w: 0.92, h: 0.25, fontFace: FONT, fontSize: 14, bold: true, color: it[2], align: "center" });
    s.addShape("rect", { x: x - 0.75, y: 3.75, w: 1.55, h: 1.05, fill: { color: C.white }, line: { color: it[2], width: 1 } });
    s.addText(it[1], { x: x - 0.58, y: 4.05, w: 1.2, h: 0.35, fontFace: FONT, fontSize: 11.5, bold: true, color: C.ink, align: "center" });
  });
  s.addShape("rect", { x: 1.25, y: 5.55, w: 10.6, h: 0.58, fill: { color: C.bgDark }, line: { color: C.bgDark, width: 0 } });
  s.addText("Kỹ năng mềm thật sự: đọc ticket rõ, hỏi sớm, update tiến độ và giao tiếp bằng tiếng Anh.", {
    x: 1.55, y: 5.78, w: 10.0, h: 0.16, fontFace: FONT, fontSize: 13.8, bold: true, color: C.white, align: "center",
  });
  addFooter(s, deckLabel, false);
  return s;
}

function outsourceProcessVisualSlide(pres, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Quy trình một dự án outsource", "Flow này giúp sinh viên thấy ticket không tự nhiên xuất hiện trong Jira.");
  const steps = [
    ["1", "Pre-sales", "proposal\nestimate"],
    ["2", "Kick-off", "scope\nteam size"],
    ["3", "Sprint", "plan -> code\nreview -> retro"],
    ["4", "UAT", "client test\nlog bug"],
    ["5", "Release", "TestFlight\nApp Store"],
    ["6", "Maintain", "fix bug\niOS update"],
  ];
  steps.forEach((st, i) => {
    const x = 0.65 + i * 2.05;
    s.addShape("rect", { x, y: 2.05, w: 1.55, h: 1.35, fill: { color: i % 2 ? "EEF5EA" : C.white }, line: { color: i % 2 ? C.accent : C.primary, width: 1 } });
    s.addText(st[0], { x: x + 0.15, y: 2.22, w: 0.3, h: 0.22, fontFace: FONT, fontSize: 12, bold: true, color: i % 2 ? C.accent : C.primary });
    s.addText(st[1], { x: x + 0.28, y: 2.52, w: 1.0, h: 0.2, fontFace: FONT, fontSize: 12.5, bold: true, color: C.primaryDk, align: "center" });
    s.addText(st[2], { x: x + 0.22, y: 2.91, w: 1.1, h: 0.32, fontFace: FONT, fontSize: 9.8, color: C.inkSoft, align: "center" });
    if (i < steps.length - 1) {
      s.addShape("rect", { x: x + 1.62, y: 2.7, w: 0.35, h: 0.05, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });
      s.addShape("rect", { x: x + 1.9, y: 2.61, w: 0.12, h: 0.22, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });
    }
  });
  s.addShape("rect", { x: 1.05, y: 4.85, w: 11.15, h: 0.85, fill: { color: C.bgDark }, line: { color: C.bgDark, width: 0 } });
  s.addText("Điểm cần nhấn mạnh: intern/fresher thường vào ở Sprint/bugfix, nhưng phải hiểu cả chuỗi để làm việc có trách nhiệm.", {
    x: 1.35, y: 5.16, w: 10.55, h: 0.22, fontFace: FONT, fontSize: 14, bold: true, color: C.white, align: "center",
  });
  addFooter(s, deckLabel, false);
  return s;
}

function careerSkillsVisualSlide(pres, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, false);
  addSlideTitle(s, "Kỹ năng cần để xin intern/fresher iOS", "Gom kỹ năng thành bản đồ ưu tiên để sinh viên tự học không bị loạn.");
  const rings = [
    ["Swift + UI", "build được màn đơn giản", 1.15, 1.85, C.primary],
    ["REST + JSON", "gọi API, parse, error state", 4.55, 1.85, C.accent],
    ["Git + Jira", "branch, PR, conflict, sprint", 7.95, 1.85, C.primary],
    ["English docs", "đọc Apple docs mỗi ngày", 2.85, 4.25, C.accent],
    ["Self-learning", "WWDC đổi liên tục", 6.25, 4.25, C.primaryDk],
  ];
  rings.forEach((r) => {
    s.addShape("ellipse", { x: r[2], y: r[3], w: 2.55, h: 1.22, fill: { color: C.white }, line: { color: r[4], width: 1.2 } });
    s.addText(r[0], { x: r[2] + 0.25, y: r[3] + 0.34, w: 2.05, h: 0.22, fontFace: FONT, fontSize: 13, bold: true, color: r[4], align: "center" });
    s.addText(r[1], { x: r[2] + 0.25, y: r[3] + 0.68, w: 2.05, h: 0.18, fontFace: FONT, fontSize: 8.8, color: C.inkSoft, align: "center" });
  });
  s.addText("Bài tập về nhà: đưa 1 app nhỏ lên GitHub, có README, ảnh màn hình, API/mock data và ít nhất 5 commit rõ nghĩa.", {
    x: 1.0, y: 6.0, w: 11.3, h: 0.3, fontFace: FONT, fontSize: 15, bold: true, color: C.ink, align: "center",
  });
  addFooter(s, deckLabel, false);
  return s;
}

// ===================================================
// SHARED — Phần kỹ thuật (dùng cho cả 2 ca, đổi sắc thái)
// ===================================================
function addCommonTechnical(pres, deckLabel, flavor /* 'enterprise' | 'indie' */) {
  // Agenda
  const agendaItems = [
    "08:00 – 08:05 · Mở đầu & khảo sát phòng",
    "08:05 – 08:20 · Tổng quan lập trình iOS",
    "08:20 – 08:35 · Objective-C → Swift",
    "08:35 – 08:50 · UIKit vs SwiftUI",
    "08:50 – 09:00 · Công cụ + tour code mẫu",
    flavor === 'enterprise' ? "09:00 – 09:30 · OUTSOURCE iOS tại Đà Nẵng + Q&A"
                            : "10:30 – 11:00 · PRODUCT / INDIE — case Flappy Bird + Q&A",
  ];
  contentSlide(pres, "Agenda — 90 phút", agendaItems, deckLabel, { bulletSize: 22 });

  // Section divider — kỹ thuật
  sectionDivider(pres, "PHẦN 1 — KỸ THUẬT", "Lập trình iOS\nhôm nay trông như thế nào?", deckLabel);

  // Tổng quan
  platformEcosystemSlide(pres, deckLabel, flavor);

  appArchitectureVisualSlide(pres, deckLabel, flavor);

  // Lịch sử
  sectionDivider(pres, "LỊCH SỬ NGÔN NGỮ", "Từ Objective-C\nđến Swift", deckLabel);

  languageHistoryVisualSlide(pres, deckLabel, flavor);

  // UIKit vs SwiftUI
  sectionDivider(pres, "FRAMEWORK", "UIKit vs SwiftUI\nhai paradigm, một hệ sinh thái", deckLabel);

  frameworkComparisonVisualSlide(pres, deckLabel);

  frameworkDecisionSlide(pres, deckLabel, flavor);

  // Công cụ
  sectionDivider(pres, "CÔNG CỤ", "Xcode, SPM,\nInstruments & tour code", deckLabel);

  toolchainVisualSlide(pres, deckLabel, flavor);

  tourCodeVisualSlide(pres, deckLabel);
}

// ===================================================
// CA 1 — OUTSOURCE
// ===================================================
function buildCa1() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
  pres.author = "Nguyễn Quế Lân";
  pres.title = "Seminar iOS — VKU — Ca 1 — Outsource";
  const LBL = "Ca 1 · 08:00–09:30 · Outsource";

  titleSlide(pres,
    "Phát triển iOS",
    "Góc nhìn OUTSOURCE",
    "Kỹ thuật · công cụ · thực trạng outsource Mobile tại Đà Nẵng",
    LBL);

  mobileTalentComparisonSlide(pres, LBL,
    "Cơ hội cho intern/fresher: iOS không ít nhu cầu; điểm nghẽn là quá ít sinh viên chọn học và có app iOS thật.");

  // About me
  contentSlide(pres, "Về diễn giả", [
    { text: "Nguyễn Quế Lân", bold: true, size: 28 },
    "Technical Leader — VNGalaxy",
    "Chuyên môn: iOS, kiến trúc mobile, lead team kỹ thuật",
    "Vai trò hôm nay: chia sẻ từ vị trí lead — chọn công nghệ, review code, tuyển dụng & mentor",
    { text: "Liên hệ: sẽ chia sẻ ở slide cuối", dim: true },
  ], LBL);

  // Mục tiêu ca 1
  contentSlide(pres, "Sau ca 1, bạn sẽ", [
    "Hiểu dòng thời gian iOS: Obj-C → Swift, UIKit → SwiftUI",
    "Biết cấu trúc một project Xcode chuẩn",
    "Hình dung được một ngày của lập trình viên iOS outsource",
    "Biết kỹ năng cần có để xin intern/fresher iOS",
    "Biết outsource Mobile tại Đà Nẵng đang trông như thế nào",
  ], LBL);

  // PHẦN KỸ THUẬT CHUNG
  addCommonTechnical(pres, LBL, 'enterprise');

  // Timeline ca 1
  timelineSlide(pres, "Timeline ca 1 — 08:00 đến 09:30", [
    ["08:00", "5'", "Mở đầu & khảo sát phòng"],
    ["08:05", "15'", "Tổng quan lập trình iOS"],
    ["08:20", "15'", "Objective-C → Swift"],
    ["08:35", "15'", "UIKit vs SwiftUI"],
    ["08:50", "10'", "Công cụ + tour code mẫu"],
    ["09:00", "25'", "OUTSOURCE iOS tại Đà Nẵng"],
    ["09:25", "5'", "Q&A & chốt 3 điểm cần nhớ"],
  ], LBL);

  // ====== PHẦN 2 — OUTSOURCE ======
  sectionDivider(pres, "PHẦN 2 — OUTSOURCE", "Đi làm thuê,\nlàm cho khách quốc tế", LBL);

  outsourceDefinitionVisualSlide(pres, LBL);

  twoColSlide(pres, "Outsource ≠ Product",
    "OUTSOURCE",
    [
      "Bán giờ công",
      "Thu nhập tuyến tính (người × giờ)",
      "Chắc chắn, ổn định",
      "Học quy trình enterprise nhanh",
      "Tiếp xúc client quốc tế",
    ],
    "PRODUCT",
    [
      "Bán sản phẩm",
      "Thu nhập có thể scale phi tuyến",
      "Rủi ro cao, upside lớn",
      "Học toàn diện hơn (PM · marketing · ops)",
      "Có 'chủ quyền' với app của mình",
    ],
    LBL);

  daNangMarketVisualSlide(pres, LBL);

  dayInLifeVisualSlide(pres, LBL);

  outsourceProcessVisualSlide(pres, LBL);

  twoColSlide(pres, "Vai trò trong team",
    "Hierarchy",
    [
      "Intern / Fresher — task nhỏ, học từ senior",
      "Junior (1–2y) — code feature có hướng dẫn",
      "Senior (3–5y) — thiết kế module, review code",
      "Tech Lead (5y+) — kiến trúc, làm việc trực tiếp client",
    ],
    "Vai trò đặc thù VN",
    [
      "BrSE / Comtor (Nhật) — cầu nối ngôn ngữ-kỹ thuật",
      "Solution Architect — design hệ thống pre-sales",
      "QA Lead — kiểm thử, viết test plan",
      "Scrum Master / PM — quản lý sprint, báo cáo client",
    ],
    LBL);

  careerSkillsVisualSlide(pres, LBL);

  twoColSlide(pres, "Outsource — Ưu & Nhược",
    "Ưu điểm",
    [
      "Dễ tìm việc đầu sự nghiệp",
      "Tiếp xúc client quốc tế",
      "Cải thiện tiếng Anh nhanh",
      "Mức lương khởi điểm tốt",
      "Học engineering chuyên nghiệp",
    ],
    "Nhược điểm",
    [
      "Ít 'chủ quyền' sản phẩm",
      "Đôi khi làm mảng lặp lại",
      "Ít cơ hội học design/PM",
      "Phụ thuộc dòng dự án từ client",
      "Career path đôi khi đóng khung",
    ],
    LBL);

  bigQuoteSlide(pres,
    "Outsource là trường học engineering tốt nhất Việt Nam — đi 2–3 năm bạn sẽ trưởng thành rất nhanh.",
    "Lời khuyên từ vị trí lead phỏng vấn",
    LBL);

  contentSlide(pres, "Trao đổi với sinh viên", [
    "Bạn nào đã thử apply intern iOS?",
    "Khó khăn lớn nhất khi đi phỏng vấn là gì?",
    "Bạn muốn hỏi gì về quy trình tuyển dụng, mức lương, công nghệ?",
    { text: "Đặt câu hỏi — không có câu nào ngớ ngẩn", bold: true },
  ], LBL);

  // Closing
  contentSlide(pres, "3 điểm cần nhớ — Ca 1", [
    { text: "1. Swift + SwiftUI là tương lai — UIKit vẫn còn cần", bold: true },
    { text: "2. Outsource = trường học engineering, không phải đích đến cuối", bold: true },
    { text: "3. Tiếng Anh + Git + tự học = bộ ba bắt buộc", bold: true },
  ], LBL, { bulletSize: 26 });

  const last = pres.addSlide();
  addBaseBg(last, true);
  last.addText("Cảm ơn!", { x: 0.8, y: 2.5, w: 12, h: 1.5, fontFace: FONT, fontSize: 72, bold: true, color: C.white });
  last.addText("Hẹn gặp lại ở ca 2 — góc nhìn Product & Indie", { x: 0.8, y: 4.0, w: 12, h: 0.8, fontFace: FONT, fontSize: 24, color: C.accent });
  last.addText("Nguyễn Quế Lân · VNGalaxy", { x: 0.8, y: 5.5, w: 12, h: 0.6, fontFace: FONT, fontSize: 20, color: C.white, bold: true });
  addFooter(last, LBL, true);

  return pres.writeFile({ fileName: "seminar-ios-vku-ca1.pptx" });
}

// ===================================================
// CA 2 — PRODUCT / INDIE
// ===================================================
function buildCa2() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "Nguyễn Quế Lân";
  pres.title = "Seminar iOS — VKU — Ca 2 — Product / Indie";
  const LBL = "Ca 2 · 09:35–11:00 · Product/Indie";

  titleSlide(pres,
    "Phát triển iOS",
    "Góc nhìn PRODUCT & INDIE",
    "Kỹ thuật · công cụ · câu chuyện Flappy Bird & con đường indie",
    LBL);

  mobileTalentComparisonSlide(pres, LBL,
    "Cơ hội cho product/indie: chọn iOS giúp bạn nổi bật nhanh hơn, vì nguồn người học mỏng hơn nhiều so với Android.");

  // About me
  contentSlide(pres, "Về diễn giả", [
    { text: "Nguyễn Quế Lân", bold: true, size: 28 },
    "Technical Leader — VNGalaxy",
    "Chuyên môn: iOS, kiến trúc mobile, lead team kỹ thuật",
    "Vai trò hôm nay: kể câu chuyện product từ góc nhìn người đã build sản phẩm thực tế",
    { text: "Liên hệ: sẽ chia sẻ ở slide cuối", dim: true },
  ], LBL);

  contentSlide(pres, "Sau ca 2, bạn sẽ", [
    "Hiểu vòng đời một product iOS từ idea đến App Store",
    "Biết các mô hình monetization phổ biến năm 2026",
    "Hiểu vì sao Flappy Bird đáng được nhắc tới mỗi năm",
    "Có một roadmap thực tế để đi đường indie từ sinh viên năm 2",
    "Phân biệt được khi nào nên outsource, khi nào nên làm product",
  ], LBL);

  // KỸ THUẬT CHUNG (flavor indie)
  addCommonTechnical(pres, LBL, 'indie');

  // Timeline ca 2
  timelineSlide(pres, "Timeline ca 2 — 09:35 đến 11:00", [
    ["09:35", "5'", "Mở đầu ca 2 — đổi tone sang 'làm sản phẩm của mình'"],
    ["09:40", "15'", "Tổng quan iOS (đi nhanh, ví dụ khác ca 1)"],
    ["09:55", "15'", "Obj-C → Swift từ góc nhìn người làm product"],
    ["10:10", "15'", "UIKit vs SwiftUI — chọn cái nào cho indie?"],
    ["10:25", "5'", "Công cụ + tour code mẫu (nhấn StoreKit, ASO)"],
    ["10:30", "25'", "PRODUCT/INDIE — case Flappy Bird"],
    ["10:55", "5'", "Q&A & lời nhắn cuối"],
  ], LBL);

  // PHẦN 2 — PRODUCT
  sectionDivider(pres, "PHẦN 2 — PRODUCT / INDIE", "Làm sản phẩm\ncủa chính bạn", LBL);

  twoColSlide(pres, "Product vs Outsource — khác biệt cốt lõi",
    "OUTSOURCE",
    [
      "Bán giờ công cho khách",
      "Thu nhập tuyến tính",
      "Chắc chắn",
      "Học engineering chuẩn",
      "Career path rõ ràng",
    ],
    "PRODUCT",
    [
      "Bán sản phẩm cho user cuối",
      "Thu nhập có thể scale phi tuyến (1 người làm, triệu user mua)",
      "Rủi ro cao, upside lớn",
      "Học toàn diện: design · marketing · ops",
      "Không loại trừ outsource — nhiều người làm cả hai",
    ],
    LBL);

  contentSlide(pres, "Vòng đời một product iOS", [
    "Idea → Validate (hỏi user thật, không hỏi bạn bè)",
    "MVP — build trong 2–6 tuần",
    "Public beta qua TestFlight",
    "Launch trên App Store — metadata, screenshot, video preview, ASO",
    "Growth — track D1/D7/D30 retention, ARPU, free→paid conversion",
    "Iterate — ra version mới mỗi 2–4 tuần",
  ], LBL);

  contentSlide(pres, "Monetization — kiếm tiền từ app như thế nào?", [
    "Paid up-front — trả 1 lần khi tải (Things 3, Procreate)",
    "Freemium + IAP — mở khoá tính năng cao cấp",
    "Subscription — mô hình lớn nhất 2020s (Duolingo, Notion, Calm)",
    "Quảng cáo — AdMob, AppLovin — phù hợp game casual (Flappy Bird)",
    "Hybrid — free + ads + IAP remove ads",
    { text: "Apple cắt 30% (15% với Small Business Program <1M USD/năm)", dim: true, size: 18 },
  ], LBL);

  // ====== FLAPPY BIRD ======
  sectionDivider(pres, "CASE STUDY", "Flappy Bird\n& Nguyễn Hà Đông", LBL);

  contentSlide(pres, "Flappy Bird — những con số", [
    "Tác giả: Nguyễn Hà Đông (Đông Nguyễn) — Hà Nội, sinh 1985",
    "Studio: dotGEARS — gần như một mình",
    "Phát hành: tháng 5/2013 — ban đầu không nổi",
    "Viral: cuối 2013–đầu 2014 nhờ độ khó 'gây nghiện gây bực'",
    "Đỉnh điểm: #1 App Store ở 100+ quốc gia",
    "Doanh thu công khai ước tính ~50.000 USD/ngày từ quảng cáo banner",
    { text: "Gỡ khỏi App Store ngày 10/02/2014 — quyết định gây sốc toàn cầu", bold: true },
  ], LBL, { bulletSize: 20 });

  bigQuoteSlide(pres,
    "Tôi không thể chịu được nữa.",
    "Nguyễn Hà Đông, giải thích lý do gỡ Flappy Bird (2014)",
    LBL);

  contentSlide(pres, "Vì sao Flappy Bird đáng nhắc với SV Việt Nam 2026?", [
    "Một người, một app đơn giản, đứng #1 App Store thế giới",
    "Phá niềm tin 'muốn nổi phải ở Silicon Valley' — App Store là sân chơi phẳng",
    { text: "Ý tưởng > Công nghệ", bold: true },
    "Flappy Bird dùng kỹ thuật cơ bản: gameplay 1 tap, art retro",
    "Yếu tố 'gây nghiện' & viral organic > số lượng tính năng",
    { text: "Bài học cá nhân: thành công vượt kỳ vọng có cái giá của nó", bold: true },
  ], LBL);

  // Indie tại VN/Đà Nẵng
  contentSlide(pres, "Indie iOS tại Việt Nam & Đà Nẵng", [
    "Cộng đồng indie/solo nhỏ nhưng tích cực — studio 1–5 người",
    "Doanh thu phổ biến: vài nghìn đến vài chục nghìn USD/tháng",
    "Lĩnh vực: game hyper-casual · utility · sticker/launcher · AI wrapper",
    "Đà Nẵng có lợi thế: chi phí sống thấp · cộng đồng dev đông · internet tốt",
    "Hạn chế: cần Wise/Payoneer để nhận tiền · cần làm quen W-8BEN · marketing là 70% công việc",
  ], LBL);

  // Roadmap
  twoColSlide(pres, "Roadmap thực tế cho sinh viên",
    "Năm 2 — bây giờ",
    [
      "Học Swift + SwiftUI cho thạo",
      "Clone 2–3 app nhỏ (Weather, Notes, Pomodoro)",
      "Bắt đầu đọc Apple Documentation",
    ],
    "Năm 3 → Tốt nghiệp",
    [
      "Năm 3: ship app đầu tay lên App Store (free OK)",
      "Năm 3–4: thử monetize 1 app utility (RevenueCat + ASO)",
      "Tốt nghiệp: chọn 1 trong 3 — outsource lấy kinh nghiệm · product company · full-time indie nếu app đã ≥ chi phí sống",
    ],
    LBL);

  // Công cụ indie
  contentSlide(pres, "Công cụ đặc thù cho indie", [
    "RevenueCat — gắn IAP/Subscription, dashboard analytics",
    "Superwall — A/B test paywall",
    "App Store Connect Analytics — funnel, retention, country breakdown",
    "Apple Search Ads — kênh paid acquisition chính trong App Store",
    "Sensor Tower / data.ai — benchmark đối thủ",
    "Twitter/X — kênh marketing số 1 của indie iOS toàn cầu",
  ], LBL);

  bigQuoteSlide(pres,
    "Cứ ship sớm, ship nhỏ, rồi học từ phản hồi của user thật.",
    "Lời nhắn cuối — Nguyễn Quế Lân",
    LBL);

  // Q&A
  contentSlide(pres, "Trao đổi với sinh viên", [
    "Nếu được phát hành 1 app miễn phí ngay tuần sau, bạn muốn làm app gì?",
    "Bạn nghĩ mình hợp với product hay outsource hơn? Vì sao?",
    "Bạn muốn hỏi gì về App Store, monetization, marketing?",
    { text: "Đặt câu hỏi — đây là phần quý nhất của seminar", bold: true },
  ], LBL);

  // 3 điểm
  contentSlide(pres, "3 điểm cần nhớ — Ca 2", [
    { text: "1. Học iOS không phải để biết Swift — là để build sản phẩm cho user thật", bold: true },
    { text: "2. Ý tưởng > Công nghệ. Flappy Bird là minh chứng.", bold: true },
    { text: "3. Outsource & Product không loại trừ — chọn con đường phù hợp từng giai đoạn", bold: true },
  ], LBL, { bulletSize: 24 });

  // Closing
  const last = pres.addSlide();
  addBaseBg(last, true);
  last.addText("Cảm ơn VKU!", { x: 0.8, y: 2.2, w: 12, h: 1.5, fontFace: FONT, fontSize: 72, bold: true, color: C.white });
  last.addText("Hẹn gặp lại trên App Store", { x: 0.8, y: 3.8, w: 12, h: 0.8, fontFace: FONT, fontSize: 28, color: C.accent, italic: true });
  last.addText("Nguyễn Quế Lân · Technical Leader · VNGalaxy", { x: 0.8, y: 5.3, w: 12, h: 0.6, fontFace: FONT, fontSize: 22, color: C.white, bold: true });
  last.addText("Liên hệ: (bổ sung email/GitHub/LinkedIn trước buổi)", { x: 0.8, y: 5.9, w: 12, h: 0.6, fontFace: FONT, fontSize: 18, color: C.accent });
  addFooter(last, LBL, true);

  return pres.writeFile({ fileName: "seminar-ios-vku-ca2.pptx" });
}

(async () => {
  await buildCa1();
  console.log("WROTE seminar-ios-vku-ca1.pptx");
  await buildCa2();
  console.log("WROTE seminar-ios-vku-ca2.pptx");
})();
