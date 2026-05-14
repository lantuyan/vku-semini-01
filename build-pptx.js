// Build 2 slide decks for the iOS seminar at VKU.
// Output: seminar-ios-vku-ca1.pptx, seminar-ios-vku-ca2.pptx
// Tone: green (forest), large readable text for a 200-person hall.

const pptxgen = require("pptxgenjs");

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
  slide.addText(`Seminar iOS · VKU · ${deckLabel}`, {
    x: 0.4, y: 7.1, w: 8, h: 0.3,
    fontFace: FONT, fontSize: 11, color: dark ? C.accent : C.inkSoft,
  });
  slide.addText("Nguyễn Quế Lân · VNGalaxy", {
    x: 8.5, y: 7.1, w: 4.5, h: 0.3,
    fontFace: FONT, fontSize: 11, color: dark ? C.accent : C.inkSoft, align: "right",
  });
}

function titleSlide(pres, line1, line2, subtitle, deckLabel) {
  const s = pres.addSlide();
  addBaseBg(s, true);
  s.addText(line1, {
    x: 0.8, y: 1.2, w: 12, h: 1.1,
    fontFace: FONT, fontSize: 56, bold: true, color: C.white,
  });
  s.addText(line2, {
    x: 0.8, y: 2.3, w: 12, h: 1.1,
    fontFace: FONT, fontSize: 56, bold: true, color: C.accent,
  });
  s.addText(subtitle, {
    x: 0.8, y: 3.7, w: 12, h: 0.8,
    fontFace: FONT, fontSize: 24, color: C.white,
  });
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
  contentSlide(pres, "Hệ sinh thái Apple Platforms", [
    "iOS · iPadOS · macOS · watchOS · tvOS · visionOS — cùng SDK, cùng Swift, cùng Xcode",
    "Kỹ năng iOS chuyển dịch được sang các nền tảng khác của Apple",
    "Phân phối duy nhất qua App Store (trừ sideload tại EU từ 2024)",
    flavor === 'enterprise'
      ? "Doanh nghiệp ưu tiên iOS: user chi cao, ít fragmentation, UX chuẩn cao"
      : "Indie ưu tiên iOS: cộng đồng sẵn sàng trả tiền, App Store sân chơi phẳng toàn cầu",
  ], deckLabel);

  contentSlide(pres, "Cấu trúc một app iOS hiện đại", [
    "UI: SwiftUI (mới) hoặc UIKit (truyền thống)",
    "Logic: Swift thuần · Combine · Swift Concurrency (async/await)",
    "Data: Core Data · SwiftData (iOS 17+) · Realm · REST/GraphQL",
    "Platform: APNs · HealthKit · MapKit · AVFoundation · StoreKit",
    flavor === 'enterprise'
      ? "Ví dụ minh hoạ: app banking, super-app — kiến trúc nhiều module"
      : "Ví dụ minh hoạ: Things 3, Bear, Procreate — kiến trúc gọn, 1–3 người maintain",
  ], deckLabel);

  // Lịch sử
  sectionDivider(pres, "LỊCH SỬ NGÔN NGỮ", "Từ Objective-C\nđến Swift", deckLabel);

  twoColSlide(pres, "Objective-C vs Swift",
    "Objective-C (1984–nay)",
    [
      "Brad Cox · Tom Love → NeXT → Apple (1996)",
      "Cú pháp [object method:arg], kế thừa C + Smalltalk",
      "Toàn bộ API gốc macOS/iOS viết bằng Obj-C",
      "Yếu: không type-safe mạnh, không optionals, dễ crash do nil messaging",
      "Còn cần khi maintain code legacy & bridge C/C++",
    ],
    "Swift (2014–nay)",
    [
      "WWDC 2014 · Chris Lattner (cha đẻ LLVM)",
      "An toàn · hiện đại · nhanh · thân thiện cho người mới",
      "Mã nguồn mở từ 2015 → server-side với Vapor",
      "Cột mốc: Swift 3 (2016), Swift 5 ABI (2019), async/await (Swift 5.5, 2021), Swift 6 strict concurrency (2024)",
      flavor === 'enterprise' ? "Khuyến nghị: học Swift là chính, biết đọc Obj-C để không 'mù'" : "Indie 2026: Swift + SwiftUI là mặc định, Obj-C chỉ là kiến thức bonus",
    ],
    deckLabel
  );

  // UIKit vs SwiftUI
  sectionDivider(pres, "FRAMEWORK", "UIKit vs SwiftUI\nhai paradigm, một hệ sinh thái", deckLabel);

  twoColSlide(pres, "UIKit vs SwiftUI",
    "UIKit (2008–nay) — Imperative",
    [
      "Framework gốc của iPhone OS 2",
      "MVC truyền thống · UIView/UIViewController",
      "Layout: frame · Auto Layout · Storyboard",
      "Ổn định, control sâu, mọi tutorial cũ đều dùng",
      "Nhược: code dài, 'massive view controller', khó test",
    ],
    "SwiftUI (2019–nay) — Declarative",
    [
      "Cảm hứng từ React và Flutter",
      "View là struct · state-driven · auto re-render",
      "Tận dụng @State · @Binding · @Observable",
      "Nhanh hơn 2–3× cho UI thường gặp",
      "Nhược: thiếu một số custom sâu, phụ thuộc iOS version",
    ],
    deckLabel
  );

  contentSlide(pres, "Chọn UIKit hay SwiftUI trong thực tế 2026?", [
    flavor === 'enterprise'
      ? "Dự án enterprise mới: ưu tiên SwiftUI cho màn hình mới"
      : "App indie mới: gần như luôn SwiftUI — tốc độ ship là vũ khí",
    "Codebase lớn cũ: vẫn UIKit, thêm SwiftUI cho màn mới (hybrid)",
    "Hai framework nhúng vào nhau: UIViewRepresentable · UIHostingController",
    "Animation phức tạp, custom drawing → UIKit còn lợi thế",
    flavor === 'enterprise' ? { text: "Bài học cho intern: học CẢ HAI — phỏng vấn outsource hay hỏi", bold: true } : { text: "Bài học cho indie: chọn SwiftUI để 1 người làm được nhanh hơn", bold: true },
  ], deckLabel);

  // Công cụ
  sectionDivider(pres, "CÔNG CỤ", "Xcode, SPM,\nInstruments & tour code", deckLabel);

  contentSlide(pres, "Bộ công cụ tối thiểu", [
    "Xcode — IDE chính thức, miễn phí, macOS only",
    "Swift Package Manager (SPM) — thay CocoaPods, tích hợp thẳng Xcode",
    "Instruments — đo memory leak · CPU · năng lượng · network",
    "TestFlight — phân phối beta trước khi lên App Store",
    "Fastlane — automation đóng gói, upload, screenshot",
    "SwiftLint · SwiftFormat — giữ code style trong team",
    flavor === 'indie' ? "RevenueCat — gắn IAP/Subscription nhanh (đặc thù indie)" : "Xcode Cloud — CI/CD chính chủ của Apple từ 2022",
  ], deckLabel, { bulletSize: 20 });

  contentSlide(pres, "Tour code mẫu (live)", [
    "Diễn giả mở Xcode — đi tour project đã chuẩn bị sẵn",
    "Cấu trúc thư mục · Info.plist · Capabilities",
    "Asset Catalog (icon · color · image set)",
    "1 View SwiftUI + 1 ViewModel + 1 Service",
    "Build & chạy trên Simulator → Instruments xem memory",
    { text: "Không gõ code live — chỉ tour cấu trúc", bold: true, dim: true },
  ], deckLabel);
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

  contentSlide(pres, "Outsource là gì?", [
    "Khách hàng (Mỹ · Nhật · EU · Singapore) thuê đội ngũ VN làm phần mềm theo yêu cầu",
    "Mô hình: ODC · Project-based · Staff augmentation",
    "Không sở hữu sản phẩm cuối, không chia user revenue",
    "Đổi lại: dòng tiền ổn định, học codebase enterprise",
  ], LBL);

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

  contentSlide(pres, "Đà Nẵng — bức tranh outsource Mobile", [
    "1 trong 3 trung tâm CNTT lớn nhất VN (cùng HN, HCM)",
    "Cụm doanh nghiệp tập trung quanh Khu CNTT tập trung & Khu CNC",
    "Đa phần phục vụ thị trường Nhật · Mỹ · EU",
    "Mobile (iOS+Android) chiếm tỷ trọng đáng kể sau Web/Backend",
    "iOS: số lượng ít hơn Android nhưng đơn giá CAO HƠN (cung không đủ cầu)",
    { text: "Số liệu tổng hợp từ nguồn công khai (VINASA · TopDev · ITviec · báo chí)", dim: true, size: 16 },
  ], LBL);

  contentSlide(pres, "Một ngày của iOS dev outsource", [
    "Sáng — pull code, đọc Jira, daily standup với client (EN/JP)",
    "Trưa — code feature theo ticket · viết unit test · review PR",
    "Chiều — meeting BA/QA · fix bug · pair-programming",
    "Tối — đôi khi sync meeting 21–22h nếu client US",
    { text: "Tiếng Anh là điều kiện cần — tiếng Nhật là điểm cộng lớn", bold: true },
  ], LBL);

  contentSlide(pres, "Quy trình một dự án outsource", [
    "Pre-sales · proposal · estimate effort",
    "Kick-off · scope · team size · sprint 2 tuần",
    "Development · planning → daily → review → retro",
    "UAT — client test, log bug",
    "Release · TestFlight → App Store · bàn giao tài liệu",
    "Maintenance · vá lỗi · cập nhật theo iOS version mới mỗi năm",
  ], LBL);

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

  contentSlide(pres, "Kỹ năng cần để xin intern/fresher iOS", [
    { text: "Cứng cơ bản", bold: true },
    "Swift cơ bản · UIKit hoặc SwiftUI build được màn đơn giản · Auto Layout · MVC/MVVM",
    { text: "Cứng nâng cao", bold: true },
    "REST API · JSON Codable · UserDefaults/CoreData · async/await",
    { text: "Mềm", bold: true },
    "Tiếng Anh đọc tài liệu Apple (bắt buộc) · giao tiếp cơ bản · tiếng Nhật = bonus",
    { text: "Quy trình", bold: true },
    "Git (branch, PR, conflict) · Jira · hiểu sprint là gì",
    { text: "Tự học là kỹ năng số 1 — iOS đổi mỗi năm tại WWDC", bold: true, size: 22 },
  ], LBL, { bulletSize: 18 });

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
