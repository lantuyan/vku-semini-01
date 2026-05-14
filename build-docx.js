// Build báo cáo seminar iOS - VKU
// Output: seminar-ios-vku-baocao.docx

const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, Header, Footer, PageBreak, TabStopType, TabStopPosition,
} = require('docx');

const GREEN = "2E7D32";        // primary
const GREEN_LIGHT = "C8E6C9";  // table header
const GRAY = "555555";

const border = { style: BorderStyle.SINGLE, size: 4, color: "BDBDBD" };
const borders = { top: border, bottom: border, left: border, right: border };

// Helpers
const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 320 },
  ...opts,
  children: Array.isArray(text)
    ? text
    : [new TextRun({ text, ...(opts.run || {}) })],
});

const H1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 200 },
  children: [new TextRun({ text, bold: true, color: GREEN, size: 36 })],
});

const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 140 },
  children: [new TextRun({ text, bold: true, color: GREEN, size: 30 })],
});

const H3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, bold: true, color: "1B5E20", size: 26 })],
});

const Bul = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 80, line: 300 },
  children: [new TextRun({ text })],
});

const BulBold = (label, text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 80, line: 300 },
  children: [
    new TextRun({ text: label + " ", bold: true }),
    new TextRun({ text }),
  ],
});

// Timeline table
function timelineTable(rows) {
  const widthTotal = 9360;
  const cols = [1700, 1700, 5960];
  const headerRow = new TableRow({
    tableHeader: true,
    children: ["Thời gian", "Thời lượng", "Nội dung"].map((t, i) => new TableCell({
      borders,
      width: { size: cols[i], type: WidthType.DXA },
      shading: { fill: GREEN_LIGHT, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, color: "1B5E20" })] })],
    })),
  });
  const bodyRows = rows.map(r => new TableRow({
    children: r.map((cell, i) => new TableCell({
      borders,
      width: { size: cols[i], type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      children: [new Paragraph({ children: [new TextRun({ text: cell })] })],
    })),
  }));
  return new Table({
    width: { size: widthTotal, type: WidthType.DXA },
    columnWidths: cols,
    rows: [headerRow, ...bodyRows],
  });
}

// =========================================================
// CONTENT
// =========================================================

const children = [];

// ====== Cover ======
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 1200, after: 200 },
    children: [new TextRun({ text: "SEMINAR", bold: true, color: GREEN, size: 56 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({ text: "PHÁT TRIỂN ỨNG DỤNG TRÊN HỆ ĐIỀU HÀNH iOS", bold: true, size: 40 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 600 },
    children: [new TextRun({ text: "Kỹ thuật, công cụ & câu chuyện ngành Mobile App tại Đà Nẵng", italics: true, size: 28, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [new TextRun({ text: "Diễn giả: Nguyễn Quế Lân", bold: true, size: 28 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [new TextRun({ text: "Technical Leader — VNGalaxy", size: 26 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 600 },
    children: [new TextRun({ text: "Đối tượng: Sinh viên năm 2 — Trường Đại học Công nghệ Thông tin và Truyền thông Việt – Hàn (VKU)", size: 24, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Thời lượng: 2 ca × 90 phút (08:00 – 11:00, không nghỉ giữa giờ)", size: 24, color: GRAY })],
  }),
  new Paragraph({ children: [new PageBreak()] }),
);

// ====== Mục lục mô tả ======
children.push(
  H1("1. Giới thiệu chung"),
  P("Tài liệu này là bản báo cáo nội dung chi tiết cho buổi seminar gồm hai ca (mỗi ca 90 phút) tại VKU vào sáng thứ Bảy, chủ đề “Phát triển ứng dụng trên hệ điều hành iOS — kỹ thuật, công cụ và thực trạng ngành Mobile App tại Đà Nẵng”."),
  P("Hai ca chia sẻ chung phần đầu (khoảng 55–60 phút) tập trung vào nội dung kỹ thuật: lịch sử ngôn ngữ, UIKit vs SwiftUI, công cụ Xcode/SPM, dạo nhanh một code mẫu chuẩn bị sẵn. Phần sau (30–35 phút) là phần khác biệt:"),
  BulBold("Ca 1 (08:00 – 09:30):", "Góc nhìn OUTSOURCE — quy trình dự án outsource iOS, client làm việc thế nào, vai trò trong team, hiện trạng outsource Mobile tại Đà Nẵng, kỹ năng để được thuê."),
  BulBold("Ca 2 (09:30 – 11:00):", "Góc nhìn PRODUCT / INDIE — vòng đời sản phẩm trên App Store, monetization, case study Nguyễn Hà Đông – Flappy Bird, indie hacker, cơ hội cho sinh viên Đà Nẵng."),
  P("Mục tiêu cuối: sinh viên năm 2 (đã có nền lập trình mobile) hiểu được bức tranh đầy đủ về nghề lập trình iOS, biết bắt đầu từ đâu, biết hai con đường khác nhau (đi làm thuê outsource hay làm sản phẩm riêng) để chủ động chọn hướng phát triển."),
);

// ====== Mục tiêu ======
children.push(
  H1("2. Mục tiêu học tập"),
  P("Sau seminar, sinh viên có thể:"),
  Bul("Giải thích được dòng thời gian phát triển của lập trình iOS: từ Objective-C → Swift, từ UIKit → SwiftUI và lý do vì sao Apple đi theo hướng đó."),
  Bul("Phân biệt được khi nào dùng UIKit, khi nào dùng SwiftUI, và hiểu khái niệm “interop” giữa hai framework."),
  Bul("Đọc hiểu được cấu trúc một project iOS chuẩn (Xcode workspace, scheme, target, Swift Package Manager, Info.plist, asset catalog)."),
  Bul("Phân biệt hai con đường nghề nghiệp iOS phổ biến tại Việt Nam — outsource và product — và biết các kỹ năng/chỉ số khác nhau mà mỗi con đường yêu cầu."),
  Bul("Có cái nhìn thực tế về thị trường Mobile App ở Đà Nẵng năm 2026, biết các vị trí intern/fresher đang tuyển và mức kỳ vọng."),
  Bul("Lấy được động lực từ câu chuyện Nguyễn Hà Đông & Flappy Bird, hiểu được rằng một sản phẩm nhỏ vẫn có thể tạo ra ảnh hưởng toàn cầu."),
);

// ====== Diễn giả ======
children.push(
  H1("3. Giới thiệu diễn giả"),
  BulBold("Họ tên:", "Nguyễn Quế Lân"),
  BulBold("Chức vụ:", "Technical Leader — VNGalaxy"),
  BulBold("Lĩnh vực:", "Phát triển ứng dụng iOS, kiến trúc ứng dụng mobile, quản lý team kỹ thuật."),
  BulBold("Vai trò trong seminar:", "Chia sẻ kinh nghiệm thực chiến từ vị trí lead — từ chọn công nghệ, review code, đến tuyển dụng và mentor lập trình viên mới."),
  P("Liên hệ / kênh trao đổi sau seminar: (sẽ bổ sung ngay trước buổi — email, GitHub, LinkedIn)."),
);

// ====== Khung tổng ======
children.push(
  H1("4. Khung thời gian tổng (08:00 – 11:00)"),
  timelineTable([
    ["08:00 – 08:05", "5 phút", "Mở đầu, giới thiệu diễn giả, khảo sát nhanh phòng (giơ tay: ai đã viết Swift, ai dùng macOS)."],
    ["08:05 – 09:00", "55 phút", "PHẦN CHUNG — Kỹ thuật: lịch sử, ngôn ngữ, UIKit vs SwiftUI, công cụ, tour code mẫu."],
    ["09:00 – 09:30", "30 phút", "PHẦN RIÊNG CA 1 — Góc nhìn OUTSOURCE + Q&A."],
    ["09:30 – 09:35", "5 phút", "Chuyển ca, sinh viên đổi nhóm (không nghỉ giải lao)."],
    ["09:35 – 10:30", "55 phút", "PHẦN CHUNG (lặp lại cho ca 2) — Kỹ thuật."],
    ["10:30 – 11:00", "30 phút", "PHẦN RIÊNG CA 2 — Góc nhìn PRODUCT / INDIE + Q&A."],
  ]),
  P(""),
  P("Lưu ý: phần “chung” lặp lại về cấu trúc nhưng diễn giả nên đổi ví dụ minh hoạ để tránh nhàm chán nếu có sinh viên dự cả hai ca."),
  new Paragraph({ children: [new PageBreak()] }),
);

// =========================================================
// CA 1
// =========================================================
children.push(
  H1("5. CA 1 — 08:00 đến 09:30 — Góc nhìn OUTSOURCE"),
  H2("5.1. Timeline chi tiết ca 1"),
  timelineTable([
    ["08:00 – 08:05", "5'", "Mở đầu & khảo sát phòng."],
    ["08:05 – 08:20", "15'", "Tổng quan về lập trình ứng dụng di động iOS."],
    ["08:20 – 08:35", "15'", "Lịch sử: Objective-C → Swift."],
    ["08:35 – 08:50", "15'", "UIKit vs SwiftUI: hai paradigm, một hệ sinh thái."],
    ["08:50 – 09:00", "10'", "Công cụ: Xcode, Simulator, SPM, Instruments + tour code mẫu (project trống chuẩn bị sẵn)."],
    ["09:00 – 09:25", "25'", "Thực trạng OUTSOURCE iOS tại Đà Nẵng + kỹ năng để được tuyển."],
    ["09:25 – 09:30", "5'", "Q&A nhanh, chốt lại 3 điểm cần nhớ."],
  ]),

  // 5.2
  H2("5.2. Tổng quan về lập trình ứng dụng iOS (15')"),
  H3("Hệ sinh thái Apple Platforms"),
  Bul("iOS (iPhone), iPadOS (iPad), macOS (Mac), watchOS (Apple Watch), tvOS (Apple TV), visionOS (Vision Pro)."),
  Bul("Cùng dùng chung SDK chính, cùng ngôn ngữ Swift, cùng IDE Xcode → kỹ năng iOS chuyển dịch được sang nền tảng khác."),
  Bul("Phân phối qua App Store; không có “sideload” như Android theo cách thông thường (trừ trường hợp EU bắt buộc từ 2024)."),

  H3("Vì sao iOS được doanh nghiệp ưu tiên?"),
  Bul("Tệp người dùng có sức chi cao → revenue per user tốt hơn Android ở nhiều quốc gia."),
  Bul("Sự đồng nhất phần cứng → ít fragmentation hơn Android."),
  Bul("Tiêu chuẩn UX của Apple cao → ứng dụng iOS thường được đầu tư kỹ hơn."),

  H3("Cấu trúc một app iOS hiện đại"),
  Bul("Tầng UI: SwiftUI (mới) hoặc UIKit (truyền thống)."),
  Bul("Tầng business logic: Swift thuần, các framework như Combine, Swift Concurrency (async/await)."),
  Bul("Tầng dữ liệu: Core Data, SwiftData (mới từ iOS 17), hoặc Realm/SQLite/REST/GraphQL."),
  Bul("Tầng nền tảng: APNs (push), HealthKit, MapKit, AVFoundation, StoreKit, …"),

  // 5.3
  H2("5.3. Lịch sử ngôn ngữ: từ Objective-C đến Swift (15')"),
  H3("Objective-C (1984 – nay)"),
  Bul("Brad Cox, Tom Love phát triển từ đầu thập niên 80, NeXT mua bản quyền, Apple thừa kế khi mua NeXT năm 1996."),
  Bul("Cú pháp “vuông” đặc trưng: [object method:arg] — kết hợp giữa C và Smalltalk."),
  Bul("Toàn bộ macOS/iOS API gốc đều viết bằng Objective-C → đến hôm nay vẫn cần đọc được khi maintain code cũ."),
  Bul("Điểm yếu: cú pháp khó với người mới, không type-safe đủ mạnh, không có optionals, dễ crash do nil messaging im lặng."),

  H3("Swift (2014 – nay)"),
  Bul("Apple công bố tại WWDC 2014, Chris Lattner là kiến trúc sư chính (cũng là người tạo LLVM)."),
  Bul("Mục tiêu: an toàn (type safety, optionals), hiện đại (closures, generics, protocol-oriented), nhanh (compile sang native), thân thiện cho người mới."),
  Bul("Mã nguồn mở từ 2015 → cộng đồng phát triển mạnh, lan sang server-side (Vapor)."),
  Bul("Các phiên bản đáng chú ý: Swift 3 (2016 — đổi cú pháp lớn), Swift 5 (2019 — ABI stability), Swift Concurrency với async/await (2021, Swift 5.5), Swift 6 (2024 — strict concurrency, data race safety mặc định)."),

  H3("Khi nào cần Objective-C trong năm 2026?"),
  Bul("Maintain dự án cũ (legacy)."),
  Bul("Đọc và bridge thư viện C/C++."),
  Bul("Hiểu runtime để debug crash phức tạp (KVO, swizzling, method dispatch)."),
  P("Khuyến nghị cho sinh viên: học Swift là chính, biết đọc Objective-C để không “mù” khi đụng code legacy."),

  // 5.4
  H2("5.4. UIKit vs SwiftUI — hai paradigm, một hệ sinh thái (15')"),
  H3("UIKit (2008 – nay) — paradigm imperative"),
  Bul("Ra mắt cùng iPhone OS 2 — framework gốc của iOS."),
  Bul("Mô hình MVC truyền thống, view kế thừa UIView, controller kế thừa UIViewController."),
  Bul("Layout: bằng frame, Auto Layout (constraint), hoặc Storyboard/XIB."),
  Bul("Ưu điểm: ổn định, đầy đủ, control rất sâu, mọi tutorial cũ đều dùng."),
  Bul("Nhược điểm: code dài, dễ bị “massive view controller”, khó test."),

  H3("SwiftUI (2019 – nay) — paradigm declarative"),
  Bul("Công bố WWDC 2019, lấy cảm hứng từ React và Flutter (declarative UI)."),
  Bul("View là struct, build UI bằng cách mô tả “trông thế nào” thay vì “làm như thế nào”."),
  Bul("State-driven: thay đổi state → UI tự render lại."),
  Bul("Tận dụng tốt Swift modern: result builders, property wrappers (@State, @Binding, @Observable)."),
  Bul("Nhược điểm: vẫn còn thiếu một số custom UI sâu, debug đôi khi khó, phụ thuộc iOS version (nhiều API chỉ có từ iOS 16/17)."),

  H3("So sánh thực tế"),
  Bul("Tốc độ phát triển: SwiftUI nhanh hơn 2–3x cho UI thường gặp (list, form, navigation)."),
  Bul("Kiểm soát: UIKit thắng khi cần animation phức tạp, custom drawing, hoặc tương tác hardware-level."),
  Bul("Tương thích: SwiftUI và UIKit có thể nhúng vào nhau qua UIViewRepresentable / UIHostingController."),
  Bul("Lựa chọn trong thực tế 2026: dự án mới ưu tiên SwiftUI, dự án enterprise lớn vẫn UIKit, hybrid (mỗi màn dùng framework phù hợp) là phổ biến."),

  // 5.5
  H2("5.5. Công cụ phát triển (10') + tour code mẫu"),
  H3("Xcode"),
  Bul("IDE chính thức, miễn phí, chỉ chạy trên macOS."),
  Bul("Tích hợp simulator (mô phỏng iPhone/iPad), debugger, profiler, asset editor, Interface Builder."),
  Bul("Xcode Cloud: CI/CD chính chủ của Apple từ 2022."),

  H3("Swift Package Manager (SPM)"),
  Bul("Trình quản lý thư viện chính thức, thay thế dần CocoaPods và Carthage."),
  Bul("Cú pháp khai báo trong Package.swift, tích hợp thẳng vào Xcode từ Xcode 11."),
  Bul("Ưu điểm: chạy nhanh, không cần Ruby, được Apple đảm bảo tương thích."),

  H3("Các công cụ thường gặp"),
  Bul("Instruments — đo memory leak, CPU, năng lượng, network."),
  Bul("TestFlight — phân phối beta nội bộ trước khi lên App Store."),
  Bul("Fastlane — automation đóng gói, upload, screenshot."),
  Bul("SwiftLint, SwiftFormat — bảo đảm code style trong team."),
  Bul("Git + GitHub/GitLab — version control (giống mọi ngôn ngữ khác)."),

  H3("Tour code mẫu (placeholder)"),
  P("Phần này diễn giả mở Xcode chiếu lên máy chiếu, đi tour qua một project mẫu (sẽ chuẩn bị nội dung cụ thể trước buổi — gợi ý: một app SwiftUI nhỏ kiểu “To-do”, “Weather giả lập” hoặc “Notes”, đủ để minh hoạ:"),
  Bul("Cấu trúc thư mục project."),
  Bul("Tệp Info.plist và Capabilities."),
  Bul("Asset Catalog (icon, color, image set)."),
  Bul("Một View SwiftUI + một ViewModel + một Service đọc dữ liệu."),
  Bul("Build & chạy trên Simulator, mở Instruments xem nhanh memory."),
  P("Không gõ code live, chỉ tour. Code mẫu chốt sau."),

  // 5.6
  H2("5.6. Phần riêng ca 1 — OUTSOURCE iOS tại Đà Nẵng (25')"),
  P("Lưu ý: số liệu trong phần này được tổng hợp từ dữ liệu công khai (báo cáo VINASA, TopDev, ITviec, VietnamWorks, các bài báo cộng đồng) — không nêu tên công ty cụ thể."),

  H3("a) Outsource là gì?"),
  Bul("Khách hàng (thường ở Mỹ, Nhật, EU, Singapore) thuê đội ngũ Việt Nam phát triển hộ phần mềm theo yêu cầu."),
  Bul("Mô hình phổ biến: ODC (Offshore Development Center — dành riêng cho một client), Project-based (làm theo dự án có deadline), Staff augmentation (cho thuê người ghép vào team khách)."),
  Bul("Khác với product: outsource KHÔNG sở hữu sản phẩm cuối, không chia sẻ doanh thu user — đổi lại có dòng tiền ổn định."),

  H3("b) Bức tranh thị trường outsource tại Đà Nẵng"),
  Bul("Đà Nẵng là một trong 3 trung tâm CNTT lớn nhất Việt Nam (cùng Hà Nội và TP.HCM), tốc độ tăng trưởng nhân lực CNTT trung bình 2 chữ số/năm."),
  Bul("Cụm doanh nghiệp outsource tập trung quanh khu công nghệ cao và khu CNTT tập trung — đa phần phục vụ thị trường Nhật, Mỹ, châu Âu."),
  Bul("Mobile (iOS/Android) chiếm tỷ trọng đáng kể trong cơ cấu outsource, sau Web/Backend; nhu cầu iOS thường ít hơn Android về số lượng nhưng đơn giá cao hơn."),
  Bul("Lý do iOS đơn giá cao: nguồn lập trình viên iOS ít hơn, yêu cầu Mac thật, cộng đồng nhỏ hơn → cung không đáp ứng đủ cầu."),

  H3("c) Một ngày của lập trình viên iOS outsource"),
  Bul("Sáng: pull code mới, đọc Jira/Trello/Linear, daily standup với client (thường tiếng Anh hoặc tiếng Nhật)."),
  Bul("Trưa: code feature theo ticket, viết unit test, review PR cho đồng nghiệp."),
  Bul("Chiều: meeting với BA/QA, fix bug từ QA, đôi khi pair-programming."),
  Bul("Tối (nếu client US): đôi khi sync meeting muộn 21–22h."),

  H3("d) Quy trình một dự án outsource điển hình"),
  Bul("Pre-sales: solution architect làm proposal, estimate effort."),
  Bul("Kick-off: thống nhất scope, team size, sprint cycle (thường 2 tuần)."),
  Bul("Development: sprint planning → daily → review → retro, lặp lại."),
  Bul("UAT (User Acceptance Testing): client test, log bug."),
  Bul("Release: build TestFlight → App Store; bàn giao tài liệu kỹ thuật."),
  Bul("Maintenance: hỗ trợ vận hành, vá lỗi, cập nhật theo iOS version mới mỗi năm."),

  H3("e) Vai trò trong team outsource iOS"),
  Bul("Intern / Fresher: học từ senior, làm task nhỏ, viết unit test, fix bug."),
  Bul("Junior (1–2 năm): tự code được feature có hướng dẫn, làm việc trực tiếp với BA."),
  Bul("Senior (3–5 năm): thiết kế module, review code, mentor."),
  Bul("Tech Lead (5+ năm): chốt kiến trúc, làm việc trực tiếp với client architect, chịu trách nhiệm chất lượng kỹ thuật toàn dự án."),
  Bul("BrSE / Comtor (Nhật): cầu nối ngôn ngữ-văn hoá-kỹ thuật, đặc thù của thị trường outsource Việt-Nhật."),

  H3("f) Kỹ năng cần có để xin việc outsource iOS (vai trò Intern/Fresher)"),
  Bul("Cứng: Swift cơ bản, hiểu UIKit hoặc SwiftUI ở mức build được màn hình đơn giản, biết Auto Layout, hiểu MVC/MVVM."),
  Bul("Cứng nâng cao: gọi REST API, parse JSON với Codable, lưu local với UserDefaults/CoreData."),
  Bul("Mềm: tiếng Anh đọc hiểu tài liệu Apple (bắt buộc), giao tiếp cơ bản; tiếng Nhật là điểm cộng lớn."),
  Bul("Quy trình: biết Git (branch, PR, conflict), hiểu sprint là gì, dùng được Jira."),
  Bul("Tinh thần: tự học là kỹ năng số 1 — iOS đổi liên tục, mỗi năm WWDC ra hàng trăm API mới."),

  H3("g) Ưu/nhược của con đường outsource"),
  Bul("Ưu: dễ tìm việc đầu sự nghiệp, được tiếp xúc khách hàng quốc tế, cải thiện tiếng Anh nhanh, mức lương khởi điểm tốt."),
  Bul("Nhược: ít “chủ quyền sản phẩm” (mình code nhưng app thuộc về khách), đôi khi làm cùng một mảng lặp lại, ít cơ hội học design/PM."),
  Bul("Lựa chọn nghề: outsource phù hợp khi muốn ổn định, học nhanh codebase lớn, build kỷ luật engineering chuyên nghiệp."),

  H3("h) Trao đổi với sinh viên (5–10 phút cuối)"),
  Bul("Câu hỏi gợi mở: “Bạn nào đã thử apply intern iOS? Khó khăn lớn nhất khi đi phỏng vấn là gì?”"),
  Bul("Sinh viên đặt câu hỏi — diễn giả chia sẻ kinh nghiệm thực tế từ vị trí lead phỏng vấn."),

  new Paragraph({ children: [new PageBreak()] }),
);

// =========================================================
// CA 2
// =========================================================
children.push(
  H1("6. CA 2 — 09:35 đến 11:00 — Góc nhìn PRODUCT / INDIE"),
  H2("6.1. Timeline chi tiết ca 2"),
  timelineTable([
    ["09:35 – 09:40", "5'", "Mở đầu ca 2, đổi tone — “bây giờ ta nói về làm ra sản phẩm của chính mình”."],
    ["09:40 – 09:55", "15'", "Tổng quan iOS (đi nhanh hơn ca 1, dùng ví dụ khác)."],
    ["09:55 – 10:10", "15'", "Lịch sử Objective-C → Swift (kể từ góc nhìn người làm product)."],
    ["10:10 – 10:25", "15'", "UIKit vs SwiftUI — chọn cái nào cho indie/product?"],
    ["10:25 – 10:30", "5'", "Công cụ + tour code mẫu (nhấn vào StoreKit, Analytics, App Store metadata)."],
    ["10:30 – 10:55", "25'", "PRODUCT / INDIE — case Flappy Bird & con đường indie tại Đà Nẵng."],
    ["10:55 – 11:00", "5'", "Q&A & lời nhắn cuối."],
  ]),

  // 6.2 - reuse condensed
  H2("6.2. Phần kỹ thuật chung (ca 2) — cách trình bày khác"),
  P("Nội dung kỹ thuật về bản chất giống ca 1 (lịch sử, UIKit vs SwiftUI, công cụ) — nhưng để tránh lặp với sinh viên dự cả hai ca, diễn giả nên:"),
  Bul("Đổi câu chuyện minh hoạ: ca 1 minh hoạ bằng app enterprise (HSBC, banking, Zalo Pay), ca 2 minh hoạ bằng app indie nổi tiếng (Flappy Bird, Things 3, Bear, Procreate)."),
  Bul("Tốc độ nhanh hơn (45–50 phút thay vì 55) để dành thêm thời gian cho phần product."),
  Bul("Khi nói UIKit vs SwiftUI, lần này tập trung vào góc “time-to-market” và “maintain bởi 1 người” — yếu tố cốt lõi của indie."),
  Bul("Khi nói công cụ, dành thêm thời gian cho StoreKit (in-app purchase, subscription), App Store Connect metadata, Apple Search Ads, App Analytics."),
  P("Diễn giả tự chủ động lựa các phần để rút gọn — báo cáo này không liệt kê lại toàn bộ vì nội dung core không đổi."),

  // 6.3 PRODUCT
  H2("6.3. Phần riêng ca 2 — PRODUCT / INDIE (25')"),

  H3("a) Product vs Outsource — khác biệt cốt lõi"),
  Bul("Outsource: bán giờ công cho khách → thu nhập tuyến tính theo số người × số giờ."),
  Bul("Product: bán sản phẩm cho người dùng cuối → thu nhập có thể scale phi tuyến (1 người làm, triệu người mua)."),
  Bul("Outsource thắng về độ chắc chắn; product thắng về upside."),
  Bul("Hai mô hình KHÔNG xung đột — nhiều studio làm outsource ban ngày, làm side-project product ngoài giờ."),

  H3("b) Vòng đời một product iOS"),
  Bul("Idea → Validate (hỏi user thật, không hỏi bạn bè): có cần thật không?"),
  Bul("MVP (minimum viable product): bản nhỏ nhất có giá trị, build trong 2–6 tuần."),
  Bul("Public beta qua TestFlight."),
  Bul("Launch trên App Store: chuẩn bị metadata (tiêu đề, mô tả, screenshot, video preview), keyword cho ASO (App Store Optimization)."),
  Bul("Growth: track retention D1/D7/D30, ARPU, conversion từ free → paid."),
  Bul("Iterate: ra version mới mỗi 2–4 tuần."),

  H3("c) Monetization — kiếm tiền từ app như thế nào?"),
  Bul("Paid up-front: trả 1 lần khi tải. Hiếm dần, chỉ còn hiệu quả với app tool chất lượng cao (Things 3, Procreate)."),
  Bul("Freemium + IAP (in-app purchase): tải miễn phí, mở khoá tính năng cao cấp."),
  Bul("Subscription: thu phí định kỳ — mô hình lớn nhất 2020s (Duolingo, Notion, Calm…)."),
  Bul("Quảng cáo: dùng Google AdMob, AppLovin, IronSource — phù hợp game casual; chính là cách Flappy Bird kiếm tiền."),
  Bul("Hybrid: kết hợp — free + ads + IAP để remove ads."),
  Bul("Apple cắt 30% (15% với subscription năm 2 hoặc app dưới 1 triệu USD/năm theo Small Business Program)."),

  H3("d) Case study — Nguyễn Hà Đông & Flappy Bird"),
  Bul("Tác giả: Nguyễn Hà Đông (Đông Nguyễn), lập trình viên người Hà Nội, sinh năm 1985."),
  Bul("Studio: dotGEARS — gần như chỉ có một mình."),
  Bul("Flappy Bird ra mắt: tháng 5/2013, ban đầu không nổi."),
  Bul("Viral bùng nổ: cuối 2013 – đầu 2014, lan trên YouTube/Twitter nhờ độ khó “gây nghiện gây bực”."),
  Bul("Đỉnh điểm: đầu tháng 2/2014, đứng số 1 App Store ở hơn 100 quốc gia, doanh thu công khai ước tính khoảng 50.000 USD/ngày từ quảng cáo banner."),
  Bul("Đông gỡ game khỏi App Store ngày 10/02/2014 vì áp lực tâm lý — quyết định gây sốc toàn cầu, càng làm game thành huyền thoại."),
  P("Vì sao Flappy Bird đáng nhắc với sinh viên Việt Nam 2026:"),
  Bul("Chứng minh: từ Việt Nam, một người làm, một app rất đơn giản về kỹ thuật vẫn có thể đứng số 1 App Store thế giới."),
  Bul("Phá vỡ niềm tin “muốn nổi phải ở Silicon Valley” — App Store là sân chơi phẳng, chỉ cần app đến được tay user."),
  Bul("Bài học kỹ thuật: ý tưởng > công nghệ. Flappy Bird dùng kỹ thuật rất cơ bản, gameplay 1 tap, art retro."),
  Bul("Bài học sản phẩm: yếu tố “gây nghiện” và viral organic quan trọng hơn tính năng nhiều."),
  Bul("Bài học cá nhân: thành công vượt kỳ vọng có cái giá của nó — sức khoẻ tâm lý, đời sống riêng tư là thật. Đông chọn rút lui, đó là một quyết định cần được tôn trọng."),

  H3("e) Indie iOS tại Việt Nam và Đà Nẵng"),
  Bul("Việt Nam có cộng đồng indie/solo developer nhỏ nhưng tích cực — nhiều studio 1–5 người, doanh thu vài nghìn đến vài chục nghìn USD/tháng từ App Store."),
  Bul("Lĩnh vực phổ biến: game hyper-casual, app utility (tool nhỏ, dịch, đếm calo, habit tracker), app sticker/launcher, app AI wrapper (GPT, image generation)."),
  Bul("Đà Nẵng có lợi thế: chi phí sống thấp, cộng đồng dev đông, internet tốt — môi trường tốt để bắt đầu indie part-time."),
  Bul("Hạn chế thực tế: thu thanh toán từ Apple cần tài khoản US bank/Wise/Payoneer; cần nộp thuế (W-8BEN, tax treaty); cần build thói quen marketing — code chỉ là 30% công việc."),

  H3("f) Một roadmap thực tế cho sinh viên muốn đi đường product"),
  Bul("Năm 2 (hiện tại): học Swift và SwiftUI cho thạo, làm 2–3 app nhỏ để clone (Weather, Notes, Pomodoro)."),
  Bul("Năm 3: ra app đầu tay trên App Store — kể cả miễn phí, mục đích là đi hết quy trình submit/review/release."),
  Bul("Năm 3–4: thử monetize bằng một app utility nhỏ (Subscription qua RevenueCat) — học analytics, ASO, marketing cơ bản."),
  Bul("Tốt nghiệp: chọn 1 trong 3 — (1) full-time outsource lấy kinh nghiệm 2–3 năm rồi quay lại indie, (2) full-time product company, (3) full-time indie nếu app đã có doanh thu ≥ chi phí sống."),
  Bul("Lưu ý: outsource và indie không loại trừ nhau. Rất nhiều indie giỏi đi lên từ outsource."),

  H3("g) Trao đổi & lời nhắn cuối (5')"),
  Bul("Câu hỏi gợi mở: “Nếu được phát hành 1 app miễn phí ngay tuần sau, bạn muốn làm app gì?”"),
  Bul("Lời nhắn: học iOS không phải để biết Swift — là để biết build sản phẩm cho người dùng thật. Cứ ship sớm, ship nhỏ, rồi học từ phản hồi."),
  Bul("Mở kênh contact: GitHub/email diễn giả để sinh viên hỏi tiếp."),

  new Paragraph({ children: [new PageBreak()] }),
);

// =========================================================
// Phụ lục
// =========================================================
children.push(
  H1("7. Phụ lục"),
  H2("7.1. Tài liệu tham khảo sinh viên nên đọc"),
  Bul("Apple Developer Documentation — developer.apple.com/documentation"),
  Bul("Swift.org — swift.org/documentation"),
  Bul("Hacking with Swift (Paul Hudson) — hackingwithswift.com"),
  Bul("Stanford CS193p — “Developing Applications for iOS using SwiftUI” (miễn phí trên YouTube)."),
  Bul("WWDC Sessions — developer.apple.com/videos (xem WWDC 2024–2026 cho cập nhật mới nhất)."),
  Bul("Cộng đồng: r/iOSProgramming, Swift Forums, Hacking with Swift+ community."),

  H2("7.2. Công cụ khuyến nghị dùng từ năm 2"),
  Bul("Xcode (luôn cập nhật bản mới nhất từ Mac App Store)."),
  Bul("SF Symbols app — kho icon chính thức của Apple."),
  Bul("Proxyman — đọc HTTP traffic khi debug API."),
  Bul("RevenueCat — gắn IAP/Subscription nhanh (cho ai làm product)."),
  Bul("TestFlight — chia sẻ app cho bạn bè test."),

  H2("7.3. Checklist chuẩn bị của diễn giả trước buổi"),
  Bul("Sạc đầy MacBook + mang adapter HDMI."),
  Bul("Tải sẵn project code mẫu (sẽ chốt nội dung sau), build thử trên Simulator."),
  Bul("Mang sẵn iPhone đã cài app demo phòng trường hợp Simulator lỗi."),
  Bul("Phông chữ slide cỡ tối thiểu 24pt — phòng 200 người."),
  Bul("Chuẩn bị 3 câu hỏi “mồi” cho phần Q&A nếu sinh viên ngại đặt câu hỏi."),

  H2("7.4. Ghi chú về số liệu"),
  P("Mọi số liệu thị trường (số lượng công ty, mức lương, thị phần) trong báo cáo này được tổng hợp từ các nguồn công khai (VINASA, TopDev, ITviec, VietnamWorks, báo Tuổi Trẻ, VnExpress, Sở TT-TT Đà Nẵng) tính đến đầu năm 2026. Sinh viên cần kiểm chứng lại trước khi sử dụng cho mục đích khác. Báo cáo cố tình KHÔNG nêu tên cụ thể công ty outsource để tránh đánh giá không cần thiết."),
);

// =========================================================
// Build document
// =========================================================

const doc = new Document({
  creator: "Nguyễn Quế Lân",
  title: "Seminar iOS — VKU",
  styles: {
    default: { document: { run: { font: "Calibri", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, color: GREEN, font: "Calibri" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, color: GREEN, font: "Calibri" },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: "1B5E20", font: "Calibri" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "Seminar iOS — VKU — Nguyễn Quế Lân", color: GRAY, size: 18 })],
      })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Trang ", color: GRAY, size: 18 }),
          new TextRun({ children: [PageNumber.CURRENT], color: GRAY, size: 18 }),
        ],
      })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("seminar-ios-vku-baocao.docx", buf);
  console.log("WROTE seminar-ios-vku-baocao.docx (" + buf.length + " bytes)");
});
