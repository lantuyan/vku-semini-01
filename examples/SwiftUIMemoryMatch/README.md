# SwiftUI Memory Match

Ví dụ app iOS SwiftUI cho seminar: người chơi lật 2 thẻ, nếu 2 hình giống nhau thì giữ mở, nếu sai thì tự úp lại sau một khoảng ngắn.

## Cách mở và chạy bằng Xcode

1. Mở Terminal tại thư mục này:

   ```bash
   cd /Users/quelannguyen/workspace/vku-seminar/new/examples/SwiftUIMemoryMatch
   ```

2. Nếu cần tạo lại project Xcode:

   ```bash
   /opt/homebrew/bin/xcodegen generate
   ```

3. Mở file `SwiftUIMemoryMatch.xcodeproj` bằng Xcode.
4. Chọn simulator iPhone bất kỳ và bấm Run.

Có thể kiểm tra build nhanh bằng lệnh:

```bash
/usr/bin/xcodebuild -project SwiftUIMemoryMatch.xcodeproj -scheme SwiftUIMemoryMatch -destination 'generic/platform=iOS Simulator' build
```

## Các điểm nên giảng dạy

- `SwiftUIMemoryMatchApp`: điểm vào của app SwiftUI với `@main` và `WindowGroup`.
- `ContentView`: màn hình chính, gồm tiêu đề, thống kê, nút chơi lại, chọn độ khó và lưới thẻ.
- `@StateObject`: giữ `GameViewModel` sống ổn định trong vòng đời của `ContentView`.
- `@State`: lưu độ khó đang chọn trong `ContentView`.
- `@Binding`: `DifficultyPicker` nhận binding để view con có thể cập nhật state của view cha.
- `ObservableObject` và `@Published`: `GameViewModel` phát tín hiệu để giao diện vẽ lại khi card, số lượt thay đổi.
- `LazyVGrid`: tạo layout lưới linh hoạt cho danh sách thẻ.
- `Identifiable`: mỗi `MemoryCard` có `id` để dùng trong `ForEach`.
- Animation: `withAnimation`, `transition`, `rotation3DEffect` tạo cảm giác lật thẻ.
- Xử lý bất đồng bộ đơn giản: `Task.sleep` giữ 2 thẻ sai trong 650ms rồi úp lại.

## Cấu trúc file

```text
SwiftUIMemoryMatch/
  project.yml
  README.md
  SwiftUIMemoryMatch.xcodeproj/
  SwiftUIMemoryMatch/
    ContentView.swift
    GameViewModel.swift
    Info.plist
    MemoryCard.swift
    SwiftUIMemoryMatchApp.swift
```
