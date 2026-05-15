# UIKit2048

Ví dụ app iOS UIKit nhỏ để minh họa cách dùng `UIViewController`, `UILabel`, `UIButton`, `UIStackView`, Auto Layout, gesture recognizer và model state đơn giản qua game 2048.

## Cách mở và chạy

1. Mở Terminal tại thư mục này:

   ```sh
   cd /Users/quelannguyen/workspace/vku-seminar/new/examples/UIKit2048
   ```

2. Nếu cần tạo lại project Xcode:

   ```sh
   /opt/homebrew/bin/xcodegen generate
   ```

3. Mở `UIKit2048.xcodeproj` bằng Xcode.
4. Chọn một iPhone Simulator và bấm Run.

Có thể build nhanh bằng lệnh:

```sh
/usr/bin/xcodebuild -project UIKit2048.xcodeproj -scheme UIKit2048 -destination 'generic/platform=iOS Simulator' build
```

## Điểm giảng dạy chính

- `AppDelegate` và `SceneDelegate`: điểm vào app UIKit không dùng storyboard.
- `GameViewController`: tạo giao diện bằng code, gồm label, button, stack view, constraint và swipe gesture.
- `LeaderboardViewController`: màn bảng xếp hạng mockup, dùng XIB để minh họa cách dựng layout bằng Interface Builder.
- `GameModel`: tách logic game khỏi UI để sinh viên thấy rõ state, move, merge, score và undo.
- `UINavigationController`: điều hướng từ màn game sang màn bảng xếp hạng bằng `pushViewController`.
- `UIStackView`: dùng stack dọc/ngang để lập grid 4x4 mà không cần tính frame thủ công.
- Auto Layout: board có tỉ lệ vuông bằng constraint `heightAnchor == widthAnchor`.
- `UISwipeGestureRecognizer`: lắng nghe 4 hướng vuốt và gọi model.
- Màu tile: mapping giá trị 2, 4, 8... sang màu nền riêng, giống tinh thần game 2048.
- UIKit animation: dùng `UIView.animate` để tile mới phóng to dần, tile merge bật nhẹ, score nhấn mạnh khi thay đổi và `CAKeyframeAnimation` để rung board khi vuốt không hợp lệ.

## Tính năng

- New game với 2 tile ban đầu.
- Vuốt 4 hướng để di chuyển và merge tile.
- Tính score theo tổng giá trị tile được merge.
- Best score trong phiên chạy app.
- Undo một bước gần nhất.
- Thông báo khi đạt 2048 hoặc hết nước đi.
- Animation cho tile mới, tile merge, score và nước đi không hợp lệ.
- Nút `BXH` mở màn bảng xếp hạng mockup từ file `LeaderboardViewController.xib`.

## Gợi ý demo trên lớp

1. Cho sinh viên xem `GameModel.move(_:)` trước để hiểu logic không phụ thuộc UIKit.
2. Sang `GameViewController.buildInterface()` để thấy UI tạo bằng component UIKit.
3. Sửa màu trong `tileColor(for:)` hoặc spacing trong `buildBoard()` và chạy lại để thay đổi giao diện.
4. Đặt breakpoint ở `handleSwipe(_:)` để quan sát gesture chuyển thành lệnh `.up`, `.down`, `.left`, `.right`.
5. Cho sinh viên sửa tham số `duration`, `delay`, `usingSpringWithDamping` trong `animateNewTile(_:)` và `animateMergedTile(_:)` để thấy tác động của animation.
6. Mở `LeaderboardViewController.xib` trong Xcode để trình bày File's Owner, root view, Auto Layout và mối liên hệ giữa XIB với `LeaderboardViewController.swift`.
