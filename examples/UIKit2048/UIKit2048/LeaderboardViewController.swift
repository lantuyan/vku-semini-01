import UIKit

final class LeaderboardViewController: UIViewController {
    override func viewDidLoad() {
        super.viewDidLoad()
        title = "Bảng xếp hạng"

        // Màn này dùng XIB để minh họa cách tách layout khỏi code Swift.
        navigationItem.rightBarButtonItem = UIBarButtonItem(
            title: "Đóng",
            style: .plain,
            target: self,
            action: #selector(closeTapped)
        )
    }

    @objc private func closeTapped() {
        navigationController?.popViewController(animated: true)
    }
}
