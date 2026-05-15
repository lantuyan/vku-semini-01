import UIKit

final class GameViewController: UIViewController {
    private var game = GameModel()
    private var tileLabels: [[UILabel]] = []
    private var isAnimatingMove = false

    private let scoreValueLabel = UILabel()
    private let bestValueLabel = UILabel()
    private let messageLabel = UILabel()
    private let undoButton = UIButton(type: .system)

    private let boardView: UIView = {
        let view = UIView()
        view.backgroundColor = UIColor(red: 0.73, green: 0.68, blue: 0.63, alpha: 1)
        view.layer.cornerRadius = 12
        view.translatesAutoresizingMaskIntoConstraints = false
        return view
    }()

    override func viewDidLoad() {
        super.viewDidLoad()
        title = "2048"
        view.backgroundColor = UIColor(red: 0.98, green: 0.96, blue: 0.90, alpha: 1)

        buildInterface()
        addSwipeGestures()
        updateInterface(animated: false)
    }

    private func buildInterface() {
        let titleLabel = UILabel()
        titleLabel.text = "2048"
        titleLabel.textColor = UIColor(red: 0.47, green: 0.43, blue: 0.39, alpha: 1)
        titleLabel.font = .systemFont(ofSize: 48, weight: .heavy)
        titleLabel.adjustsFontSizeToFitWidth = true
        titleLabel.minimumScaleFactor = 0.75
        titleLabel.setContentCompressionResistancePriority(.defaultLow, for: .horizontal)

        let scoreBox = makeScoreBox(title: "SCORE", valueLabel: scoreValueLabel)
        let bestBox = makeScoreBox(title: "BEST", valueLabel: bestValueLabel)
        let scoreStack = UIStackView(arrangedSubviews: [scoreBox, bestBox])
        scoreStack.axis = .horizontal
        scoreStack.spacing = 8
        scoreStack.distribution = .fillEqually

        let headerStack = UIStackView(arrangedSubviews: [titleLabel, scoreStack])
        headerStack.axis = .horizontal
        headerStack.alignment = .center
        headerStack.spacing = 16

        let newButton = makeCommandButton(title: "New", action: #selector(newGameTapped))
        undoButton.setTitle("Undo", for: .normal)
        undoButton.addTarget(self, action: #selector(undoTapped), for: .touchUpInside)
        styleCommandButton(undoButton)

        let leaderboardButton = makeCommandButton(title: "BXH", action: #selector(leaderboardTapped))

        let buttonStack = UIStackView(arrangedSubviews: [newButton, undoButton, leaderboardButton])
        buttonStack.axis = .horizontal
        buttonStack.spacing = 10
        buttonStack.distribution = .fillEqually

        messageLabel.textColor = UIColor(red: 0.47, green: 0.43, blue: 0.39, alpha: 1)
        messageLabel.font = .systemFont(ofSize: 16, weight: .semibold)
        messageLabel.textAlignment = .center
        messageLabel.numberOfLines = 0

        let mainStack = UIStackView(arrangedSubviews: [headerStack, buttonStack, boardView, messageLabel])
        mainStack.axis = .vertical
        mainStack.spacing = 18
        mainStack.translatesAutoresizingMaskIntoConstraints = false

        view.addSubview(mainStack)
        NSLayoutConstraint.activate([
            mainStack.leadingAnchor.constraint(equalTo: view.safeAreaLayoutGuide.leadingAnchor, constant: 20),
            mainStack.trailingAnchor.constraint(equalTo: view.safeAreaLayoutGuide.trailingAnchor, constant: -20),
            mainStack.centerYAnchor.constraint(equalTo: view.safeAreaLayoutGuide.centerYAnchor),
            boardView.heightAnchor.constraint(equalTo: boardView.widthAnchor)
        ])

        buildBoard()
    }

    private func buildBoard() {
        let boardStack = UIStackView()
        boardStack.axis = .vertical
        boardStack.spacing = 10
        boardStack.distribution = .fillEqually
        boardStack.translatesAutoresizingMaskIntoConstraints = false

        boardView.addSubview(boardStack)
        NSLayoutConstraint.activate([
            boardStack.leadingAnchor.constraint(equalTo: boardView.leadingAnchor, constant: 10),
            boardStack.trailingAnchor.constraint(equalTo: boardView.trailingAnchor, constant: -10),
            boardStack.topAnchor.constraint(equalTo: boardView.topAnchor, constant: 10),
            boardStack.bottomAnchor.constraint(equalTo: boardView.bottomAnchor, constant: -10)
        ])

        for _ in 0..<4 {
            var rowLabels: [UILabel] = []
            let rowStack = UIStackView()
            rowStack.axis = .horizontal
            rowStack.spacing = 10
            rowStack.distribution = .fillEqually

            for _ in 0..<4 {
                let tileLabel = UILabel()
                tileLabel.textAlignment = .center
                tileLabel.adjustsFontSizeToFitWidth = true
                tileLabel.minimumScaleFactor = 0.5
                tileLabel.font = .systemFont(ofSize: 32, weight: .heavy)
                tileLabel.layer.cornerRadius = 8
                tileLabel.layer.masksToBounds = true
                rowLabels.append(tileLabel)
                rowStack.addArrangedSubview(tileLabel)
            }

            tileLabels.append(rowLabels)
            boardStack.addArrangedSubview(rowStack)
        }
    }

    private func addSwipeGestures() {
        // Gesture recognizer biến thao tác vuốt trên màn hình thành lệnh cho model.
        let gestures: [(UISwipeGestureRecognizer.Direction, MoveDirection)] = [
            (.up, .up),
            (.down, .down),
            (.left, .left),
            (.right, .right)
        ]

        for (swipeDirection, moveDirection) in gestures {
            let gesture = UISwipeGestureRecognizer(target: self, action: #selector(handleSwipe(_:)))
            gesture.direction = swipeDirection
            gesture.name = String(describing: moveDirection)
            view.addGestureRecognizer(gesture)
        }
    }

    @objc private func handleSwipe(_ gesture: UISwipeGestureRecognizer) {
        guard !isAnimatingMove else { return }

        let direction: MoveDirection
        switch gesture.direction {
        case .up:
            direction = .up
        case .down:
            direction = .down
        case .left:
            direction = .left
        case .right:
            direction = .right
        default:
            return
        }

        let previousBoard = game.board
        let previousScore = game.score
        if game.move(direction) {
            updateInterface(animated: true, previousBoard: previousBoard, previousScore: previousScore)
        } else {
            shakeBoard()
        }
    }

    @objc private func newGameTapped() {
        let emptyBoard = Array(repeating: Array(repeating: 0, count: 4), count: 4)
        game.startNewGame()
        updateInterface(animated: true, previousBoard: emptyBoard, previousScore: 0)
    }

    @objc private func undoTapped() {
        let previousBoard = game.board
        let previousScore = game.score
        game.undo()
        updateInterface(animated: true, previousBoard: previousBoard, previousScore: previousScore)
    }

    @objc private func leaderboardTapped() {
        let viewController = LeaderboardViewController(nibName: "LeaderboardViewController", bundle: nil)
        navigationController?.pushViewController(viewController, animated: true)
    }

    private func updateInterface(
        animated: Bool,
        previousBoard: [[Int]]? = nil,
        previousScore: Int? = nil
    ) {
        scoreValueLabel.text = "\(game.score)"
        bestValueLabel.text = "\(game.bestScore)"
        undoButton.isEnabled = game.canUndo
        undoButton.alpha = game.canUndo ? 1 : 0.45

        for row in 0..<4 {
            for column in 0..<4 {
                let value = game.board[row][column]
                let label = tileLabels[row][column]
                label.text = value == 0 ? "" : "\(value)"
                label.backgroundColor = tileColor(for: value)
                label.textColor = value <= 4 ? UIColor(red: 0.47, green: 0.43, blue: 0.39, alpha: 1) : .white
            }
        }

        if game.isGameOver {
            messageLabel.text = "Hết nước đi. Bấm New để chơi lại."
        } else if game.has2048Tile {
            messageLabel.text = "Đã tạo được 2048. Có thể tiếp tục chơi."
        } else {
            messageLabel.text = "Vuốt lên, xuống, trái, phải để di chuyển các ô."
        }

        guard animated, let previousBoard else { return }
        animateTileChanges(from: previousBoard)

        if let previousScore, previousScore != game.score {
            animateScoreChange()
        }
    }

    private func animateTileChanges(from previousBoard: [[Int]]) {
        isAnimatingMove = true

        for row in 0..<4 {
            for column in 0..<4 {
                let oldValue = previousBoard[row][column]
                let newValue = game.board[row][column]
                let label = tileLabels[row][column]

                if oldValue == 0, newValue != 0 {
                    animateNewTile(label)
                } else if oldValue != 0, newValue != 0, oldValue != newValue {
                    animateMergedTile(label)
                }
            }
        }

        UIView.animate(withDuration: 0.05, delay: 0.22, options: [.allowUserInteraction]) {
            self.boardView.layoutIfNeeded()
        } completion: { _ in
            self.isAnimatingMove = false
        }
    }

    private func animateNewTile(_ label: UILabel) {
        label.alpha = 0
        label.transform = CGAffineTransform(scaleX: 0.2, y: 0.2)

        UIView.animate(
            withDuration: 0.24,
            delay: 0.06,
            usingSpringWithDamping: 0.72,
            initialSpringVelocity: 0.3,
            options: [.curveEaseOut, .allowUserInteraction]
        ) {
            label.alpha = 1
            label.transform = .identity
        }
    }

    private func animateMergedTile(_ label: UILabel) {
        label.transform = CGAffineTransform(scaleX: 0.86, y: 0.86)

        UIView.animate(
            withDuration: 0.18,
            delay: 0.02,
            usingSpringWithDamping: 0.58,
            initialSpringVelocity: 0.8,
            options: [.curveEaseOut, .allowUserInteraction]
        ) {
            label.transform = CGAffineTransform(scaleX: 1.12, y: 1.12)
        } completion: { _ in
            UIView.animate(withDuration: 0.12) {
                label.transform = .identity
            }
        }
    }

    private func animateScoreChange() {
        scoreValueLabel.transform = CGAffineTransform(scaleX: 1.16, y: 1.16)
        UIView.animate(withDuration: 0.18, delay: 0, options: [.curveEaseOut]) {
            self.scoreValueLabel.transform = .identity
        }
    }

    private func shakeBoard() {
        let animation = CAKeyframeAnimation(keyPath: "transform.translation.x")
        animation.values = [0, -10, 8, -6, 4, 0]
        animation.duration = 0.22
        animation.timingFunction = CAMediaTimingFunction(name: .easeOut)
        boardView.layer.add(animation, forKey: "invalidMoveShake")
    }

    private func makeScoreBox(title: String, valueLabel: UILabel) -> UIView {
        let titleLabel = UILabel()
        titleLabel.text = title
        titleLabel.textColor = UIColor(red: 0.93, green: 0.89, blue: 0.82, alpha: 1)
        titleLabel.font = .systemFont(ofSize: 12, weight: .bold)
        titleLabel.textAlignment = .center

        valueLabel.textColor = .white
        valueLabel.font = .systemFont(ofSize: 20, weight: .bold)
        valueLabel.textAlignment = .center

        let stack = UIStackView(arrangedSubviews: [titleLabel, valueLabel])
        stack.axis = .vertical
        stack.spacing = 2
        stack.alignment = .fill
        stack.translatesAutoresizingMaskIntoConstraints = false

        let container = UIView()
        container.backgroundColor = UIColor(red: 0.73, green: 0.68, blue: 0.63, alpha: 1)
        container.layer.cornerRadius = 8
        container.addSubview(stack)

        NSLayoutConstraint.activate([
            stack.leadingAnchor.constraint(equalTo: container.leadingAnchor, constant: 12),
            stack.trailingAnchor.constraint(equalTo: container.trailingAnchor, constant: -12),
            stack.topAnchor.constraint(equalTo: container.topAnchor, constant: 8),
            stack.bottomAnchor.constraint(equalTo: container.bottomAnchor, constant: -8),
            container.widthAnchor.constraint(greaterThanOrEqualToConstant: 78)
        ])

        return container
    }

    private func makeCommandButton(title: String, action: Selector) -> UIButton {
        let button = UIButton(type: .system)
        button.setTitle(title, for: .normal)
        button.addTarget(self, action: action, for: .touchUpInside)
        styleCommandButton(button)
        return button
    }

    private func styleCommandButton(_ button: UIButton) {
        button.backgroundColor = UIColor(red: 0.56, green: 0.48, blue: 0.40, alpha: 1)
        button.tintColor = .white
        button.titleLabel?.font = .systemFont(ofSize: 17, weight: .bold)
        button.layer.cornerRadius = 8
        button.heightAnchor.constraint(equalToConstant: 48).isActive = true
    }

    private func tileColor(for value: Int) -> UIColor {
        switch value {
        case 0:
            return UIColor(red: 0.80, green: 0.76, blue: 0.70, alpha: 1)
        case 2:
            return UIColor(red: 0.93, green: 0.89, blue: 0.82, alpha: 1)
        case 4:
            return UIColor(red: 0.93, green: 0.87, blue: 0.76, alpha: 1)
        case 8:
            return UIColor(red: 0.95, green: 0.69, blue: 0.47, alpha: 1)
        case 16:
            return UIColor(red: 0.96, green: 0.58, blue: 0.39, alpha: 1)
        case 32:
            return UIColor(red: 0.96, green: 0.48, blue: 0.37, alpha: 1)
        case 64:
            return UIColor(red: 0.93, green: 0.35, blue: 0.23, alpha: 1)
        case 128:
            return UIColor(red: 0.93, green: 0.81, blue: 0.45, alpha: 1)
        case 256:
            return UIColor(red: 0.93, green: 0.80, blue: 0.38, alpha: 1)
        case 512:
            return UIColor(red: 0.93, green: 0.78, blue: 0.31, alpha: 1)
        case 1024:
            return UIColor(red: 0.93, green: 0.76, blue: 0.25, alpha: 1)
        default:
            return UIColor(red: 0.24, green: 0.23, blue: 0.20, alpha: 1)
        }
    }
}
