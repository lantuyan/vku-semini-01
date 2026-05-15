import Foundation

enum MoveDirection {
    case up
    case down
    case left
    case right
}

struct GameSnapshot {
    let board: [[Int]]
    let score: Int
}

struct GameModel {
    private(set) var board: [[Int]]
    private(set) var score: Int
    private(set) var bestScore: Int
    private var previousSnapshot: GameSnapshot?

    init() {
        board = Array(repeating: Array(repeating: 0, count: 4), count: 4)
        score = 0
        bestScore = 0
        startNewGame()
    }

    mutating func startNewGame() {
        board = Array(repeating: Array(repeating: 0, count: 4), count: 4)
        score = 0
        previousSnapshot = nil
        addRandomTile()
        addRandomTile()
    }

    mutating func move(_ direction: MoveDirection) -> Bool {
        let beforeMove = GameSnapshot(board: board, score: score)
        var newBoard = board
        var gainedScore = 0

        for index in 0..<4 {
            let originalLine = line(at: index, direction: direction)
            let result = mergedLine(from: originalLine)
            gainedScore += result.gainedScore
            write(result.values, to: &newBoard, at: index, direction: direction)
        }

        guard newBoard != board else { return false }

        previousSnapshot = beforeMove
        board = newBoard
        score += gainedScore
        bestScore = max(bestScore, score)
        addRandomTile()
        return true
    }

    mutating func undo() {
        guard let previousSnapshot else { return }
        board = previousSnapshot.board
        score = previousSnapshot.score
        self.previousSnapshot = nil
    }

    var canUndo: Bool {
        previousSnapshot != nil
    }

    var has2048Tile: Bool {
        board.flatMap { $0 }.contains(2048)
    }

    var isGameOver: Bool {
        if board.flatMap({ $0 }).contains(0) {
            return false
        }

        for row in 0..<4 {
            for column in 0..<4 {
                let value = board[row][column]
                if row < 3, board[row + 1][column] == value {
                    return false
                }
                if column < 3, board[row][column + 1] == value {
                    return false
                }
            }
        }

        return true
    }

    private func line(at index: Int, direction: MoveDirection) -> [Int] {
        switch direction {
        case .left:
            return board[index]
        case .right:
            return board[index].reversed()
        case .up:
            return (0..<4).map { board[$0][index] }
        case .down:
            return (0..<4).reversed().map { board[$0][index] }
        }
    }

    private func mergedLine(from line: [Int]) -> (values: [Int], gainedScore: Int) {
        let compacted = line.filter { $0 != 0 }
        var merged: [Int] = []
        var gainedScore = 0
        var index = 0

        // Quy tắc 2048: mỗi ô chỉ được merge một lần trong một lần vuốt.
        while index < compacted.count {
            if index + 1 < compacted.count, compacted[index] == compacted[index + 1] {
                let newValue = compacted[index] * 2
                merged.append(newValue)
                gainedScore += newValue
                index += 2
            } else {
                merged.append(compacted[index])
                index += 1
            }
        }

        return (merged + Array(repeating: 0, count: 4 - merged.count), gainedScore)
    }

    private func write(_ line: [Int], to board: inout [[Int]], at index: Int, direction: MoveDirection) {
        switch direction {
        case .left:
            board[index] = line
        case .right:
            board[index] = line.reversed()
        case .up:
            for row in 0..<4 {
                board[row][index] = line[row]
            }
        case .down:
            for (offset, row) in (0..<4).reversed().enumerated() {
                board[row][index] = line[offset]
            }
        }
    }

    private mutating func addRandomTile() {
        let emptyCells = board.enumerated().flatMap { row, values in
            values.enumerated().compactMap { column, value in
                value == 0 ? (row, column) : nil
            }
        }

        guard let cell = emptyCells.randomElement() else { return }
        board[cell.0][cell.1] = Int.random(in: 0..<10) == 0 ? 4 : 2
    }
}
