import Foundation

struct MemoryCard: Identifiable, Equatable {
    let id = UUID()
    let symbol: String
    var isFaceUp = false
    var isMatched = false
}

enum GameDifficulty: String, CaseIterable, Identifiable {
    case easy = "Dễ"
    case medium = "Vừa"

    var id: String { rawValue }

    var title: String {
        switch self {
        case .easy: "Dễ"
        case .medium: "Vừa"
        }
    }

    var pairCount: Int {
        switch self {
        case .easy: 4
        case .medium: 6
        }
    }
}
