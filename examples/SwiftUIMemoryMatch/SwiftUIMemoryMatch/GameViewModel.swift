import SwiftUI

@MainActor
final class GameViewModel: ObservableObject {
    @Published private(set) var cards: [MemoryCard] = []
    @Published private(set) var moves = 0

    private var firstSelectedIndex: Int?
    private var isCheckingPair = false
    private var gameID = UUID()
    private let allSymbols = ["star.fill", "heart.fill", "bolt.fill", "leaf.fill", "flame.fill", "moon.fill"]

    var matchedPairs: Int {
        cards.filter(\.isMatched).count / 2
    }

    func startNewGame(difficulty: GameDifficulty) {
        gameID = UUID()
        let symbols = Array(allSymbols.prefix(difficulty.pairCount))

        cards = (symbols + symbols)
            .map { MemoryCard(symbol: $0) }
            .shuffled()

        moves = 0
        firstSelectedIndex = nil
        isCheckingPair = false
    }

    func choose(_ card: MemoryCard) {
        guard let index = cards.firstIndex(of: card),
              !cards[index].isFaceUp,
              !cards[index].isMatched,
              !isCheckingPair else {
            return
        }

        cards[index].isFaceUp = true

        if let firstIndex = firstSelectedIndex {
            checkPair(firstIndex: firstIndex, secondIndex: index)
        } else {
            firstSelectedIndex = index
        }
    }

    private func checkPair(firstIndex: Int, secondIndex: Int) {
        moves += 1
        isCheckingPair = true
        let currentGameID = gameID

        if cards[firstIndex].symbol == cards[secondIndex].symbol {
            cards[firstIndex].isMatched = true
            cards[secondIndex].isMatched = true
            firstSelectedIndex = nil
            isCheckingPair = false
        } else {
            Task {
                try? await Task.sleep(for: .milliseconds(650))
                guard currentGameID == gameID else { return }
                cards[firstIndex].isFaceUp = false
                cards[secondIndex].isFaceUp = false
                firstSelectedIndex = nil
                isCheckingPair = false
            }
        }
    }
}
