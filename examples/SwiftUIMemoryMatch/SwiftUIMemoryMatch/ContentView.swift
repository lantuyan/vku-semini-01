import SwiftUI

struct ContentView: View {
    @StateObject private var game = GameViewModel()
    @State private var difficulty: GameDifficulty = .medium

    private let columns = [
        GridItem(.adaptive(minimum: 82), spacing: 12)
    ]

    var body: some View {
        NavigationStack {
            VStack(spacing: 18) {
                header

                DifficultyPicker(difficulty: $difficulty)
                    .onChange(of: difficulty) { _, newValue in
                        withAnimation(.spring(duration: 0.35)) {
                            game.startNewGame(difficulty: newValue)
                        }
                    }

                LazyVGrid(columns: columns, spacing: 12) {
                    ForEach(game.cards) { card in
                        CardView(card: card) {
                            withAnimation(.spring(response: 0.35, dampingFraction: 0.75)) {
                                game.choose(card)
                            }
                        }
                    }
                }

                Spacer(minLength: 0)
            }
            .padding()
            .navigationTitle("Memory Match")
            .background(Color(.systemGroupedBackground))
            .onAppear {
                game.startNewGame(difficulty: difficulty)
            }
        }
    }

    private var header: some View {
        HStack(spacing: 12) {
            StatBox(title: "Lượt thử", value: "\(game.moves)")
            StatBox(title: "Đã tìm", value: "\(game.matchedPairs)/\(difficulty.pairCount)")

            Button {
                withAnimation(.spring(duration: 0.35)) {
                    game.startNewGame(difficulty: difficulty)
                }
            } label: {
                Label("Chơi lại", systemImage: "arrow.clockwise")
                    .font(.headline)
            }
            .buttonStyle(.borderedProminent)
        }
    }
}

struct DifficultyPicker: View {
    // Binding cho phép view con đọc và sửa state của view cha.
    @Binding var difficulty: GameDifficulty

    var body: some View {
        Picker("Độ khó", selection: $difficulty) {
            ForEach(GameDifficulty.allCases) { level in
                Text(level.title).tag(level)
            }
        }
        .pickerStyle(.segmented)
    }
}

struct StatBox: View {
    let title: String
    let value: String

    var body: some View {
        VStack(alignment: .leading, spacing: 2) {
            Text(title)
                .font(.caption)
                .foregroundStyle(.secondary)
            Text(value)
                .font(.title3.bold())
        }
        .frame(maxWidth: .infinity, alignment: .leading)
        .padding(12)
        .background(.background)
        .clipShape(RoundedRectangle(cornerRadius: 8))
    }
}

struct CardView: View {
    let card: MemoryCard
    let onTap: () -> Void

    private var isVisible: Bool {
        card.isFaceUp || card.isMatched
    }

    var body: some View {
        Button(action: onTap) {
            ZStack {
                RoundedRectangle(cornerRadius: 8)
                    .fill(isVisible ? Color.white : Color.indigo)

                RoundedRectangle(cornerRadius: 8)
                    .strokeBorder(isVisible ? Color.indigo.opacity(0.25) : Color.indigo, lineWidth: 2)

                if isVisible {
                    Image(systemName: card.symbol)
                        .font(.system(size: 34, weight: .bold))
                        .foregroundStyle(card.isMatched ? .green : .indigo)
                        .transition(.scale.combined(with: .opacity))
                } else {
                    Image(systemName: "questionmark")
                        .font(.system(size: 28, weight: .bold))
                        .foregroundStyle(.white)
                }
            }
            .aspectRatio(1, contentMode: .fit)
            .rotation3DEffect(.degrees(isVisible ? 0 : 180), axis: (x: 0, y: 1, z: 0))
        }
        .buttonStyle(.plain)
        .disabled(card.isMatched)
        .accessibilityLabel(isVisible ? "Thẻ \(card.symbol)" : "Thẻ chưa lật")
    }
}

#Preview {
    ContentView()
}
