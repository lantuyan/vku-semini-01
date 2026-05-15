import UIKit

final class SceneDelegate: UIResponder, UIWindowSceneDelegate {
    var window: UIWindow?

    func scene(
        _ scene: UIScene,
        willConnectTo session: UISceneSession,
        options connectionOptions: UIScene.ConnectionOptions
    ) {
        guard let windowScene = scene as? UIWindowScene else { return }

        let window = UIWindow(windowScene: windowScene)
        let navigationController = UINavigationController(rootViewController: GameViewController())
        navigationController.navigationBar.tintColor = UIColor(red: 0.56, green: 0.48, blue: 0.40, alpha: 1)
        window.rootViewController = navigationController
        window.makeKeyAndVisible()
        self.window = window
    }
}
