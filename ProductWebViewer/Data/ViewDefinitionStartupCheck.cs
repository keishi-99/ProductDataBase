namespace ProductWebViewer.Data {
    // 起動時にリポジトリを構築させ（DB疎通確認）、DB上のビュー定義がこのアプリの想定と一致するか検証する
    // IHostedServiceとして実装することで、Program.cs側でGetRequiredServiceを手動呼び出しする必要をなくす
    internal class ViewDefinitionStartupCheck(
        ProductRecordRepository productRepository,
        SubstrateRecordRepository substrateRepository) : IHostedService {

        public Task StartAsync(CancellationToken cancellationToken) {
            _ = substrateRepository; // コンストラクタ実行（DB疎通確認）のためだけに解決させる
            ViewDefinitionVerifier.Verify(productRepository.ConnectionString);
            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
    }
}
