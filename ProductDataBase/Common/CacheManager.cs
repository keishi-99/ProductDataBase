namespace ProductDatabase.Common {
    // TTL（Time-To-Live）ベースのジェネリックキャッシング機構
    internal class CacheManager<T> {
        private T? _cachedData;
        private DateTime? _lastLoadTime;
        private readonly TimeSpan _ttl;
        private readonly object _lock = new();

        // TTL のデフォルト値（5 分）
        private static readonly TimeSpan DefaultTtl = TimeSpan.FromMinutes(5);

        public CacheManager(TimeSpan? ttl = null) {
            _ttl = ttl ?? DefaultTtl;
        }

        // キャッシュが有効期限内かどうかを判定する
        public bool IsCacheValid() {
            lock (_lock) {
                return _lastLoadTime.HasValue && DateTime.Now - _lastLoadTime.Value < _ttl;
            }
        }

        // キャッシュが有効な場合はデータを取得、無効な場合は null を返す
        public T? GetCachedData() {
            lock (_lock) {
                if (IsCacheValid()) {
                    return _cachedData;
                }
                return null;
            }
        }

        // データをキャッシュに保存し、タイムスタンプを更新する
        public void SetCache(T data) {
            lock (_lock) {
                _cachedData = data;
                _lastLoadTime = DateTime.Now;
            }
        }

        // キャッシュをクリアし、タイムスタンプをリセットする
        public void ClearCache() {
            lock (_lock) {
                _cachedData = default(T);
                _lastLoadTime = null;
            }
        }

        // デバッグ・テスト用：キャッシュ状態を取得する
        public (bool IsValid, DateTime? LastLoadTime, TimeSpan Ttl) GetCacheStatus() {
            lock (_lock) {
                return (IsCacheValid(), _lastLoadTime, _ttl);
            }
        }
    }
}
