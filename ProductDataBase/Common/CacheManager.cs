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

        // キャッシュが有効期限内かどうかを判定する（ロックなしの内部用）
        private bool IsCacheValidInternal() {
            return _lastLoadTime.HasValue && DateTime.UtcNow - _lastLoadTime.Value < _ttl;
        }

        // キャッシュが有効期限内かどうかを判定する
        public bool IsCacheValid() {
            lock (_lock) {
                return IsCacheValidInternal();
            }
        }

        // キャッシュが有効な場合はデータを取得する（out パラメータで戻す）
        public bool TryGetCachedData(out T cachedData) {
            lock (_lock) {
                if (IsCacheValidInternal()) {
                    cachedData = _cachedData!;
                    return true;
                }
                // キャッシュ無効時に明示的にメモリを解放
                _cachedData = default;
                _lastLoadTime = null;
                cachedData = default!;
                return false;
            }
        }

        // データをキャッシュに保存し、タイムスタンプを更新する
        public void SetCache(T data) {
            lock (_lock) {
                _cachedData = data;
                _lastLoadTime = DateTime.UtcNow;
            }
        }

        // キャッシュをクリアし、タイムスタンプをリセットする
        public void ClearCache() {
            lock (_lock) {
                _cachedData = default;
                _lastLoadTime = null;
            }
        }

        // デバッグ・テスト用：キャッシュ状態を取得する
        public (bool IsValid, DateTime? LastLoadTime, TimeSpan Ttl) GetCacheStatus() {
            lock (_lock) {
                return (IsCacheValidInternal(), _lastLoadTime, _ttl);
            }
        }
    }
}
