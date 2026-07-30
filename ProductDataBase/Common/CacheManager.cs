using System.Diagnostics.CodeAnalysis;

namespace ProductDatabase.Common {
    // TTL（Time-To-Live）ベースのジェネリックキャッシング機構
    internal class CacheManager<T>(TimeSpan? ttl = null) {
        private T? _cachedData;
        private DateTime? _lastLoadTime;
        private readonly TimeSpan _ttl = ttl ?? _defaultTtl;
        private readonly Lock _lock = new();

        // TTL のデフォルト値（5 分）
        private static readonly TimeSpan _defaultTtl = TimeSpan.FromMinutes(5);

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
        // MaybeNullWhen(false) で「false を返す場合、cachedData が null の可能性」を示す
        public bool TryGetCachedData([MaybeNullWhen(false)] out T cachedData) {
            lock (_lock) {
                if (IsCacheValidInternal()) {
                    cachedData = _cachedData!;
                    return true;
                }
                // キャッシュ無効時に明示的にメモリを解放
                _cachedData = default;
                _lastLoadTime = null;
                cachedData = default;
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
