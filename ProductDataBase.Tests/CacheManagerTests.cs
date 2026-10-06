using ProductDatabase.Common;

namespace ProductDataBase.Tests {
    public class CacheManagerTests {

        // 期限切れを待つテストで使う短いTTLと、それを確実に超える待ち時間
        private static readonly TimeSpan ShortTtl = TimeSpan.FromMilliseconds(50);
        private static readonly TimeSpan WaitPastTtl = TimeSpan.FromMilliseconds(300);

        [Fact]
        public void TryGetCachedData_未設定_取得できない() {
            var cache = new CacheManager<string>();

            var found = cache.TryGetCachedData(out var data);

            Assert.False(found);
            Assert.Null(data);
        }

        [Fact]
        public void TryGetCachedData_期限内_保存した値を取得できる() {
            var cache = new CacheManager<string>();
            cache.SetCache("abc");

            var found = cache.TryGetCachedData(out var data);

            Assert.True(found);
            Assert.Equal("abc", data);
        }

        [Fact]
        public void TryGetCachedData_期限切れ_取得できず無効になる() {
            var cache = new CacheManager<string>(ShortTtl);
            cache.SetCache("abc");

            Thread.Sleep(WaitPastTtl);

            Assert.False(cache.IsCacheValid());
            Assert.False(cache.TryGetCachedData(out var data));
            Assert.Null(data);
        }

        [Fact]
        public void SetCache_2回呼ぶ_後から保存した値で上書きされる() {
            var cache = new CacheManager<string>();
            cache.SetCache("first");
            cache.SetCache("second");

            cache.TryGetCachedData(out var data);

            Assert.Equal("second", data);
        }

        [Fact]
        public void ClearCache_呼ぶ_取得できなくなる() {
            var cache = new CacheManager<string>();
            cache.SetCache("abc");

            cache.ClearCache();

            Assert.False(cache.IsCacheValid());
            Assert.False(cache.TryGetCachedData(out _));
        }

        [Fact]
        public void GetCacheStatus_TTL未指定_既定の5分になる() {
            var cache = new CacheManager<string>();

            var status = cache.GetCacheStatus();

            Assert.Equal(TimeSpan.FromMinutes(5), status.Ttl);
            Assert.False(status.IsValid);
            Assert.Null(status.LastLoadTime);
        }
    }
}
