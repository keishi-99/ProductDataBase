using System.Security.Cryptography;
using System.Text;

namespace ProductDatabase.Common {
    // ビューのSQL定義文を比較するためのハッシュ計算（空白の書式差異は無視する）
    internal static class ViewDefinitionHash {
        public static string Compute(string sql) {
            if (string.IsNullOrEmpty(sql)) { return string.Empty; }

            var normalized = string.Join(' ', sql.ToUpperInvariant().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
            var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(normalized));
            return Convert.ToHexString(bytes);
        }
    }
}
