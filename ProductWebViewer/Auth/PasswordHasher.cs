using System.Security.Cryptography;

namespace ProductWebViewer.Auth {
    // PBKDF2によるパスワードのハッシュ化・検証を行う（平文パスワードを設定ファイルに保存しないため）
    public static class PasswordHasher {
        private const int SaltSize = 16;
        private const int HashSize = 32;
        private const int Iterations = 100_000;

        // "ソルト.ハッシュ" 形式の文字列を生成する
        public static string Hash(string password) {
            var salt = RandomNumberGenerator.GetBytes(SaltSize);
            var hash = Rfc2898DeriveBytes.Pbkdf2(password, salt, Iterations, HashAlgorithmName.SHA256, HashSize);
            return $"{Convert.ToBase64String(salt)}.{Convert.ToBase64String(hash)}";
        }

        // Hash() が生成した文字列に対して入力パスワードを検証する
        public static bool Verify(string password, string stored) {
            var parts = stored.Split('.');
            if (parts.Length != 2) return false;

            byte[] salt, expectedHash;
            try {
                salt = Convert.FromBase64String(parts[0]);
                expectedHash = Convert.FromBase64String(parts[1]);
            } catch (FormatException) {
                return false;
            }

            var actualHash = Rfc2898DeriveBytes.Pbkdf2(password, salt, Iterations, HashAlgorithmName.SHA256, expectedHash.Length);
            // タイミング攻撃を防ぐため定数時間比較を使う
            return CryptographicOperations.FixedTimeEquals(actualHash, expectedHash);
        }
    }
}
