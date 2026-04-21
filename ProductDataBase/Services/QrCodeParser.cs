using System.Text.RegularExpressions;

namespace ProductDatabase.Services {
    internal static partial class QrCodeParser {

        // QRコードテキストを解析して製番・品目番号・数量・注番を取得する
        public static QrCodeParseResult Parse(string input) {
            string[] separator = ["//"];
            var arr = input.Split(separator, StringSplitOptions.None);

            if (arr.Length == 1) {
                return new QrCodeParseResult { ProductModel = input };
            }
            if (arr.Length != 4) {
                throw new Exception("QRコードが正しくありません。");
            }
            return new QrCodeParseResult {
                ProductNumber = arr[0],
                ProductModel = arr[1],
                Quantity = int.TryParse(arr[2], out var qty) ? qty : throw new Exception("数量に数値以外が入力されています。"),
                OrderNumber = arr[3]
            };
        }

        [GeneratedRegex(@"-(?:SMT|H|GH).*")]
        private static partial Regex SuffixRegex();

        // 品目番号から不要なサフィックスを除去して正規化する
        public static string NormalizeProductModel(string productModel) {
            var result = SuffixRegex().Replace(productModel, string.Empty);
            return result
                .Replace("-ACGH", "-AC")
                .Replace("-DCGH", "-DC");
        }
    }

    internal sealed class QrCodeParseResult {
        public string ProductNumber { get; init; } = string.Empty;
        public string ProductModel { get; init; } = string.Empty;
        public int Quantity { get; init; }
        public string OrderNumber { get; init; } = string.Empty;
    }
}
