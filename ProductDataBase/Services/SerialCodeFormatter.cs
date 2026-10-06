using ProductDatabase.Common;

namespace ProductDatabase.Services {
    internal static class SerialCodeFormatter {

        // 書式文字列のプレースホルダーを値に置換してラベル印字コードを生成する
        // {T}接頭 {OT}O-Les接頭 {Y}製造年(2桁) {MM}製造月(2桁) {R}リビジョン {M}月コード(1桁) {S}シリアル {SA}O-Les接尾
        public static string Format(
            string format, string? initial, string? oLesInitial, DateTime regDate,
            string revision, int serialNumber, int serialDigit, string oLesSuffix) {

            var monthCode = CommonUtils.ToMonthCode(regDate);

            var map = new Dictionary<string, string> {
                ["{T}"] = initial ?? string.Empty,
                ["{OT}"] = oLesInitial ?? string.Empty,
                ["{Y}"] = regDate.ToString("yy"),
                ["{MM}"] = regDate.ToString("MM"),
                ["{R}"] = revision,
                ["{M}"] = monthCode[^1..],
                ["{S}"] = serialNumber.ToString($"D{serialDigit}"),
                ["{SA}"] = oLesSuffix
            };

            var outputCode = format;
            foreach (var kv in map) {
                outputCode = outputCode.Replace(kv.Key, kv.Value);
            }

            return outputCode;
        }

        // O-Lesシリアル接尾の次の文字を返す
        public static string GetNextOLesSuffix(string? current) {
            if (string.IsNullOrWhiteSpace(current)) return "A";
            var c = char.ToUpperInvariant(current[0]);
            if (c < 'A' || c >= 'Z') return "A";  // 'Z' は 'A' へ循環、範囲外文字も 'A' にフォールバック
            return ((char)(c + 1)).ToString();
        }
    }
}
