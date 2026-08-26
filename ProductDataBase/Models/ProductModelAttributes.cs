namespace ProductDatabase.Models {
    // ProductModel（型式）末尾3桁のコードから仕様を分解する。PA2K固有のコード規則:
    // 百の位(1-8): 1引いた値を3bitとして 0bit=電源(0:AC/1:DC) 1bit=設置形態(0:一体形/1:別置形) 2bit=ADI(0:無し/1:有り)
    // 十の位: 1=HART / 2=MODBUS
    // 一の位: 0=無し / 1=Ver.C / 2=Ver.D
    public class ProductModelAttributes {
        public string Power { get; init; } = string.Empty;        // AC / DC
        public string Mount { get; init; } = string.Empty;        // 一体形 / 別置形
        public bool HasAdi { get; init; }                          // ADI有無
        public string Communication { get; init; } = string.Empty; // HART / MODBUS
        public string? Version { get; init; }                      // C / D / null(無し)

        // 末尾3桁が規則に一致しないProductModelはnullを返す
        public static ProductModelAttributes? TryParse(string productModel) {
            if (productModel.Length < 3) { return null; }

            var suffix = productModel[^3..];
            if (!int.TryParse(suffix.AsSpan(0, 1), out var powerCode) ||
                !int.TryParse(suffix.AsSpan(1, 1), out var commCode) ||
                !int.TryParse(suffix.AsSpan(2, 1), out var verCode)) {
                return null;
            }
            if (powerCode is < 1 or > 8) { return null; }
            if (verCode is not (0 or 1 or 2)) { return null; }

            var communication = commCode switch {
                1 => "HART",
                2 => "MODBUS",
                _ => null
            };
            if (communication is null) { return null; }

            var index = powerCode - 1;

            return new ProductModelAttributes {
                Power = (index & 1) == 0 ? "AC" : "DC",
                Mount = (index & 2) == 0 ? "一体形" : "別置形",
                HasAdi = (index & 4) != 0,
                Communication = communication,
                Version = verCode switch { 1 => "C", 2 => "D", _ => null }
            };
        }
    }
}
