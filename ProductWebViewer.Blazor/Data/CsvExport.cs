using System.Text;
using ProductWebViewer.Blazor.Models;

namespace ProductWebViewer.Blazor.Data {
    // 本家 ProductWebViewer の IndexModel.BuildCsvBytes/CsvField と同じロジック
    internal static class CsvExport {
        // BOM付きUTF-8 バイト列を返す（Excelで日本語が文字化けしないようにするため）
        public static byte[] BuildCsvBytes(IEnumerable<string> lines) {
            using var ms = new MemoryStream();
            using (var writer = new StreamWriter(ms, new UTF8Encoding(true))) {
                writer.NewLine = "\r\n";
                foreach (var line in lines)
                    writer.WriteLine(line);
            }
            return ms.ToArray();
        }

        // カンマ・ダブルクォート・改行を含むフィールドをエスケープする
        // 数式インジェクション対策: =,+,-,@ 始まりの値にはシングルクォートを付加する
        public static string Field(string? value) {
            if (string.IsNullOrEmpty(value)) return "";
            value = value.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ");
            if (value.Length > 0 && (value[0] is '=' or '+' or '-' or '@'))
                value = "'" + value;
            if (value.Contains(',') || value.Contains('"'))
                return $"\"{value.Replace("\"", "\"\"")}\"";
            return value;
        }

        public static IEnumerable<string> BuildProductLines(IReadOnlyList<ProductRecord> records) {
            yield return "ID,カテゴリ,製品名,種別,型式,注文番号,製造番号,O-Les番号,数量,担当者,登録日,Revision,シリアル開始,シリアル終了,コメント,登録日時";
            foreach (var r in records)
                yield return string.Join(",",
                    Field(r.Id.ToString()), Field(r.CategoryName), Field(r.ProductName),
                    Field(r.ProductType), Field(r.ProductModel), Field(r.OrderNumber),
                    Field(r.ProductNumber), Field(r.OLesNumber), Field(r.Quantity?.ToString()),
                    Field(r.PersonInfo), Field(r.RegDate), Field(r.Revision),
                    Field(r.SerialFirst), Field(r.SerialLast), Field(r.Comment), Field(r.CreatedAt));
        }

        public static IEnumerable<string> BuildSubstrateLines(IReadOnlyList<SubstrateRecord> records) {
            yield return "ID,カテゴリ,製品名,基板名,基板型式,注文番号,製造番号,入庫,出庫,不良,使用製品名,使用注文番号,使用製造番号,担当者,登録日,コメント,登録日時";
            foreach (var r in records)
                yield return string.Join(",",
                    Field(r.Id.ToString()), Field(r.CategoryName), Field(r.ProductName),
                    Field(r.SubstrateName), Field(r.SubstrateModel), Field(r.OrderNumber),
                    Field(r.SubstrateNumber), Field(r.Increase?.ToString()), Field(r.Decrease?.ToString()),
                    Field(r.Defect?.ToString()), Field(r.UseProductName), Field(r.UseOrderNumber),
                    Field(r.UseProductNumber), Field(r.PersonInfo), Field(r.RegDate),
                    Field(r.Comment), Field(r.CreatedAt));
        }
    }
}
