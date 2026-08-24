using System.Text;
using System.Text.RegularExpressions;
using ProductWebViewer.Blazor.Models;

namespace ProductWebViewer.Blazor.Data {
    // db/logs/log_web_yyyyMM.csv を読み取り専用で参照する（このアプリはこのファイルにしか書き込まないため、
    // 本家のように log_*.csv と log_web_*.csv をマージする必要はない）。
    public partial class LogRecordRepository {
        private readonly string _logDirectory;

        [GeneratedRegex(@"\[([^\]]*)\]$")]
        private static partial Regex BracketRegex();

        public LogRecordRepository(IConfiguration configuration) {
            var dbPath = configuration["DatabasePath"] ?? "db/ProductRegistry.Practice.db";
            var dbFullPath = Path.IsPathRooted(dbPath) ? dbPath : Path.Combine(AppContext.BaseDirectory, dbPath);
            _logDirectory = Path.Combine(Path.GetDirectoryName(dbFullPath) ?? AppContext.BaseDirectory, "logs");
        }

        // 存在する年月（新しい順）を "yyyyMM" の形式で返す
        public IReadOnlyList<string> GetAvailableMonths() {
            if (!Directory.Exists(_logDirectory)) return [];

            return Directory.GetFiles(_logDirectory, "log_web_*.csv")
                .Select(f => Path.GetFileNameWithoutExtension(f).Replace("log_web_", ""))
                .Where(s => s.Length == 6 && s.All(char.IsDigit))
                .Distinct()
                .OrderByDescending(s => s)
                .ToList();
        }

        // 指定年月の操作ログを返す
        public IReadOnlyList<LogEntry> GetLogEntries(string yearMonth) {
            // yearMonthがファイルパスの一部になるため、意図しないパス（ディレクトリ区切り等）を含まないことを確認する
            if (string.IsNullOrEmpty(yearMonth) || yearMonth.Length != 6 || !yearMonth.All(char.IsAsciiDigit))
                return [];

            return ReadEntriesFromFile(Path.Combine(_logDirectory, $"log_web_{yearMonth}.csv")).ToList();
        }

        private static IEnumerable<LogEntry> ReadEntriesFromFile(string filePath) {
            if (!File.Exists(filePath)) yield break;

            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(fs, Encoding.UTF8);
            while (reader.ReadLine() is string line) {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var f = ParseCsvLine(line);
                yield return new LogEntry {
                    Timestamp = GetField(f, 0),
                    OperationType = GetField(f, 1),
                    Category = ExtractValue(GetField(f, 2)),
                    Id = ExtractValue(GetField(f, 3)),
                    OrderNumber = ExtractValue(GetField(f, 4)),
                    ProductNumber = ExtractValue(GetField(f, 5)),
                    OLesNumber = ExtractValue(GetField(f, 6)),
                    ProductName = ExtractValue(GetField(f, 7)),
                    ProductType = ExtractValue(GetField(f, 8)),
                    ProductModel = ExtractValue(GetField(f, 9)),
                    Quantity = ExtractValue(GetField(f, 10)),
                    SerialFirst = ExtractValue(GetField(f, 11)),
                    SerialLast = ExtractValue(GetField(f, 12)),
                    Revision = ExtractValue(GetField(f, 13)),
                    RegDate = ExtractValue(GetField(f, 14)),
                    PersonInfo = ExtractValue(GetField(f, 15)),
                    Comment = ExtractValue(GetField(f, 16)),
                };
            }
        }

        private static string GetField(string[] fields, int index) => index < fields.Length ? fields[index] : string.Empty;

        // "ラベル[値]" 形式から値だけを取り出す（該当しなければそのまま返す）
        private static string ExtractValue(string raw) {
            if (string.IsNullOrEmpty(raw)) return raw;
            var match = BracketRegex().Match(raw);
            return match.Success ? match.Groups[1].Value : raw;
        }

        // ダブルクォートで囲まれたCSVの1行を分割する（フィールド内のカンマ・改行・エスケープされた""に対応）
        private static string[] ParseCsvLine(string line) {
            var fields = new List<string>();
            var current = new StringBuilder();
            var inQuotes = false;
            var i = 0;
            while (i < line.Length) {
                var c = line[i];
                if (inQuotes) {
                    if (c == '"' && i + 1 < line.Length && line[i + 1] == '"') {
                        current.Append('"');
                        i += 2;
                    } else if (c == '"') {
                        inQuotes = false;
                        i++;
                    } else {
                        current.Append(c);
                        i++;
                    }
                } else {
                    if (c == '"') {
                        inQuotes = true;
                        i++;
                    } else if (c == ',') {
                        fields.Add(current.ToString());
                        current.Clear();
                        i++;
                    } else {
                        current.Append(c);
                        i++;
                    }
                }
            }
            fields.Add(current.ToString());
            return [.. fields];
        }
    }
}
