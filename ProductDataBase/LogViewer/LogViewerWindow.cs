using System.Data;
using System.Text;
using System.Text.RegularExpressions;

namespace ProductDatabase.LogViewer {
    public partial class LogViewerWindow : Form {
        // AppDomain.CurrentDomain.BaseDirectory を使用してファイルダイアログ等による CurrentDirectory の変化を回避する
        private static readonly string LogDirectory =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "db", "logs");

        private static readonly string[] ColumnHeaders = [
            "日時", "操作種別", "カテゴリ", "ID",
            "注文番号", "製造番号", "OLes番号", "製品名",
            "タイプ", "型式", "数量", "シリアル先頭",
            "シリアル末尾", "Revision", "登録日", "担当者", "コメント"
        ];

        private static readonly Regex BracketRegex = new(@"\[([^\]]*)\]$", RegexOptions.Compiled);

        private DataTable _logTable = new();
        private DataView _logView = new();

        public LogViewerWindow() {
            InitializeComponent();
            InitializeDataGridView();
            InitializeYearMonthComboBox();
        }

        private void InitializeDataGridView() {
            LogDataGridView.AllowUserToAddRows = false;
            LogDataGridView.AllowUserToDeleteRows = false;
            LogDataGridView.AllowUserToResizeRows = false;
            LogDataGridView.ReadOnly = true;
            LogDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(LogDataGridView.Font, FontStyle.Bold);
            LogDataGridView.RowTemplate.DefaultCellStyle.Padding = new Padding(5);
            LogDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            LogDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // 表示中のセルだけ処理するため DataBindingComplete より CellFormatting の方が大量データで高速
            LogDataGridView.CellFormatting += LogDataGridView_CellFormatting;
        }

        private void InitializeYearMonthComboBox() {
            YearMonthComboBox.Items.Clear();

            if (!Directory.Exists(LogDirectory)) {
                CountLabel.Text = "ログフォルダが見つかりません";
                return;
            }

            var logFiles = Directory.GetFiles(LogDirectory, "log_*.csv")
                .Select(f => Path.GetFileNameWithoutExtension(f).Replace("log_", ""))
                .Where(s => s.Length == 6 && s.All(char.IsDigit))
                .OrderByDescending(s => s)
                .ToList();

            foreach (var ym in logFiles)
                YearMonthComboBox.Items.Add(new YearMonthItem(ym, $"{ym[..4]}年{ym[4..]}月"));

            if (YearMonthComboBox.Items.Count > 0)
                YearMonthComboBox.SelectedIndex = 0;
            else
                CountLabel.Text = "ログファイルがありません";
        }

        private async Task LoadSelectedMonth() {
            if (YearMonthComboBox.SelectedItem is not YearMonthItem item) return;
            _logTable = await Task.Run(() => LoadLogFile(item.Value));
            _logView = _logTable.DefaultView;
            LogDataGridView.DataSource = _logView;
            RefreshOperationTypeFilter();
            ApplyFilter();
        }

        private DataTable LoadLogFile(string yearMonth) {
            var table = new DataTable();
            foreach (var header in ColumnHeaders)
                table.Columns.Add(header);

            var filePath = Path.Combine(LogDirectory, $"log_{yearMonth}.csv");
            if (!File.Exists(filePath)) return table;

            try {
                foreach (var line in File.ReadLines(filePath, Encoding.UTF8)) {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var fields = ParseCsvLine(line);
                    var row = table.NewRow();
                    for (int i = 0; i < ColumnHeaders.Length; i++) {
                        var raw = i < fields.Length ? fields[i] : string.Empty;
                        row[i] = i <= 1 ? raw : ExtractValue(raw);
                    }
                    table.Rows.Add(row);
                }
            } catch (Exception ex) {
                MessageBox.Show(
                    $"ログファイルの読み込み中にエラーが発生しました: {ex.Message}",
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return table;
        }

        private static string ExtractValue(string raw) {
            if (string.IsNullOrEmpty(raw)) return raw;
            var match = BracketRegex.Match(raw);
            return match.Success ? match.Groups[1].Value : raw;
        }

        private static string[] ParseCsvLine(string line) {
            var fields = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;
            int i = 0;
            while (i < line.Length) {
                char c = line[i];
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

        private void RefreshOperationTypeFilter() {
            var current = OperationTypeComboBox.SelectedItem?.ToString();
            OperationTypeComboBox.Items.Clear();
            OperationTypeComboBox.Items.Add("(すべて)");

            foreach (var t in _logTable.AsEnumerable()
                .Select(r => r["操作種別"].ToString() ?? string.Empty)
                .Distinct()
                .OrderBy(s => s)) {
                OperationTypeComboBox.Items.Add(t);
            }

            OperationTypeComboBox.SelectedItem = current ?? "(すべて)";
            if (OperationTypeComboBox.SelectedIndex < 0)
                OperationTypeComboBox.SelectedIndex = 0;
        }

        private void ApplyFilter() {
            var opFilter = OperationTypeComboBox.SelectedItem?.ToString();
            var keyword = SearchTextBox.Text.Trim();

            var conditions = new List<string>();

            if (!string.IsNullOrEmpty(opFilter) && opFilter != "(すべて)") {
                conditions.Add($"[操作種別] = '{opFilter.Replace("'", "''")}'");
            }

            if (!string.IsNullOrEmpty(keyword)) {
                var escaped = keyword.Replace("'", "''").Replace("[", "[[]").Replace("*", "[*]").Replace("%", "[%]");
                var colConditions = ColumnHeaders.Select(h => $"[{h}] LIKE '%{escaped}%'");
                conditions.Add($"({string.Join(" OR ", colConditions)})");
            }

            _logView.RowFilter = string.Join(" AND ", conditions);
            CountLabel.Text = $"{_logView.Count} 件表示中";
        }

        private static Color GetRowColor(string operationType) {
            if (operationType.Contains("削除")) return Color.MistyRose;
            if (operationType.Contains("編集")) return Color.LightGoldenrodYellow;
            return operationType switch {
                "[製品登録]" => Color.LightSkyBlue,
                "[基板登録]" => Color.LightGreen,
                "[基板変更]" => Color.LightYellow,
                "[再印刷]" => Color.LightSalmon,
                "[Rev変更]" => Color.Plum,
                _ => Color.White
            };
        }

        private void LogDataGridView_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e) {
            if (e.RowIndex < 0 || e.CellStyle is null) return;
            if (LogDataGridView.Rows[e.RowIndex].DataBoundItem is DataRowView rowView) {
                var opType = rowView["操作種別"]?.ToString() ?? string.Empty;
                e.CellStyle.BackColor = GetRowColor(opType);
            }
        }

        private async void YearMonthComboBox_SelectedIndexChanged(object sender, EventArgs e) {
            await LoadSelectedMonth();
        }

        private void OperationTypeComboBox_SelectedIndexChanged(object sender, EventArgs e) {
            ApplyFilter();
        }

        private void SearchButton_Click(object sender, EventArgs e) {
            ApplyFilter();
        }

        private void SearchTextBox_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter) ApplyFilter();
        }

        private sealed record YearMonthItem(string Value, string Label) {
            public override string ToString() => Label;
        }
    }
}
