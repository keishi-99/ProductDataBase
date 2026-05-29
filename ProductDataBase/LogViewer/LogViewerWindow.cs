using ProductDatabase.Common;
using ProductDatabase.Models;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;

namespace ProductDatabase.LogViewer {
    public partial class LogViewerWindow : Form {
        // Logger と同じパスを参照して書き込み先と読み込み先のズレを防ぐ
        private static string LogDirectory => Logger._logDirectory;

        private static readonly string[] _columnHeaders = [
            "日時", "操作種別", "カテゴリ", "ID",
            "注文番号", "製造番号", "OLes番号", "製品名",
            "タイプ", "型式", "数量", "シリアル先頭",
            "シリアル末尾", "Revision", "登録日", "担当者", "コメント"
        ];

        private static readonly string[] _errorColumnHeaders = [
            "日時", "メソッド名", "例外タイプ", "エラーメッセージ", "スタックトレース", "追加情報"
        ];

        [GeneratedRegex(@"\[([^\]]*)\]$")]
        private static partial Regex BracketRegex();

        private DataTable _logTable = new();
        private DataView _logView = new();
        private CancellationTokenSource? _loadCts;

        private DataTable _errorLogTable = new();
        private DataView _errorLogView = new();
        private CancellationTokenSource? _errorLoadCts;

        public LogViewerWindow(AppSettings appSettings) {
            InitializeComponent();
            Font = new Font(appSettings.FontName, appSettings.FontSize);
            InitializeDataGridView();
            InitializeYearMonthComboBox();
            InitializeErrorDataGridView();
            InitializeErrorYearMonthComboBox();
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

            // 前回の読み込みをキャンセルして最新の選択結果だけをUIに反映する
            _loadCts?.Cancel();
            _loadCts = new CancellationTokenSource();
            var cts = _loadCts;

            try {
                // ファイル読み込みと操作種別リストの抽出をまとめてバックグラウンドで実行
                var (table, types) = await Task.Run(() => {
                    var t = LoadLogFile(item.Value, cts.Token);
                    var ops = t.AsEnumerable()
                        .Select(r => r["操作種別"].ToString() ?? string.Empty)
                        .Distinct()
                        .OrderBy(s => s)
                        .ToList();
                    return (t, ops);
                }, cts.Token);

                if (cts.IsCancellationRequested) return;

                _logTable = table;
                _logView = _logTable.DefaultView;
                LogDataGridView.DataSource = _logView;
                RefreshOperationTypeFilter(types);
                ApplyFilter();
            } catch (OperationCanceledException) {
                // 新しい選択に切り替わったためキャンセル済み
            } catch (Exception ex) {
                if (!cts.IsCancellationRequested) {
                    MessageBox.Show(
                        $"ログファイルの読み込み中にエラーが発生しました: {ex.Message}",
                        "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private static DataTable LoadLogFile(string yearMonth, CancellationToken token) {
            var table = new DataTable();
            foreach (var header in _columnHeaders)
                table.Columns.Add(header);

            var filePath = Path.Combine(LogDirectory, $"log_{yearMonth}.csv");
            if (!File.Exists(filePath)) return table;

            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(fs, Encoding.UTF8);
            while (reader.ReadLine() is string line) {
                token.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(line)) continue;
                var fields = ParseCsvLine(line);
                var row = table.NewRow();
                for (int i = 0; i < _columnHeaders.Length; i++) {
                    var raw = i < fields.Length ? fields[i] : string.Empty;
                    var header = _columnHeaders[i];
                    row[i] = (header == "日時" || header == "操作種別") ? raw : ExtractValue(raw);
                }
                table.Rows.Add(row);
            }

            return table;
        }

        private static string ExtractValue(string raw) {
            if (string.IsNullOrEmpty(raw)) return raw;
            var match = BracketRegex().Match(raw);
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
                    }
                    else if (c == '"') {
                        inQuotes = false;
                        i++;
                    }
                    else {
                        current.Append(c);
                        i++;
                    }
                }
                else {
                    if (c == '"') {
                        inQuotes = true;
                        i++;
                    }
                    else if (c == ',') {
                        fields.Add(current.ToString());
                        current.Clear();
                        i++;
                    }
                    else {
                        current.Append(c);
                        i++;
                    }
                }
            }
            fields.Add(current.ToString());
            return [.. fields];
        }

        private void RefreshOperationTypeFilter(List<string> types) {
            var current = OperationTypeComboBox.SelectedItem?.ToString();
            OperationTypeComboBox.Items.Clear();
            OperationTypeComboBox.Items.Add("(すべて)");

            foreach (var t in types)
                OperationTypeComboBox.Items.Add(t);

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
                var escaped = keyword
                    .Replace("'", "''")
                    .Replace("[", "[[]")
                    .Replace("]", "[]]")
                    .Replace("*", "[*]")
                    .Replace("%", "[%]");
                var colConditions = _columnHeaders.Select(h => $"[{h}] LIKE '%{escaped}%'");
                conditions.Add($"({string.Join(" OR ", colConditions)})");
            }

            _logView.RowFilter = string.Join(" AND ", conditions);
            CountLabel.Text = $"{_logView.Count} 件表示中";
        }

        private static Color GetRowColor(string operationType) {
            if (operationType is LogOperationTypes.SubstrateHistoryDelete
                or LogOperationTypes.ProductHistoryDelete
                or LogOperationTypes.ProductRelatedSubstrateDelete
                or LogOperationTypes.ProductRelatedSerialDelete
                or LogOperationTypes.SerialHistoryDelete)
                return Color.MistyRose;

            if (operationType is LogOperationTypes.SubstrateHistoryEditBefore
                or LogOperationTypes.SubstrateHistoryEditAfter
                or LogOperationTypes.ProductHistoryEditBefore
                or LogOperationTypes.ProductHistoryEditAfter)
                return Color.LightGoldenrodYellow;

            return operationType switch {
                LogOperationTypes.ProductRegistration => Color.LightSkyBlue,
                LogOperationTypes.SubstrateRegistration => Color.LightGreen,
                LogOperationTypes.SubstrateChange => Color.LightYellow,
                LogOperationTypes.RePrint => Color.LightSalmon,
                LogOperationTypes.RevChange => Color.Plum,
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

        private void MainTabControl_SelectedIndexChanged(object sender, EventArgs e) {
            CountLabel.Text = MainTabControl.SelectedTab == ErrorLogTabPage
                ? $"{_errorLogView.Count} 件表示中"
                : $"{_logView.Count} 件表示中";
        }

        private void InitializeErrorDataGridView() {
            ErrorDataGridView.AllowUserToAddRows = false;
            ErrorDataGridView.AllowUserToDeleteRows = false;
            ErrorDataGridView.AllowUserToResizeRows = false;
            ErrorDataGridView.ReadOnly = true;
            ErrorDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(ErrorDataGridView.Font, FontStyle.Bold);
            ErrorDataGridView.RowTemplate.DefaultCellStyle.Padding = new Padding(5);
            ErrorDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            ErrorDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            ErrorDataGridView.CellFormatting += ErrorDataGridView_CellFormatting;
        }

        private void InitializeErrorYearMonthComboBox() {
            ErrorYearMonthComboBox.Items.Clear();

            if (!Directory.Exists(LogDirectory)) return;

            var errorFiles = Directory.GetFiles(LogDirectory, "error_*.csv")
                .Select(f => Path.GetFileNameWithoutExtension(f).Replace("error_", ""))
                .Where(s => s.Length == 6 && s.All(char.IsDigit))
                .OrderByDescending(s => s)
                .ToList();

            foreach (var ym in errorFiles)
                ErrorYearMonthComboBox.Items.Add(new YearMonthItem(ym, $"{ym[..4]}年{ym[4..]}月"));

            if (ErrorYearMonthComboBox.Items.Count > 0)
                ErrorYearMonthComboBox.SelectedIndex = 0;
        }

        private async Task LoadErrorSelectedMonth() {
            if (ErrorYearMonthComboBox.SelectedItem is not YearMonthItem item) return;

            _errorLoadCts?.Cancel();
            _errorLoadCts = new CancellationTokenSource();
            var cts = _errorLoadCts;

            try {
                var table = await Task.Run(() => LoadErrorLogFile(item.Value, cts.Token), cts.Token);
                if (cts.IsCancellationRequested) return;

                _errorLogTable = table;
                _errorLogView = _errorLogTable.DefaultView;
                ErrorDataGridView.DataSource = _errorLogView;

                // スタックトレースは長いため幅を固定する
                if (ErrorDataGridView.Columns.Contains("スタックトレース")) {
                    var col = ErrorDataGridView.Columns["スタックトレース"];
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    col.Width = 300;
                }

                ApplyErrorFilter();
            } catch (OperationCanceledException) {
            } catch (Exception ex) {
                if (!cts.IsCancellationRequested) {
                    MessageBox.Show(
                        $"エラーログファイルの読み込み中にエラーが発生しました: {ex.Message}",
                        "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private static DataTable LoadErrorLogFile(string yearMonth, CancellationToken token) {
            var table = new DataTable();
            foreach (var header in _errorColumnHeaders)
                table.Columns.Add(header);

            var filePath = Path.Combine(LogDirectory, $"error_{yearMonth}.csv");
            if (!File.Exists(filePath)) return table;

            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(fs, Encoding.UTF8);
            while (reader.ReadLine() is string line) {
                token.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(line)) continue;
                var fields = ParseCsvLine(line);
                var row = table.NewRow();
                for (int i = 0; i < _errorColumnHeaders.Length; i++)
                    row[i] = i < fields.Length ? fields[i] : string.Empty;
                table.Rows.Add(row);
            }

            return table;
        }

        private void ApplyErrorFilter() {
            var keyword = ErrorSearchTextBox.Text.Trim();

            if (string.IsNullOrEmpty(keyword)) {
                _errorLogView.RowFilter = null;
            } else {
                var escaped = keyword
                    .Replace("'", "''")
                    .Replace("[", "[[]")
                    .Replace("]", "[]]")
                    .Replace("*", "[*]")
                    .Replace("%", "[%]");
                var colConditions = _errorColumnHeaders.Select(h => $"[{h}] LIKE '%{escaped}%'");
                _errorLogView.RowFilter = string.Join(" OR ", colConditions);
            }

            CountLabel.Text = $"{_errorLogView.Count} 件表示中";
        }

        private void ErrorDataGridView_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e) {
            if (e.RowIndex < 0 || e.CellStyle is null) return;
            if (ErrorDataGridView.Rows[e.RowIndex].DataBoundItem is not DataRowView rowView) return;
            var exType = rowView["例外タイプ"]?.ToString() ?? string.Empty;
            // TODO(human): exType の値に応じて e.CellStyle.BackColor を設定する（2〜5行）
            // 例外タイプ例: SqliteException, IOException, InvalidOperationException など
            // 参考: GetRowColor() で使われている Color 定数（MistyRose, LightYellow 等）を活用する
        }

        private async void ErrorYearMonthComboBox_SelectedIndexChanged(object sender, EventArgs e) {
            await LoadErrorSelectedMonth();
        }

        private void ErrorSearchButton_Click(object sender, EventArgs e) {
            ApplyErrorFilter();
        }

        private void ErrorSearchTextBox_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter) ApplyErrorFilter();
        }

        private sealed record YearMonthItem(string Value, string Label) {
            public override string ToString() => Label;
        }
    }
}
