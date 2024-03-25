using ProductDatabase.Substrate;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProductDatabase {
    public partial class SubstrateChange1 : Form {

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public DataTable DtHistoryTable { get; set; } = new();

        private readonly List<string> ListColFilter = new();

        public int IntRadioBtnFlg { get; set; }

        public string StrProductName { get; set; } = string.Empty;
        public string StrProductModel { get; set; } = string.Empty;
        public string StrProductType { get; set; } = string.Empty;
        public string StrSubstrateName { get; set; } = string.Empty;
        public string StrSubstrateModel { get; set; } = string.Empty;

        public SubstrateChange1() => InitializeComponent();

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(StrFontName, IntFontSize);

                switch (IntRadioBtnFlg) {
                    case 4:
                        using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                            DtHistoryTable.Clear();
                            using SQLiteDataAdapter _adapter = new($"SELECT _rowid_, * FROM 'Product_Reg_{StrProductName}' WHERE col_Product_Model = '{StrProductModel}' ORDER BY _rowid_ DESC", _con);
                            _adapter.Fill(DtHistoryTable);

                            SubstrateChangeDataGridView.DataSource = DtHistoryTable;

                            ListColFilter.Add("");
                            for (int _i = 0; _i < SubstrateChangeDataGridView.ColumnCount; _i++) {
                                string _headerValue = SubstrateChangeDataGridView.Columns[_i].HeaderCell.Value?.ToString() ?? string.Empty;
                                if (_headerValue != null) { ListColFilter.Add(_headerValue); }
                            }

                            SubstrateChangeDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(SubstrateChangeDataGridView.Font, FontStyle.Bold);
                            SubstrateChangeDataGridView.RowHeadersVisible = true;
                            SubstrateChangeDataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            SubstrateChangeDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Gainsboro;

                            SubstrateChangeDataGridView.Columns[0].HeaderCell.Value = "ID";
                            SubstrateChangeDataGridView.Columns[0].Width = 40;
                            SubstrateChangeDataGridView.Columns[1].HeaderCell.Value = "注文番号";
                            SubstrateChangeDataGridView.Columns[2].HeaderCell.Value = "製造番号";
                            SubstrateChangeDataGridView.Columns[2].Width = 130;
                            SubstrateChangeDataGridView.Columns[3].HeaderCell.Value = "製品名";
                            SubstrateChangeDataGridView.Columns[4].HeaderCell.Value = "製品型式";
                            SubstrateChangeDataGridView.Columns[5].HeaderCell.Value = "数量";
                            SubstrateChangeDataGridView.Columns[5].Width = 40;
                            SubstrateChangeDataGridView.Columns[6].HeaderCell.Value = "担当者";
                            SubstrateChangeDataGridView.Columns[6].Width = 70;
                            SubstrateChangeDataGridView.Columns[7].HeaderCell.Value = "登録日";
                            SubstrateChangeDataGridView.Columns[7].Width = 80;
                            SubstrateChangeDataGridView.Columns[8].HeaderCell.Value = "Rev";
                            SubstrateChangeDataGridView.Columns[8].Width = 40;
                            SubstrateChangeDataGridView.Columns[9].HeaderCell.Value = "シリアル先頭";
                            SubstrateChangeDataGridView.Columns[10].HeaderCell.Value = "シリアル末尾";
                            SubstrateChangeDataGridView.Columns[11].HeaderCell.Value = "末番";
                            SubstrateChangeDataGridView.Columns[11].Width = 40;
                            SubstrateChangeDataGridView.Columns[12].HeaderCell.Value = "コメント";
                            SubstrateChangeDataGridView.Columns[13].HeaderCell.Value = "使用基板";

                        }
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }

        private void OpenSubstrateChangeWindow() {

            int _i = SubstrateChangeDataGridView.SelectedCells[0].RowIndex;

            string? orderNumber = SubstrateChangeDataGridView.Rows[_i].Cells[1].Value.ToString();
            string? productNumber = SubstrateChangeDataGridView.Rows[_i].Cells[2].Value.ToString();
            string? productType = SubstrateChangeDataGridView.Rows[_i].Cells[3].Value.ToString();
            string? productModel = SubstrateChangeDataGridView.Rows[_i].Cells[4].Value.ToString();
            string? revision = SubstrateChangeDataGridView.Rows[_i].Cells[8].Value.ToString();
            string? comment = SubstrateChangeDataGridView.Rows[_i].Cells[12].Value.ToString();
            string? useSubstrate = SubstrateChangeDataGridView.Rows[_i].Cells[13].Value.ToString();

            int intQuantity = Convert.ToInt32(SubstrateChangeDataGridView.Rows[_i].Cells[5].Value);
            int intSerialFirstNumber = Convert.ToInt32(SubstrateChangeDataGridView.Rows[_i].Cells[9].Value);

            using SubstrateChange2 _substrateChange2 = new() {
                StrFontName = StrFontName,
                IntFontSize = IntFontSize,
                StrProductName = StrProductName,
                StrOrderNumber = orderNumber ?? string.Empty,
                StrProductNumber = productNumber ?? string.Empty,
                StrProductType = productType ?? string.Empty,
                StrProductModel = productModel ?? string.Empty,
                IntQuantity = intQuantity,
                StrRevision = revision ?? string.Empty,
                IntSerialFirstNumber = intSerialFirstNumber,
                StrComment = comment ?? string.Empty,
                StrUseSubstrate = useSubstrate ?? string.Empty
            };
            _substrateChange2.ShowDialog(this);

        }

        private void SubstrateChange1_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OKButton_Click(object sender, EventArgs e) { OpenSubstrateChangeWindow(); }
    }
}
