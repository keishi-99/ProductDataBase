using ProductDatabase.Models;
using System.ComponentModel;
using System.Data;

namespace ProductDatabase {
    // 属性の組み合わせが多い製品カテゴリ向けに、電源/設置形態/ADI/通信方式/バージョンを
    // 1つずつ選ばせて候補を1件に絞り込むダイアログ
    public partial class ProductAttributeSelectWindow : Form {

        private sealed record Candidate(DataRow Row, ProductModelAttributes Attributes);

        private static readonly ListItem<string> Unselected = new() { Id = "", Name = "(未選択)" };

        private static readonly ListItem<string>[] PowerOptions = [
            Unselected,
            new() { Id = "AC", Name = "AC" },
            new() { Id = "DC", Name = "DC" },
        ];
        private static readonly ListItem<string>[] MountOptions = [
            Unselected,
            new() { Id = "一体形", Name = "一体形" },
            new() { Id = "別置形", Name = "別置形" },
        ];
        private static readonly ListItem<string>[] AdiOptions = [
            Unselected,
            new() { Id = "NONE", Name = "無し" },
            new() { Id = "ADI", Name = "ADI" },
        ];
        private static readonly ListItem<string>[] CommOptions = [
            Unselected,
            new() { Id = "HART", Name = "HART" },
            new() { Id = "MODBUS", Name = "MODBUS" },
        ];
        private static readonly ListItem<string>[] VerOptions = [
            Unselected,
            new() { Id = "NONE", Name = "無し" },
            new() { Id = "C", Name = "Ver.C" },
            new() { Id = "D", Name = "Ver.D" },
        ];

        private readonly List<Candidate> _allCandidates;
        private Candidate? _narrowedCandidate;
        private bool _isUpdating;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public long SelectedProductId { get; private set; }

        // candidateRowsには同一カテゴリ・同一製品名で仕様(ProductType)違いの複数行を渡す
        public ProductAttributeSelectWindow(IEnumerable<DataRow> candidateRows) {
            InitializeComponent();

            _allCandidates = candidateRows
                .Select(row => (Row: row, Attributes: ProductModelAttributes.TryParse(row["ProductModel"]?.ToString() ?? string.Empty)))
                .Where(x => x.Attributes is not null)
                .Select(x => new Candidate(x.Row, x.Attributes!))
                .ToList();
        }

        // 各リストボックスの現在の選択値を返す（未選択ならnull）
        private static string? SelectedId(ListBox listBox) =>
            listBox.SelectedItem is ListItem<string> { Id.Length: > 0 } item ? item.Id : null;

        // ADIリストボックスの選択値(未選択/ADI/NONE)と候補の一致判定
        private static bool MatchesAdi(Candidate candidate, string? selectedAdi) => selectedAdi switch {
            null => true,
            "ADI" => candidate.Attributes.HasAdi,
            _ => !candidate.Attributes.HasAdi
        };

        // バージョンリストボックスの選択値(未選択/NONE/C/D)と候補の一致判定
        private static bool MatchesVersion(Candidate candidate, string? selectedVer) => selectedVer switch {
            null => true,
            "NONE" => candidate.Attributes.Version is null,
            _ => candidate.Attributes.Version == selectedVer
        };

        // 指定した属性条件(未選択はnullで無視)で_allCandidatesを絞り込む
        private IEnumerable<Candidate> FilterCandidates(string? power, string? mount, string? adi, string? comm, string? ver) =>
            _allCandidates.Where(c =>
                (power is null || c.Attributes.Power == power) &&
                (mount is null || c.Attributes.Mount == mount) &&
                MatchesAdi(c, adi) &&
                (comm is null || c.Attributes.Communication == comm) &&
                MatchesVersion(c, ver));

        // 選択状態に応じて候補を絞り込み、各リストボックスの選択肢を更新する
        private void RefreshAvailableOptions() {
            var selectedPower = SelectedId(PowerListBox);
            var selectedMount = SelectedId(MountListBox);
            var selectedAdi = SelectedId(AdiListBox);
            var selectedComm = SelectedId(CommListBox);
            var selectedVer = SelectedId(VerListBox);

            var narrowed = FilterCandidates(selectedPower, selectedMount, selectedAdi, selectedComm, selectedVer).ToList();
            _narrowedCandidate = narrowed.Count == 1 ? narrowed[0] : null;
            OKButton.Enabled = _narrowedCandidate is not null;

            _isUpdating = true;

            UpdateListBoxOptions(PowerListBox, PowerOptions,
                FilterCandidates(null, selectedMount, selectedAdi, selectedComm, selectedVer).Select(c => c.Attributes.Power));
            UpdateListBoxOptions(MountListBox, MountOptions,
                FilterCandidates(selectedPower, null, selectedAdi, selectedComm, selectedVer).Select(c => c.Attributes.Mount));
            UpdateListBoxOptions(AdiListBox, AdiOptions,
                FilterCandidates(selectedPower, selectedMount, null, selectedComm, selectedVer).Select(c => c.Attributes.HasAdi ? "ADI" : "NONE"));
            UpdateListBoxOptions(CommListBox, CommOptions,
                FilterCandidates(selectedPower, selectedMount, selectedAdi, null, selectedVer).Select(c => c.Attributes.Communication));
            UpdateListBoxOptions(VerListBox, VerOptions,
                FilterCandidates(selectedPower, selectedMount, selectedAdi, selectedComm, null).Select(c => c.Attributes.Version ?? "NONE"));

            _isUpdating = false;
        }

        // 実在する値(availableIds)＋未選択を選択肢として詰め直す。それまでの選択がまだ含まれていれば維持する
        private static void UpdateListBoxOptions(ListBox listBox, ListItem<string>[] allOptions, IEnumerable<string> availableIds) {
            var availableSet = availableIds.ToHashSet();
            var currentSelection = SelectedId(listBox);

            var newItems = allOptions.Where(o => o == Unselected || availableSet.Contains(o.Id)).ToArray();

            listBox.Items.Clear();
            listBox.Items.AddRange(newItems);
            listBox.SelectedItem = newItems.FirstOrDefault(o => o.Id == currentSelection) ?? Unselected;
        }

        private void ProductAttributeSelectWindow_Load(object sender, EventArgs e) {
            _isUpdating = true;
            PowerListBox.Items.AddRange(PowerOptions);
            MountListBox.Items.AddRange(MountOptions);
            AdiListBox.Items.AddRange(AdiOptions);
            CommListBox.Items.AddRange(CommOptions);
            VerListBox.Items.AddRange(VerOptions);
            PowerListBox.SelectedIndex = 0;
            MountListBox.SelectedIndex = 0;
            AdiListBox.SelectedIndex = 0;
            CommListBox.SelectedIndex = 0;
            VerListBox.SelectedIndex = 0;
            _isUpdating = false;
            RefreshAvailableOptions();
        }

        private void AttributeListBox_SelectedIndexChanged(object sender, EventArgs e) {
            if (_isUpdating) { return; }
            RefreshAvailableOptions();
        }

        private void OKButton_Click(object sender, EventArgs e) {
            if (_narrowedCandidate is null) { return; }
            SelectedProductId = _narrowedCandidate.Row.Field<long>("ProductID");
            DialogResult = DialogResult.OK;
            Close();
        }

        private void CancelButton_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
