# 使用基板フィルター機能 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 製品マスター登録ダイアログの「使用基板」セクションで全基板を表示し、製品名・基板型式によるリアルタイムフィルター機能を追加する。

**Architecture:** `ProductMasterEditDialog.cs` にチェック状態辞書 (`_substrateCheckState`) を追加し、フィルターテキストボックスの `TextChanged` イベントで `LoadSubstrateCheckedList()` を再描画する。チェック済みアイテムはフィルター後も常に表示されるよう辞書で状態を管理する。

**Tech Stack:** C# WinForms (.NET 8), SQLite (Dapper)

---

## ファイル構成

| ファイル | 変更内容 |
|---------|---------|
| `ProductDataBase/MasterManagement/ProductMasterEditDialog.Designer.cs` | フィルター用コントロール追加、既存コントロールのサイズ・位置調整 |
| `ProductDataBase/MasterManagement/ProductMasterEditDialog.cs` | チェック状態辞書追加、LoadSubstrateCheckedList変更、イベントハンドラ追加、保存ロジック変更 |

---

## Task 1: Designer.cs - フィルター用コントロールの追加と位置調整

**Files:**
- Modify: `ProductDataBase/MasterManagement/ProductMasterEditDialog.Designer.cs`

### 変更の全体像

UseSubstrateGroupBox を縦に50px拡張し、内部上部にフィルター用ラベルとテキストボックスを2組追加する。その分 SubstrateCheckedListBox を下にずらし、ボタン類とウィンドウ全体の高さも50px拡張する。

- [ ] **Step 1: UseSubstrateGroupBox の設定を変更する**

`Size = new Size(736, 200)` を以下に変更:

```csharp
this.UseSubstrateGroupBox.Size = new Size(736, 250);
```

- [ ] **Step 2: SubstrateCheckedListBox の位置とサイズを変更する**

```csharp
// 変更前
this.SubstrateCheckedListBox.Location = new Point(8, 22);
this.SubstrateCheckedListBox.Size = new Size(718, 166);

// 変更後
this.SubstrateCheckedListBox.Location = new Point(8, 77);
this.SubstrateCheckedListBox.Size = new Size(718, 161);
```

- [ ] **Step 3: SubstrateCheckedListBox に ItemCheck イベントを追加する**

`TabIndex = 28` の行の直後に追加:

```csharp
this.SubstrateCheckedListBox.ItemCheck += this.SubstrateCheckedListBox_ItemCheck;
```

- [ ] **Step 4: UseSubstrateGroupBox.Controls.Add でフィルター用コントロールを追加する**

`UseSubstrateGroupBox` の Controls.Add 行（現在 `this.UseSubstrateGroupBox.Controls.Add(this.SubstrateCheckedListBox);` のみ）を以下に変更:

```csharp
this.UseSubstrateGroupBox.Controls.Add(this.SubstrateProductNameFilterLabel);
this.UseSubstrateGroupBox.Controls.Add(this.SubstrateProductNameFilterTextBox);
this.UseSubstrateGroupBox.Controls.Add(this.SubstrateModelFilterLabel);
this.UseSubstrateGroupBox.Controls.Add(this.SubstrateModelFilterTextBox);
this.UseSubstrateGroupBox.Controls.Add(this.SubstrateCheckedListBox);
```

- [ ] **Step 5: フィルター用コントロールの初期化コードを追加する**

`UseSubstrateGroupBox` の設定ブロック（`this.UseSubstrateGroupBox.Text = ...` の行の直後）にフィルター用コントロールの設定を追加:

```csharp
//
// SubstrateProductNameFilterLabel
//
this.SubstrateProductNameFilterLabel.AutoSize = true;
this.SubstrateProductNameFilterLabel.Location = new Point(8, 27);
this.SubstrateProductNameFilterLabel.Name = "SubstrateProductNameFilterLabel";
this.SubstrateProductNameFilterLabel.Text = "製品名:";
//
// SubstrateProductNameFilterTextBox
//
this.SubstrateProductNameFilterTextBox.Location = new Point(56, 23);
this.SubstrateProductNameFilterTextBox.Name = "SubstrateProductNameFilterTextBox";
this.SubstrateProductNameFilterTextBox.Size = new Size(200, 23);
this.SubstrateProductNameFilterTextBox.TabIndex = 29;
this.SubstrateProductNameFilterTextBox.TextChanged += this.SubstrateProductNameFilterTextBox_TextChanged;
//
// SubstrateModelFilterLabel
//
this.SubstrateModelFilterLabel.AutoSize = true;
this.SubstrateModelFilterLabel.Location = new Point(268, 27);
this.SubstrateModelFilterLabel.Name = "SubstrateModelFilterLabel";
this.SubstrateModelFilterLabel.Text = "基板型式:";
//
// SubstrateModelFilterTextBox
//
this.SubstrateModelFilterTextBox.Location = new Point(332, 23);
this.SubstrateModelFilterTextBox.Name = "SubstrateModelFilterTextBox";
this.SubstrateModelFilterTextBox.Size = new Size(200, 23);
this.SubstrateModelFilterTextBox.TabIndex = 30;
this.SubstrateModelFilterTextBox.TextChanged += this.SubstrateModelFilterTextBox_TextChanged;
```

- [ ] **Step 6: SaveButton と DialogCancelButton の Y 位置を50px下に移動する**

```csharp
// 変更前
this.SaveButton.Location = new Point(538, 785);
// 変更後
this.SaveButton.Location = new Point(538, 835);

// 変更前
this.DialogCancelButton.Location = new Point(648, 785);
// 変更後
this.DialogCancelButton.Location = new Point(648, 835);
```

- [ ] **Step 7: ClientSize の高さを50px拡張する**

```csharp
// 変更前
this.ClientSize = new Size(760, 832);
// 変更後
this.ClientSize = new Size(760, 882);
```

- [ ] **Step 8: フィールド宣言をセクション5に追加する**

ファイル末尾のフィールド宣言部（`// セクション5` のブロック）を以下に変更:

```csharp
// セクション5
private GroupBox        UseSubstrateGroupBox;
private Label           SubstrateProductNameFilterLabel;
private TextBox         SubstrateProductNameFilterTextBox;
private Label           SubstrateModelFilterLabel;
private TextBox         SubstrateModelFilterTextBox;
private CheckedListBox  SubstrateCheckedListBox;
```

- [ ] **Step 9: ビルドしてエラーがないことを確認する**

Visual Studio または以下のコマンドでビルド:

```
dotnet build ProductDataBase/ProductDatabase.csproj
```

期待結果: ビルド成功（エラー0件）

- [ ] **Step 10: コミットする**

```bash
git add ProductDataBase/MasterManagement/ProductMasterEditDialog.Designer.cs
git commit -m "feat: 使用基板フィルター用コントロールをDesignerに追加"
```

---

## Task 2: ProductMasterEditDialog.cs - ロジック変更

**Files:**
- Modify: `ProductDataBase/MasterManagement/ProductMasterEditDialog.cs`

### 変更の全体像

チェック状態辞書の追加、`LoadSubstrateCheckedList()` の全件表示化、イベントハンドラの追加、保存ロジックの辞書参照化を行う。

- [ ] **Step 1: `_substrateCheckState` フィールドを追加する**

既存フィールド宣言の末尾（`private readonly DataRow? _sourceRow;` の直後）に追加:

```csharp
// 基板チェック状態管理（フィルタリング時もチェック状態を保持するため辞書で管理）
private Dictionary<long, bool> _substrateCheckState = [];
```

- [ ] **Step 2: `InitSubstrateCheckState()` メソッドを追加する**

`LoadSubstrateCheckedList()` メソッドの直前に追加:

```csharp
// チェック状態辞書を初期化する（全基板IDをキーに、編集時は既存紐づきをtrue）
private void InitSubstrateCheckState() {
    _substrateCheckState = _repository.SubstrateDataTable.AsEnumerable()
        .ToDictionary(r => r.Field<long>("SubstrateID"), _ => false);

    if (!_isNewRecord && _sourceRow != null) {
        var productId = _sourceRow.Field<long>("ProductID");
        var linkedIds = new HashSet<long>(_repository.ProductUseSubstrate.AsEnumerable()
            .Where(r => r.Field<long?>("P_ProductID") == productId
                     && r.Field<long?>("S_SubstrateID").HasValue)
            .Select(r => r.Field<long>("S_SubstrateID")));

        foreach (var id in linkedIds)
            if (_substrateCheckState.ContainsKey(id))
                _substrateCheckState[id] = true;
    }
}
```

- [ ] **Step 3: `ProductMasterEditDialog_Load` の呼び出し順を修正する**

現在の `LoadSubstrateCheckedList()` 呼び出しと `ProductNameTextBox.TextChanged` イベント登録を、`InitSubstrateCheckState()` 呼び出しに変更する。

変更前（37〜41行目付近）:

```csharp
// フィールド反映後に製品名でフィルタしてリストを構築する
LoadSubstrateCheckedList();

// 新規追加時：製品名が変わったらリストを再フィルタする
ProductNameTextBox.TextChanged += (s, e) => LoadSubstrateCheckedList();
```

変更後（`ProductNameTextBox.TextChanged` イベント登録も削除する）:

```csharp
// チェック状態辞書を初期化してからリストを構築する
InitSubstrateCheckState();
LoadSubstrateCheckedList();
```

- [ ] **Step 4: `LoadSubstrateCheckedList()` を全件表示ロジックに置き換える**

メソッド全体を以下に置き換える:

```csharp
// 全基板マスターをCheckedListBoxに描画する（チェック済みは常に表示、未チェックはフィルターで絞り込み）
private void LoadSubstrateCheckedList() {
    var productNameFilter = SubstrateProductNameFilterTextBox.Text.Trim();
    var modelFilter = SubstrateModelFilterTextBox.Text.Trim();

    SubstrateCheckedListBox.Items.Clear();

    foreach (DataRow row in _repository.SubstrateDataTable.Rows) {
        var substrateId = row.Field<long>("SubstrateID");
        var productName = row["ProductName"]?.ToString() ?? string.Empty;
        var substrateName = row["SubstrateName"]?.ToString() ?? string.Empty;
        var substrateModel = row["SubstrateModel"]?.ToString() ?? string.Empty;

        var isChecked = _substrateCheckState.GetValueOrDefault(substrateId, false);

        var matchesFilter =
            (string.IsNullOrEmpty(productNameFilter) ||
             productName.Contains(productNameFilter, StringComparison.OrdinalIgnoreCase)) &&
            (string.IsNullOrEmpty(modelFilter) ||
             substrateModel.Contains(modelFilter, StringComparison.OrdinalIgnoreCase));

        if (!isChecked && !matchesFilter) continue;

        var item = new ListItem<long> {
            Id = substrateId,
            Name = $"{productName} - {substrateName} [{substrateModel}]"
        };
        SubstrateCheckedListBox.Items.Add(item, isChecked);
    }
}
```

- [ ] **Step 5: イベントハンドラを追加する**

`LoadSubstrateCheckedList()` メソッドの直後に追加:

```csharp
// チェック状態の変化を辞書に同期する
private void SubstrateCheckedListBox_ItemCheck(object sender, ItemCheckEventArgs e) {
    var item = (ListItem<long>)SubstrateCheckedListBox.Items[e.Index];
    _substrateCheckState[item.Id] = (e.NewValue == CheckState.Checked);
}

// フィルターテキスト変更時にリストを再描画する
private void SubstrateProductNameFilterTextBox_TextChanged(object sender, EventArgs e)
    => LoadSubstrateCheckedList();

private void SubstrateModelFilterTextBox_TextChanged(object sender, EventArgs e)
    => LoadSubstrateCheckedList();
```

- [ ] **Step 6: 保存処理のチェック済みID取得を辞書参照に変更する**

`SaveButton_Click` 内の以下のコード（175〜180行目付近）:

```csharp
// 使用基板の紐づけ更新
var selectedSubstrateIds = SubstrateCheckedListBox.CheckedItems
    .Cast<ListItem<long>>()
    .Select(item => item.Id)
    .ToList();
_repository.UpdateProductUseSubstrates(productId, selectedSubstrateIds);
```

を以下に変更:

```csharp
// 使用基板の紐づけ更新（辞書から取得することでフィルター非表示のチェック済みも含める）
var selectedSubstrateIds = _substrateCheckState
    .Where(kvp => kvp.Value)
    .Select(kvp => kvp.Key)
    .ToList();
_repository.UpdateProductUseSubstrates(productId, selectedSubstrateIds);
```

- [ ] **Step 7: ビルドしてエラーがないことを確認する**

```
dotnet build ProductDataBase/ProductDatabase.csproj
```

期待結果: ビルド成功（エラー0件）

- [ ] **Step 8: 動作確認チェックリスト**

以下を手動で確認する:

1. ダイアログを開いたとき全基板が表示される
2. 製品名フィルターに文字を入力すると一致する基板のみ表示される
3. 基板型式フィルターに文字を入力すると一致する基板のみ表示される
4. フィルター後にチェックを入れ、フィルターをクリアしても該当基板がチェック済みのまま表示される
5. チェック済みの基板はフィルターで非一致になっても消えない
6. 保存後、紐づきが正しく更新されている
7. 編集ダイアログを再度開いたとき既存の紐づきがチェック済みで表示される

- [ ] **Step 9: コミットする**

```bash
git add ProductDataBase/MasterManagement/ProductMasterEditDialog.cs
git commit -m "feat: 使用基板選択を全基板表示＋フィルター機能に変更"
```
