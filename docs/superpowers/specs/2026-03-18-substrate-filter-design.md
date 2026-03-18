# 設計ドキュメント: 製品マスター登録ダイアログ 使用基板フィルター機能

作成日: 2026-03-18

## 概要

製品マスター登録ダイアログの「使用基板」セクションにおいて、全基板を対象に選択できるよう変更し、製品名・基板型式によるフィルター機能を追加する。

## 背景・動機

現状の `LoadSubstrateCheckedList()` は入力した製品名と一致する基板のみを表示している。全基板から選択できるようにすることで、柔軟な基板紐づけが可能になる。

## 要件

1. 全基板（M_SubstrateDef の全レコード）を表示対象にする（製品名による自動絞り込みを廃止）
2. UseSubstrateGroupBox 内の上部にフィルター欄（テキストボックス2つ）を配置する
   - 製品名フィルター用テキストボックス
   - 基板型式フィルター用テキストボックス
3. 各項目の表示フォーマット: `{ProductName} - {SubstrateName} [{SubstrateModel}]`
4. フィルターはリアルタイム（TextChanged イベント）で反映する
5. チェック済みのアイテムはフィルター結果に関わらず常に表示する

## UI 変更

### UseSubstrateGroupBox 内レイアウト

```
[製品名フィルター: ________________] [基板型式フィルター: ________________]
[                                                                        ]
[  SubstrateCheckedListBox（全基板リスト）                                ]
[                                                                        ]
```

追加コントロール:
- `SubstrateProductNameFilterLabel` (Label)
- `SubstrateProductNameFilterTextBox` (TextBox)
- `SubstrateModelFilterLabel` (Label)
- `SubstrateModelFilterTextBox` (TextBox)

既存コントロールの変更:
- `SubstrateCheckedListBox` のY位置を約50px下へ移動
- `SubstrateCheckedListBox` の高さを約50px縮小（グループボックス内に収める）
- `UseSubstrateGroupBox` の高さを約50px拡張

## ロジック変更

### 追加フィールド (ProductMasterEditDialog.cs)

```csharp
private Dictionary<long, bool> _substrateCheckState = [];
```

全基板IDをキー、チェック状態を値として管理する。フィルタリングでリストを再描画してもチェック状態が失われない。

### InitSubstrateCheckState() の追加

ダイアログロード時に呼び出す。全基板IDで辞書を初期化し、編集時は既存の紐づき基板IDをチェック済みにする。

```csharp
private void InitSubstrateCheckState()
{
    _substrateCheckState = _repository.SubstrateDataTable.AsEnumerable()
        .ToDictionary(r => r.Field<long>("SubstrateID"), _ => false);

    if (!_isNewRecord && _sourceRow != null)
    {
        var productId = _sourceRow.Field<long>("ProductID");
        var linkedIds = _repository.ProductUseSubstrate.AsEnumerable()
            .Where(r => r.Field<long?>("P_ProductID") == productId
                     && r.Field<long?>("S_SubstrateID").HasValue)
            .Select(r => r.Field<long>("S_SubstrateID"))
            .ToHashSet();

        foreach (var id in linkedIds)
            if (_substrateCheckState.ContainsKey(id))
                _substrateCheckState[id] = true;
    }
}
```

### LoadSubstrateCheckedList() の変更

製品名による自動絞り込みを廃止し、全件を対象にする。表示条件は「チェック済み OR フィルターに一致」。

```csharp
private void LoadSubstrateCheckedList()
{
    var productNameFilter = SubstrateProductNameFilterTextBox.Text.Trim();
    var modelFilter = SubstrateModelFilterTextBox.Text.Trim();

    SubstrateCheckedListBox.Items.Clear();

    foreach (DataRow row in _repository.SubstrateDataTable.Rows)
    {
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

        var item = new ListItem<long>
        {
            Id = substrateId,
            Name = $"{productName} - {substrateName} [{substrateModel}]"
        };
        SubstrateCheckedListBox.Items.Add(item, isChecked);
    }
}
```

### SubstrateCheckedListBox_ItemCheck イベントハンドラの追加

チェック状態の変化を辞書に同期する。

```csharp
private void SubstrateCheckedListBox_ItemCheck(object sender, ItemCheckEventArgs e)
{
    var item = (ListItem<long>)SubstrateCheckedListBox.Items[e.Index];
    _substrateCheckState[item.Id] = (e.NewValue == CheckState.Checked);
}
```

### フィルターテキストボックスの TextChanged イベントハンドラ

```csharp
private void SubstrateProductNameFilterTextBox_TextChanged(object sender, EventArgs e)
    => LoadSubstrateCheckedList();

private void SubstrateModelFilterTextBox_TextChanged(object sender, EventArgs e)
    => LoadSubstrateCheckedList();
```

### 保存処理の変更

保存時は CheckedListBox のアイテムではなく辞書からチェック済みIDを取得する。

```csharp
var selectedSubstrateIds = _substrateCheckState
    .Where(kvp => kvp.Value)
    .Select(kvp => kvp.Key)
    .ToList();
_repository.UpdateProductUseSubstrates(productId, selectedSubstrateIds);
```

## 変更対象ファイル

| ファイル | 変更内容 |
|---------|---------|
| `ProductMasterEditDialog.cs` | `_substrateCheckState` フィールド追加、`InitSubstrateCheckState()` 追加、`LoadSubstrateCheckedList()` 変更、イベントハンドラ追加、保存処理変更 |
| `ProductMasterEditDialog.Designer.cs` | フィルター用ラベル・テキストボックス追加、既存コントロールの位置・サイズ調整 |

## 影響範囲

- 既存の保存・削除ロジックへの影響なし
- `UpdateProductUseSubstrates()` の呼び出し元変更のみ（インタフェースは同一）
- 新規登録・編集の両方で動作する
