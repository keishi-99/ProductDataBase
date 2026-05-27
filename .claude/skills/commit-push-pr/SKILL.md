---
name: commit-push-pr
description: このプロジェクトの規約（日本語Conventional Commits、feature/fixブランチ命名）に従ってコミット・プッシュ・PRを作成する。バンドル版commit-push-prのプロジェクト特化版。
disable-model-invocation: true
---

# commit-push-pr スキル（プロジェクト固有版）

## $ARGUMENTS

PRタイトルや変更内容の補足（省略可）

---

## 実行手順

### 1. ブランチ確認

```bash
git branch --show-current
```

- `feature/<名前>` または `fix/<名前>` 形式であること
- `main` ブランチに直接コミットしない（禁止）

ブランチが `main` の場合は作業を中止し、ユーザーに適切なブランチへの切り替えを促す。

### 2. 変更内容の確認

```bash
git status
git diff --staged
git diff
```

コミットに含める変更を把握する。

### 3. コミットメッセージの作成

**形式:** `<type>: <日本語の説明>`

**typeの選択:**
- `feat:` — 新機能追加
- `fix:` — バグ修正
- `refactor:` — 動作を変えないコード改善
- `docs:` — ドキュメントのみの変更
- `chore:` — ビルド設定・ツールなどの変更
- `test:` — テストの追加・修正
- `style:` — フォーマット・空白などの変更（動作に影響なし）

**例:**
- `feat: T_Productに備考カラムを追加`
- `fix: 再印刷時にSerialPrintTypeが正しく保存されない問題を修正`
- `refactor: CSVエクスポートのメモリ効率を改善`

説明は日本語で、変更の「何を」ではなく「なぜ・何のために」を重視する。

### 4. ステージング・コミット

```bash
git add <対象ファイル>
git commit -m "<type>: <日本語の説明>"
```

`.env`、`appsettings.Development.json`（ハードコードされたパスが含まれる）は**コミットしない**。

### 5. プッシュ

```bash
git push -u origin <ブランチ名>
```

### 6. PR 作成

```bash
gh pr create \
  --title "<type>: <日本語の説明>" \
  --body "$(cat <<'EOF'
## 変更内容
- <箇条書きで変更点>

## 確認方法
- <テスト・動作確認の手順>

🤖 Generated with [Claude Code](https://claude.ai/code)
EOF
)"
```

PR本文は日本語で記述する。

### 7. PR URL の表示

作成されたPR URLをユーザーに表示する。
