# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

製品・基板の登録、シリアル番号生成・印刷、Excel帳票出力を行う Windows デスクトップアプリ（WinForms）。バーコード/QRスキャンとラベル印刷にも対応。

## ソリューション構成

`ProductDataBase.sln` はマルチプロジェクト構成:

| プロジェクト | フレームワーク | 役割 |
|---|---|---|
| `ProductDataBase` | net8.0-windows | メインアプリ（WinForms） |
| `Launcher` | net10.0-windows | 起動ランチャー（WinForms） |
| `ProductWebViewer` | net10.0 | Web閲覧コンポーネント（ASP.NET Core） |

**注意:** `ProductDataBase` は .NET 8、他2プロジェクトは .NET 10。API互換性に注意。

## ビルド・実行コマンド

```bash
# ソリューション全体のビルド
dotnet build ProductDataBase.sln

# メインアプリ起動
dotnet run --project ProductDataBase/ProductDatabase.csproj

# ランチャー起動
dotnet run --project Launcher/Launcher.csproj

# Webビューア起動（ポート5009）
dotnet run --project ProductWebViewer
```

テストプロジェクトは存在しない。

## データアクセス

- **DB:** SQLite（`ProductRegistry.db`）
- **ORM:** Dapper（EF Core は不使用）
- **マイグレーション:** EF Core マイグレーションは使わない。DDL は直接 SQL で管理。
- **リポジトリ:** `Data/` 配下に各エンティティごとのリポジトリクラス。

スキーマ変更が必要な場合は、直接 SQL を書いてリポジトリを更新する。

## COM インターオップ依存

以下の COM コンポーネントが必要（実行環境にインストール済み前提）:

- **Brother bPac SDK** — ラベル印刷（`Print/` 配下の `PrintManager.cs`）
- **Microsoft.Office.Interop.Excel** — 生成した Excel ファイルを開く

これらは開発環境にインストールされていない場合、実行時エラーになる。COM 関連コードを変更するときは慎重に。

## 設定ファイル

- **メインアプリ:** `Config/General/appsettings.json`（アプリ実行ディレクトリ以下）
- **Webビューア:** `ProductWebViewer/appsettings.json`（`DatabasePath` でSQLiteパス指定）
- `appsettings.Development.json` にはハードコードされたローカルパスが含まれる。コミットに注意。

主要設定キー: `BackupFolderPath`, `Persons`, `Administrators`, `AuthorizedUsers`

## シリアル番号のロジック

`SerialType` と `SerialPrintType` はビットフラグを使った特殊な値体系。`sql.md` に定義あり。変更前に必ず参照すること。

## ブランチ・コミット規約

- `main` への直接コミット・プッシュは禁止。PR 経由でマージ。
- ブランチ命名: `feature/<名前>` / `fix/<名前>`
- コミットメッセージ: 日本語、Conventional Commits 形式（`feat:`, `fix:`, `refactor:` など）

## Windows 専用制約

`net8.0-windows` / `net10.0-windows` のため、Linux/macOS では動作しない。WinForms・Windows API 依存のコードを追加する際は TFM を確認すること。
