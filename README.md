# SharePoint Online ファイルアップローダー

SharePoint Onlineにファイルをアップロードするコマンドラインツールです。

## 機能

- コマンドラインから簡単にファイルをSharePoint Onlineにアップロード
- 設定ファイルでアップロード先やアカウント情報を管理
- 自動的にフォルダを作成
- ファイルサイズの表示
- エラーハンドリング

## 必要な環境

- .NET 6.0 以降
- SharePoint Onlineのサイト
- Azure AD アプリケーションの登録

## セットアップ

### 1. Azure ADアプリケーションの登録

SharePoint Onlineにアクセスするため、Azure ADにアプリケーションを登録する必要があります。

1. [Azure Portal](https://portal.azure.com) にアクセス
2. 「Azure Active Directory」→「アプリの登録」→「新規登録」
3. アプリケーション名を入力（例：SharePointUploader）
4. 「登録」をクリック
5. 「証明書とシークレット」→「新しいクライアントシークレット」を作成
6. 「APIのアクセス許可」で以下を追加：
   - SharePoint → Application permissions → Sites.ReadWrite.All
   - 「管理者の同意を付与」をクリック

以下の情報をメモしてください：
- **テナントID** (Azure ADの概要ページで確認)
- **クライアントID** (アプリケーションIDとも呼ばれる)
- **クライアントシークレット** (作成時にのみ表示されます)

### 2. 設定ファイルの準備

```bash
cd SharePointUploader
cp config.sample.json config.json
```

`config.json` を編集して、SharePointとAzure ADの情報を設定します：

```json
{
  "SharePoint": {
    "SiteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
    "LibraryName": "Documents",
    "FolderPath": "Shared Documents/UploadFolder",
    "TenantId": "your-tenant-id",
    "ClientId": "your-client-id",
    "ClientSecret": "your-client-secret"
  }
}
```

#### 設定項目の説明

- **SiteUrl**: SharePointサイトのURL
- **LibraryName**: ドキュメントライブラリ名（通常は "Documents"）
- **FolderPath**: アップロード先のフォルダパス
- **TenantId**: Azure ADテナントID
- **ClientId**: Azure ADアプリケーションID
- **ClientSecret**: Azure ADクライアントシークレット

### 3. ビルド

```bash
cd SharePointUploader
dotnet build
```

または、リリースビルド：

```bash
dotnet build -c Release
```

## 使用方法

### 基本的な使い方

```bash
dotnet run --project SharePointUploader <ファイルパス>
```

または、ビルド後の実行ファイルを直接実行：

```bash
./SharePointUploader/bin/Debug/net6.0/SharePointUploader <ファイルパス>
```

### 使用例

```bash
# カレントディレクトリのファイルをアップロード
dotnet run --project SharePointUploader document.pdf

# 絶対パスでファイルを指定
dotnet run --project SharePointUploader /home/user/reports/monthly-report.docx

# 相対パスでファイルを指定
dotnet run --project SharePointUploader ../files/presentation.pptx
```

### 出力例

```
SharePoint Online ファイルアップローダー
=========================================

ファイルをアップロード中: document.pdf
アップロード先: https://yourtenant.sharepoint.com/sites/yoursite/Shared Documents/UploadFolder
✓ アップロード成功: document.pdf
  サイズ: 1.25 MB
```

## トラブルシューティング

### 認証エラー

- Azure ADアプリケーションの設定を確認してください
- クライアントシークレットが正しいか確認してください
- APIアクセス許可が正しく設定され、管理者の同意が付与されているか確認してください

### アップロードエラー

- SharePointサイトのURLが正しいか確認してください
- フォルダパスが正しいか確認してください
- ネットワーク接続を確認してください

### 設定ファイルが見つからない

```
エラー: 設定ファイルが見つかりません: config.json
```

このエラーが表示された場合、`config.sample.json` を `config.json` にコピーして設定を編集してください。

## セキュリティに関する注意

- `config.json` には機密情報（クライアントシークレット）が含まれています
- `config.json` はGitにコミットしないでください（.gitignoreに追加済み）
- クライアントシークレットは定期的に更新することを推奨します
- 本番環境では、Azure Key Vaultなどのシークレット管理サービスの使用を推奨します

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。

## 使用しているライブラリ

- [PnP.Framework](https://github.com/pnp/pnpframework) - SharePoint接続とファイル操作
- [Microsoft.Identity.Client](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet) - Azure AD認証
- [Microsoft.Extensions.Configuration](https://github.com/dotnet/runtime) - 設定ファイル読み込み