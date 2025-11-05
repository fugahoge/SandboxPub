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
5. 「APIのアクセス許可」で以下を追加：
   - SharePoint → Application permissions → Sites.ReadWrite.All
   - 「管理者の同意を付与」をクリック

以下の情報をメモしてください：
- **テナントID** (Azure ADの概要ページで確認)
- **クライアントID** (アプリケーションIDとも呼ばれる)

#### 認証方法の選択

このツールは2つの認証方法をサポートしています：

**A. 証明書認証（推奨）**

セキュリティが高く、本番環境での使用に推奨されます。

**B. クライアントシークレット認証**

開発・テスト環境での使用に適しています。

### 2. 証明書の作成（証明書認証を使用する場合）

#### 自己署名証明書の作成

PowerShell（Windows）またはOpenSSL（Linux/Mac）で証明書を作成できます。

**PowerShellの場合（Windows）:**

```powershell
# 証明書を作成
$cert = New-SelfSignedCertificate -Subject "CN=SharePointUploader" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# 証明書のThumbprintを表示
$cert.Thumbprint

# PFXファイルとしてエクスポート（パスワード保護）
$password = ConvertTo-SecureString -String "YourPassword123" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath ".\SharePointUploader.pfx" -Password $password

# 公開鍵（.cer）をエクスポート（Azure ADにアップロード用）
Export-Certificate -Cert $cert -FilePath ".\SharePointUploader.cer"
```

**OpenSSLの場合（Linux/Mac）:**

```bash
# 秘密鍵と証明書を作成
openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 730 -nodes \
    -subj "/CN=SharePointUploader"

# PFXファイルを作成（パスワード保護）
openssl pkcs12 -export -out SharePointUploader.pfx -inkey key.pem -in cert.pem \
    -password pass:YourPassword123

# 公開鍵（.cer）を作成（Azure ADにアップロード用）
openssl x509 -outform der -in cert.pem -out SharePointUploader.cer
```

#### Azure ADに証明書をアップロード

1. Azure Portalで作成したアプリケーションを開く
2. 「証明書とシークレット」→「証明書」タブ
3. 「証明書のアップロード」をクリック
4. 作成した `.cer` ファイルをアップロード

### 3. クライアントシークレットの作成（クライアントシークレット認証を使用する場合）

1. Azure Portalで作成したアプリケーションを開く
2. 「証明書とシークレット」→「新しいクライアントシークレット」を作成
3. **クライアントシークレットの値**をメモ（作成時にのみ表示されます）

### 4. 設定ファイルの準備

```bash
cd SharePointUploader
cp config.sample.json config.json
```

`config.json` を編集して、SharePointとAzure ADの情報を設定します。

#### 証明書認証を使用する場合（推奨）:

```json
{
  "SharePoint": {
    "SiteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
    "LibraryName": "Documents",
    "FolderPath": "Shared Documents/UploadFolder",
    "TenantId": "your-tenant-id",
    "ClientId": "your-client-id",
    "CertificatePath": "SharePointUploader.pfx",
    "CertificatePassword": "YourPassword123"
  }
}
```

#### クライアントシークレット認証を使用する場合:

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
- **CertificatePath**: クライアント証明書ファイルのパス（.pfx形式）※証明書認証の場合
- **CertificatePassword**: クライアント証明書のパスワード ※証明書認証の場合
- **ClientSecret**: Azure ADクライアントシークレット ※クライアントシークレット認証の場合

**注意**: CertificatePathが設定されている場合、証明書認証が優先されます。

### 5. ビルド

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

- `config.json` には機密情報（証明書パスワードまたはクライアントシークレット）が含まれています
- `config.json` はGitにコミットしないでください（.gitignoreに追加済み）
- 証明書ファイル（.pfx）も機密情報なので適切に管理してください
- 本番環境では証明書認証の使用を強く推奨します
- クライアントシークレットや証明書は定期的に更新することを推奨します
- 本番環境では、Azure Key Vaultなどのシークレット管理サービスの使用を推奨します

### 証明書認証 vs クライアントシークレット認証

| 項目 | 証明書認証 | クライアントシークレット認証 |
|------|-----------|---------------------------|
| セキュリティ | 高い（秘密鍵が漏洩しにくい） | 中程度（文字列が漏洩しやすい） |
| 推奨環境 | 本番環境 | 開発・テスト環境 |
| 有効期限 | 証明書の有効期限（通常1-2年） | シークレットの有効期限（最大2年） |
| 管理の複雑さ | やや複雑 | シンプル |

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。

## 使用しているライブラリ

- [Microsoft.Graph](https://github.com/microsoftgraph/msgraph-sdk-dotnet) - Microsoft Graph API経由でSharePointにアクセス
- [Azure.Identity](https://github.com/Azure/azure-sdk-for-net/tree/main/sdk/identity/Azure.Identity) - Azure AD認証
- [Microsoft.Extensions.Configuration](https://github.com/dotnet/runtime) - 設定ファイル読み込み

すべてMicrosoft公式のライブラリを使用しています。