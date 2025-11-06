namespace SharePointUploader;

/// <summary>
/// SharePoint接続設定を保持するクラス
/// </summary>
public class SharePointConfig
{
  /// <summary>
  /// SharePointサイトのURL
  /// </summary>
  public string SiteUrl { get; set; } = string.Empty;

  /// <summary>
  /// ドキュメントライブラリ名
  /// </summary>
  public string LibraryName { get; set; } = string.Empty;

  /// <summary>
  /// アップロード先フォルダパス
  /// </summary>
  public string FolderPath { get; set; } = string.Empty;

  /// <summary>
  /// Azure AD テナントID
  /// </summary>
  public string TenantId { get; set; } = string.Empty;

  /// <summary>
  /// Azure AD クライアントID (アプリケーションID)
  /// </summary>
  public string ClientId { get; set; } = string.Empty;

  /// <summary>
  /// Azure AD クライアントシークレット (証明書認証を使用しない場合)
  /// </summary>
  public string? ClientSecret { get; set; }

  /// <summary>
  /// クライアント証明書ファイルのパス (.pfx形式)
  /// </summary>
  public string? CertificatePath { get; set; }

  /// <summary>
  /// クライアント証明書のパスワード
  /// </summary>
  public string? CertificatePassword { get; set; }
}

/// <summary>
/// 設定ファイルのルート構造
/// </summary>
public class AppConfig
{
  public SharePointConfig SharePoint { get; set; } = new SharePointConfig();
}
