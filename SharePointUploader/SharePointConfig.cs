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
    /// Azure AD クライアントシークレット
    /// </summary>
    public string ClientSecret { get; set; } = string.Empty;
}

/// <summary>
/// 設定ファイルのルート構造
/// </summary>
public class AppConfig
{
    public SharePointConfig SharePoint { get; set; } = new SharePointConfig();
}
