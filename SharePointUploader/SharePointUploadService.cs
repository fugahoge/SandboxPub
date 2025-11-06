using Microsoft.Graph;
using Microsoft.Graph.Models;
using Azure.Identity;
using System.Security.Cryptography.X509Certificates;

namespace SharePointUploader;

/// <summary>
/// SharePoint Onlineへのファイルアップロードサービス (Microsoft Graph API使用)
/// </summary>
public class SharePointUploadService
{
  private readonly SharePointConfig _config;
  private readonly GraphServiceClient _graphClient;

  public SharePointUploadService(SharePointConfig config)
  {
    _config = config ?? throw new ArgumentNullException(nameof(config));
    ValidateConfig();
    _graphClient = CreateGraphClient();
  }

  /// <summary>
  /// 設定の妥当性を検証
  /// </summary>
  private void ValidateConfig()
  {
    if (string.IsNullOrWhiteSpace(_config.SiteUrl))
      throw new InvalidOperationException("SiteUrl が設定されていません。");

    if (string.IsNullOrWhiteSpace(_config.TenantId))
      throw new InvalidOperationException("TenantId が設定されていません。");

    if (string.IsNullOrWhiteSpace(_config.ClientId))
      throw new InvalidOperationException("ClientId が設定されていません。");

    // 証明書認証またはクライアントシークレット認証のいずれかが必要
    bool hasCertificate = !string.IsNullOrWhiteSpace(_config.CertificatePath);
    bool hasClientSecret = !string.IsNullOrWhiteSpace(_config.ClientSecret);

    if (!hasCertificate && !hasClientSecret)
      throw new InvalidOperationException(
        "CertificatePath または ClientSecret のいずれかを設定してください。");

    if (hasCertificate && !System.IO.File.Exists(_config.CertificatePath))
      throw new InvalidOperationException(
        $"証明書ファイルが見つかりません: {_config.CertificatePath}");

    if (string.IsNullOrWhiteSpace(_config.FolderPath))
      throw new InvalidOperationException("FolderPath が設定されていません。");
  }

  /// <summary>
  /// Graph APIクライアントを作成
  /// </summary>
  private GraphServiceClient CreateGraphClient()
  {
    var scopes = new[] { "https://graph.microsoft.com/.default" };

    // 証明書認証を優先
    if (!string.IsNullOrWhiteSpace(_config.CertificatePath))
    {
      Console.WriteLine("認証方法: クライアント証明書");
      Console.WriteLine($"証明書パス: {_config.CertificatePath}");

      // 証明書をロード
      var certificate = LoadCertificate(_config.CertificatePath, _config.CertificatePassword);

      var credential = new ClientCertificateCredential(
        _config.TenantId,
        _config.ClientId,
        certificate);

      return new GraphServiceClient(credential, scopes);
    }
    else
    {
      Console.WriteLine("認証方法: クライアントシークレット");

      var credential = new ClientSecretCredential(
        _config.TenantId,
        _config.ClientId,
        _config.ClientSecret);

      return new GraphServiceClient(credential, scopes);
    }
  }

  /// <summary>
  /// 証明書ファイルをロード
  /// </summary>
  private X509Certificate2 LoadCertificate(string certificatePath, string? password)
  {
    try
    {
      // パスワードがある場合とない場合で処理を分ける
      if (!string.IsNullOrWhiteSpace(password))
      {
        return new X509Certificate2(certificatePath, password);
      }
      else
      {
        return new X509Certificate2(certificatePath);
      }
    }
    catch (Exception ex)
    {
      throw new InvalidOperationException(
        $"証明書のロードに失敗しました: {certificatePath}", ex);
    }
  }

  /// <summary>
  /// ファイルをSharePoint Onlineにアップロード
  /// </summary>
  /// <param name="filePath">アップロードするファイルのパス</param>
  /// <returns>成功した場合true</returns>
  public async Task<bool> UploadFileAsync(string filePath)
  {
    if (!System.IO.File.Exists(filePath))
    {
      Console.WriteLine($"エラー: ファイルが見つかりません: {filePath}");
      return false;
    }

    try
    {
      var fileName = Path.GetFileName(filePath);
      Console.WriteLine($"ファイルをアップロード中: {fileName}");
      Console.WriteLine($"アップロード先: {_config.SiteUrl}/{_config.FolderPath}");

      // SharePointサイトのホスト名とサイトパスを取得
      var siteInfo = ParseSiteUrl(_config.SiteUrl);

      // サイトIDを取得
      Site? site = null;

      try
      {
        // サイトパスが指定されている場合（例: /sites/sitename）
        if (!string.IsNullOrEmpty(siteInfo.sitePath))
        {
          // サイトパスから先頭のスラッシュを除去
          var sitePath = siteInfo.sitePath.TrimStart('/');
          
          // サイト識別子を構築: hostname:/sites/sitename
          var siteIdentifier = $"{siteInfo.hostName}:/{sitePath}";
          
          // サブサイトを取得
          site = await _graphClient.Sites[siteIdentifier]
            .GetAsync(requestConfig =>
            {
              requestConfig.QueryParameters.Select = new[] { "id", "webUrl" };
            });
        }
        else
        {
          // ルートサイトを取得
          site = await _graphClient.Sites[siteInfo.hostName]
            .GetAsync(requestConfig =>
            {
              requestConfig.QueryParameters.Select = new[] { "id", "webUrl" };
            });
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine($"エラー: SharePointサイトの取得に失敗しました: {ex.Message}");
        return false;
      }

      if (site == null || string.IsNullOrEmpty(site.Id))
      {
        Console.WriteLine($"エラー: SharePointサイトが見つかりません: {_config.SiteUrl}");
        return false;
      }

      Console.WriteLine($"  サイトID: {site.Id}");

      // ドライブ（ドキュメントライブラリ）を取得
      var drive = await GetDriveAsync(site.Id, _config.LibraryName);
      if (drive == null)
      {
        Console.WriteLine($"エラー: ドキュメントライブラリが見つかりません: {_config.LibraryName}");
        return false;
      }

      Console.WriteLine($"  ドライブID: {drive.Id}");

      // フォルダパスを正規化
      var targetFolderPath = _config.FolderPath.Trim('/');

      // フォルダが存在することを確認（存在しない場合は作成）
      var folderId = await EnsureFolderAsync(drive.Id!, targetFolderPath);

      // ファイルを読み込み
      using var fileStream = System.IO.File.OpenRead(filePath);
      var fileSize = new FileInfo(filePath).Length;

      // ファイルをアップロード
      DriveItem? uploadedItem;

      if (fileSize < 4 * 1024 * 1024) // 4MB未満は通常アップロード
      {
        uploadedItem = await _graphClient.Drives[drive.Id]
          .Items[folderId]
          .ItemWithPath(fileName)
          .Content
          .PutAsync(fileStream);
      }
      else // 4MB以上は大容量アップロード
      {
        var uploadSession = await _graphClient.Drives[drive.Id]
          .Items[folderId]
          .ItemWithPath(fileName)
          .CreateUploadSession
          .PostAsync(new Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession.CreateUploadSessionPostRequestBody
          {
            Item = new DriveItemUploadableProperties
            {
              AdditionalData = new Dictionary<string, object>
              {
                { "@microsoft.graph.conflictBehavior", "replace" }
              }
            }
          });

        if (uploadSession?.UploadUrl == null)
        {
          Console.WriteLine("エラー: アップロードセッションの作成に失敗しました");
          return false;
        }

        // 大容量ファイルのアップロード（チャンク単位）
        var maxChunkSize = 320 * 1024; // 320KB（推奨チャンクサイズ）
        var uploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxChunkSize);
        
        // アップロードの進行状況を表示
        IProgress<long> progress = new Progress<long>(uploadedBytes =>
        {
          var progressPercent = (double)uploadedBytes / fileSize * 100;
          Console.Write($"\r  アップロード進捗: {progressPercent:F1}% ({FormatFileSize(uploadedBytes)} / {FormatFileSize(fileSize)})");
        });
        
        var uploadResult = await uploadTask.UploadAsync(progress);
        Console.WriteLine(); // 改行

        if (uploadResult.UploadSucceeded)
        {
          uploadedItem = uploadResult.ItemResponse;
        }
        else
        {
          Console.WriteLine("エラー: 大容量ファイルのアップロードに失敗しました");
          return false;
        }
      }

      if (uploadedItem != null)
      {
        Console.WriteLine($"✓ アップロード成功: {fileName}");
        Console.WriteLine($"  サイズ: {FormatFileSize(fileSize)}");
        return true;
      }
      else
      {
        Console.WriteLine("エラー: アップロードに失敗しました");
        return false;
      }
    }
    catch (Exception ex)
    {
      Console.WriteLine($"エラー: アップロードに失敗しました");
      Console.WriteLine($"  詳細: {ex.Message}");

      if (ex.InnerException != null)
      {
        Console.WriteLine($"  内部エラー: {ex.InnerException.Message}");
      }

      return false;
    }
  }

  /// <summary>
  /// サイトURLを解析してホスト名とサイトパスを取得
  /// </summary>
  private (string hostName, string sitePath) ParseSiteUrl(string siteUrl)
  {
    var uri = new Uri(siteUrl);
    var hostName = uri.Host;
    var sitePath = uri.AbsolutePath;

    return (hostName, sitePath);
  }

  /// <summary>
  /// ドライブ（ドキュメントライブラリ）を取得
  /// </summary>
  private async Task<Drive?> GetDriveAsync(string siteId, string libraryName)
  {
    var drives = await _graphClient.Sites[siteId].Drives.GetAsync();

    if (drives?.Value == null)
    {
      return null;
    }

    // ライブラリ名で検索（Documentsの場合はデフォルトドライブを返す）
    var drive = drives.Value.FirstOrDefault(d =>
      d.Name?.Equals(libraryName, StringComparison.OrdinalIgnoreCase) == true);

    // 見つからない場合はデフォルトドライブを使用
    if (drive == null && libraryName.Equals("Documents", StringComparison.OrdinalIgnoreCase))
    {
      drive = await _graphClient.Sites[siteId].Drive.GetAsync();
    }

    return drive;
  }

  /// <summary>
  /// フォルダが存在することを確認し、存在しない場合は作成
  /// </summary>
  private async Task<string> EnsureFolderAsync(string driveId, string folderPath)
  {
    var folders = folderPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
    var currentParentId = "root";

    foreach (var folderName in folders)
    {
      try
      {
        // フォルダが存在するか確認
        var existingFolder = await _graphClient.Drives[driveId]
          .Items[currentParentId]
          .ItemWithPath(folderName)
          .GetAsync();

        if (existingFolder?.Id != null)
        {
          currentParentId = existingFolder.Id;
          continue;
        }
      }
      catch
      {
        // フォルダが存在しない場合は作成
      }

      // フォルダを作成
      var newFolder = new DriveItem
      {
        Name = folderName,
        Folder = new Folder(),
        AdditionalData = new Dictionary<string, object>
        {
          { "@microsoft.graph.conflictBehavior", "fail" }
        }
      };

      try
      {
        var createdFolder = await _graphClient.Drives[driveId]
          .Items[currentParentId]
          .Children
          .PostAsync(newFolder);

        if (createdFolder?.Id != null)
        {
          Console.WriteLine($"  フォルダを作成しました: {folderName}");
          currentParentId = createdFolder.Id;
        }
        else
        {
          throw new InvalidOperationException($"フォルダの作成に失敗しました: {folderName}");
        }
      }
      catch (Exception ex)
      {
        // 既に存在する場合は無視して取得
        if (ex.Message.Contains("nameAlreadyExists") || ex.Message.Contains("resourceAlreadyExists"))
        {
          var existingFolder = await _graphClient.Drives[driveId]
            .Items[currentParentId]
            .ItemWithPath(folderName)
            .GetAsync();

          if (existingFolder?.Id != null)
          {
            currentParentId = existingFolder.Id;
          }
          else
          {
            throw;
          }
        }
        else
        {
          throw;
        }
      }
    }

    return currentParentId;
  }

  /// <summary>
  /// ファイルサイズを人間が読みやすい形式にフォーマット
  /// </summary>
  private string FormatFileSize(long bytes)
  {
    string[] sizes = { "B", "KB", "MB", "GB", "TB" };
    double len = bytes;
    int order = 0;

    while (len >= 1024 && order < sizes.Length - 1)
    {
      order++;
      len = len / 1024;
    }

    return $"{len:0.##} {sizes[order]}";
  }
}
