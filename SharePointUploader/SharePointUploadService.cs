using Microsoft.SharePoint.Client;
using PnP.Framework;

namespace SharePointUploader;

/// <summary>
/// SharePoint Onlineへのファイルアップロードサービス
/// </summary>
public class SharePointUploadService
{
    private readonly SharePointConfig _config;

    public SharePointUploadService(SharePointConfig config)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        ValidateConfig();
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

        if (string.IsNullOrWhiteSpace(_config.ClientSecret))
            throw new InvalidOperationException("ClientSecret が設定されていません。");

        if (string.IsNullOrWhiteSpace(_config.FolderPath))
            throw new InvalidOperationException("FolderPath が設定されていません。");
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
            Console.WriteLine($"ファイルをアップロード中: {Path.GetFileName(filePath)}");
            Console.WriteLine($"アップロード先: {_config.SiteUrl}/{_config.FolderPath}");

            // 認証マネージャーを使用してClientContextを取得
            var authManager = new AuthenticationManager(
                _config.ClientId,
                _config.ClientSecret,
                _config.TenantId);

            using (var context = await authManager.GetContextAsync(_config.SiteUrl))
            {
                var web = context.Web;
                context.Load(web);
                await context.ExecuteQueryAsync();

                // ファイル情報を取得
                var fileName = Path.GetFileName(filePath);
                var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);

                // フォルダパスを正規化
                var targetFolderUrl = _config.FolderPath.TrimStart('/');

                // ターゲットフォルダを取得または作成
                var folder = await EnsureFolderAsync(context, web, targetFolderUrl);

                // ファイルをアップロード
                var fileCreationInfo = new FileCreationInformation
                {
                    Content = fileBytes,
                    Url = fileName,
                    Overwrite = true
                };

                var uploadFile = folder.Files.Add(fileCreationInfo);
                context.Load(uploadFile);
                await context.ExecuteQueryAsync();

                Console.WriteLine($"✓ アップロード成功: {fileName}");
                Console.WriteLine($"  サイズ: {FormatFileSize(fileBytes.Length)}");
                return true;
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
    /// フォルダが存在することを確認し、存在しない場合は作成
    /// </summary>
    private async Task<Folder> EnsureFolderAsync(ClientContext context, Web web, string folderPath)
    {
        var folders = folderPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var currentFolder = web.RootFolder;
        context.Load(currentFolder);
        await context.ExecuteQueryAsync();

        foreach (var folderName in folders)
        {
            var folderCollection = currentFolder.Folders;
            context.Load(folderCollection);
            await context.ExecuteQueryAsync();

            var existingFolder = folderCollection.FirstOrDefault(f =>
                f.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase));

            if (existingFolder != null)
            {
                currentFolder = existingFolder;
            }
            else
            {
                currentFolder = folderCollection.Add(folderName);
                context.Load(currentFolder);
                await context.ExecuteQueryAsync();
                Console.WriteLine($"  フォルダを作成しました: {folderName}");
            }
        }

        return currentFolder;
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
