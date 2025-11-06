using Microsoft.Extensions.Configuration;
using SharePointUploader;

class Program
{
  static async Task<int> Main(string[] args)
  {
    Console.WriteLine("SharePoint Online ファイルアップローダー");
    Console.WriteLine("=========================================");
    Console.WriteLine();

    // 引数チェック
    if (args.Length == 0)
    {
      ShowUsage();
      return 1;
    }

    var filePath = args[0];

    // ファイルの存在確認
    if (!File.Exists(filePath))
    {
      Console.WriteLine($"エラー: ファイルが見つかりません: {filePath}");
      return 1;
    }

    try
    {
      // 設定ファイルの読み込み
      var config = LoadConfiguration();

      if (config?.SharePoint == null)
      {
        Console.WriteLine("エラー: 設定ファイルの読み込みに失敗しました。");
        return 1;
      }

      // SharePointアップロードサービスを初期化
      var uploadService = new SharePointUploadService(config.SharePoint);

      // ファイルをアップロード
      var success = await uploadService.UploadFileAsync(filePath);

      return success ? 0 : 1;
    }
    catch (Exception ex)
    {
      Console.WriteLine($"予期しないエラーが発生しました: {ex.Message}");
      if (ex.InnerException != null)
      {
        Console.WriteLine($"詳細: {ex.InnerException.Message}");
      }
      return 1;
    }
  }

  /// <summary>
  /// 設定ファイルを読み込む
  /// </summary>
  static AppConfig? LoadConfiguration()
  {
    try
    {
      var configPath = Path.Combine(AppContext.BaseDirectory, "config.json");

      if (!File.Exists(configPath))
      {
        Console.WriteLine($"エラー: 設定ファイルが見つかりません: {configPath}");
        Console.WriteLine("config.sample.json を config.json にコピーして設定を編集してください。");
        return null;
      }

      var configuration = new ConfigurationBuilder()
        .SetBasePath(AppContext.BaseDirectory)
        .AddJsonFile("config.json", optional: false, reloadOnChange: false)
        .Build();

      var config = new AppConfig();
      configuration.Bind(config);

      return config;
    }
    catch (Exception ex)
    {
      Console.WriteLine($"設定ファイルの読み込みエラー: {ex.Message}");
      return null;
    }
  }

  /// <summary>
  /// 使用方法を表示
  /// </summary>
  static void ShowUsage()
  {
    Console.WriteLine("使用方法:");
    Console.WriteLine("  SharePointUploader <ファイルパス>");
    Console.WriteLine();
    Console.WriteLine("例:");
    Console.WriteLine("  SharePointUploader document.pdf");
    Console.WriteLine("  SharePointUploader C:\\Documents\\report.docx");
    Console.WriteLine();
    Console.WriteLine("注意:");
    Console.WriteLine("  - config.json にSharePointの接続情報を設定してください");
    Console.WriteLine("  - Azure ADアプリケーションの登録が必要です");
  }
}
