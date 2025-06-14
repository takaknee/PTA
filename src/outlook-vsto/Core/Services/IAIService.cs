using System.Threading.Tasks;

namespace OutlookPTAAddin.Core.Services
{
    /// <summary>
    /// AI サービスの抽象インターフェース
    /// 複数のAIプロバイダー（OpenAI、Claude、Azure OpenAI等）に対応
    /// </summary>
    public interface IAIService
    {
        /// <summary>
        /// AIプロバイダー名
        /// </summary>
        string ProviderName { get; }

        /// <summary>
        /// テキスト解析を実行する
        /// </summary>
        /// <param name="systemPrompt">システムプロンプト</param>
        /// <param name="userContent">ユーザーコンテンツ</param>
        /// <returns>解析結果</returns>
        Task<string> AnalyzeTextAsync(string systemPrompt, string userContent);

        /// <summary>
        /// テキスト生成を実行する
        /// </summary>
        /// <param name="systemPrompt">システムプロンプト</param>
        /// <param name="userPrompt">ユーザープロンプト</param>
        /// <returns>生成されたテキスト</returns>
        Task<string> GenerateTextAsync(string systemPrompt, string userPrompt);

        /// <summary>
        /// API接続テストを実行する
        /// </summary>
        /// <returns>テスト結果メッセージ</returns>
        Task<string> TestConnectionAsync();

        /// <summary>
        /// サービスが利用可能かどうかを確認する
        /// </summary>
        /// <returns>利用可能な場合はtrue</returns>
        bool IsAvailable();
    }
}