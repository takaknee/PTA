using System;

namespace OutlookPTAAddin.Core.Models
{
    /// <summary>
    /// メール解析処理中に発生する例外
    /// </summary>
    public class EmailAnalysisException : Exception
    {
        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        public EmailAnalysisException(string message) : base(message)
        {
        }

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        /// <param name="innerException">内部例外</param>
        public EmailAnalysisException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// メール作成処理中に発生する例外
    /// </summary>
    public class EmailCompositionException : Exception
    {
        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        public EmailCompositionException(string message) : base(message)
        {
        }

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        /// <param name="innerException">内部例外</param>
        public EmailCompositionException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// OpenAI API処理中に発生する例外
    /// </summary>
    public class OpenAIException : Exception
    {
        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        public OpenAIException(string message) : base(message)
        {
        }

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        /// <param name="innerException">内部例外</param>
        public OpenAIException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// 設定処理中に発生する例外
    /// </summary>
    public class ConfigurationException : Exception
    {
        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        public ConfigurationException(string message) : base(message)
        {
        }

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        /// <param name="innerException">内部例外</param>
        public ConfigurationException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}