using System;

namespace OutlookPTAAddin.Core.Models
{
    /// &lt;summary&gt;
    /// メール解析処理中に発生する例外
    /// &lt;/summary&gt;
    public class EmailAnalysisException : Exception
    {
        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="message"&gt;エラーメッセージ&lt;/param&gt;
        public EmailAnalysisException(string message) : base(message)
        {
        }

        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="message"&gt;エラーメッセージ&lt;/param&gt;
        /// &lt;param name="innerException"&gt;内部例外&lt;/param&gt;
        public EmailAnalysisException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    /// &lt;summary&gt;
    /// メール作成処理中に発生する例外
    /// &lt;/summary&gt;
    public class EmailCompositionException : Exception
    {
        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="message"&gt;エラーメッセージ&lt;/param&gt;
        public EmailCompositionException(string message) : base(message)
        {
        }

        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="message"&gt;エラーメッセージ&lt;/param&gt;
        /// &lt;param name="innerException"&gt;内部例外&lt;/param&gt;
        public EmailCompositionException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    /// &lt;summary&gt;
    /// OpenAI API処理中に発生する例外
    /// &lt;/summary&gt;
    public class OpenAIException : Exception
    {
        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="message"&gt;エラーメッセージ&lt;/param&gt;
        public OpenAIException(string message) : base(message)
        {
        }

        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="message"&gt;エラーメッセージ&lt;/param&gt;
        /// &lt;param name="innerException"&gt;内部例外&lt;/param&gt;
        public OpenAIException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    /// &lt;summary&gt;
    /// 設定処理中に発生する例外
    /// &lt;/summary&gt;
    public class ConfigurationException : Exception
    {
        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="message"&gt;エラーメッセージ&lt;/param&gt;
        public ConfigurationException(string message) : base(message)
        {
        }

        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="message"&gt;エラーメッセージ&lt;/param&gt;
        /// &lt;param name="innerException"&gt;内部例外&lt;/param&gt;
        public ConfigurationException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}