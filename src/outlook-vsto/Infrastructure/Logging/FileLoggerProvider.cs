using System;
using System.IO;
using Microsoft.Extensions.Logging;

namespace OutlookPTAAddin.Infrastructure.Logging
{
    /// &lt;summary&gt;
    /// ファイルログプロバイダー
    /// &lt;/summary&gt;
    public class FileLoggerProvider : ILoggerProvider
    {
        private readonly string _filePath;
        private readonly object _lock = new object();

        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="fileName"&gt;ログファイル名&lt;/param&gt;
        public FileLoggerProvider(string fileName)
        {
            var logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PTA",
                "Logs"
            );

            Directory.CreateDirectory(logDirectory);
            _filePath = Path.Combine(logDirectory, fileName);
        }

        /// &lt;summary&gt;
        /// ロガーを作成する
        /// &lt;/summary&gt;
        /// &lt;param name="categoryName"&gt;カテゴリ名&lt;/param&gt;
        /// &lt;returns&gt;ロガー&lt;/returns&gt;
        public ILogger CreateLogger(string categoryName)
        {
            return new FileLogger(categoryName, _filePath, _lock);
        }

        /// &lt;summary&gt;
        /// リソースを解放する
        /// &lt;/summary&gt;
        public void Dispose()
        {
            // 特にリソース解放は不要
        }
    }

    /// &lt;summary&gt;
    /// ファイルログ実装
    /// &lt;/summary&gt;
    public class FileLogger : ILogger
    {
        private readonly string _categoryName;
        private readonly string _filePath;
        private readonly object _lock;

        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="categoryName"&gt;カテゴリ名&lt;/param&gt;
        /// &lt;param name="filePath"&gt;ファイルパス&lt;/param&gt;
        /// &lt;param name="lockObject"&gt;ロックオブジェクト&lt;/param&gt;
        public FileLogger(string categoryName, string filePath, object lockObject)
        {
            _categoryName = categoryName;
            _filePath = filePath;
            _lock = lockObject;
        }

        /// &lt;summary&gt;
        /// ログスコープを開始する
        /// &lt;/summary&gt;
        /// &lt;typeparam name="TState"&gt;状態型&lt;/typeparam&gt;
        /// &lt;param name="state"&gt;状態&lt;/param&gt;
        /// &lt;returns&gt;スコープ&lt;/returns&gt;
        public IDisposable BeginScope&lt;TState&gt;(TState state)
        {
            return null;
        }

        /// &lt;summary&gt;
        /// ログレベルが有効かどうかを判定する
        /// &lt;/summary&gt;
        /// &lt;param name="logLevel"&gt;ログレベル&lt;/param&gt;
        /// &lt;returns&gt;有効かどうか&lt;/returns&gt;
        public bool IsEnabled(LogLevel logLevel)
        {
            return logLevel != LogLevel.None;
        }

        /// &lt;summary&gt;
        /// ログを出力する
        /// &lt;/summary&gt;
        /// &lt;typeparam name="TState"&gt;状態型&lt;/typeparam&gt;
        /// &lt;param name="logLevel"&gt;ログレベル&lt;/param&gt;
        /// &lt;param name="eventId"&gt;イベントID&lt;/param&gt;
        /// &lt;param name="state"&gt;状態&lt;/param&gt;
        /// &lt;param name="exception"&gt;例外&lt;/param&gt;
        /// &lt;param name="formatter"&gt;フォーマッター&lt;/param&gt;
        public void Log&lt;TState&gt;(LogLevel logLevel, EventId eventId, TState state, Exception exception, Func&lt;TState, Exception, string&gt; formatter)
        {
            if (!IsEnabled(logLevel))
            {
                return;
            }

            try
            {
                lock (_lock)
                {
                    var message = formatter(state, exception);
                    var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{GetLogLevelString(logLevel)}] [{_categoryName}] {message}";
                    
                    if (exception != null)
                    {
                        logEntry += Environment.NewLine + exception.ToString();
                    }

                    logEntry += Environment.NewLine;

                    File.AppendAllText(_filePath, logEntry);
                }
            }
            catch
            {
                // ログ出力でエラーが発生しても、アプリケーションの動作に影響させない
            }
        }

        /// &lt;summary&gt;
        /// ログレベルを文字列に変換する
        /// &lt;/summary&gt;
        /// &lt;param name="logLevel"&gt;ログレベル&lt;/param&gt;
        /// &lt;returns&gt;ログレベル文字列&lt;/returns&gt;
        private static string GetLogLevelString(LogLevel logLevel)
        {
            return logLevel switch
            {
                LogLevel.Trace =&gt; "TRACE",
                LogLevel.Debug =&gt; "DEBUG",
                LogLevel.Information =&gt; "INFO ",
                LogLevel.Warning =&gt; "WARN ",
                LogLevel.Error =&gt; "ERROR",
                LogLevel.Critical =&gt; "FATAL",
                _ =&gt; "UNKNOW"
            };
        }
    }
}