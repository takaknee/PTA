using System;
using System.IO;
using Microsoft.Extensions.Logging;

namespace OutlookPTAAddin.Infrastructure.Logging
{
    /// <summary>
    /// ファイルログプロバイダー
    /// </summary>
    public class FileLoggerProvider : ILoggerProvider
    {
        private readonly string _filePath;
        private readonly object _lock = new object();

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="fileName">ログファイル名</param>
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

        /// <summary>
        /// ロガーを作成する
        /// </summary>
        /// <param name="categoryName">カテゴリ名</param>
        /// <returns>ロガー</returns>
        public ILogger CreateLogger(string categoryName)
        {
            return new FileLogger(categoryName, _filePath, _lock);
        }

        /// <summary>
        /// リソースを解放する
        /// </summary>
        public void Dispose()
        {
            // 特にリソース解放は不要
        }
    }

    /// <summary>
    /// ファイルログ実装
    /// </summary>
    public class FileLogger : ILogger
    {
        private readonly string _categoryName;
        private readonly string _filePath;
        private readonly object _lock;

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="categoryName">カテゴリ名</param>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="lockObject">ロックオブジェクト</param>
        public FileLogger(string categoryName, string filePath, object lockObject)
        {
            _categoryName = categoryName;
            _filePath = filePath;
            _lock = lockObject;
        }

        /// <summary>
        /// ログスコープを開始する
        /// </summary>
        /// <typeparam name="TState">状態型</typeparam>
        /// <param name="state">状態</param>
        /// <returns>スコープ</returns>
        public IDisposable BeginScope<TState>(TState state)
        {
            return null;
        }

        /// <summary>
        /// ログレベルが有効かどうかを判定する
        /// </summary>
        /// <param name="logLevel">ログレベル</param>
        /// <returns>有効かどうか</returns>
        public bool IsEnabled(LogLevel logLevel)
        {
            return logLevel != LogLevel.None;
        }

        /// <summary>
        /// ログを出力する
        /// </summary>
        /// <typeparam name="TState">状態型</typeparam>
        /// <param name="logLevel">ログレベル</param>
        /// <param name="eventId">イベントID</param>
        /// <param name="state">状態</param>
        /// <param name="exception">例外</param>
        /// <param name="formatter">フォーマッター</param>
        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception exception, Func<TState, Exception, string> formatter)
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

        /// <summary>
        /// ログレベルを文字列に変換する
        /// </summary>
        /// <param name="logLevel">ログレベル</param>
        /// <returns>ログレベル文字列</returns>
        private static string GetLogLevelString(LogLevel logLevel)
        {
            return logLevel switch
            {
                LogLevel.Trace => "TRACE",
                LogLevel.Debug => "DEBUG",
                LogLevel.Information => "INFO ",
                LogLevel.Warning => "WARN ",
                LogLevel.Error => "ERROR",
                LogLevel.Critical => "FATAL",
                _ => "UNKNOW"
            };
        }
    }
}