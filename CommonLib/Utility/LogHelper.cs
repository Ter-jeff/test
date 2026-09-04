using System;
using System.IO;

using NLog;
using NLog.Config;
using NLog.Targets;

namespace CommonLib.Utility
{
    public enum EnumLogTarget
    {
        File,
        Console,
        FileAndConsole
    }

    public class LogHelper
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private static string _logFile = string.Empty;

        public static void SetNLog(string filePath)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                DeleteOldFileOnStartup = false,
                FileName = filePath,
                Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}"
            };
            config.AddTarget("file", fileTarget);
            var rule = new LoggingRule("*", LogLevel.Trace, fileTarget);
            config.LoggingRules.Add(rule);
            LogManager.Configuration = config;
        }
        public static void DeleteLogFile()
        {
            if (File.Exists(_logFile))
            {
                File.Delete(_logFile);
            }
        }
        public static void Init(string fileName)
        {
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            if (!File.Exists(folder))
            {
                folder = Directory.GetCurrentDirectory();
            }

            _logFile = Path.Combine(folder, "Teradyne", fileName);

            _ = new LoggingConfiguration();
            SetupConfig(_logFile);
            LogManager.ThrowExceptions = true;
            InitLog();
        }

        public static void Init(string filePath, string fileName)
        {
            _logFile = Path.Combine(filePath, fileName);
            SetupConfig(_logFile);
            InitLog();
        }

        private static void SetupConfig(string fileName)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                DeleteOldFileOnStartup = false,
                FileName = fileName,
                Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}"
            };
            config.AddTarget("file", fileTarget);
            var rule = new LoggingRule("*", LogLevel.Trace, fileTarget);
            config.LoggingRules.Add(rule);
            LogManager.Configuration = config;
        }

        public static void Copy(string destFileName)
        {
            string dir = Path.GetDirectoryName(destFileName);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            File.Copy(_logFile, destFileName, true);
        }

        private static void InitLog()
        {
            try
            {
                string processName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
                _logger.Info("********New Session( " + processName + " )*********");
            }
            catch (Exception)
            {
                if (_logger != null)
                {
                    _logger.Info("InitLog(): error");
                }
            }
        }

        private static void Write(EnumLogTarget type, Action logAction, string message, Exception exception = null)
        {
            if (type == EnumLogTarget.Console || type == EnumLogTarget.FileAndConsole)
            {
                Console.WriteLine(message);
                if (exception != null)
                {
                    Console.WriteLine(exception.Message);
                }
            }

            if (type == EnumLogTarget.File || type == EnumLogTarget.FileAndConsole)
            {
                logAction();
                if (exception != null)
                {
                    _logger.Error(exception);
                }
            }
        }

        public static void Debug(string message, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Debug(message), message);

        public static void Debug(string message, Exception ex, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Debug(message), message, ex);

        public static void Info(string message, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Info(message), message);

        public static void Info(string message, Exception ex, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Info(message), message, ex);

        public static void Error(string message, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Error(message), message);

        public static void Error(string message, Exception ex, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Error(message), message, ex);

        public static void Fatal(string message, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Fatal(message), message);

        public static void Fatal(string message, Exception ex, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Fatal(message), message, ex);

        public static void Warn(string message, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Warn(message), message);

        public static void Warn(string message, Exception ex, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Warn(message), message, ex);

        public static void LogException(Exception e, EnumLogTarget type = EnumLogTarget.File) =>
            Write(type, () => _logger.Error(e), e.Message, e);
    }
}
