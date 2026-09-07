using System;
using System.Diagnostics;
using System.IO;

using LogLib.Enums;

using NLog;
using NLog.Config;
using NLog.Targets;

namespace LogLib.Utility
{
    public static class LogHelper
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private static string _logFile = string.Empty;

        public static void SetNLog(string filePath)
        {
            _logFile = filePath;
            ConfigureLogging(filePath, deleteOldOnStartup: true, traceLevel: true, addErrorRule: true);
            LogManager.ThrowExceptions = true;
            PrintProcessName();
        }

        public static void Copy(string destFileName)
        {
            string? dir = Path.GetDirectoryName(destFileName);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            File.Copy(_logFile, destFileName, true);
        }

        public static void DeleteLogFile()
        {
            if (File.Exists(_logFile))
            {
                File.Delete(_logFile);
            }
        }

        private static void ConfigureLogging(string fileName, bool deleteOldOnStartup, bool traceLevel, bool addErrorRule)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                DeleteOldFileOnStartup = deleteOldOnStartup,
                FileName = fileName,
                Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}"
            };
            config.AddTarget("file", fileTarget);

            if (addErrorRule)
            {
                var errorRule = new LoggingRule("*", LogLevel.Error, fileTarget);
                config.LoggingRules.Add(errorRule);
            }

            if (traceLevel)
            {
                var traceRule = new LoggingRule("*", LogLevel.Trace, fileTarget);
                config.LoggingRules.Add(traceRule);
            }

            LogManager.Configuration = config;
        }

        private static void PrintProcessName()
        {
            try
            {
                string processName = Process.GetCurrentProcess().ProcessName;
                _logger.Info("********New Session( " + processName + " )*********");
            }
            catch (Exception)
            {
                _logger?.Info("InitLog(): error");
            }
        }

        private static void Write(EnumLogTarget enumLogTarget, Action action, string message, Exception? exception = null)
        {
            if (enumLogTarget == EnumLogTarget.Console || enumLogTarget == EnumLogTarget.FileAndConsole)
            {
                Console.WriteLine(message);
                if (exception != null)
                {
                    Console.WriteLine(exception.Message);
                }
            }

            if (enumLogTarget == EnumLogTarget.File || enumLogTarget == EnumLogTarget.FileAndConsole)
            {
                action();
                if (exception != null)
                {
                    _logger.Error(exception);
                }
            }
        }

        public static void Info(string message, EnumLogTarget enumLogTarget = EnumLogTarget.File) => Log(_logger.Info, message, enumLogTarget);

        public static void Error(string message, EnumLogTarget enumLogTarget = EnumLogTarget.File) => Log(_logger.Error, message, enumLogTarget);

        public static void Warn(string message, EnumLogTarget enumLogTarget = EnumLogTarget.File) => Log(_logger.Warn, message, enumLogTarget);

        public static void Debug(string message, EnumLogTarget enumLogTarget = EnumLogTarget.File) => Log(_logger.Debug, message, enumLogTarget);

        private static void Log(Action<string> logAction, string message, EnumLogTarget enumLogTarget) => Write(enumLogTarget, () => logAction(message), message);
    }
}
