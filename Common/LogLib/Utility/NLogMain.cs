using NLog;
using NLog.Config;
using NLog.Targets;

namespace LogLib.Utility
{
    public class NLogMain
    {
        public static void SetNLog(string fileName)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                DeleteOldFileOnStartup = true,
                FileName = fileName,
                Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}"//@"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}";
            };
            config.AddTarget("file", fileTarget);
            var rule1 = new LoggingRule("*", LogLevel.Error, fileTarget);
            config.LoggingRules.Add(rule1);
            var rule2 = new LoggingRule("*", LogLevel.Trace, fileTarget);
            config.LoggingRules.Add(rule2);
            LogManager.Configuration = config;
        }

        public static void SetNLogInTrace(string fileName)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                DeleteOldFileOnStartup = true,
                FileName = fileName,
                Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}"//@"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}";
            };
            config.AddTarget("file", fileTarget);
            var rule = new LoggingRule("*", LogLevel.Trace, fileTarget);
            config.LoggingRules.Add(rule);
            LogManager.Configuration = config;
        }

        public static void SetNLog()
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                DeleteOldFileOnStartup = false,
                FileName = "${basedir}/logs/${shortdate}.log",
                Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}"
            };
            config.AddTarget("file", fileTarget);
            var rule = new LoggingRule("*", LogLevel.Trace, fileTarget);
            config.LoggingRules.Add(rule);
            LogManager.Configuration = config;
        }
    }
}
