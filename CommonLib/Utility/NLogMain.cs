using NLog;
using NLog.Config;
using NLog.Targets;

namespace CommonLib.Utility
{
    public class NLogMain
    {
        public void SetNLog(string fileName)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget();
            fileTarget.CreateDirs = true;
            fileTarget.DeleteOldFileOnStartup = true;
            fileTarget.FileName = fileName;
            fileTarget.Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}";//@"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}";
            config.AddTarget("file", fileTarget);
            var rule1 = new LoggingRule("*", LogLevel.Error, fileTarget);
            config.LoggingRules.Add(rule1);
            var rule2 = new LoggingRule("*", LogLevel.Trace, fileTarget);
            config.LoggingRules.Add(rule2);
            LogManager.Configuration = config;
        }
        public void SetNLogInTrace(string fileName)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget();
            fileTarget.CreateDirs = true;
            fileTarget.DeleteOldFileOnStartup = true;
            fileTarget.FileName = fileName;
            fileTarget.Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}";//@"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}";
            config.AddTarget("file", fileTarget);
            var rule = new LoggingRule("*", LogLevel.Trace, fileTarget);
            config.LoggingRules.Add(rule);
            LogManager.Configuration = config;
        }
        public void SetNLog()
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget();
            fileTarget.CreateDirs = true;
            fileTarget.DeleteOldFileOnStartup = false;
            fileTarget.FileName = "${basedir}/logs/${shortdate}.log";
            fileTarget.Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}";
            config.AddTarget("file", fileTarget);
            var rule = new LoggingRule("*", LogLevel.Trace, fileTarget);
            config.LoggingRules.Add(rule);
            LogManager.Configuration = config;
        }
    }
}
