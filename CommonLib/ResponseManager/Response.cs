using System;
using System.Threading;

using CommonLib.DataStructure;
using CommonLib.Utility;

namespace CommonLib.ResponseManager
{
    public static class Response
    {
        public static IProgress<ProgressStatus> Progress { get; set; }
        private static string _lastMessage = "";
        public static bool Disable { get; set; } = false;

        #region Member Function

        public static void WriteMessage(string message, MessageLevel messageLevel = MessageLevel.General, int percentage = -1)
        {
            if (Disable)
            {
                if (messageLevel == MessageLevel.General)
                {
                    LogHelper.Info(message, EnumLogTarget.Console);
                }
                else if (messageLevel == MessageLevel.Error)
                {
                    LogHelper.Error("[Error]: " + message, EnumLogTarget.Console);
                }
                else if (messageLevel == MessageLevel.Warning)
                {
                    LogHelper.Warn(message, EnumLogTarget.Console);
                }
                else
                {
                    LogHelper.Info(message, EnumLogTarget.Console);
                }
            }
            else
            {
                if (_lastMessage == message)
                {
                    return;
                }

                var progressStatus = new ProgressStatus
                {
                    Result = message
                };
                if (percentage > 0)
                {
                    progressStatus.Percentage = percentage;
                }

                progressStatus.Level = messageLevel;
                if (Progress != null)
                {
                    Progress.Report(progressStatus);
                    Thread.Sleep(50);
                }
                _lastMessage = message;

                if (messageLevel == MessageLevel.General)
                {
                    LogHelper.Info(message, EnumLogTarget.FileAndConsole);
                }
                else if (messageLevel == MessageLevel.Error)
                {
                    LogHelper.Error("[Error]: " + message, EnumLogTarget.FileAndConsole);
                }
                else if (messageLevel == MessageLevel.Warning)
                {
                    LogHelper.Warn(message, EnumLogTarget.FileAndConsole);
                }
                else
                {
                    LogHelper.Info(message, EnumLogTarget.FileAndConsole);
                }
            }
        }

        public static void Report(string message, MessageLevel messageLevel = MessageLevel.General, int percentage = -1,
            string status = "", bool showUiOnly = false)
        {
            if (_lastMessage == message)
            {
                return;
            }

            _lastMessage = message;
            if (Disable)
            {
                if (messageLevel == MessageLevel.General)
                {
                    LogHelper.Info(message, EnumLogTarget.Console);
                }
                else if (messageLevel == MessageLevel.Error)
                {
                    LogHelper.Error("[Error]: " + message, EnumLogTarget.Console);
                }
                else if (messageLevel == MessageLevel.Warning)
                {
                    LogHelper.Warn(message, EnumLogTarget.Console);
                }
                else
                {
                    LogHelper.Info(message, EnumLogTarget.Console);
                }
            }
            else
            {
                if (!showUiOnly)
                {
                    var progressStatus = new ProgressStatus
                    {
                        Result = message
                    };
                    if (percentage > 0)
                    {
                        progressStatus.Percentage = percentage;
                    }

                    progressStatus.Level = messageLevel;
                    progressStatus.Status = status;
                    if (Progress != null)
                    {
                        Progress.Report(progressStatus);
                    }

                    Thread.Sleep(50);
                }

                if (messageLevel == MessageLevel.General)
                {
                    LogHelper.Info(message, EnumLogTarget.FileAndConsole);
                }
                else if (messageLevel == MessageLevel.Error)
                {
                    LogHelper.Error("[Error]: " + message, EnumLogTarget.FileAndConsole);
                }
                else if (messageLevel == MessageLevel.Warning)
                {
                    LogHelper.Warn(message, EnumLogTarget.FileAndConsole);
                }
                else
                {
                    LogHelper.Info(message, EnumLogTarget.FileAndConsole);
                }
            }
        }

        public static void ReportLite(string message, MessageLevel messageLevel = MessageLevel.General, int percentage = -1, string status = "")
        {
            var progressStatus = new ProgressStatus();
            if (percentage > 0)
            {
                progressStatus.Percentage = percentage;
            }

            progressStatus.Status = status;
            if (Progress != null)
            {
                Progress.Report(progressStatus);
            }

            Thread.Sleep(50);
        }

        public static void Initialize()
        {
            var progressStatus = new ProgressStatus
            {
                Percentage = 0
            };
            Thread.Sleep(100);
            if (Progress != null)
            {
                Progress.Report(progressStatus);
            }
        }

        public static void Clear()
        {
            _lastMessage = "";
        }
        #endregion
    }
}
