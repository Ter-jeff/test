using CommonLib.Enums;

using LogLib.Enums;
using LogLib.Utility;

namespace LogLib.Static
{
    public static class Response
    {
        private static string _lastMessage = "";

        public static void Report(string message, EnumMessageLevel enumMessageLevel = EnumMessageLevel.General, int percentage = -1, string status = "")
        {
            if (_lastMessage == message)
            {
                return;
            }

            _lastMessage = message;

            if (enumMessageLevel == EnumMessageLevel.General)
            {
                LogHelper.Info(message, EnumLogTarget.FileAndConsole);
            }
            else if (enumMessageLevel == EnumMessageLevel.Error)
            {
                LogHelper.Error("[Error]: " + message, EnumLogTarget.FileAndConsole);
            }
            else if (enumMessageLevel == EnumMessageLevel.Warning)
            {
                LogHelper.Warn(message, EnumLogTarget.FileAndConsole);
            }
            else
            {
                LogHelper.Info(message, EnumLogTarget.FileAndConsole);
            }
        }

        // for ui
        public static void Report(string message, int enumMessageLevel, int percentage, string status = "")
        {
            Report(message, (EnumMessageLevel)enumMessageLevel, percentage, status);
        }

        public static void Clear()
        {
            _lastMessage = "";
        }
    }
}
