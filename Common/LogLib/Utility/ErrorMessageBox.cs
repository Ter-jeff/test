using NLog;

namespace LogLib.Utility
{
    public static class ErrorMessageBox
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public static void Show(string text) => _logger.Trace(text);

        public static void Show(string text, string type) => _logger.Trace(type + " : " + text);
    }
}
