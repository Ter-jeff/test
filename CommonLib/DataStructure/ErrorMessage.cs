using System.Windows;

using NLog;

namespace CommonLib.DataStructure
{
    public static class ErrorMessage
    {
        public static bool CmdMode { set; get; }

        public static void SetMode(bool isCmd)
        {
            CmdMode = isCmd;
        }

        public static void Show(string text)
        {
            if (CmdMode)
            {
                Logger logger = LogManager.GetCurrentClassLogger();
                logger.Trace(string.Format(text));
            }
            else
            {
                MessageBox.Show(text);
            }
        }

        public static MessageBoxResult Show(string text, string type, MessageBoxButton button, MessageBoxImage icon)
        {
            if (CmdMode)
            {
                Logger logger = LogManager.GetCurrentClassLogger();
                logger.Trace(string.Format(text));
                return MessageBoxResult.Yes;
            }
            else
            {
                return MessageBox.Show(text, type, button, icon);
            }
        }

        public static void Show(string type, string text)
        {
            if (CmdMode)
            {
                Logger logger = LogManager.GetCurrentClassLogger();
                logger.Trace(string.Format(type + " : " + text));
            }
            else
            {
                MessageBox.Show(type, text);
            }
        }

        public static MessageBoxResult Show(string text, string caption, MessageBoxButton buttons)
        {
            if (CmdMode)
            {
                Logger logger = LogManager.GetCurrentClassLogger();
                logger.Trace(string.Format(caption + " : " + text));
                return MessageBoxResult.Yes;
            }
            else
            {
                return MessageBox.Show(text, caption, buttons);
            }
        }
    }
}
