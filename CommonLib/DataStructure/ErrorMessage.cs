using System.Runtime.InteropServices;

using NLog;

namespace CommonLib.DataStructure
{
    public enum MessageBoxButton
    {
        OK = 0x0,
        OKCancel = 0x1,
        YesNoCancel = 0x3,
        YesNo = 0x4,
    }

    public enum MessageBoxImage
    {
        None = 0x0,
        Error = 0x10,
        Question = 0x20,
        Warning = 0x30,
        Information = 0x40,
    }

    public enum MessageBoxResult
    {
        None = 0,
        OK = 1,
        Cancel = 2,
        Yes = 6,
        No = 7,
    }

    public static class ErrorMessage
    {
        public static bool CmdMode { set; get; }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBox(nint hWnd, string text, string caption, uint type);

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
                MessageBox(0, text, string.Empty, (uint)MessageBoxButton.OK);
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
                int result = MessageBox(0, text, type, (uint)button | (uint)icon);
                return (MessageBoxResult)result;
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
                MessageBox(0, type, text, (uint)MessageBoxButton.OK);
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
                int result = MessageBox(0, text, caption, (uint)buttons);
                return (MessageBoxResult)result;
            }
        }
    }
}
