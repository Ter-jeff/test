using System;
using System.IO;

using CommonLib.Enums;

namespace Cautogen.AutoCZ.CharPreProcessor.Utility
{
    public class MessageWriter
    {
        private delegate void DelegateSetMessageText();

        private delegate void DelegateAddMessageText(string text, int colorindex);

        public static IProgress<string> Progress = null;

        public static void SetMessageText()
        {
            if (!Directory.Exists(UtilityMain.UtilityData.InputParam.TarDic))
            {
                Directory.CreateDirectory(UtilityMain.UtilityData.InputParam.TarDic);
            }

            string fileName = Path.Combine(UtilityMain.UtilityData.InputParam.TarDic, "CharPreProcessor.log");
            var sw = new StreamWriter(fileName, false);
            sw.Write("");
            sw.Close();
        }

        public static void WriteMessage(string text, EnumMessageLevel meslevel)
        {
            if (Progress != null)
            {
                Progress.Report(text);
            }

            switch (meslevel)
            {
                case EnumMessageLevel.Error:
                    AddMessageText(text, 2);
                    break;
                case EnumMessageLevel.Warning:
                    AddMessageText(text, 3);
                    break;
                case EnumMessageLevel.Info:
                    AddMessageText(text, 1);
                    break;
                case EnumMessageLevel.Result:
                    AddMessageText(text, 4);
                    break;
            }
        }

        private static void AddMessageText(string text, int colorindex)
        {
            WriteLogs(text, colorindex);
        }

        private static void WriteLogs(string text, int colorindex)
        {
            try
            {
                if (!Directory.Exists(UtilityMain.UtilityData.InputParam.TarDic))
                {
                    Directory.CreateDirectory(UtilityMain.UtilityData.InputParam.TarDic);
                }

                string fileName = Path.Combine(UtilityMain.UtilityData.InputParam.TarDic, "CharPreProcessor.log");
                var sw = new StreamWriter(fileName, true);
                string messageLevel = "";
                switch (colorindex)
                {
                    case 2:
                        messageLevel = "Error: ";
                        break;
                    case 3:
                        messageLevel = "Warning: ";
                        break;
                }
                sw.WriteLine(messageLevel + text);
                sw.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
