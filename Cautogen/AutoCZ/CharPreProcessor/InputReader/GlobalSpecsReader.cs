using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader;

using CommonLib.Enums;

using LogLib.Utility;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader
{
    public class GlobalSpecsReader : TextInputReader
    {
        /* properties */
        public static Dictionary<string, double> GlobalSpecs = new Dictionary<string, double>();

        /* constructor */
        public GlobalSpecsReader(string filePath)
            : base(filePath)
        {
        }

        /* methods */
        protected override void _Read(StreamReader textReader)
        {
            MessageWriter.WriteMessage("Reading GlobalSpecs file...", EnumMessageLevel.Info);
            LogHelper.Info("Reading GlobalSpecs file...");

            GlobalSpecs = new Dictionary<string, double>();

            // read file

            string line = "";

            // fast forward until seeing the header line contains "Symbol"
            while (line != null && !line.Contains("Symbol"))
            {
                line = textReader.ReadLine();
            }
            while ((line = textReader.ReadLine()) != null)
            {
                string key = "";
                string value = "";
                if (Regex.IsMatch(line, @"^\t\w.*"))
                {
                    key = line.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)[0].Replace("_", "").ToUpper().Replace("GLB", "");
                    key = Regex.Replace(key, "P[0]+", "P");

                    // todo: if there is job filled, then will get the wrong filed
                    string[] values = line.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    if (values.Length > 1)
                    {
                        value = values[1].Replace("=", "");
                    }
                }

                // update utility data for type(value) is numeric
                if (!GlobalSpecs.Keys.Contains(key) && double.TryParse(value, out double val))
                {
                    GlobalSpecs.Add(key, val);
                }
            }
        }
    }
}
