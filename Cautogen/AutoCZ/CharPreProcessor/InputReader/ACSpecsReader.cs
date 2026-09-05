using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader;

using CommonLib.Enums;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader
{
    internal class AcSpecsReader : TextInputReader
    {
        /* properties */
        public static Dictionary<string, double> AcSpecs = new Dictionary<string, double>();

        /* constructor */
        public AcSpecsReader(string filePath)
            : base(filePath)
        {
        }

        /* methods */

        protected override void _Read(StreamReader textReader)
        {
            MessageWriter.WriteMessage("Reading ACSpecs file...", EnumMessageLevel.Info);

            AcSpecs = new Dictionary<string, double>();

            // read file

            string line = "";

            // fast forward until seeing the header line contains "Symbol"
            while (line != null && !line.Contains("Selectors"))
            {
                line = textReader.ReadLine();
            }

            if (!string.IsNullOrEmpty(line))
            {
                UtilityMain.UtilityData.AcCategories = line.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries).Skip(1).ToList();
            }
        }
    }
}
