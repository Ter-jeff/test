using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;

using IgxlLib.Enums;

using LocalSpecs = Cautogen.AutoCZ.CharPostProcessor.LocalSpec.LocalSpecs;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class OtherSheetGenerator
    {

        public static void Generate()
        {
            if (LocalSpecs.FileStructure.Count == 0)
            {
                return;
            }

            GeneralFunc.WriteMessage("Generating Other sheet... ");

            string[] sheetTypeNames = Enum.GetNames(typeof(EnumSheetType));

            foreach (KeyValuePair<string, string> file in LocalSpecs.FileStructure)
            {
                if (file.Value.Contains(@"\Module\Library\", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!File.Exists(file.Value))
                {
                    continue;
                }

                string firstLine = File.ReadLines(file.Value).FirstOrDefault() ?? string.Empty;

                bool igxlSheetType = sheetTypeNames.Contains(firstLine.Split(',')[0].Trim(), StringComparer.OrdinalIgnoreCase);

                if (!igxlSheetType)
                {
                    TestProgram.NonIgxlSheetsList.Add(Path.GetDirectoryName(file.Value), Path.GetFileNameWithoutExtension(file.Value));
                }
            }
        }

        public static void WriteConfig(string fileName, string oriPath)
        {
            List<string> lines = File.ReadLines(oriPath).ToList();

            File.WriteAllLines(fileName, lines);
            TestProgram.NonIgxlSheetsList.Add(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName));
            LocalSpecs.GenOthers.Add(fileName);
        }
    }
}
