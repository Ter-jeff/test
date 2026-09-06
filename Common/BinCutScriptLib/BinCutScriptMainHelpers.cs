using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

namespace BinCutScriptLib
{
    public class BinCutScriptMainHelpers
    {
        internal static string CopyToTemp(string fileName)
        {
            string temp = Path.Combine(Path.GetDirectoryName(fileName)!, Path.GetFileNameWithoutExtension(fileName) + "_TEMP" + Path.GetExtension(fileName));
            if (File.Exists(temp))
            {
                File.Delete(temp);
            }

            File.Copy(fileName, temp);
            return temp;
        }

        internal static void RemoveTemp(string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
        }

        public static string GetProjectName(string binCutSpec, string postbinCutSpec, string testPlan, string testProgram)
        {
            if (binCutSpec != null && Path.GetFileName(binCutSpec).Contains('_'))
            {
                if (Equals(Path.GetFileName(binCutSpec).Split('_')[0], "Jade-S"))
                {
                    return "JC-Chop";
                }

                return Path.GetFileName(binCutSpec).Split('_')[0];
            }

            if (postbinCutSpec != null && Path.GetFileName(postbinCutSpec).Contains('_'))
            {
                return Path.GetFileName(postbinCutSpec).Split('_')[0];
            }

            if (testPlan != null && Path.GetFileName(testPlan).Contains('_'))
            {
                if (Equals(Path.GetFileName(testPlan).Split('_')[0], "Jade-S"))
                {
                    return "JC-Chop";
                }

                return Path.GetFileName(testPlan).Split('_')[0];
            }

            if (testProgram != null && Path.GetFileName(testProgram).Contains('_'))
            {
                return GetProjectNameByTestProgram(testProgram);
            }
            return "";
        }

        private static string GetProjectNameByTestProgram(string testProgram)
        {
            if (string.IsNullOrWhiteSpace(testProgram))
            {
                return string.Empty;
            }

            string fileName = Path.GetFileName(testProgram);
            if (!fileName.Contains('_'))
            {
                return string.Empty;
            }

            string codeName = fileName.Split('_')[0];

            var projectMap = new Dictionary<string, string>(StringExtensions.IgnoreCase)
            {
                { "NM28", "Rhodes" }, { "NM31", "Rhodes" },
                { "NV37", "Crete" },  { "NV38", "Crete" },  { "QK41", "Crete" },
                { "NK84", "Staten" },
                { "TMNB26", "Bora" }, { "NB26", "Bora" },   { "TMNB27", "Bora" }, { "NB27", "Bora" },
                { "MU71", "Ellis" },  { "MU72", "Ellis" },
                { "LV83", "Jadecdie" }, { "ND61", "Jadecdie" },
                { "LV82", "JC-Chop" },  { "NA46", "JC-Chop" }, { "Jade-S", "JC-Chop" },
                { "LR66", "Sicily" },   { "LR67", "Sicily" },  { "LibBasGolden", "Sicily" },
                { "PJ80", "Leda" },
                { "LR68", "Tonga" },
                { "KF89", "Cebu" },
                { "NS98", "Prinia" },
                { "QF25", "Ibiza" },  { "RJ66", "Ibiza" },
                { "QW67", "Caicos" },
                { "QG40", "Palma" },  { "RH37", "Palma" },
                { "RV93", "Donan" },
                { "SJ94", "Brava_C" },
                { "SJ93", "Brava_S" },
                { "QG39", "Lobos" },
                { "SQ70", "Tahiti" }
            };

            return projectMap.TryGetValue(codeName, out string? projectName) ? projectName : codeName;
        }
    }
}
