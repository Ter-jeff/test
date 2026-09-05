using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

using CommonLib.Enums;

namespace Automation.GenerateIgxl.PostAction.GenIgxlProj
{
    /// <summary>
    /// Injects <c>#Const isUFP = True</c> / <c>False</c> at line 2 of every
    /// .bas source (right after the IGLinkBase-emitted <c>Attribute VB_Name</c>
    /// header) so library code can branch on the equipment family.
    /// Lives outside <see cref="IIgxlPackager"/> on purpose — it must run
    /// regardless of whether IGXL packaging is available.
    /// </summary>
    public static class UfpBasMutator
    {
        public static void Apply(IEnumerable<string> sources)
        {
            string isUfp = TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlexPlus)
                ? "#Const isUFP = True"
                : "#Const isUFP = False";

            foreach (string data in sources)
            {
                if (data.IndexOf(".bas", StringComparison.OrdinalIgnoreCase) == -1)
                {
                    continue;
                }
                if (Path.GetFileName(data).StartsWith("DSP_", StringComparison.InvariantCultureIgnoreCase))
                {
                    continue;
                }

                string allText = File.ReadAllText(data);

                if (allText.Contains("#Const isUFP = "))
                {
                    string[] lines = File.ReadAllLines(data);
                    string isUfpFromBas = "";
                    foreach (string line in lines)
                    {
                        if (line.Contains("#Const isUFP = "))
                        {
                            isUfpFromBas = line;
                            break;
                        }
                    }
                    allText = allText.Replace(isUfpFromBas, isUfp);
                    File.WriteAllText(data, allText);
                }
                else
                {
                    // Insert at index 1 so the line lands AFTER the
                    // "Attribute VB_Name = ..." header that Cautogen / IGLinkBase
                    // writes at line 1. Prepending at line 0 would push the
                    // attribute down and break VBA module identification.
                    List<string> lines = File.ReadAllLines(data).ToList();
                    lines.Insert(1, isUfp);
                    File.WriteAllLines(data, lines);
                }
            }
        }
    }
}
