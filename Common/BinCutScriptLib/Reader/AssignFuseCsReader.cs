using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.Reader
{
    internal class AssignFuseCsReader
    {
        private const string AssignFuse = "Assign read back value";

        public static void GetAssignFuseCs(ref OneTouchDown oneTouchDown, ref SiteInfo[] siteInfoArray, EnumJob enumJob, List<PinInfo> pinInfos)
        {
            //[INFO]  [Site 0] Assign read back value '610' for fuse 'VDD_ECPU_ME001' to cache.
            //[INFO]  [Site 0] Lock fuse 'VDD_ECPU_ME001' in cache.
            //[INFO]  [Site 2] Assign read back value '590' for fuse 'VDD_ECPU_ME001' to cache.
            //[INFO]  [Site 2] Lock fuse 'VDD_ECPU_ME001' in cache.

            List<BinCutLineBase> assignFuseLines = GetAssignFuseLines(oneTouchDown);
            List<BinCutLineBase> lines = GetFuseLines(assignFuseLines, pinInfos);

            foreach (BinCutLineBase line in lines)
            {
                Match match = Reg.RegexAssignFuse.Match(line.Line);
                if (!match.Success)
                {
                    continue;
                }

                int site = assignFuseLines[lines.IndexOf(line)].GetSite();
                if (!double.TryParse(match.Groups["value"].ToString(), out double value))
                {
                    continue;
                }

                string mode = BinCutAlgorithmService.GetModeByName(match.Groups["pmode"].ToString());
                double gb = SiteInfoHelpers.GetEfuseGb(BinCutData.BinningTables[siteInfoArray[site].Bin - 1], mode, enumJob);

                EFuseRow? eFuse = siteInfoArray[site].EFuseValues.FirstOrDefault(x => x.Name.EqualsIgnoreCase(mode));
                if (eFuse == null)
                {
                    siteInfoArray[site].EFuseValues.Add(new EFuseRow { Name = mode, Value = value, Gb = gb });
                }
                else
                {
                    eFuse.Value = value;
                    eFuse.Gb = gb;
                }
            }

        }

        public static List<BinCutLineBase> GetAssignFuseLines(OneTouchDown oneTouchDown)
        {
            return [.. oneTouchDown.Lines.Where(line => line.Line.Contains(AssignFuse))];
        }

        private static List<BinCutLineBase> GetFuseLines(List<BinCutLineBase> binCutLineBases, List<PinInfo> pinInfos)
        {
            return [.. binCutLineBases
                .Where(line => Reg.RegexAssignFuse.Match(line.Line).Success)
                .Where(line => pinInfos.Exists(pin => pin.PinMode.EqualsIgnoreCase(Reg.RegexAssignFuse.Match(line.Line).Groups["pmode"].ToString())))];
        }
    }
}
