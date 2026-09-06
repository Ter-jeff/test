using System;
using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

namespace BinCutScriptLib.Reader
{
    internal class CfgFuseMain
    {
        private const string Getreadvalue = "GetReadValue";

        public static bool GetCfgFuse(ref OneTouchDown oneTouchDown, ref SiteInfo[] siteInfoArray, EnumJob enumJob)
        {
            bool? returnFlag = null;
            int oneTouchIndex = GetStartIndex(oneTouchDown, ref returnFlag);
            if (returnFlag != null)
            {
                return (bool)returnFlag;
            }

            for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (IsStopPoint(oneTouchDown.Lines[oneTouchIndex]))
                {
                    break;
                }

                if (!oneTouchDown.Lines[oneTouchIndex].Line.Contains(Getreadvalue))
                {
                    continue;
                }

                //Serial   Site  items
                //Site(3)  CFGFuse GetReadValue             Product_Identifier = 0
                //Site(4)  CFGFuse GetReadValue             Product_Identifier = 0
                //Site(3)  CFGFuse GetReadValue             VDD_SOC_MS001 = 637.5
                string[] spt = oneTouchDown.Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (spt.Length != 6)
                {
                    continue;
                }

                //if (!Regex.IsMatch(oneTouchDown.Lines[oneTouchIndex].Line, @"\(|\)"))
                if (!Reg.RegexSite2.IsMatch(oneTouchDown.Lines[oneTouchIndex].Line))
                {
                    continue;
                }

                string[] data = spt[0].Split(['(', ')'], StringSplitOptions.RemoveEmptyEntries);
                int site = 0;
                if (data.Length > 1)
                {
                    site = int.Parse(data[1]);
                }

                //Get Bin
                if (spt[3].Contains("Product_Identifier"))
                {
                    siteInfoArray[site].Bin = (int)double.Parse(spt[5]) + 1;
                }

                //Start to search efuse items
                if (Reg.RegexPerformance.IsMatch(spt[3]))
                {
                    string mode = Reg.RegexPerformance.Match(spt[3]).Groups["pmode"].ToString();
                    double gb = SiteInfoHelpers.GetEfuseGb(BinCutData.BinningTables[siteInfoArray[site].Bin - 1], mode, enumJob);
                    if (!siteInfoArray[site].EFuseValues.Exists(x => x.Name.EqualsIgnoreCase(mode)))
                    {
                        var row = new EFuseRow { Name = mode, Value = double.Parse(spt[5]), Gb = gb };
                        siteInfoArray[site].EFuseValues.Add(row);
                    }
                    else
                    {
                        IEnumerable<EFuseRow> eFuses = siteInfoArray[site].EFuseValues.Where(x => x.Name.EqualsIgnoreCase(mode));
                        foreach (EFuseRow eFuse in eFuses)
                        {
                            eFuse.Value = double.Parse(spt[5]);
                            eFuse.Gb = gb;
                        }
                    }
                }
            }
            return true;
        }

        private static int GetStartIndex(OneTouchDown oneTouchDown, ref bool? returnFlag)
        {
            bool isFoundCfg = false;
            int oneTouchIndex;
            for (oneTouchIndex = 0; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (oneTouchDown.Lines[oneTouchIndex].Line.Contains(Getreadvalue))
                {
                    isFoundCfg = true;
                    break;
                }
            }

            if (!isFoundCfg)
            {
                returnFlag = false;
            }
            return oneTouchIndex;
        }

        public static bool GetCp2PowerNames(OneTouchDown oneTouchDown, out List<string> powerNames)
        {
            powerNames = [];
            bool? returnFlag = null;
            int oneTouchIndex = GetStartIndex(oneTouchDown, ref returnFlag);
            if (returnFlag != null)
            {
                return (bool)returnFlag;
            }

            for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (IsStopPoint(oneTouchDown.Lines[oneTouchIndex]))
                {
                    break;
                }

                if (!oneTouchDown.Lines[oneTouchIndex].Line.Contains(Getreadvalue))
                {
                    continue;
                }

                //Serial   Site  items
                //Site(3)  CFGFuse GetReadValue             Product_Identifier = 0
                //Site(4)  CFGFuse GetReadValue             Product_Identifier = 0
                //Site(3)  CFGFuse GetReadValue             VDD_SOC_MS001 = 637.5
                string[] spt = oneTouchDown.Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (spt.Length != 6)
                {
                    continue;
                }

                if (Reg.RegexPerformance.IsMatch(spt[3]))
                {
                    powerNames.Add(spt[3]);
                }
            }
            powerNames = [.. powerNames.Distinct()];
            return true;
        }

        private static bool IsStopPoint(BinCutLineBase binCutLineBase)
        {
            return binCutLineBase.IsCfgFuseStopPoint();
        }
    }
}
