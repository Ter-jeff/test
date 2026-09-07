using System;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using LogLib.Utility;

using TestPlanLib.BinCut.BinCutConfig;
using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib
{
    internal static class BinCutAlgorithmService
    {
        public static EnumPowerType GetTypeByPowerName(string powerName)
        {
            //eg. VDD_SOC_MS001 => CORE_POWER
            EnumPowerType pwrNameType = EnumPowerType.Others;
            string pPowerToken = "";
            string[] spt = powerName.Split(['_'], StringSplitOptions.RemoveEmptyEntries);
            foreach (string tok in spt)
            {
                if (Reg.RegexRegexPerformance.IsMatch(tok))
                {
                    pPowerToken = tok;
                    break;
                }
            }
            if (pPowerToken.Length != 0)
            {
                powerName = powerName.Replace("_" + pPowerToken, "");
            }

            if (BinCutConfig.PowerType.ContainsKey(powerName.ToUpper()))
            {
                pwrNameType = BinCutConfig.PowerType[powerName.ToUpper()];
            }
            else
            {
                ErrorMessageBox.Show("The power name : " + powerName + " can't be found in the config setting.");
            }
            return pwrNameType;
        }

        public static string GetModeByName(string name)
        {
            //eg. VDD_SOC_MS001 => MS001
            //eg. BV_VDD_SOC_MS001 => MS001
            //eg. MC602 E1 Voltage => MC602
            string pPowerToken = "";
            string[] spt = name.Split(['_', ' ', '(', ')', ','], StringSplitOptions.RemoveEmptyEntries);
            foreach (string tok in spt)
            {
                if (Reg.RegexMode.IsMatch(tok.Trim('\'')))
                {
                    pPowerToken = tok;
                    break;
                }
            }
            return pPowerToken.Trim('\'');
        }

        public static string GetModeByNameCsharp(string name)
        {
            //eg. MS001_TD_Mbist_BV_Csharp => MS001
            //eg. VDD_SOC_MS001 => MS001
            //eg. BV_VDD_SOC_MS001 => MS001
            //eg. MC602 E1 Voltage => MC602
            string pPowerToken = "";
            string[] spt = name.Split(['_', ' ', '(', ')'], StringSplitOptions.RemoveEmptyEntries);
            foreach (string tok in spt)
            {
                if (Reg.RegexMode.IsMatch(tok))
                {
                    pPowerToken = tok;
                    break;
                }
            }
            return pPowerToken;
        }

        public static string GetPowerByName(string name)
        {
            //eg. VDD_SOC_MS001 => MS001
            //eg. BV_VDD_SOC_MS001 => MS001
            //eg. MC602 E1 Voltage => MC602
            string mode = "";
            if (name.Split([','], StringSplitOptions.RemoveEmptyEntries).Length > 1)
            {
                string[] spt = name.Split([','], StringSplitOptions.RemoveEmptyEntries);
                mode = GetModeByName(spt[0]);
            }
            else
            {
                mode = GetModeByName(name);
            }
            if (BinCutData.PinInfos.Exists(x => x.Mode.EqualsIgnoreCase(mode)))
            {
                return BinCutData.PinInfos.Find(x => x.Mode.EqualsIgnoreCase(mode))!.PinMode;
            }

            if (string.IsNullOrEmpty(mode))
            {
                if (BinCutData.PinInfos.Exists(x => x.Pin.EqualsIgnoreCase(name)))
                {
                    return BinCutData.PinInfos.Find(x => x.Pin.EqualsIgnoreCase(name))!.Pin;
                }
            }

            return "";
        }

        public static int GetPowerIndex(string powerName, SiteInfo siteInfo)
        {
            int pPowerIdx = -1;
            for (int i = 0; i < siteInfo.AllPowers.Count; i++)
            {
                if (siteInfo.AllPowers[i].PinMode.EqualsIgnoreCase(powerName))
                {
                    pPowerIdx = i;
                    break;
                }
            }
            return pPowerIdx;
        }

        public static double Floor3P125(double dVal)
        {
            if (BinCutData.BinningTables[0].StepSize == 5)
            {
                int quotient = (int)(dVal / 5.0);
                return quotient * 5.0;
            }
            else
            {
                int quotient = (int)(dVal / 3.125);
                return quotient * 3.125;
            }

        }

        public static double Ceiling3P125(double dVal)
        {
            if (BinCutData.BinningTables[0].StepSize == 5)
            {
                int quotient = (int)(dVal / 5.0);
                if (quotient * 5.0 < dVal)
                {
                    quotient++;
                }

                return quotient * 5.0;
            }
            else
            {
                int quotient = (int)(dVal / 3.125);
                if (quotient * 3.125 < dVal)
                {
                    quotient++;
                }

                return quotient * 3.125;
            }
        }

        public static double FloorIds(double dVal, double idsLsb)
        {
            const double tolerance = 1e-10;
            int cnt = 1;
            while ((idsLsb * cnt) - dVal < 0)
            {
                if (Math.Abs((idsLsb * cnt) - dVal) < tolerance)
                {
                    return Math.Round(dVal, 3);
                }

                cnt++;
            }
            return Math.Round(idsLsb * (cnt - 1), 3);
        }

        public static double CeilingIds(double dVal, double idsLsb)
        {
            const double tolerance = 1e-10;
            int cnt = 1;
            while ((idsLsb * cnt) - dVal < 0)
            {
                if (Math.Abs((idsLsb * cnt) - dVal) < tolerance)
                {
                    return Math.Round(dVal, 3);
                }

                cnt++;
            }
            return Math.Round(idsLsb * cnt, 3);
        }

        public static int GetOtherIndex(string dictExpress, BinningTable binningTable)
        {
            if (string.IsNullOrEmpty(dictExpress))
            {
                return -1;
            }

            dictExpress = BinCutConfig.DomainInOtherRail2Power.ContainsValue(dictExpress.ToUpper()) ?
                BinCutConfig.DomainInOtherRail2Power.FirstOrDefault(x => x.Value.EqualsIgnoreCase(dictExpress)).Key : dictExpress.Replace("VDD_", "");

            return binningTable.GetOtherIndex(dictExpress);
        }
    }
}
