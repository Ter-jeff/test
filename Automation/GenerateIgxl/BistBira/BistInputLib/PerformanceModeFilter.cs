using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BistBira.Base;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;

using ScghLib.Reader;

namespace Automation.GenerateIgxl.BistBira.BistInputLib
{
    public class PerformanceModeFilter
    {
        public void WorkFlow(BistProdFlowSheet prodFlow)
        {
            GetPerformanceMode(prodFlow);
        }

        internal void GetPerformanceMode(BistProdFlowSheet prodFlow)
        {
            string lStrCurrentMode = "";
            var currentModeList = new HashSet<string>();
            string otherModeInfoForNaming = "";  // To Save other performance mode in instance name by JN 20190617

            string lStrNextPayLoadVoltage = "";
            BistNaming naming = new BistNaming(new MbistConfig());
            for (int i = 0; i < prodFlow.Rows.Count; i++)
            {
                string lStrMode = "";
                string nextPayLoadPattern = "";
                bool isDsscRow = false;

                if (naming.IsSelDssc(prodFlow.Rows[i].Pattern))
                {
                    isDsscRow = true;
                }
                else if (naming.GetPatternType(prodFlow.Rows[i].Pattern) == BistConst.ConInit)
                {
                    lStrMode = RecogniseMode(prodFlow.Rows[i]);

                    otherModeInfoForNaming = "";

                    if (RecogniseResetMode(prodFlow.Rows[i]))
                    {
                        lStrCurrentMode = "";
                        currentModeList = new HashSet<string>();
                    }

                }
                if (lStrMode != "")
                {
                    lStrCurrentMode = lStrMode;
                    currentModeList.Add(lStrMode);
                }

                if (i > 0 && (currentModeList.Any() || !prodFlow.Rows[i - 1].IsDsscRow) && naming.GetPatternType(prodFlow.Rows[i - 1].Pattern) == BistConst.ConInit && naming.GetPatternType(prodFlow.Rows[i].Pattern) == BistConst.ConPayload)
                {
                    ResolveModeAtPayloadBoundary(prodFlow, naming, ref currentModeList, ref lStrCurrentMode, ref otherModeInfoForNaming);
                }



                int serIndex = i;

                if (isDsscRow && serIndex < prodFlow.Rows.Count - 1)
                {
                    // same dssc row original setting
                    prodFlow.Rows[i].IsDsscRow = true;
                    prodFlow.Rows[i].OriVoltage = prodFlow.Rows[i].Voltage;

                    SetOriPerformance(prodFlow.Rows[i], lStrCurrentMode);

                    DetermineNextPayload(prodFlow, naming, ref serIndex, lStrCurrentMode, ref lStrNextPayLoadVoltage, ref nextPayLoadPattern, out string lStrNextInitMode);


                    prodFlow.Rows[i].Voltage = lStrNextPayLoadVoltage;
                    prodFlow.Rows[i].NextPayLoadPattern = nextPayLoadPattern;



                    SetDsscVoltageMode(prodFlow.Rows[i], lStrNextPayLoadVoltage, lStrNextInitMode, otherModeInfoForNaming);

                }
                else
                {
                    SetNonDsscVoltageMode(prodFlow.Rows[i], lStrCurrentMode, otherModeInfoForNaming);
                }
            }
        }

        private void ResolveModeAtPayloadBoundary(BistProdFlowSheet prodFlow, BistNaming naming, ref HashSet<string> currentModeList, ref string lStrCurrentMode, ref string otherModeInfoForNaming)
        {
            string sheetModule = naming.GetModule(prodFlow.MbistSheet.SheetName);

            foreach (string mode in currentModeList)
            {
                string module = naming.GetPatternModule(mode, sheetModule);
                if (module == sheetModule)
                {
                    lStrCurrentMode = mode;
                }
            }

            currentModeList.Remove(lStrCurrentMode);
            otherModeInfoForNaming = string.Join("_", currentModeList);
            currentModeList = new HashSet<string>();
        }

        private void SetOriPerformance(BistProdFlowRow row, string lStrCurrentMode)
        {
            if (Regex.IsMatch(row.OriVoltage, BistConst.ConVmargin, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = row.OriVoltage;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConVdst, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = row.OriVoltage;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConVdisturb, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = BistConst.ConVdst;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConMhv1, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = BistConst.ConMhv1;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConMhv2, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = BistConst.ConMhv2;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConMhv3, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = BistConst.ConMhv3;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConMhv, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = BistConst.ConVmargin1;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConMlv1, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = BistConst.ConMlv1;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConMlv2, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = BistConst.ConMlv2;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConMlv3, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = BistConst.ConMlv3;
            }
            else if (Regex.IsMatch(row.OriVoltage, BistConst.ConMlv, RegexOptions.IgnoreCase))
            {
                row.OriPerformance = BistConst.ConVmargin3;
            }
            else if (BistNaming.IsNormalVoltage(row.OriVoltage))
            {
                row.OriPerformance = lStrCurrentMode;
            }
            else
            {
                if (row.OriVoltage.Contains(":"))
                {
                    row.OriPerformance = row.OriVoltage.Split(':').First();
                }
                else
                {
                    row.OriPerformance = row.OriVoltage;
                }
            }
        }

        private void DetermineNextPayload(BistProdFlowSheet prodFlow, BistNaming naming, ref int serIndex, string lStrCurrentMode, ref string lStrNextPayLoadVoltage, ref string nextPayLoadPattern, out string lStrNextInitMode)
        {
            while (prodFlow.Rows[serIndex + 1].Pattern == "" && serIndex < prodFlow.Rows.Count - 1)
            {
                serIndex++;
            }
            string nextLineType = naming.GetPatternType(prodFlow.Rows[serIndex + 1].Pattern);
            if (nextLineType == BistConst.ConInit)
            {
                string lStrMode = RecogniseMode(prodFlow.Rows[serIndex + 1]);
                if (lStrMode != "")
                {
                    lStrNextInitMode = lStrMode;
                }
                else
                {
                    lStrNextInitMode = lStrCurrentMode;
                }

                for (int j = serIndex + 1; j < prodFlow.Rows.Count; j++)
                {
                    if (prodFlow.Rows[j].Pattern == "")
                    {
                        continue;
                    }

                    if (naming.GetPatternType(prodFlow.Rows[j].Pattern) == BistConst.ConPayload)
                    {
                        lStrNextPayLoadVoltage = prodFlow.Rows[j].Voltage;
                        nextPayLoadPattern = prodFlow.Rows[j].Pattern;
                        break;
                    }
                }
            }

            else
            {
                lStrNextInitMode = lStrCurrentMode;
                lStrNextPayLoadVoltage = prodFlow.Rows[serIndex + 1].Voltage;
                nextPayLoadPattern = prodFlow.Rows[serIndex + 1].Pattern;
            }
        }

        private void SetDsscVoltageMode(BistProdFlowRow row, string lStrNextPayLoadVoltage, string lStrNextInitMode, string otherModeInfoForNaming)
        {
            if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConVmargin1, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVmargin1;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConVmargin3, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVmargin3;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConVmargin, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = row.Voltage;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConVdst, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = row.Voltage;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConVdisturb, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVdst;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConMhv1, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMhv1;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConMhv2, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMhv2;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConMhv3, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMhv3;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConMhv, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVmargin1;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConMlv1, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMlv1;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConMlv2, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMlv2;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConMlv3, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMlv3;
            }
            else if (Regex.IsMatch(lStrNextPayLoadVoltage, BistConst.ConMlv, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVmargin3;
            }
            else if (string.Equals(lStrNextPayLoadVoltage, BistConst.ConVrs, StringComparison.OrdinalIgnoreCase))  //2018/6/5 add VRS
            {
                row.VoltageMode = lStrNextPayLoadVoltage;    // if Next PayLoad level is VRS should keep VRS performance
            }
            else if (BistNaming.IsNormalVoltage(lStrNextPayLoadVoltage))
            {
                row.VoltageMode = lStrNextInitMode;
                row.OtherPModeInfoStr = otherModeInfoForNaming;
            }
            else if (BistNaming.IsNormalVoltage(lStrNextPayLoadVoltage.Split(',').Last()))
            {
                row.VoltageMode = PerformanceModeSingleton.Instance().FindPerformanceMode(lStrNextPayLoadVoltage);
                row.OtherPModeInfoStr = otherModeInfoForNaming;
            }
        }

        private void SetNonDsscVoltageMode(BistProdFlowRow row, string lStrCurrentMode, string otherModeInfoForNaming)
        {
            if (IsTargetPattern(row.Pattern, 6, "EFU"))
            {

            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConVmargin1, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVmargin1;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConVmargin3, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVmargin3;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConVmargin, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = row.Voltage;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConVdst, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = row.Voltage;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConVdisturb, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVdst;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConMhv1, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMhv1;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConMhv2, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMhv2;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConMhv3, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMhv3;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConMhv, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVmargin1;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConMlv1, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMlv1;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConMlv2, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMlv2;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConMlv3, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConMlv3;
            }
            else if (Regex.IsMatch(row.Voltage, BistConst.ConMlv, RegexOptions.IgnoreCase))
            {
                row.VoltageMode = BistConst.ConVmargin3;
            }
            else if (BistNaming.IsNormalVoltage(row.Voltage))
            {
                row.VoltageMode = lStrCurrentMode;
                row.PerformanceMode = lStrCurrentMode;
                row.OtherPModeInfoStr = otherModeInfoForNaming;
            }
            else
            {
                if (row.Voltage.Contains(":"))
                {
                    row.VoltageMode = row.Voltage.Split(':').First();
                }
                else
                {
                    row.VoltageMode = row.Voltage;
                }
            }
        }

        internal bool IsTargetPattern(string pPattern, int pPosition, string pKeyword)
        {

            string[] subName = pPattern.Split('_');
            if (pPosition >= subName.Length || pPosition < 0)
            {
                return false;
            }
            if (!subName[pPosition].Equals(pKeyword, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            return true;
        }

        private string RecogniseMode(BistProdFlowRow pRow)
        {

            string lStrMode = "";

            string[] subStrings = pRow.Pattern.Split('_');
            if (subStrings.Length > 9)
            {
                lStrMode = PerformanceModeSingleton.Instance().FindPerformanceMode(subStrings[9]);
            }

            return lStrMode;
        }

        private bool RecogniseResetMode(BistProdFlowRow pRow)
        {

            string lStrMatchPattern = "^F00";


            string[] subStrings = pRow.Pattern.Split('_');
            if (subStrings.Length > 9)
            {
                return Regex.IsMatch(subStrings[9], lStrMatchPattern, RegexOptions.IgnoreCase);
            }
            return false;


        }
    }
}
