using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Static;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class MultipleInitChecker : HardIpPrecheckBase
    {
        public MultipleInitChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        public override void Check()
        {
            if (Pattern.Pattern.IsMultiple())
            {
                #region check whether patterns match to the rule of multiple init is match(only use timeset)
                foreach (List<string> patternList in Pattern.Pattern.PatternSetList)
                {
                    if (patternList.Count == 0)
                    {
                        continue;
                    }

                    string payload = patternList.Last();
                    List<HardIpInfo> infos = SearchInfo.GetHardIpInfos(patternList, LocalSpecs.HardIpInfos);
                    try
                    {
                        string payloadTimeset = infos.Last().TimeSet;
                        foreach (HardIpInfo init in infos)
                        {
                            if (payloadTimeset != init.TimeSet)
                            {
                                string errorMessage = string.Format("The Payload of multiple init : " + payload + " is not the same timeSet with pattern : " + init.Payload + ", please check");
                                Response.Report(errorMessage, EnumMessageLevel.Error, 0);
                                ErrorReportManager.AddError(
                                    HardIpErrorType.E_WrongTimeSet_01,
                                    Pattern.SheetName,
                                    Pattern.RowNum,
                                    0,
                                    [payload, init.Payload]
                                );
                            }
                        }
                    }
                    catch (Exception)
                    {
                        //pass
                    }
                }
                #endregion
            }
        }
    }
}
