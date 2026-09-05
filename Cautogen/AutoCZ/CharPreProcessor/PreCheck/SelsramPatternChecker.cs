using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class SelsramPatternChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            bool isExistSelMappingTalbe = SelSrmReader.SelMappingList.Count > 0;
            /* check if selsram pattern is always the last init pattern in a single char row*/
            foreach (Characterization item in charList.Where(p => p.IsUse && p.IsSelSram))
            {
                var patterns = item.AllPatterns.Where(x => Regex.IsMatch(x.Key, "init")).Select(x => x.Value).ToList();
                int dssc = 0;
                for (int i = 0; i < patterns.Count; i++)
                {
                    if (Regex.IsMatch(patterns[i], "DSSC"))
                    {
                        dssc = i;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(patterns[i]) && i > dssc && dssc > 0)
                        {
                            ErrorManager.AddWarning(ErrorType.InitPatternbehindSelsram, item.SheetName, item.RowNum,
                                item.ColNum("init" + (i + 1)), "use", "Exist init pattern behind selsram pattern");
                            foreach (int col in item.ColNum("init" + (i + 1)))
                            {
                                ErrorReportManager.AddError(CharErrorType.W_InitPatternbehindSelsram_01, item.SheetName, item.RowNum, col, []);
                            }
                        }
                    }

                }
                if (isExistSelMappingTalbe)
                {
                    SelSrmItem target = null;
                    foreach (KeyValuePair<string, string> pattern in item.AllPatterns)
                    {
                        if (target != null)
                        {
                            break;
                        }

                        foreach (SelSrmItem selsrmItem in SelSrmReader.SelMappingList)
                        {
                            bool isMatch = true;
                            foreach (string regPattern in selsrmItem.Pattern.Split('*'))
                            {
                                if (!pattern.Value.ToUpper().Contains(regPattern.ToUpper()))
                                {
                                    isMatch = false;
                                }
                            }
                            if (isMatch)
                            {
                                target = selsrmItem;
                                break;//Find and break, do not need to continue search
                            }
                        }
                    }
                    if (target == null)// cannot find pattern from mapping table => notice to user
                    {
                        ErrorManager.AddError(ErrorType.SelsramPatternNotFoundInMappingTable, item.SheetName, item.RowNum,
    item.ColNum("userdef9"), "use", "Cannot find any init patterns in mapping table");
                        foreach (int col in item.ColNum("userdef9"))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_SelsramPatternNotFoundInMappingTable_01, item.SheetName, item.RowNum, col, []);
                        }
                    }
                    else // get related selsram set and check bit width with userdef9
                    {
                        string regbit = @"selsr[a]*m(?<data>\w+)";
                        string data = Regex.Match(item.UserDef9, regbit, RegexOptions.IgnoreCase).Groups["data"].Value;
                        if (data.Length != target.Rows.Count)
                        {
                            ErrorManager.AddError(ErrorType.SelsramBitMismatch, item.SheetName, item.RowNum,
    item.ColNum("userdef9"), "use", "UserDef9 data mismatch with mapping table");
                            foreach (int col in item.ColNum("userdef9"))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_SelsramBitMismatch_01, item.SheetName, item.RowNum, col, []);
                            }
                        }
                    }

                    //Adding gap gating between logic/sram pins.
                    //if (target!=null)
                    //{
                    //    //Comparing logic and sram pin list length

                    //    if(target.SramPins.Count == target.LogicPins.Count)
                    //    {
                    //        var deltaDict = new Dictionary<string, string>();
                    //        var logicPinStartVoltage = new Dictionary<string, string>();
                    //        var logicPinEndVoltage = new Dictionary<string, string>();


                    //        for (int i = 0; i< target.SramPins.Count; ++i) 
                    //        { 
                    //            var rowSramPin = item.PowerSupplyX.Where(p => p.Name.ToUpper() == string.Join("", target.SramPins[i].Split('_'))).FirstOrDefault();
                    //            var rowLogicPin = item.PowerSupplyX.Where(p => p.Name.ToUpper() == string.Join("", target.LogicPins[i].Split('_'))).FirstOrDefault();

                    //            double.TryParse(rowSramPin.Start, out var sramStartVoltageDouble);
                    //            double.TryParse(rowLogicPin.Start, out var logicStartVoltageDouble);
                    //            double.TryParse(rowLogicPin.Stop, out var logicEndVoltageDouble);

                    //            var deltaStart = Math.Abs(sramStartVoltageDouble - logicStartVoltageDouble);
                    //            var deltaStartEnd = Math.Abs(sramStartVoltageDouble - logicEndVoltageDouble);

                    //            if (deltaStart > 0.3 || deltaStartEnd > 0.3)
                    //                deltaDict.Add($"{rowSramPin.Name}@{rowLogicPin.Name}", "Start or End delta > 0.3");

                    //        }


                    //    }
                    //}
                }
            }
        }
    }
}
