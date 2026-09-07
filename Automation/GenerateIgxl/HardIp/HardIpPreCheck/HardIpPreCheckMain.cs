using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Extension;

using TestPlanLib.HardIpDc.BaseData;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class HardIpPreCheckMain
    {
        public void Check(Dictionary<string, HardIpSheet> planDic, HardIpDcSheet hardIpDcSheet)
        {
            // Check HardIpEfuseField
            new HardIpEfuseFieldChecker(planDic).Check();

            // Check wrong HardIp sheet name
            var wrongHardIpChecker = new WrongHardIpSheetNameChecker();
            wrongHardIpChecker.CheckWrongHardIpSheetName(EpWorkbook.TestPlanWorkbook);

            // Check missing Hardip dc pins in PinMap file
            var missingHardIpDcPinInPinMapChecker = new MissingHardIpDcPinInPinMapChecker();
            if (TestProgram.IgxlWorkBk.PinMapPair.Value.PinList.Count > 0 && hardIpDcSheet != null)
            {
                missingHardIpDcPinInPinMapChecker.CheckMissingPinInPinMap(hardIpDcSheet);
            }

            // Check patInfo send information
            if (LocalSpecs.HardIpInfos != null && LocalSpecs.HardIpInfos.Count > 0)
            {
                var patInfoChecker = new PatInfoChecker();
                patInfoChecker.CheckPatInfo(planDic);
            }

            // Check duplicate Store name
            var storeNameCheck = new DuplicateStoreNameChecker();
            storeNameCheck.CheckStoreName(planDic.SelectMany(block => block.Value.Rows).ToList());

            var subblockChecker = new SubBlockChecker(planDic);
            var list = planDic.SelectMany(x => x.Value.Rows).ToList();
            var refPatterns = list.GroupBy(x => x.SubBlockCopy).ToDictionary(g => g.Key, g => g.First());
            var refPatternsGlobal = list.GroupBy(x => x.SheetSubBlockName).ToDictionary(g => g.Key, g => g.First());
            foreach (KeyValuePair<string, HardIpSheet> patterns in planDic)
            {
                subblockChecker.ResetSubBlockList();
                foreach (HardIpPattern pattern in patterns.Value.Rows)
                {
                    subblockChecker.Check(patterns.Value, pattern);
                }
                bool patternBurst = patterns.Value.Rows.Any(x => x.IsBurst);
                subblockChecker.CheckNoBurstItem(patterns, patternBurst, refPatterns, refPatternsGlobal);
            }
            Dictionary<string, string> dir = GetParameterDictionary(planDic);

            Parallel.ForEach(planDic, patterns =>
            //foreach (var patterns in planDic)
            {
                foreach (HardIpPattern pattern in patterns.Value.Rows)
                {
                    new RegisterAssignmentChecker(patterns.Value, pattern, dir).Check();
                    new CalculationEquationChecker(patterns.Value, pattern, dir).Check();
                    new MultipleInitChecker(patterns.Value, pattern).Check();
                    new SelsramChecker(patterns.Value, pattern).Check();

                    if (pattern.IsBurst)
                    {
                        continue;
                    }

                    var precheckList = new List<HardIpPrecheckBase>
                    {
                        new MissingPinNameChecker(patterns.Value,pattern),
                        new MissingPinInPinMapChecker(patterns.Value,pattern),
                        new WrongMeasTypeChecker(patterns.Value,pattern),
                        new OppositeLimitChecker(patterns.Value,pattern),
                        new ForceConditionChecker(patterns.Value,pattern),
                        new PinSeqChecker(patterns.Value,pattern),
                        new ManualItemsChecker(patterns.Value,pattern),
                        new MeasCCaptureChecker(patterns.Value,pattern),
                        new DuplicateInstanceChecker(patterns.Value,pattern),
                        new MeasCPinChecker(patterns.Value,pattern)
                    };
                    foreach (HardIpPrecheckBase precheck in precheckList)
                    {
                        precheck.Check();
                    }
                }
            }
            );
        }

        private Dictionary<string, string> GetParameterDictionary(Dictionary<string, HardIpSheet> planDic)
        {
            var paraDic = new Dictionary<string, string>();
            foreach (string sheetName in planDic.Keys)
            {
                foreach (HardIpPattern planItem in planDic[sheetName].Rows)
                {
                    if (planItem.CalcEqn.ContainsIgnoreCase("alg"))
                    {
                        string regCalc = @"\((?<parameter>.*)\)";
                        foreach (string eqn in planItem.CalcEqn.Split(';'))
                        {
                            string paras = Regex.Match(eqn, regCalc, RegexOptions.IgnoreCase).Groups["parameter"].Value;
                            foreach (string parameter in Regex.Split(paras, @"[\,\+\&\:\@]", RegexOptions.IgnoreCase))
                            {
                                if (parameter != "" && !int.TryParse(parameter, out _) && !paraDic.ContainsKey(parameter.ToUpper()))
                                {
                                    paraDic.Add(parameter.ToUpper(), sheetName);
                                }
                            }
                        }
                    }

                    List<string> trimStoreNames = SearchInfo.GetTrimStoreNameByMiscInfo(planItem.MiscInfo);
                    if (trimStoreNames.Count > 0)
                    {
                        foreach (string trimStoreName in trimStoreNames)
                        {
                            if (!paraDic.ContainsKey(trimStoreName.ToUpper()))
                            {
                                paraDic.Add(trimStoreName.ToUpper(), sheetName);
                            }
                        }

                    }
                    string digCapNam = planItem.GetDigCapNameByMiscInfo();
                    if (digCapNam != "" && !paraDic.ContainsKey(digCapNam.ToUpper()))
                    {
                        paraDic.Add(digCapNam.ToUpper(), sheetName);
                    }
                    foreach (MeasPin pin in planItem.MeasPins)
                    {
                        if (pin.CusStr != "" && !paraDic.ContainsKey(pin.CusStr.ToUpper()))
                        {
                            paraDic.Add(pin.CusStr.ToUpper(), sheetName);
                        }
                    }
                }
            }
            return paraDic;
        }
    }
}
