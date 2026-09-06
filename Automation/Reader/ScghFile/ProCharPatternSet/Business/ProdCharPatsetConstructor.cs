using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Reader.ScghFile.ProCharPatternSet.Base;
using Automation.Utility.Atpg;

using CommonLib.Enums;
using CommonLib.Utility;

using LogLib.Static;

using ScghLib.Base;

namespace Automation.Reader.ScghFile.ProCharPatternSet.Business
{
    public interface IProdCharPatsetConstructor<T> where T : ProdCharPatternSetBase
    {
        List<T> WorkFlow(bool removeNonUsage = false);
    }

    public abstract class ProdCharPatSetConstructorBase
    {
        internal virtual List<ProdCharPatternSetBase> GetPatSetFromProdChar(IEnumerable<IProdCharItem> pInitList, IEnumerable<IProdCharItem> pPayoadList)
        {
            var missingInitNameList = new List<string>();
            var patsetList = new List<ProdCharPatternSetBase>();
            var initList = pInitList.ToList();
            var payloadList = pPayoadList.ToList();
            var all = new List<IProdCharItem>();
            all.AddRange(initList);
            all.AddRange(payloadList);
            bool hasInti = initList.Count != 0;
            foreach (IProdCharItem dataRow in payloadList)
            {
                if (IsPatternIgnored(dataRow.PayloadValue))
                {
                    continue;
                }

                if (hasInti)
                {
                    List<ProdCharPatternSetBase> patsets = CartersianGenerate(dataRow, all, missingInitNameList);
                    foreach (ProdCharPatternSetBase patset in patsets)
                    {
                        patset.RowNum = dataRow.RowNum;
                        patset.SheetName = dataRow.GetSourceSheet();
                    }
                    patsetList.AddRange(patsets);
                }
                else
                {
                    ProdCharPatternSetBase patset = DirectGenerate(dataRow, all);
                    patset.RowNum = dataRow.RowNum;
                    patset.SheetName = dataRow.GetSourceSheet();
                    patsetList.Add(patset);
                }
            }
            missingInitNameList = missingInitNameList.Distinct().ToList();
            foreach (string missingInit in missingInitNameList)
            {
                Response.Report($"Can not Find Init :{missingInit} ", EnumMessageLevel.Error, 0);
            }

            //PP_SCYA0_S_IN01_BI_XXXX_XXX_JTG_UNS_MS001_SI
            //PP_SCYA0_S_IN01_BI_XXXX_XXX_JTG_UNS_MS001_SI
            //PP_SCYA0_S_IN01_BI_XXXX_TDR_JTG_UNS_ALLFV_SI_DSRMXX_DSSC
            //Remove 2nd the same pattern
            List<ProdCharPatternSetBase> newPatsetList = DeleteSamePattern(patsetList);

            //PP_SCYA0_S_IN01_BI_XXXX_XXX_JTG_UNS_MS001_SI
            //PP_SCYA0_S_IN01_BI_XXXX_PLL_JTG_UNS_MS002_SI_FROM_F0024
            //Don't gen due to MS001 and From mode F0024 is not the same
            //newPatsetList = FilterProdCharPatternSetByFrom(newPatsetList);

            //PP_SCYA0_S_IN01_BI_XXXX_XXX_JTG_UNS_MS001_SI
            //PP_SCYA0_S_IN01_BI_XXXX_PLL_JTG_UNS_MS001_SI_FROM_F0024
            //PP_SCYA0_S_IN01_BI_XXXX_TDR_JTG_UNS_ALLFV_SI_DSRMXX_DSSC
            //Remove 2nd init pattern due to MS001 and MS001_X_FROM_F0024 are the same
            //newPatsetList = RemoveSameModeInit(newPatsetList);
            //Remove the same init and payload 
            //newPatsetList = RemoveSamePatterns(newPatsetList);
            newPatsetList = RemoveSamePatternsHashSet(newPatsetList);
            return newPatsetList;
        }

        internal List<ProdCharPatternSetBase> RemoveSamePatternsHashSet(List<ProdCharPatternSetBase> patsetList)
        {
            var prodCharPatternsetBases = new HashSet<ProdCharPatternSetBase>();
            foreach (ProdCharPatternSetBase row in patsetList)
            {
                prodCharPatternsetBases.Add(row);
            }
            return prodCharPatternsetBases.ToList();
        }

        internal List<ProdCharPatternSetBase> DeleteSamePattern(List<ProdCharPatternSetBase> patsetList)
        {
            foreach (ProdCharPatternSetBase patset in patsetList)
            {
                var newInitList = new Dictionary<int, PatternWithMode>();
                int i = -1;
                foreach (PatternWithMode init in patset.InitList.Values)
                {
                    i++;
                    if (string.IsNullOrEmpty(init.PatternName))
                    {
                        continue;
                    }

                    if (newInitList.Count == 0)
                    {
                        newInitList.Add(i, init);
                    }
                    else
                    {
                        if (init.PatternName != newInitList.Values.Last().PatternName)
                        {
                            newInitList.Add(i, init);
                        }
                    }
                }
                patset.InitList = newInitList;

                var newPayloadList = new List<PatternWithMode>();
                foreach (PatternWithMode payLoad in patset.PayloadList)
                {
                    if (string.IsNullOrEmpty(payLoad.PatternName))
                    {
                        continue;
                    }

                    if (newPayloadList.Count == 0)
                    {
                        newPayloadList.Add(payLoad);
                    }
                    else
                    {
                        if (payLoad != newPayloadList.Last())
                        {
                            newPayloadList.Add(payLoad);
                        }
                    }
                }
                patset.PayloadList = newPayloadList;
            }
            return patsetList;
        }

        protected ProdCharPatternSetBase DirectGenerate(IProdCharItem row, List<IProdCharItem> rowList)
        {
            var instanceTemp = new ProdCharPatternSetBase(row);
            int i = -1;
            foreach (string initName in row.GetInitList())
            {
                i++;
                //init name is null
                if (initName != "" && initName != "NA" && initName != "N/A")
                {
                    instanceTemp.InitList.Add(i, new PatternWithMode { PatternName = initName, Mode = GetMode(initName) });
                }
                else
                {
                    instanceTemp.InitList.Add(i, new PatternWithMode { PatternName = "", Mode = "" });
                }

            }
            foreach (string payload in GetPayloadList(row, rowList))
            {
                instanceTemp.PayloadList.Add(new PatternWithMode { PatternName = payload, Mode = GetMode(payload) });
            }
            return instanceTemp;
        }

        internal List<string> GetPayloadList(IProdCharItem row, List<IProdCharItem> rowList)
        {

            var currentPayloadList = new List<string>();
            foreach (string currentPayload in row.GetPayloadList())
            {
                if (currentPayload != "" && currentPayload != "NA" && currentPayload != "N/A")
                {
                    List<string> payloads = FindAllInMode(currentPayload, rowList);
                    if (payloads.Count != 0)
                    {
                        row.PayLoads += string.Join(",", payloads) + ";";
                    }

                    if (payloads.Count == 0)
                    {
                        payloads = FindAllInItem(currentPayload, rowList);
                        if (payloads.Count != 0)
                        {
                            row.PayLoads += string.Join(",", payloads) + ";";
                        }
                    }

                    if (payloads.Count == 0)
                    {
                        payloads.Add(currentPayload);
                        row.PayLoads += currentPayload + ",";
                    }
                    currentPayloadList.AddRange(payloads);
                }
                else
                {
                    currentPayloadList.Add("");
                }
            }
            return currentPayloadList;
        }

        protected virtual List<ProdCharPatternSetBase> CartersianGenerate(IProdCharItem row, List<IProdCharItem> rowList, List<string> missingInitiList)
        {
            var patsetList = new List<ProdCharPatternSetBase>();

            if (row.GetInitList().Any())
            {
                var initLists = new List<List<string>>();
                foreach (string init in row.GetInitList())
                {
                    var initList = new List<string>();
                    if (init != "" && init != "NA" && init != "N/A")
                    {
                        List<string> inits = FindInitInMode(init, rowList);
                        if (inits.Count == 0)
                        {
                            inits = FindInitInItem(init, rowList);
                        }

                        if (inits.Count == 0)
                        {
                            inits.Add(init);
                        }

                        initList.AddRange(inits);
                    }
                    else
                    {
                        initList.Add("");
                    }
                    initLists.Add(initList);
                }

                List<List<string>> combos = Combination.GetAllPossibleCombos(initLists);

                foreach (List<string> combo in combos)
                {
                    var patsetTemp = new ProdCharPatternSetBase(row);
                    for (int i = 0; i < combo.Count; i++)
                    {
                        patsetTemp.InitList.Add(i, new PatternWithMode { PatternName = combo[i], Mode = GetMode(combo[i]) });
                    }

                    patsetTemp.InitAliasList.AddRange(row.GetInitList());

                    foreach (string payload in GetPayloadList(row, rowList))
                    {
                        patsetTemp.PayloadList.Add(new PatternWithMode { PatternName = payload, Mode = GetMode(payload) });
                    }

                    patsetTemp.PayloadAliasList.AddRange(row.GetPayloadList());
                    patsetList.Add(patsetTemp);
                }
            }
            else
            {
                var patsetTemp = new ProdCharPatternSetBase(row);
                foreach (string payload in GetPayloadList(row, rowList))
                {
                    patsetTemp.PayloadList.Add(new PatternWithMode { PatternName = payload, Mode = GetMode(payload) });
                }

                patsetTemp.PayloadAliasList.AddRange(row.GetPayloadList());
                patsetList.Add(patsetTemp);
            }

            return patsetList;
        }

        protected List<string> FindInitInMode(string init, List<IProdCharItem> rowList)
        {
            var rows = rowList.FindAll(p => p.Mode.Equals(init, StringComparison.OrdinalIgnoreCase)).Select(p => p.PayloadValue).Distinct().ToList();
            return rows;
        }

        protected List<string> FindInitInItem(string init, List<IProdCharItem> rowList)
        {
            return rowList.FindAll(p => p.Item.Equals(init, StringComparison.OrdinalIgnoreCase)).Select(p => p.PayloadValue).Distinct().ToList();
        }

        protected List<string> FindAllInMode(string init, List<IProdCharItem> rowList)
        {
            var allList = rowList.FindAll(p => p.Mode.Equals(init, StringComparison.OrdinalIgnoreCase)).SelectMany(p => p.GetInitList()).ToList();
            allList.AddRange(rowList.FindAll(p => p.Mode.Equals(init, StringComparison.OrdinalIgnoreCase)).Select(p => p.PayloadValue).ToList());
            return allList;
        }

        protected List<string> FindAllInItem(string init, List<IProdCharItem> rowList)
        {
            var allList = rowList.FindAll(p => p.Item.Equals(init, StringComparison.OrdinalIgnoreCase)).SelectMany(p => p.GetInitList()).ToList();
            allList.AddRange(rowList.FindAll(p => p.Item.Equals(init, StringComparison.OrdinalIgnoreCase)).Select(p => p.PayloadValue).ToList());
            return allList;
        }

        protected List<IProdCharItem> GetAllInitiRows(IEnumerable<IProdCharItem> prodcharSheetRowList)
        {
            return prodcharSheetRowList.Where(p => IsInitPattern(p.PayloadValue)).ToList();
        }

        protected List<IProdCharItem> GetAllPayloadRows(IEnumerable<IProdCharItem> prodcharSheetRowList)
        {
            return prodcharSheetRowList.Where(p => !IsInitPattern(p.PayloadValue)).ToList();
        }

        protected virtual bool IsPatternIgnored(string pattern)
        {
            return false;
        }

        internal List<IProdCharItem> FilterProChar(IEnumerable<IProdCharItem> inputList)
        {
            return inputList.Where(a => a.Usage != "0").ToList();
        }

        /// <summary>
        /// Make a judgement whether the payload is a init payload
        /// </summary>
        /// <param name="pattern">pattern</param>
        /// <returns>whether the payload is a init</returns>
        internal bool IsInitPattern(string pattern)
        {
            string[] numbersStrings = pattern.Split('_');
            if (numbersStrings.Length < 4)
            {
                return false;
            }


            string lStrMatchPattern = "IN.*";
            if (Regex.IsMatch(numbersStrings[3], lStrMatchPattern, RegexOptions.IgnoreCase))
            {
                return true;
            }

            if (numbersStrings.Length < 7)
            {
                return false;
            }

            // for gfx SPC0 1
            if (Regex.IsMatch(numbersStrings[2], "L", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(numbersStrings[6], "SPC", RegexOptions.IgnoreCase))
            {
                return true;
            }

            // for RSCR_CLEAR
            if (Regex.IsMatch(pattern, "L_FULP_BI_CL00_EFU_JTG_REP_ALLFV_SI", RegexOptions.IgnoreCase))
            {
                return true;
            }

            // for RBOX_RESTOREs
            if (Regex.IsMatch(pattern, "L_FULP_BI_CL00_BHR_JTG_UNS_ALLFV_SI", RegexOptions.IgnoreCase))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Get sub-names from payloads according to naming rule
        /// </summary>
        /// <param name="name"></param>
        /// <param name="rule"></param>
        /// <returns></returns>
        internal string GetSubName(string name, string rule)
        {
            return AtpgService.GenGetSubName(name, rule);
        }

        internal string GetMode(string pattern)
        {
            string performanceMode = "";
            string[] subStrings = pattern.Split('_');

            if (subStrings.Length > 9 && Regex.IsMatch(subStrings[9], @"^M[a-zA-Z]+\d+[a-zA-Z0-9]$"))
            {
                if (!Regex.IsMatch(subStrings[9], "999|010"))
                {
                    performanceMode = subStrings[9];
                }
            }

            return performanceMode;
        }
    }
}
