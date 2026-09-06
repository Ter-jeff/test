using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.InputManager.Data;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Reader.ConfigFile.NamingRule.Business;
using Automation.Reader.ScghFile.ProCharPatternSet.Base;
using Automation.Singleton;

using CommonLib.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using ScghLib.Base;
using ScghLib.Reader;

namespace Automation.Reader.ScghFile.ProCharPatternSet.Business
{
    public class HardIpPatSetConstructor : ProdCharPatSetConstructorBase, IProdCharPatsetConstructor<ProdCharPatternSetHardIp>
    {
        public string SheetName;
        public List<PatSet> MissingPatSets = new List<PatSet>();
        public Dictionary<string, int> PatSetNames = new Dictionary<string, int>();

        public IEnumerable<IProdCharItem> InitList;
        public IEnumerable<IProdCharItem> PayloadList;
        protected List<string> InitNamingFields;
        protected List<string> PayloadNamingFields;
        protected List<int> InitSequence = new List<int>();
        protected List<string> PatternListUsage;

        public HardIpPatSetConstructor(IEnumerable<IProdCharItem> inputList, List<string> patterns)
        {
            InitList = GetAllInitiRows(inputList);
            PayloadList = GetAllPayloadRows(inputList);
            PatternListUsage = patterns;
        }

        public HardIpPatSetConstructor(IEnumerable<IProdCharItem> initList, IEnumerable<IProdCharItem> payloadList)
        {
            InitList = initList;
            PayloadList = payloadList;
        }

        public HardIpPatSetConstructor(HardIpInputData hardIpInputData)
        {
            var configReader = new HardIpConfigFileReader();
            HardIpConfig config = configReader.ReadConfig();
            Dictionary<string, RtosConfig> namingRules = config.NamingRulesList;
            if (namingRules.ContainsKey("HardIP"))
            {
                InitNamingFields = namingRules["HardIP"].NamingRuleInitList;
                PayloadNamingFields = namingRules["HardIP"].NamingRulePayloadList;
                InitSequence = namingRules["HardIP"].NamingRuleInitSequenceList;
            }
        }

        public List<ProdCharPatternSetHardIp> WorkFlow(bool removeNonUsage = false)
        {
            if (removeNonUsage)
            {
                InitList = FilterProChar(InitList);
                PayloadList = FilterProChar(PayloadList);
            }

            var hardipPatsetList = new List<ProdCharPatternSetHardIp>();
            List<ProdCharPatternSetBase> basePatsetList = GetPatSetFromProdChar(InitList, PayloadList);
            foreach (ProdCharPatternSetBase patset in basePatsetList)
            {
                var hardIpPatset = new ProdCharPatternSetHardIp(patset) { SheetName = patset.SheetName, RowNum = patset.RowNum };
                hardipPatsetList.Add(hardIpPatset);
            }
            return hardipPatsetList;
        }

        public string GenPatSetName(ScghData scghData, List<string> list, string miscinfo, string bBblockName, PatSetSheet patsetSheet)
        {
            string patSetName = "";
            ProdCharSheetRow row;
            if (scghData == null || scghData.ConvertedPatternRowListByHardip.Count == 0)
            {
                row = null;
            }
            else
            {
                row = bBblockName.Equals("BB", StringComparison.OrdinalIgnoreCase) ? scghData.GetProdCharSheetRow(list) : null;
            }

            if (row != null && !string.IsNullOrEmpty(row.Item))
            {
                patSetName = row.Item;
            }
            else
            {
                var initList = new List<string>();
                initList.AddRange(list.GetRange(0, list.Count - 1));
                string payLoad = list.Last();
                string blockName = CommonGenerator.GetBlockNameFromSheetName(SheetName);
                string subblockName = CommonGenerator.GetSubBlockName(payLoad, miscinfo,
                    blockName).Replace("-", "").Trim();
                string initName = "";
                for (int i = 0; i < initList.Count; i++)
                {
                    //Naming rule has the corresponding number for the init
                    if (InitNamingFields.Count > i)
                    {
                        //compose all the init and payload to obtain the new name
                        string subName = GetSubName(initList[i], InitNamingFields[i]);
                        initName = Combination.CombineByUnderLine(blockName, subName);
                    }
                }

                string payLoadName = GetSubName(payLoad, PayloadNamingFields.First());
                patSetName = Combination.CombineByUnderLine(initName, payLoadName);
            }

            //Check and add PatSet
            int cnt = 1;
            var newPatSet = new PatSet();
            bool isMissingPat = false;
            string regInit = @"((init)|(in\d+))";
            foreach (string pat in list)
            {
                if (string.IsNullOrEmpty(pat))
                {
                    continue;
                }

                var patSet = new PatSetRow
                {
                    Burst = "Yes",
                    File = pat
                };
                if (!AcTSetCategoryMapSingleton.Instance().PatternList.ContainsKey(pat))
                {
                    patSet.Comment = "NotExist";
                    if (!Regex.IsMatch(pat, regInit, RegexOptions.IgnoreCase))
                    {
                        isMissingPat = true;
                    }
                }

                newPatSet.AddRow(patSet);
            }
            //AcTSetCategoryMapSingleton.Instance().PatternList

            newPatSet.PatSetName = patSetName;

            if (patsetSheet.Rows.Exists(x => x.PatSetName.Equals(newPatSet.PatSetName, StringComparison.OrdinalIgnoreCase)))
            {
                string exiPatSet = string.Empty;
                if (!patsetSheet.IsExistTheSamePatSet(newPatSet, out exiPatSet))
                {
                    do
                    {
                        newPatSet.PatSetName = patSetName + "_" + cnt;
                        cnt++;
                    }
                    while (patsetSheet.Rows.Exists(x =>
                        x.PatSetName.Equals(newPatSet.PatSetName, StringComparison.OrdinalIgnoreCase)) &&
                        !patsetSheet.IsExistTheSamePatSet(newPatSet, out _));
                    {
                        if (isMissingPat)
                        {
                            MissingPatSets.Add(newPatSet);
                        }

                        if (!patsetSheet.Rows.Exists(x => x.PatSetName.Equals(newPatSet.PatSetName, StringComparison.OrdinalIgnoreCase)))
                        {
                            patsetSheet.AddRow(newPatSet);
                        }
                    }
                }
                return string.IsNullOrEmpty(exiPatSet) ? newPatSet.PatSetName.ToUpper() : exiPatSet.ToUpper();
            }
            if (isMissingPat)
            {
                MissingPatSets.Add(newPatSet);
            }
            else
            {
                patsetSheet.AddRow(newPatSet);
            }

            return newPatSet.PatSetName.ToUpper();
        }
    }
}
