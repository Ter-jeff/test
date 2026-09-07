using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Enums;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EFusePreProcess
    {
        private List<EfuseReadRow> _efuseReadRows;

        public List<BitDefTable> FilterBankForBitDefTableByArraySizeSheet(List<BitDefTable> eFuseBitDefTables, EfuseArraySizeSheet efuseArraySizeSheet)
        {
            var bitDefTables = new List<BitDefTable>();
            foreach (BitDefTable bitDefTable in eFuseBitDefTables)
            {
                if (bitDefTable.BlockName.ContainsIgnoreCase("bira"))
                {
                    ErrorReportManager.AddError(EFuseErrorType.I_RuleViolationBank_01, bitDefTable.SheetName, 0, 0, [bitDefTable.BlockName]);
                    continue;
                }

                EfuseArraySizeRow arraySizeRow = efuseArraySizeSheet?.Rows.Find(x => x.FuseArray.EndsWith(bitDefTable.BlockName, StringComparison.CurrentCultureIgnoreCase));
                if (arraySizeRow != null &&
                    arraySizeRow.ProgrammedAt.Equals("SLT", StringComparison.CurrentCultureIgnoreCase))
                {
                    ErrorReportManager.AddError(EFuseErrorType.I_RuleViolationBank_02, bitDefTable.SheetName, 0, 0, [bitDefTable.BlockName]);
                    continue;
                }

                bitDefTables.Add(bitDefTable);
            }
            return bitDefTables;
        }

        public Dictionary<string, List<string>> CollectEachBankPowerPin(EfuseArraySizeSheet efuseArraySizeSheet)
        {
            var bankGroups = efuseArraySizeSheet.Rows.GroupBy(x => EFuseConst.GetBankName(x.FuseArray)).ToList();
            var dictionary = new Dictionary<string, List<string>>();
            foreach (IGrouping<string, EfuseArraySizeRow> bankGroup in bankGroups)
            {
                foreach (EfuseArraySizeRow efuseArraySizeRow in bankGroup)
                {
                    if (!efuseArraySizeRow.SupplyPinList.Any())
                    {
                        continue;
                    }

                    var powerPins = new List<string>();
                    powerPins.AddRange(efuseArraySizeRow.SupplyPinList);
                    if (!dictionary.ContainsKey(bankGroup.Key))
                    {
                        dictionary.Add(bankGroup.Key, powerPins);
                    }
                    else
                    {
                        dictionary[bankGroup.Key].AddRange(powerPins);
                    }
                }
            }
            return dictionary;
        }

        public EfuseFinalInstanceRows GetEfuseInstanceRows(List<EfusePatternRow> efusePatternRows, List<EfuseReadRow> efuseReadRows, bool hasEarly)
        {
            _efuseReadRows = efuseReadRows.Any() ? efuseReadRows : new List<EfuseReadRow>();
            EfuseFinalInstanceRows efuseInstanceRows = InitialInstance(efusePatternRows, efuseReadRows, hasEarly);
            efuseInstanceRows = efuseInstanceRows.ReTestNameDuplicateRows();
            return efuseInstanceRows.RePatternNameDuplicateRows();
        }

        public EfuseFinalInstanceRows GetEfuseInstanceRows(List<BinCutInstanceSheet> efuseInstanceSheets, List<EfusePatternRow> efusePatternRows)
        {
            var instanceSheetRows = new List<BinCutInstanceRow>();
            var instanceRows = new EfuseFinalInstanceRows();
            foreach (BinCutInstanceSheet sheet in efuseInstanceSheets)
            {
                instanceSheetRows.AddRange(sheet.Rows);
            }
            foreach (BinCutInstanceRow row in instanceSheetRows)
            {
                string patSetName = "";
                EfusePatternRow patternRow = efusePatternRows.Find(x => x.RowNum.Equals(row.RowNum) && x.SheetName.Equals(row.SheetName));
                var siteVar = new EfuseFlowSiteVar();
                if (!string.IsNullOrEmpty(row.SiteVar))
                {
                    string siteVarCondition = row.SiteVar.Split(' ').First();
                    string siteVarValue = row.SiteVar.Split(' ')[1] + ' ' + row.SiteVar.Split(' ')[2];
                    siteVar.SiteVarCondition = siteVarCondition;
                    siteVar.SiteVarValue = siteVarValue;
                }
                if (patternRow != null)
                {
                    if (patternRow.InitList.Any())
                    {
                        for (int i = 0; i < patternRow.InitList.Count; i++)
                        {
                            if (i == patternRow.InitList.Count - 1)
                            {
                                patSetName = Combination.CombineByUnderLine(new List<string> { "Pats", row.Instance });
                            }
                        }
                    }
                    else
                    {
                        if (patternRow.PatternType.IsDvrv)
                        {
                            string dvrvMode = patternRow.BankName;
                            dvrvMode = patternRow.PatternType.IsEarly ? "E" + dvrvMode : dvrvMode;
                            dvrvMode = !string.IsNullOrEmpty(patternRow.PatternType.PatJob) ? patternRow.PatternType.PatJob + dvrvMode : dvrvMode;
                            patSetName = Combination.CombineByUnderLine(patternRow.BankName, dvrvMode);
                        }
                        else
                        {
                            patSetName = Combination.CombineByUnderLine(patternRow.BankName, patternRow.PatternType.TestMode.ToString());
                        }
                    }
                }
                patSetName = string.IsNullOrWhiteSpace(row.PatSetNameOrange) ? patSetName : row.PatSetNameOrange;
                instanceRows.Add(GetBinCutFinalInstanceRow(patSetName, row.Instance, "", patternRow, EfuseExtraType.Normal, false, siteVar, "", null, "", row));
            }
            return instanceRows;
        }

        public EfuseFinalInstanceRows InitialInstance(List<EfusePatternRow> efusePatternRows, List<EfuseReadRow> efuseReadRows, bool hasEarly)
        {
            var instDataRows = new EfuseFinalInstanceRows();
            var bankGroups = efusePatternRows.GroupBy(x => x.BankName).ToList();
            //include hard code generate some test item
            foreach (IGrouping<string, EfusePatternRow> bankTable in bankGroups)
            {
                if (bankTable.Key.Equals(BankType.Unknow))
                {
                    continue;
                }

                //set integer
                if (!LocalSpecs.EfuseFlowUsedInteger.Contains(bankTable.Key + "Chk_Var"))
                {
                    LocalSpecs.EfuseFlowUsedInteger.Add(bankTable.Key + "Chk_Var");
                }

                //merge test item
                List<EfusePatternRow> testItems = MergeSamePatternType(bankTable);
                if (bankTable.Key.Equals(BankType.Cfg) || bankTable.Key.Equals(BankType.Sw0) || bankTable.Key.Equals(BankType.Sw1) || bankTable.Key.Equals(BankType.Sw2) || bankTable.Key.Equals(BankType.Sw3) || bankTable.Key.Equals(BankType.Sw4) || bankTable.Key.Equals(BankType.Sw5))
                {
                    instDataRows.AddRange(GenerateCfgEcidInstanceRow(bankTable.Key, testItems, EfuseExtraType.Normal));
                    if (hasEarly)
                    {
                        instDataRows.AddRange(GenerateCfgEcidInstanceRow(bankTable.Key, testItems, EfuseExtraType.Early));
                    }
                }
                else if (bankTable.Key.Equals(BankType.Ecid))
                {
                    instDataRows.AddRange(GenerateCfgEcidInstanceRow(bankTable.Key, testItems, EfuseExtraType.Normal));
                    instDataRows.AddRange(GenerateCfgEcidInstanceRow(bankTable.Key, testItems, EfuseExtraType.Deid));
                    instDataRows.AddRange(GenerateCfgEcidInstanceRow(bankTable.Key, testItems, EfuseExtraType.NonDeid));
                }
                else if (bankTable.Key.Equals(BankType.Mon))
                {
                    instDataRows.AddRange(GenerateMonInstanceRow(bankTable.Key, testItems, EfuseExtraType.Normal));
                }
                else
                {
                    instDataRows.AddRange(GenerateUdrInstanceRow(bankTable.Key, testItems, EfuseExtraType.Normal));
                }
            }
            return instDataRows;
        }

        internal List<EfusePatternRow> MergeSamePatternType(IEnumerable<EfusePatternRow> patternRows)
        {
            var newPatRowList = new List<EfusePatternRow>();
            var groupPats = patternRows.Where(x => x.PatternType.TestMode != EfuseTestMode.Unknow).GroupBy(x => x.PatternType.TestMode + "&" + x.PatternType.PatJob + "&" + x.PatternType.IsEarly + "&" + x.PatternType.IsDvrv).ToList();
            foreach (IGrouping<string, EfusePatternRow> group in groupPats)
            {
                var patRow = new EfusePatternRow
                {
                    BankName = group.First().BankName,
                    PatternType = group.First().PatternType,
                    RowNum = group.First().RowNum,
                    SheetName = group.First().SheetName,
                    PatList = new List<string>(),
                    PayloadList = new List<string>(),
                    InitList = new List<string>()
                };
                if (group.Count() > 1)
                {
                    var readWriteList = new List<string>();
                    foreach (EfusePatternRow row in group)
                    {
                        patRow.PayloadList.AddRange(row.PayloadList);
                        patRow.InitList.AddRange(row.InitList);
                        patRow.PatList.AddRange(row.PatList);
                        if (!string.IsNullOrEmpty(row.ReadWritePin))
                        {
                            readWriteList.Add(row.ReadWritePin);
                        }
                    }
                    if (readWriteList.Any())
                    {
                        readWriteList = readWriteList.Distinct().ToList();
                        patRow.ReadWritePin = string.Join(",", readWriteList);
                    }
                    patRow.InitList = patRow.InitList.Distinct().ToList();
                }
                else
                {
                    patRow.PayloadList.AddRange(group.First().PayloadList);
                    patRow.InitList.AddRange(group.First().InitList);
                    patRow.PatList.AddRange(group.First().PatList);
                    patRow.ReadWritePin = group.First().ReadWritePin;
                }
                newPatRowList.Add(patRow);
            }
            return newPatRowList;
        }

        private IEnumerable<EfuseFinalInstanceRow> GenerateCompareWrAndSyntaxCheckItem(string bankName, EfuseExtraType extraType, string extraName, EfuseFlowSiteVar siteVar, string mode = "", bool isDvRv = false, EfuseReadRow efuseReadRow = null, string voltage = "")
        {
            var instDataRows = new EfuseFinalInstanceRows();
            var emptyPat = new EfusePatternRow { BankName = bankName };
            string compareWrName = EFuseConst.CompareWr;
            const string syntaxName = EFuseConst.SyntaxCheck;

            if (isDvRv)
            {
                compareWrName = Combination.CombineByUnderLine(compareWrName, "RV");
            }

            string testItemName = Combination.CombineByUnderLine(new List<string> { EFuseConst.ConvertToFullBankName(bankName), compareWrName, extraName });
            instDataRows.Add(GetBinCutFinalInstanceRow("", testItemName, "", emptyPat, extraType, false, siteVar, "", efuseReadRow, voltage));

            mode = string.IsNullOrEmpty(mode) ? syntaxName : Combination.CombineByUnderLine(mode, syntaxName);
            testItemName = Combination.CombineByUnderLine(new List<string> { EFuseConst.ConvertToFullBankName(bankName), mode, extraName, isDvRv ? "RV" : "" });
            instDataRows.Add(GetBinCutFinalInstanceRow("", testItemName, "", emptyPat, extraType, false, new EfuseFlowSiteVar { SiteVarCondition = EFuseConst.SiteVarGr, SiteVarValue = bankName + EFuseConst.ChkVar0 }, "", efuseReadRow, voltage)); //hardcode...
            return instDataRows;
        }

        private IEnumerable<EfuseFinalInstanceRow> GenerateUdrInstanceRow(string bankName, List<EfusePatternRow> bankRows, EfuseExtraType extraType)
        {
            var instDataRows = new EfuseFinalInstanceRows();
            string extraName = extraType != EfuseExtraType.Normal ? EFuseConst.ConvertToExtraName(extraType) : "";
            var siteVarEq = new EfuseFlowSiteVar { SiteVarCondition = EFuseConst.SiteVarEq, SiteVarValue = bankName + EFuseConst.ChkVar1 };
            var siteVarGr = new EfuseFlowSiteVar { SiteVarCondition = EFuseConst.SiteVarGr, SiteVarValue = bankName + EFuseConst.ChkVar0 };

            //UFR
            //Pattern test mode == UFR
            List<EfusePatternRow> ufrItems = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.Ufr));
            instDataRows.AddRange(SetTestItemFinalInstanceRows(ufrItems, EFuseConst.UfrBlank, extraType, extraName));

            //USO
            //Pattern test mode == USO
            List<EfusePatternRow> usoItems = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.Uso));
            instDataRows.AddRange(SetTestItemFinalInstanceRows(usoItems, EFuseConst.UsoBlank, extraType, extraName));

            //USI
            //Pattern test mode == USI
            List<EfusePatternRow> usiItems = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.Usi));
            instDataRows.AddRange(SetTestItemFinalInstanceRows(usiItems, EFuseConst.Usi, extraType, extraName, siteVarEq));

            //UFP
            //Pattern test mode == UFP
            List<EfusePatternRow> ufpItems = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.Ufp));
            instDataRows.AddRange(SetTestItemFinalInstanceRows(ufpItems, EFuseConst.Ufp, extraType, extraName, siteVarEq));

            //UFR
            //Pattern test type == UFR 
            instDataRows.AddRange(SetTestItemFinalInstanceRows(ufrItems, EFuseConst.Ufr, extraType, extraName, siteVarGr));

            //USO
            //Pattern test type == USO
            IEnumerable<EfuseFinalInstanceRow> usoRows = SetTestItemFinalInstanceRows(usoItems, EFuseConst.UsoRead, extraType, extraName, siteVarGr);
            if (usoItems.Any())
            {
                foreach (EfuseFinalInstanceRow row in usoRows)
                {
                    instDataRows.Add(row);
                    //CompareWR and SyntaxCheck
                    string voltage = row.TestName.Split('_').Last();
                    if (!voltage.Equals("LV", StringComparison.CurrentCultureIgnoreCase) && !voltage.Equals("HV", StringComparison.CurrentCultureIgnoreCase) && !voltage.Equals("NV", StringComparison.CurrentCultureIgnoreCase))
                    {
                        voltage = "";
                    }

                    instDataRows.AddRange(GenerateCompareWrAndSyntaxCheckItem(bankName, extraType, extraName, siteVarEq, "USO", false, row.EfuseReadRow, voltage));
                }
            }

            //VER1
            //Pattern test mode == VER1
            List<EfusePatternRow> ver1Items = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.Ver1));
            instDataRows.AddRange(SetTestItemFinalInstanceRows(ver1Items, EFuseConst.Ver1, extraType, extraName));

            //VER2
            //Pattern test mode == VER2
            List<EfusePatternRow> ver2Items = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.Ver2));
            IEnumerable<EfuseFinalInstanceRow> ver2Rows = SetTestItemFinalInstanceRows(ver2Items, EFuseConst.Ver2Read, extraType, extraName);
            if (ver2Items.Any())
            {
                foreach (EfuseFinalInstanceRow row in ver2Rows)
                {
                    instDataRows.Add(row);
                    string voltage = row.TestName.Split('_').Last();
                    if (!voltage.Equals("LV", StringComparison.CurrentCultureIgnoreCase) && !voltage.Equals("HV", StringComparison.CurrentCultureIgnoreCase) && !voltage.Equals("NV", StringComparison.CurrentCultureIgnoreCase))
                    {
                        voltage = "";
                    }

                    string testItemName = Combination.CombineByUnderLine(new List<string> { EFuseConst.ConvertToFullBankName(bankName), EFuseConst.Ver2, EFuseConst.SyntaxCheck });
                    var efusePatternRow = new EfusePatternRow { BankName = bankName };
                    instDataRows.Add(GetBinCutFinalInstanceRow("", testItemName, "", efusePatternRow, extraType, false, null, "", row.EfuseReadRow, voltage));
                }
            }
            return instDataRows;
        }

        private IEnumerable<EfuseFinalInstanceRow> GenerateCfgEcidInstanceRow(string bankName, List<EfusePatternRow> bankRows, EfuseExtraType extraType)
        {
            var instDataRows = new EfuseFinalInstanceRows();
            string extraName = extraType != EfuseExtraType.Normal ? EFuseConst.ConvertToExtraName(extraType) : "";
            var siteVarEq = new EfuseFlowSiteVar { SiteVarCondition = EFuseConst.SiteVarEq, SiteVarValue = bankName + EFuseConst.ChkVar1 };
            var siteVarGr = new EfuseFlowSiteVar { SiteVarCondition = EFuseConst.SiteVarGr, SiteVarValue = bankName + EFuseConst.ChkVar0 };
            bool isCfg = bankName.Equals(BankType.Cfg);

            //BlankCheck
            //Pattern test mode == MarginRead || MarginRead_Full
            // full is the first priority
            List<EfusePatternRow> efusePatternRows = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.MarginReadFull));
            #region DVRV
            if (isCfg)
            {
                AddDvrvInstanceRows(instDataRows, bankName, bankRows, extraType, extraName, efusePatternRows, siteVarEq, siteVarGr);
            }
            #endregion

            //BlankCheck
            if (efusePatternRows.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(efusePatternRows, EFuseConst.BlankCheck, extraType, extraName));
            }
            else
            {
                efusePatternRows = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.MarginRead));
            }

            //WriteDSSC
            AddWriteDsscRows(instDataRows, bankRows, extraType, extraName, siteVarEq);

            //Read
            AddReadRows(instDataRows, bankName, efusePatternRows, extraType, extraName, siteVarEq, siteVarGr);

            //JTAGRead
            //Pattern type == JTAGRead
            List<EfusePatternRow> jtagReadItems = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.JtagRead));
            if (jtagReadItems.Any() && extraType != EfuseExtraType.Deid)
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(jtagReadItems, EFuseConst.JtagRead, extraType, ""));
            }

            //BKM and IEDA
            //SwitchFlagForPseudoFuse
            if (bankName.Equals(BankType.Cfg))
            {
                var efusePatRow = new EfusePatternRow { BankName = bankName };
                if (extraType == EfuseExtraType.Normal)
                {
                    //BKM and IEDA
                    string testItemName = Combination.CombineByUnderLine(EFuseConst.ConvertToFullBankName(bankName), EFuseConst.BkmFt);
                    instDataRows.Add(GetBinCutFinalInstanceRow("", testItemName, "", new EfusePatternRow { BankName = bankName }, extraType, false));
                    instDataRows.Add(GetBinCutFinalInstanceRow("", EFuseConst.IedaBkmSet, "", new EfusePatternRow { BankName = bankName }, extraType, false));
                    instDataRows.Add(GetBinCutFinalInstanceRow("", EFuseConst.IedaBkmGet, "", new EfusePatternRow { BankName = bankName }, extraType, false));
                }
                else
                {
                    //SwitchFlagForPseudoFuse
                    instDataRows.Add(GetBinCutFinalInstanceRow("", EFuseConst.SwitchFlag, "", efusePatRow, extraType, false));
                }
            }

            //ShowECIDData
            if (bankName.Equals(BankType.Ecid) &&
                (extraType == EfuseExtraType.Normal || extraType == EfuseExtraType.Deid))
            {
                instDataRows.Add(GetBinCutFinalInstanceRow("", EFuseConst.ShowEcid, "", new EfusePatternRow { BankName = bankName }, extraType, false));
            }

            return instDataRows;
        }

        private void AddDvrvInstanceRows(EfuseFinalInstanceRows instDataRows, string bankName, List<EfusePatternRow> bankRows, EfuseExtraType extraType, string extraName, List<EfusePatternRow> efusePatternRows, EfuseFlowSiteVar siteVarEq, EfuseFlowSiteVar siteVarGr)
        {
            if (extraType == EfuseExtraType.Normal)
            {
                AddDvrvNormalRows(instDataRows, bankName, bankRows, extraType, extraName, efusePatternRows, siteVarEq, siteVarGr);
            }
            else if (extraType == EfuseExtraType.Early)
            {
                AddDvrvEarlyRows(instDataRows, bankRows, extraType, extraName, siteVarEq, siteVarGr);
            }
        }

        private void AddDvrvNormalRows(EfuseFinalInstanceRows instDataRows, string bankName, List<EfusePatternRow> bankRows, EfuseExtraType extraType, string extraName, List<EfusePatternRow> efusePatternRows, EfuseFlowSiteVar siteVarEq, EfuseFlowSiteVar siteVarGr)
        {
            var dvrvBcs = new List<EfusePatternRow>();
            List<EfusePatternRow> flatCheck = bankRows.FindAll(x => x.PatternType.IsDvrv && !x.PatternType.IsEarly && x.PatternType.TestMode.Equals(EfuseTestMode.FlatCheck));
            if (flatCheck.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(flatCheck, nameof(EfuseTestMode.FlatCheck), extraType, extraName));
            }

            List<EfusePatternRow> dvWrites = bankRows.FindAll(x => x.PatternType.IsDvrv && !x.PatternType.IsEarly && x.PatternType.TestMode.Equals(EfuseTestMode.DvWrite));
            if (dvWrites.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(dvWrites, nameof(EfuseTestMode.DvWrite), extraType, extraName, siteVarEq));
            }

            List<EfusePatternRow> dvReads = bankRows.FindAll(x => x.PatternType.IsDvrv && !x.PatternType.IsEarly && x.PatternType.TestMode.Equals(EfuseTestMode.DvRead));
            if (dvReads.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(dvReads, nameof(EfuseTestMode.DvRead), extraType, extraName, siteVarGr));
            }

            if (efusePatternRows.Any())
            {
                dvrvBcs.AddRange(GetPatForBlankCheckRv(efusePatternRows));
                instDataRows.AddRange(SetTestItemFinalInstanceRows(dvrvBcs, EFuseConst.BlankCheck + "_RV", extraType, extraName));
            }

            List<EfusePatternRow> rvWrite = bankRows.FindAll(x => x.PatternType.IsDvrv && !x.PatternType.IsEarly && x.PatternType.TestMode.Equals(EfuseTestMode.RvWrite));
            if (rvWrite.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(rvWrite, nameof(EfuseTestMode.RvWrite), extraType, extraName, siteVarEq));
            }

            if (dvrvBcs.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(dvrvBcs, EFuseConst.Read + "_Real", extraType, extraName, siteVarGr));
                //CompareWR and SyntaxCheck
                instDataRows.AddRange(GenerateCompareWrAndSyntaxCheckItem(bankName, extraType, extraName, siteVarEq, "", true));
            }
        }

        private void AddDvrvEarlyRows(EfuseFinalInstanceRows instDataRows, List<EfusePatternRow> bankRows, EfuseExtraType extraType, string extraName, EfuseFlowSiteVar siteVarEq, EfuseFlowSiteVar siteVarGr)
        {
            List<EfusePatternRow> flatCheck = bankRows.FindAll(x => x.PatternType.IsDvrv && x.PatternType.IsEarly && x.PatternType.TestMode.Equals(EfuseTestMode.FlatCheck));
            if (flatCheck.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(flatCheck, nameof(EfuseTestMode.FlatCheck), extraType, extraName, siteVarEq));
            }

            List<EfusePatternRow> dvWrites = bankRows.FindAll(x => x.PatternType.IsDvrv && x.PatternType.IsEarly && x.PatternType.TestMode.Equals(EfuseTestMode.DvWrite));
            if (dvWrites.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(dvWrites, nameof(EfuseTestMode.DvWrite), extraType, extraName, siteVarEq));
            }

            List<EfusePatternRow> dvReads = bankRows.FindAll(x => x.PatternType.IsDvrv && x.PatternType.IsEarly && x.PatternType.TestMode.Equals(EfuseTestMode.DvRead));
            if (dvReads.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(dvReads, nameof(EfuseTestMode.DvRead), extraType, extraName, siteVarGr));
            }
        }


        private void AddWriteDsscRows(EfuseFinalInstanceRows instDataRows, List<EfusePatternRow> bankRows, EfuseExtraType extraType, string extraName, EfuseFlowSiteVar siteVarEq)
        {
            //WriteDSSC
            //Pattern test mode == WriteDSSC_H||WriteDSSC_Q||WriteDSSC || WriteDSSC_HF || WriteDSSC_Full
            //WriteDSSC_H is the first priority
            //WriteDSSC is the first priority if any
            List<EfusePatternRow> findAllFull =
                bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.WriteDssc));
            if (findAllFull.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(findAllFull, nameof(EfuseTestMode.WriteDssc), extraType, extraName, siteVarEq));
            }

            List<EfusePatternRow> findAll =
                bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.WriteDsscH));
            if (findAll.Any())
            {
                instDataRows.AddRange(SetTestItemFinalInstanceRows(findAll, nameof(EfuseTestMode.WriteDsscH), extraType, extraName, siteVarEq));
            }
            else
            {
                findAll = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.WriteDsscQ));
                if (findAll.Any())
                {
                    instDataRows.AddRange(SetTestItemFinalInstanceRows(findAll, nameof(EfuseTestMode.WriteDsscQ), extraType, extraName, siteVarEq));
                }
                else
                {
                    findAll = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.WriteDsscHf));
                    if (findAll.Any())
                    {
                        instDataRows.AddRange(SetTestItemFinalInstanceRows(findAll, nameof(EfuseTestMode.WriteDsscHf), extraType, extraName, siteVarEq));
                    }
                }
            }
        }

        private void AddReadRows(EfuseFinalInstanceRows instDataRows, string bankName, List<EfusePatternRow> efusePatternRows, EfuseExtraType extraType, string extraName, EfuseFlowSiteVar siteVarEq, EfuseFlowSiteVar siteVarGr)
        {
            //Read
            //Pattern test mode == MarginRead || MarginRead_Full
            if (efusePatternRows.Any())
            {
                IEnumerable<EfuseFinalInstanceRow> rows = SetTestItemFinalInstanceRows(efusePatternRows, EFuseConst.Read, extraType, extraName, siteVarGr);
                foreach (EfuseFinalInstanceRow row in rows)
                {
                    instDataRows.Add(row);
                    //CompareWR and SyntaxCheck
                    string voltage = row.TestName.Split('_').Last();
                    if (!voltage.Equals("LV", StringComparison.CurrentCultureIgnoreCase) && !voltage.Equals("HV", StringComparison.CurrentCultureIgnoreCase) && !voltage.Equals("NV", StringComparison.CurrentCultureIgnoreCase))
                    {
                        voltage = "";
                    }

                    instDataRows.AddRange(GenerateCompareWrAndSyntaxCheckItem(bankName, extraType, extraName, siteVarEq, "", false, row.EfuseReadRow, voltage));
                }
            }
        }

        private IEnumerable<EfusePatternRow> GetPatForBlankCheckRv(IEnumerable<EfusePatternRow> pats)
        {
            var bcPats = new List<EfusePatternRow>();
            foreach (EfusePatternRow pat in pats)
            {
                var bcPat = new EfusePatternRow
                {
                    SheetName = pat.SheetName,
                    RowNum = pat.RowNum,
                    BankName = pat.BankName,
                    PatternType = new EfusePatternType { IsDvrv = true, TestMode = pat.PatternType.TestMode },
                    InitList = pat.InitList,
                    PayloadList = pat.PayloadList,
                    PatList = pat.PatList,
                    ReadWritePin = pat.ReadWritePin
                };
                bcPats.Add(bcPat);
            }
            return bcPats;
        }

        private IEnumerable<EfuseFinalInstanceRow> GenerateMonInstanceRow(string bankName, List<EfusePatternRow> bankRows, EfuseExtraType extraType)
        {
            var instDataRows = new EfuseFinalInstanceRows();
            string extraName = extraType != EfuseExtraType.Normal ? EFuseConst.ConvertToExtraName(extraType) : "";
            var siteVarEq = new EfuseFlowSiteVar { SiteVarCondition = EFuseConst.SiteVarEq, SiteVarValue = bankName + EFuseConst.ChkVar1 };
            var siteVarGr = new EfuseFlowSiteVar { SiteVarCondition = EFuseConst.SiteVarGr, SiteVarValue = bankName + EFuseConst.ChkVar0 };

            //BlankCheck
            //Pattern test mode == JTAGRead
            List<EfusePatternRow> blankChecks = bankRows.FindAll(x => !x.PatternType.IsDvrv && (x.PatternType.TestMode.Equals(EfuseTestMode.JtagRead) || x.PatternType.TestMode.Equals(EfuseTestMode.ApbRead)));
            instDataRows.AddRange(SetTestItemFinalInstanceRows(blankChecks, ExtraNameWithApb(EFuseConst.BlankCheck), extraType, extraName));

            //WriteDSSC
            //Pattern test mode == JTAGWrite
            List<EfusePatternRow> findAll = bankRows.FindAll(x => !x.PatternType.IsDvrv && (x.PatternType.TestMode.Equals(EfuseTestMode.JtagWrite) || x.PatternType.TestMode.Equals(EfuseTestMode.ApbWrite)));
            instDataRows.AddRange(SetTestItemFinalInstanceRows(findAll, ExtraNameWithApb(EFuseConst.WriteDssc), extraType, extraName, siteVarEq));

            //Read
            //Pattern test mode == JTAGRead
            if (blankChecks.Any())
            {
                IEnumerable<EfuseFinalInstanceRow> rows = SetTestItemFinalInstanceRows(blankChecks, ExtraNameWithApb(EFuseConst.Read), extraType, extraName, siteVarGr);
                foreach (EfuseFinalInstanceRow row in rows)
                {
                    instDataRows.Add(row);
                    //CompareWR and SyntaxCheck
                    string voltage = row.TestName.Split('_').Last();
                    if (!voltage.Equals("LV", StringComparison.CurrentCultureIgnoreCase) && !voltage.Equals("HV", StringComparison.CurrentCultureIgnoreCase) && !voltage.Equals("NV", StringComparison.CurrentCultureIgnoreCase))
                    {
                        voltage = "";
                    }

                    instDataRows.AddRange(GenerateCompareWrAndSyntaxCheckItem(bankName, extraType, extraName, siteVarEq, "", false, row.EfuseReadRow, voltage));
                }
            }

            //CRC
            //Pattern test mode == CRC
            List<EfusePatternRow> crcItems = bankRows.FindAll(x => !x.PatternType.IsDvrv && x.PatternType.TestMode.Equals(EfuseTestMode.Crc));
            instDataRows.AddRange(SetTestItemFinalInstanceRows(crcItems, "CRC_Func", extraType, extraName, siteVarGr));

            return instDataRows;
        }

        private IEnumerable<EfuseFinalInstanceRow> SetTestItemFinalInstanceRows(List<EfusePatternRow> patRows, string mode, EfuseExtraType extraType, string extraName, EfuseFlowSiteVar siteVar = null)
        {
            var instDataRows = new EfuseFinalInstanceRows();
            for (int i = 0; i < patRows.Count; i++)
            {
                var bankReadVoltageList = new List<string>();
                EfusePatternRow patRow = patRows[i];
                string patSetName;
                string testItemName;
                EfuseReadRow readItem = ResolveReadItem(patRow, mode, extraType, ref bankReadVoltageList);
                string readMrName = GetEfuseReadMrName(readItem);
                List<string> voltageList = GetVoltageList(readItem, bankReadVoltageList);

                if (voltageList.Any())
                {
                    //bool isInitNv = true;
                    foreach (string voltage in voltageList)
                    {
                        if (patRow.PatternType.IsDvrv)
                        {
                            string dvrvMode = mode;
                            dvrvMode = patRow.PatternType.IsEarly ? "E" + dvrvMode : dvrvMode;
                            dvrvMode = !string.IsNullOrEmpty(patRow.PatternType.PatJob) ? patRow.PatternType.PatJob + dvrvMode : dvrvMode;
                            patSetName = Combination.CombineByUnderLine(patRow.BankName, dvrvMode);
                            if (patRow.PatternType.DvrvType == DvRvType.Dv)
                            {
                                dvrvMode += "Func";
                            }

                            testItemName =
                                Combination.CombineByUnderLine(new List<string> { EFuseConst.ConvertToFullBankName(patRow.BankName), dvrvMode, readMrName, voltage });
                            if (i == patRows.Count - 1)
                            {
                                patRow.PatternType.IsFinal = true;
                            }
                        }
                        else
                        {
                            patSetName = Combination.CombineByUnderLine(patRow.BankName, patRow.PatternType.TestMode.ToString());
                            testItemName = Combination.CombineByUnderLine(new List<string> { EFuseConst.ConvertToFullBankName(patRow.BankName), mode, extraName, readMrName, voltage });
                        }

                        instDataRows.AddRange(SetFinalInitInstanceRows(patRow, extraType, siteVar, testItemName, readItem, voltage));
                        instDataRows.Add(GetBinCutFinalInstanceRow(patSetName, testItemName, "", patRow, extraType, false, siteVar, "", readItem, voltage));
                    }
                }
                else
                {
                    if (patRow.PatternType.IsDvrv)
                    {
                        string dvrvMode = mode;
                        dvrvMode = patRow.PatternType.IsEarly ? "E" + dvrvMode : dvrvMode;
                        dvrvMode = !string.IsNullOrEmpty(patRow.PatternType.PatJob) ? patRow.PatternType.PatJob + dvrvMode : dvrvMode;
                        patSetName = Combination.CombineByUnderLine(patRow.BankName, dvrvMode);
                        if (patRow.PatternType.DvrvType == DvRvType.Dv)
                        {
                            dvrvMode += "Func";
                        }

                        testItemName =
                            Combination.CombineByUnderLine(new List<string> { EFuseConst.ConvertToFullBankName(patRow.BankName), dvrvMode });
                        if (i == patRows.Count - 1)
                        {
                            patRow.PatternType.IsFinal = true;
                        }
                    }
                    else
                    {
                        patSetName = Combination.CombineByUnderLine(patRow.BankName, patRow.PatternType.TestMode.ToString());
                        testItemName = Combination.CombineByUnderLine(new List<string> { EFuseConst.ConvertToFullBankName(patRow.BankName), mode, extraName, });
                    }

                    instDataRows.AddRange(SetFinalInitInstanceRows(patRow, extraType, siteVar, testItemName));
                    instDataRows.Add(GetBinCutFinalInstanceRow(patSetName, testItemName, "", patRow, extraType, false, siteVar));
                }

            }
            return instDataRows;
        }

        private EfuseReadRow ResolveReadItem(EfusePatternRow patRow, string mode, EfuseExtraType extraType, ref List<string> bankReadVoltageList)
        {
            EfuseReadRow readItem = null;
            if (mode == EFuseConst.Read || mode == "APB_" + EFuseConst.Read || mode == EFuseConst.Ufr || mode == EFuseConst.UsoRead)
            {
                readItem = ResolveSyntaxCheckReadItem(patRow, extraType);
                EfuseReadRow bankReadItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("BankRead", StringComparison.CurrentCultureIgnoreCase));
                bankReadVoltageList = GetVoltageList(bankReadItem, new List<string>());
            }
            else if (mode == EFuseConst.BlankCheck || mode == "APB_" + EFuseConst.BlankCheck)
            {
                readItem = ResolveBlankCheckReadItem(patRow, extraType);
            }
            else
            {
                readItem = ResolveOtherModeReadItem(patRow, mode, extraType);
            }
            return readItem;
        }

        private EfuseReadRow ResolveOtherModeReadItem(EfusePatternRow patRow, string mode, EfuseExtraType extraType)
        {
            EfuseReadRow readItem = null;
            if (mode == EFuseConst.JtagRead)
            {
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("TAPRead", StringComparison.CurrentCultureIgnoreCase) && (patRow.PatternType.IsEarly || extraType == EfuseExtraType.Early ? x.UserDefined.Equals("early", StringComparison.CurrentCultureIgnoreCase) : !x.UserDefined.Equals("early", StringComparison.CurrentCultureIgnoreCase)));
            }
            else if (mode == EFuseConst.UfrBlank)
            {
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("Blankcheck", StringComparison.CurrentCultureIgnoreCase) && (patRow.PatternType.IsEarly || extraType == EfuseExtraType.Early ? x.UserDefined.Equals("early", StringComparison.CurrentCultureIgnoreCase) : !x.UserDefined.Equals("early", StringComparison.CurrentCultureIgnoreCase)));
            }
            else if (mode == EFuseConst.UsoBlank)
            {
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("Blankcheck", StringComparison.CurrentCultureIgnoreCase) && (patRow.PatternType.IsEarly || extraType == EfuseExtraType.Early ? x.UserDefined.Equals("early", StringComparison.CurrentCultureIgnoreCase) : !x.UserDefined.Equals("early", StringComparison.CurrentCultureIgnoreCase)));
            }
            else if (mode == EFuseConst.Ver1)
            {
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.UserDefined.ContainsIgnoreCase("VER1"));
            }
            else if (mode == EFuseConst.Ver2 || mode == EFuseConst.Ver2Read)
            {
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.UserDefined.ContainsIgnoreCase("VER2"));
            }
            return readItem;
        }

        private EfuseReadRow ResolveSyntaxCheckReadItem(EfusePatternRow patRow, EfuseExtraType extraType)
        {
            EfuseReadRow readItem = null;
            if (extraType == EfuseExtraType.Deid)
            { //should add syntax check or blank check?
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("Syntaxcheck", StringComparison.CurrentCultureIgnoreCase) && x.UserDefined.Equals(EFuseConst.Deid, StringComparison.CurrentCultureIgnoreCase));
            }
            else if (extraType == EfuseExtraType.NonDeid)
            { //should add syntax check or blank check?
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("Syntaxcheck", StringComparison.CurrentCultureIgnoreCase) && x.UserDefined.Equals("X", StringComparison.CurrentCultureIgnoreCase));
            }
            else
            {
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("Syntaxcheck", StringComparison.CurrentCultureIgnoreCase) && (patRow.PatternType.IsEarly || extraType == EfuseExtraType.Early ? x.UserDefined.Equals("early", StringComparison.CurrentCultureIgnoreCase) : x.UserDefined.Equals("X", StringComparison.CurrentCultureIgnoreCase)));
            }
            return readItem;
        }

        private EfuseReadRow ResolveBlankCheckReadItem(EfusePatternRow patRow, EfuseExtraType extraType)
        {
            EfuseReadRow readItem = null;
            if (extraType == EfuseExtraType.Deid)
            { //should add syntax check or blank check?
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("BlankCheck", StringComparison.CurrentCultureIgnoreCase) && x.UserDefined.Equals("deid", StringComparison.CurrentCultureIgnoreCase));
            }
            else if (extraType == EfuseExtraType.NonDeid)
            {
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("BlankCheck", StringComparison.CurrentCultureIgnoreCase) && x.UserDefined.Equals("X", StringComparison.CurrentCultureIgnoreCase));
            }
            else if (extraType == EfuseExtraType.Early)
            {
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("BlankCheck", StringComparison.CurrentCultureIgnoreCase) && x.UserDefined.Equals("early", StringComparison.CurrentCultureIgnoreCase));
            }
            else
            {
                readItem = _efuseReadRows.FirstOrDefault(x => EFuseConst.GetBankName(x.Bank).Equals(patRow.BankName) && x.Purpose.Equals("BlankCheck", StringComparison.CurrentCultureIgnoreCase) && (patRow.PatternType.IsEarly ? x.UserDefined.Equals("early", StringComparison.CurrentCultureIgnoreCase) : x.UserDefined.Equals("X", StringComparison.CurrentCultureIgnoreCase)));
            }
            return readItem;
        }

        internal List<string> GetVoltageList(EfuseReadRow readItem, List<string> bankReadVoltage)
        {
            if (readItem == null)
            {
                return new List<string>();
            }

            var retList = readItem.JobDictionary.GroupBy(x => x.Value).ToDictionary(x => x.Key, x => x.ToList()).Select(x => x.Key).Where(x => !x.Contains('/') && !string.IsNullOrEmpty(x)).ToList();
            if (bankReadVoltage.Any())
            {
                retList.AddRange(bankReadVoltage);
            }

            return retList;

        }

        internal string GetEfuseReadMrName(EfuseReadRow readItem)
        {
            string ret = "";
            if (readItem == null)
            {
                return ret;
            }

            if (readItem.WriteRead.Equals("MarginRead", StringComparison.CurrentCultureIgnoreCase))
            {
                ret += "MR1";
            }
            else if (readItem.WriteRead.Equals("NormalRead", StringComparison.CurrentCultureIgnoreCase))
            {
                ret += "MR0";
            }
            return ret;
        }

        private IEnumerable<EfuseFinalInstanceRow> SetFinalInitInstanceRows(EfusePatternRow patRow, EfuseExtraType extraType, EfuseFlowSiteVar siteVar = null, string payloadType = null, EfuseReadRow readItem = null, string voltage = "")
        {
            bool isNeedMergeInitPats = LocalSpecs.Options.MergeInitPatterns;
            var instDataRows = new EfuseFinalInstanceRows();
            bool isFinalInit = false;
            for (int i = 0; i < patRow.InitList.Count; i++)
            {
                if (isNeedMergeInitPats && patRow.InitList.Any())
                {
                    if (i == patRow.InitList.Count - 1)
                    {
                        isFinalInit = true;

                        string patSetName =
                            Combination.CombineByUnderLine(new List<string>
                        {
                            EFuseConst.ConvertToFullBankName(patRow.BankName),
                            "Init",
                            "ALL"
                        });
                        string testItemName =
                            Combination.CombineByUnderLine(new List<string>
                        {
                            EFuseConst.ConvertToFullBankName(patRow.BankName),
                            "Init",
                            "ALL",
                            "NV"
                        });
                        if (!instDataRows.Exists(x => x.TestName.Equals(testItemName)))
                        {
                            instDataRows.Add(GetBinCutFinalInstanceRow(patSetName, testItemName, patRow.InitList.FirstOrDefault(), patRow, extraType, true,
                                siteVar, payloadType));
                        }
                    }
                }
                else
                {
                    if (i == patRow.InitList.Count - 1)
                    {
                        isFinalInit = true;
                    }

                    string testItemName = Combination.CombineByUnderLine(new List<string> { EFuseConst.ConvertToFullBankName(patRow.BankName), "Init" + i, "NV" });
                    if (!instDataRows.Exists(x => x.TestName.Equals(testItemName)))
                    {
                        instDataRows.Add(GetBinCutFinalInstanceRow("", testItemName, patRow.InitList[i], patRow, extraType, isFinalInit, siteVar, payloadType, readItem, voltage));
                    }
                }
            }
            if (readItem != null && readItem.InitList.Any())
            {
                string mrPat = readItem.InitList.Last();
                string[] payloadList = mrPat.Split('_');
                if (payloadList.Length > 7 && payloadList[6].ContainsIgnoreCase("MR"))
                {
                    string testItemName = Combination.CombineByUnderLine(new List<string> { EFuseConst.ConvertToFullBankName(patRow.BankName), "Init", "MR1", "NV" });
                    if (!instDataRows.Exists(x => x.TestName.Equals(testItemName)))
                    {
                        EfuseFinalInstanceRow beforeMrItem = instDataRows[patRow.InitList.Count - 1];
                        beforeMrItem.IsFinalInit = false;
                        instDataRows.Add(GetBinCutFinalInstanceRow("", testItemName, mrPat, patRow, extraType, isFinalInit, siteVar, payloadType, readItem, voltage));
                    }
                }
            }

            return instDataRows;
        }

        protected EfuseFinalInstanceRow GetBinCutFinalInstanceRow(string patSetName, string testName, string initPatternName, EfusePatternRow efusePatRow, EfuseExtraType extraType, bool isFinalInit, EfuseFlowSiteVar siteVar = null, string payloadType = null, EfuseReadRow readItem = null, string voltage = "", BinCutInstanceRow instanceSheetRow = null)
        {
            var row = new EfuseFinalInstanceRow();
            if (efusePatRow != null)
            {
                row.BankName = efusePatRow.BankName;
            }

            row.TestName = testName;
            row.TestNameTemp = testName;
            row.InitPatName = initPatternName;
            if (!string.IsNullOrEmpty(row.InitPatName))
            {
                row.PayloadType = payloadType;
            }

            row.ExtraType = extraType;
            row.PatSetName = patSetName;
            row.PatSetNameTemp = patSetName;
            row.EfusePatternRow = efusePatRow;
            row.IsFinalInit = isFinalInit;
            if (siteVar == null)
            {
                siteVar = new EfuseFlowSiteVar();
            }

            row.SiteVar = siteVar;
            if (instanceSheetRow != null)
            {
                row.EfuseInstanceRow = instanceSheetRow;
            }

            if (readItem != null)
            {
                row.EfuseReadRow = readItem;
            }

            if (!string.IsNullOrEmpty(voltage))
            {
                row.Voltage = voltage;
            }

            return row;
        }

        internal static string ExtraNameWithApb(string extraName)
        {
            return "APB_" + extraName;
        }
    }
}
