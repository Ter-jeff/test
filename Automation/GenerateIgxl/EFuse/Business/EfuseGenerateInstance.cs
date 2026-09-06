using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Enums;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EfuseGenerateInstance
    {
        public List<InstanceRow> GenerateInstanceRows(IEnumerable<EfuseFinalInstanceRow> bankTable, List<string> powerPin = null)
        {
            bool isNeedMergeInitPats = LocalSpecs.Options.MergeInitPatterns;
            var instanceRows = new List<InstanceRow>();
            foreach (EfuseFinalInstanceRow bankItem in bankTable)
            {
                if (isNeedMergeInitPats && bankItem.TestName.Contains("Init"))
                {
                    continue;
                }
                if (!isNeedMergeInitPats && bankItem.TestName.Contains("Init"))
                {
                    //avoid duplicate 1.init 2.init+payload case
                    instanceRows.Add(GenerateInitInstanceRow(bankItem));
                    continue;
                }
                InstanceRow instanceRow = BuildBankInstanceRow(bankItem, powerPin);
                if (instanceRow != null)
                {
                    instanceRows.Add(instanceRow);
                }
            }

            return instanceRows;

        }

        private InstanceRow BuildBankInstanceRow(EfuseFinalInstanceRow bankItem, List<string> powerPin)
        {
            if (!bankItem.EfusePatternRow.PatternType.IsDvrv && bankItem.EfusePatternRow.PatternType.TestMode.Equals(EfuseTestMode.Crc))
            {
                return GenerateCrcInstanceRow(bankItem);
            }
            else if (bankItem.EfusePatternRow.PatternType.IsDvrv && bankItem.EfusePatternRow.PatternType.DvrvType == DvRvType.Dv)
            {
                return GenerateDvInstanceRow(bankItem, powerPin);
            }
            else if (bankItem.EfusePatternRow.PatternType.IsDvrv && bankItem.EfusePatternRow.PatternType.TestMode == EfuseTestMode.FlatCheck)
            {
                return GenerateFlatCheckInstanceRow(bankItem);
            }
            else if (bankItem.TestName.Contains(EFuseConst.CompareWr))
            {
                return GenerateCompareWrInstanceRow(bankItem);
            }
            else if (bankItem.TestName.Contains(EFuseConst.SyntaxCheck))
            {
                return GenerateSyntaxInstanceRow(bankItem);
            }
            else if (bankItem.TestName.Contains(EFuseConst.ShowEcid))
            {
                return AddInstanceItem(bankItem.TestName, "auto_ShowECIDData");
            }
            else if (bankItem.TestName.Contains(EFuseConst.EcidSorting))
            {
                return AddInstanceItem(bankItem.TestName, "BinSorting_Compare_FT_ECID_S");
            }
            return BuildBankInstanceRowExtra(bankItem, powerPin);
        }

        private InstanceRow BuildBankInstanceRowExtra(EfuseFinalInstanceRow bankItem, List<string> powerPin)
        {
            if (bankItem.TestName.Contains(EFuseConst.BkmFt))
            {
                return GenerateBkmFtInstanceRow(bankItem);
            }
            else if (bankItem.TestName.Contains(EFuseConst.IedaBkmSet) || bankItem.TestName.Contains(EFuseConst.IedaBkmGet))
            {
                return GenerateIedaInstanceRow(bankItem);
            }
            else if (bankItem.TestName.ContainsIgnoreCase(EFuseConst.Write) || (!bankItem.EfusePatternRow.PatternType.IsDvrv && bankItem.EfusePatternRow.PatternType.TestMode.Equals(EfuseTestMode.Usi)))
            {
                return GenerateWriteInstanceRow(bankItem, powerPin);
            }
            else if (bankItem.TestName.Equals(EFuseConst.SwitchFlag))
            {
                return AddInstanceItem(bankItem.TestName, "SwitchFlag");
            }
            else if (!bankItem.EfusePatternRow.PatternType.IsDvrv && bankItem.EfusePatternRow.PatternType.TestMode.Equals(EfuseTestMode.Ver1))
            {
                return GenerateVer1InstanceRow(bankItem);
            }
            else if (EFuseConst.BankIsUdr(bankItem.BankName) && !bankItem.EfusePatternRow.PatternType.IsDvrv && bankItem.EfusePatternRow.PatternType.TestMode.Equals(EfuseTestMode.Ufp))
            {
                return GenerateUdrUfpInstanceRow(bankItem, powerPin);
            }
            else if (EFuseConst.BankIsUdr(bankItem.BankName) && !bankItem.EfusePatternRow.PatternType.IsDvrv && bankItem.EfusePatternRow.PatternType.TestMode.Equals(EfuseTestMode.Ufr))
            {
                return GenerateUdrUfrInstanceRow(bankItem);
            }
            else if (bankItem.TestName.Contains(EFuseConst.BlankCheck) || bankItem.TestName.Contains(EFuseConst.Read))
            {
                return GenerateReadInstanceRow(bankItem);
            }
            return null;
        }

        public virtual InstanceRow GenerateAllBlankCheckInstance(string instanceName, EfuseFinalInstanceRow ecidItem)
        {
            var row = new InstanceRow { TestName = instanceName, VbtType = "VBT", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = ecidItem.EfusePatternRow.PayloadList.Any() ? ecidItem.GetTimeSet() : "";
            row.AcCategory = ecidItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = ecidItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.VbtFuncNameBankRead, "efuse");
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("ReadPatSet", ecidItem.GetPatSetNameForArgument());
            function.SetParamValue("PrePatSet", ecidItem.GetInitPatSetNameForArgument());
            function.SetParamValue("DigSource", GetTestAutoSwitchJtagTdi(ecidItem));


            if (!string.IsNullOrEmpty(ecidItem.EfusePatternRow.ReadWritePin))
            {
                function.SetParamValue("PinRead", ecidItem.EfusePatternRow.ReadWritePin);
            }

            function.SetParamValue("bank", ecidItem.BankName);
            if (ecidItem.BankName.Equals(BankType.Ecid, StringComparison.CurrentCultureIgnoreCase))
            {
                function.SetParamValue("ecid", "TRUE");
            }

            if (ecidItem.ExtraType == EfuseExtraType.Early)
            {
                function.SetParamValue("earlyfuse", "TRUE");
            }

            if (ecidItem.TestName.Contains(EFuseConst.BlankCheck))
            {
                function.SetParamValue("blankCheck", "TRUE");
            }

            function.SetParamValue("printdecode", "TRUE");
            function.SetParamValue("PrintDspWave", "TRUE");
            function.SetParamValue("blankCheckAll", "TRUE");
            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateFlatCheckInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("EFUSE_Flat_Pattern_Check", "efuse");
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;

            function.SetParamValue("Flat_Pattern", bankItem.GetPatSetNameForArgument());
            function.SetParamValue("bank", bankItem.BankName);

            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateDvInstanceRow(EfuseFinalInstanceRow bankItem, List<string> powerPin)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            string vbtName = "";
            if (bankItem.TestName.ContainsIgnoreCase("write"))
            {
                vbtName = "auto_ConfigWrite_CFG_DV";
            }
            else if (bankItem.TestName.ContainsIgnoreCase("read"))
            {
                vbtName = "Functional_T_updated";
            }

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "efuse");
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;

            //auto_ConfigWrite_CFG_DV
            function.SetParamValue("CFG_DV_pat", bankItem.GetPatSetNameForArgument());
            bool isPrg = false;
            foreach (string payload in bankItem.EfusePatternRow.PayloadList)
            {
                if (Regex.IsMatch(payload, @"\w*PRG\w*", RegexOptions.IgnoreCase))
                {
                    isPrg = true;
                }
            }
            if (isPrg)
            {
                string pins = powerPin.Any() ? string.Join(",", powerPin) : "";
                function.SetParamValue("PwrPin", pins);
                function.SetParamValue("vpwr", "1.8");
            }

            //Functional_T_updated
            function.SetParamValue("Patterns", bankItem.GetPatSetNameForArgument());

            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateInitInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = string.IsNullOrEmpty(bankItem.InitPatName) ? "" : bankItem.GetTimeSet(bankItem.InitPatName);
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.VbtFuncNameFunctionalTUpdated, "efuse");
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;

            function.SetParamValue("Patterns", bankItem.GetInitPatSetNameForArgument(true));
            function.SetParamValue("DigSource", GetTestAutoSwitchJtagTdi(bankItem));
            function.SetParamValue("ResultMode", "0");
            function.SetParamValue("RelayMode", "1");

            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateCrcInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow
            {
                TestName = bankItem.TestName,
                VbtType = "VBT",
                DcCategory = GetEfuseDcCategory(),
                DcSelector = "Typ",
                TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : ""
            };
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.VbtFuncNameFunctionalTUpdated, "efuse");
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("Patterns", bankItem.GetPatSetNameForArgument());
            function.SetParamValue("ResultMode", "0");
            function.SetParamValue("RelayMode", "1");
            function.SetParamValue("DigSource", GetTestAutoSwitchJtagTdi(bankItem));

            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateCompareWrInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT" };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_CompareWRData", "efuse");

            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("bank", bankItem.BankName);

            if (bankItem.ExtraType == EfuseExtraType.Early || bankItem.ExtraType == EfuseExtraType.Deid)
            {
                function.SetParamValue("earlyfuse", "TRUE");
            }

            if (bankItem.TestName.EndsWith("_RV", StringComparison.CurrentCultureIgnoreCase))
            {
                function.SetParamValue("RvOnly", "TRUE");
            }

            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateSyntaxInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT" };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_SyntaxCheck", "efuse");

            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            if (bankItem.TestName.Contains(EFuseConst.Ver2))
            {
                if (bankItem.BankName.Equals(BankType.UdrE))
                {
                    function.SetParamValue("bank", "CMP_E");
                }
                else if (bankItem.BankName.Equals(BankType.UdrP))
                {
                    function.SetParamValue("bank", "CMP_P");
                }
                else if (bankItem.BankName.Equals(BankType.UdrP0))
                {
                    function.SetParamValue("bank", "CMP_P0");
                }
                else if (bankItem.BankName.Equals(BankType.UdrP1))
                {
                    function.SetParamValue("bank", "CMP_P1");
                }
            }
            else
            {
                function.SetParamValue("bank", bankItem.BankName);
                function.SetParamValue("compareWR", "TRUE");
            }

            if (EFuseConst.BankIsUdr(bankItem.BankName) || bankItem.TestName.EndsWith("_RV", StringComparison.CurrentCultureIgnoreCase))
            {
                function.SetParamValue("RvOnly", "TRUE");
            }

            function.SetParamValue(bankItem.ExtraType == EfuseExtraType.Early || bankItem.ExtraType == EfuseExtraType.Deid ? "earlyfuse" : "checkAll", "TRUE");
            if (bankItem.BankName.Equals(BankType.Ecid, StringComparison.CurrentCultureIgnoreCase))
            {
                function.SetParamValue("ecid", "TRUE");
            }

            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateBkmFtInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName };

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Set_BKM", "efuse", true);
            if (function.Type == ".NET")
            {
                row.VbtType = ".NET";
                row.VbtName = function.FullFunctionName;
                row.ArgList = function.Parameters;
                return row;
            }

            function = TestProgram.VbtFunctionLib.GetFunctionByName("GetFusedBKMData", "efuse");
            row.VbtType = "VBT";
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("FuseType", bankItem.BankName);
            bool hasBkmProcess = LocalSpecs.HasBkmProcess;
            function.SetParamValue("cateName", hasBkmProcess ? "bkm_process" : "bkm_package");
            row.Args = function.ArgList;
            return row;
        }

        protected virtual InstanceRow GenerateIedaInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT" };

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("WriteIEDARegistry", "efuse");
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("FuseType", bankItem.BankName);
            if (bankItem.TestName.Equals(EFuseConst.IedaBkmGet))
            {
                function.SetParamValue("RegistryName", "BKM_Fuse");
            }
            else if (bankItem.TestName.Equals(EFuseConst.IedaBkmSet))
            {
                function.SetParamValue("RegistryName", "BKM");
            }

            bool hasBkmProcess = LocalSpecs.HasBkmProcess;
            function.SetParamValue("cateName", hasBkmProcess ? "bkm_process" : "bkm_package");
            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateWriteInstanceRow(EfuseFinalInstanceRow bankItem, List<string> powerPin)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_write", "efuse");
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("WritePatSet", bankItem.GetPatSetNameForArgument());
            if (!string.IsNullOrEmpty(bankItem.EfusePatternRow.ReadWritePin))
            {
                function.SetParamValue("PinWrite", bankItem.EfusePatternRow.ReadWritePin);
            }

            function.SetParamValue("bank", bankItem.BankName);
            if (bankItem.ExtraType == EfuseExtraType.Early || bankItem.ExtraType == EfuseExtraType.Deid)
            {
                function.SetParamValue("earlyfuse", "TRUE");
            }

            if (bankItem.BankName.Equals(BankType.Ecid))
            {
                function.SetParamValue("ecid", "TRUE");
            }

            bool isPrg = false;
            foreach (string payload in bankItem.EfusePatternRow.PayloadList)
            {
                if (Regex.IsMatch(payload, @"\w*PRG\w*", RegexOptions.IgnoreCase))
                {
                    isPrg = true;
                }
            }
            if (isPrg)
            {
                string pins = powerPin.Any() ? string.Join(",", powerPin) : "";
                function.SetParamValue("PwrPin", pins);
                function.SetParamValue("vpwr", "1.8");
            }
            function.SetParamValue("printdecode", "TRUE");
            function.SetParamValue("PrintDspWave", "TRUE");
            if (bankItem.EfusePatternRow.PatternType.IsDvrv)
            {
                function.SetParamValue("RvOnly", "TRUE");
            }
            function.SetParamValue("PrePatSet", bankItem.GetInitPatSetNameForArgument());
            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateVer1InstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("auto_Function_Test", "efuse");
            row.VbtType = "VBT";
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("bank", bankItem.BankName);

            function.SetParamValue("patset", bankItem.GetPatSetNameForArgument());
            function.SetParamValue("PrePatSet", bankItem.GetInitPatSetNameForArgument());
            function.SetParamValue("DigSource", GetTestAutoSwitchJtagTdi(bankItem));
            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateUdrUfpInstanceRow(EfuseFinalInstanceRow bankItem, List<string> powerPin)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("auto_UDR_UFP", "efuse");
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("UFP_pat", bankItem.GetPatSetNameForArgument());
            function.SetParamValue("PrePatSet", bankItem.GetInitPatSetNameForArgument());
            function.SetParamValue("DigSource", GetTestAutoSwitchJtagTdi(bankItem));

            bool isPrg = false;
            foreach (string payload in bankItem.EfusePatternRow.PayloadList)
            {
                if (Regex.IsMatch(payload, @"\w*PRG\w*", RegexOptions.IgnoreCase))
                {
                    isPrg = true;
                }
            }
            if (isPrg)
            {
                string pins = powerPin != null && powerPin.Any() ? string.Join(",", powerPin) : "";
                function.SetParamValue("PwrPin", pins);
                function.SetParamValue("vpwr", "1.8");
            }
            row.Args = function.ArgList;
            return row;
        }

        internal virtual InstanceRow GenerateUdrUfrInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("auto_UDR_UFR", "efuse");
            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("UFR_pat", bankItem.GetPatSetNameForArgument());
            function.SetParamValue("PrePatSet", bankItem.GetInitPatSetNameForArgument());
            function.SetParamValue("DigSource", GetTestAutoSwitchJtagTdi(bankItem));
            row.Args = function.ArgList;

            return row;
        }

        internal virtual InstanceRow GenerateReadInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = "VBT", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";

            row.PinLevels = bankItem.GenerateLevel();
            Function function;
            bool isJtag = false;
            if (!bankItem.EfusePatternRow.PatternType.IsDvrv &&
                (bankItem.EfusePatternRow.PatternType.TestMode.Equals(EfuseTestMode.JtagRead) || bankItem.EfusePatternRow.PatternType.TestMode.Equals(EfuseTestMode.JtagReadDap))
                && !bankItem.BankName.Equals(BankType.Mon))
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_TapRead", "efuse");
                isJtag = true;
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.VbtFuncNameBankRead, "efuse");
            }

            row.VbtName = function.FunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("ReadPatSet", bankItem.GetPatSetNameForArgument());
            function.SetParamValue("PrePatSet", bankItem.GetInitPatSetNameForArgument());
            function.SetParamValue("DigSource", GetTestAutoSwitchJtagTdi(bankItem));
            if (!string.IsNullOrEmpty(bankItem.EfusePatternRow.ReadWritePin))
            {
                function.SetParamValue("PinRead", bankItem.EfusePatternRow.ReadWritePin);
            }

            if (bankItem.TestName.Contains(EFuseConst.Ver2))
            {
                if (bankItem.BankName.Equals(BankType.UdrE))
                {
                    function.SetParamValue("bank", "CMP_E");
                }
                else if (bankItem.BankName.Equals(BankType.UdrP))
                {
                    function.SetParamValue("bank", "CMP_P");
                }
                else if (bankItem.BankName.Equals(BankType.UdrP0))
                {
                    function.SetParamValue("bank", "CMP_P0");
                }
                else if (bankItem.BankName.Equals(BankType.UdrP1))
                {
                    function.SetParamValue("bank", "CMP_P1");
                }
            }
            else
            {
                function.SetParamValue("bank", bankItem.BankName);
            }

            if (bankItem.BankName.Equals(BankType.Ecid, StringComparison.CurrentCultureIgnoreCase))
            {
                function.SetParamValue("ecid", "TRUE");
            }
            if (!isJtag)
            {
                if (bankItem.ExtraType == EfuseExtraType.Early || bankItem.ExtraType == EfuseExtraType.Deid)
                {
                    function.SetParamValue("earlyfuse", "TRUE");
                }

                if (bankItem.TestName.Contains(EFuseConst.BlankCheck))
                {
                    function.SetParamValue("blankCheck", "TRUE");
                }

                function.SetParamValue("printdecode", "TRUE");
                function.SetParamValue("PrintDspWave", "TRUE");
            }

            if (bankItem.EfusePatternRow.PatternType.IsDvrv)
            {
                function.SetParamValue("RvOnly", "TRUE");
            }
            row.Args = function.ArgList;
            return row;
        }

        public InstanceRow AddInstanceItem(string pTestName, string vbtName)
        {
            var instanceRow = new InstanceRow { TestName = pTestName, VbtType = "VBT" };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "efuse");
            instanceRow.VbtName = function.FunctionName;
            instanceRow.ArgList = function.Parameters;
            if (vbtName.Equals("BKM_Update"))
            {
                bool hasBkmProcess = LocalSpecs.HasBkmProcess;
                function.SetParamValue("cateName", hasBkmProcess ? "bkm_process" : "bkm_package");
                function.SetParamValue("bank", "CFG");
            }
            if (vbtName.Equals(EFuseConst.PseudoFuseWriteItem) || vbtName.Equals(EFuseConst.PseudoFuseReadItem))
            {
                string filePathConfig = LocalSpecs.Options.PseudoFuseFilePath;
#pragma warning disable Ter402 // IGXL network drive path for eFuse files
                string filePath = string.IsNullOrEmpty(filePathConfig) ? @"U:\TP-to-C651\OTC1_EXTERNALFILE\" : filePathConfig;
#pragma warning restore Ter402
                function.SetParamValue("FilePath", filePath);
            }
            instanceRow.Args = function.ArgList;
            return instanceRow;
        }

        public InstanceRow AddCsharpInstanceItem(string pTestName, string vbtName)
        {
            var instanceRow = new InstanceRow { TestName = pTestName, VbtType = ".NET" };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "efuse", true);
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            if (vbtName.Contains("Set_BKM"))
            {
                function.SetParamValue("bkmFolderPath", @"X:\BKM\");
                function.SetParamValue("bkmFolderPathOffline", @"D:\BKM\");
            }
            if (vbtName.Contains(EFuseConst.PseudoFuseWriteItem) || vbtName.Contains(EFuseConst.PseudoFuseReadItem))
            {
                string filePathConfig = LocalSpecs.Options.PseudoFuseFilePath;
#pragma warning disable Ter402 // IGXL network drive path for eFuse files
                string filePath = string.IsNullOrEmpty(filePathConfig) ? @"U:\TP-to-C651\OTC1_EXTERNALFILE\" : filePathConfig;
#pragma warning restore Ter402
                function.SetParamValue("filePath", filePath);
            }
            instanceRow.Args = function.ArgList;
            return instanceRow;
        }

        protected string GetEfuseDcCategory()
        {
            return "Efuse_X_X_X";
        }

        protected string GetEfuseDcSelector(string testName)
        {
            if (string.IsNullOrEmpty(testName))
            {
                return "Typ";
            }

            string voltage = testName.Split('_').LastOrDefault();
            if (voltage != null && voltage.Equals("LV", StringComparison.CurrentCultureIgnoreCase))
            {
                return "Min";
            }

            if (voltage != null && voltage.Equals("HV", StringComparison.CurrentCultureIgnoreCase))
            {
                return "Max";
            }

            return "Typ";
        }

        protected string GetTestAutoSwitchJtagTdi(EfuseFinalInstanceRow bankItem)
        {
            bool isNeedMergeInitPats = LocalSpecs.Options.MergeInitPatterns;
            string dsscPat = "";
            string ret = "";
            if (isNeedMergeInitPats)
            {
                if (bankItem.EfusePatternRow.InitList.Any(p => Regex.IsMatch(p, @"\w*DSSC\w*", RegexOptions.IgnoreCase)))
                {
                    dsscPat = Function.TestAutoSwitchJtagTdi;
                }
            }
            else
            {
                if (Regex.IsMatch(bankItem.InitPatName, @"\w*DSSC\w*", RegexOptions.IgnoreCase))
                {
                    dsscPat = Function.TestAutoSwitchJtagTdi;
                }
            }

            if (!string.IsNullOrEmpty(dsscPat))
            {
                string sendPinName = "JTAG_TDI";
                if (LocalSpecs.HardIpInfos != null)
                {
                    HardIpInfo target = LocalSpecs.HardIpInfos.GetHardIpInfo(dsscPat);
                    if (target != null && !string.IsNullOrEmpty(target.SendPinName))
                    {
                        sendPinName = target.SendPinName;
                    }
                }
                ret = "Test_AutoSwitch:" + sendPinName.ToUpper();
            }
            return ret;
        }
    }
}
