using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Enums;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.Utility;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EfuseGenerateInstanceCs : EfuseGenerateInstance
    {
        public InstanceRow GenerateInstanceRowByInstanceSheet(EfuseFinalInstanceRow instanceSheetRow)
        {
            InstanceRow row = new InstanceRow
            {
                TestName = instanceSheetRow.EfuseInstanceRow.Instance,
                VbtType = ".NET",
                DcCategory = instanceSheetRow.EfuseInstanceRow.DCcategory.Split(' ').First()
            };
            if (!string.IsNullOrEmpty(row.DcCategory))
            {
                string selector = BinCutInstanceRowUtility.GetTypeByFlowNameOrDcCategory(instanceSheetRow.EfuseInstanceRow.DCcategory);
                row.DcSelector = GetDcSelector(selector);
            }
            row.TimeSets = instanceSheetRow.EfusePatternRow != null && (instanceSheetRow.EfusePatternRow.PayloadList.Any() || instanceSheetRow.EfuseInstanceRow.InitList.Any()) ? instanceSheetRow.GetTimeSet() : "";
            row.AcCategory = !string.IsNullOrEmpty(row.TimeSets) ? instanceSheetRow.GetAcCategory(row.TimeSets) : "";
            row.AcSelector = !string.IsNullOrEmpty(row.AcCategory) ? "Typ" : "";
            row.PinLevels = !string.IsNullOrEmpty(row.DcCategory) ? instanceSheetRow.GenerateLevel() : "";
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(instanceSheetRow.EfuseInstanceRow.FunctionName, "efuse", true);
            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            row.RowNum = instanceSheetRow.EfuseInstanceRow.RowNum;
            row.SheetName = instanceSheetRow.EfuseInstanceRow.SheetName;
            if (instanceSheetRow.EfuseInstanceRow != null && !string.IsNullOrEmpty(instanceSheetRow.EfuseInstanceRow.SheetName))
            {
                row.ColumnA = instanceSheetRow.EfuseInstanceRow.GetColumnA();
                row.ColumnA += ";DC Category:" + row.DcCategory;
            }

            if (function.PatternDic.Keys.Count != 0)
            {
                foreach (string patKey in function.PatternDic.Keys)
                {
                    if (Regex.IsMatch(patKey, @"^pattern", RegexOptions.IgnoreCase))
                    {
                        function.SetParamValue(patKey, instanceSheetRow.GetPatSetNameForArgument());
                    }
                }
            }

            if (!string.IsNullOrEmpty(instanceSheetRow.EfuseInstanceRow.UserFunction))
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(instanceSheetRow.EfuseInstanceRow.UserFunction, function);
            }

            row.Args = function.ArgList;
            return row;
        }

        public List<string> GenerateBinCheckInstance(InstanceSheet instanceSheet)
        {
            var binCheckInstanceList = new List<string>();
            var resultList = TestPlanStatic.HarvestingTruthTableSheets.SelectMany(x => x.Rows).Select(y => y.GetDecisionResult()).Distinct().ToList();

            foreach (string hardBin in resultList)
            {
                var row = new InstanceRow { TestName = "EFUSE_BinCheck_" + hardBin, VbtType = ".NET" };
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Bin_Fuse_Check", "efuse", true);
                row.VbtName = function.FullFunctionName;
                row.ArgList = function.Parameters;
                row.ColumnA = "FuseCheck instance for " + hardBin + " from Harvest truth tables";
                function.SetParamValue("hardBinNum", hardBin.Replace("Bin", ""));
                row.Args = function.ArgList;

                instanceSheet.AddRow(row);
                binCheckInstanceList.Add(row.TestName);
            }
            return binCheckInstanceList;
        }

        public override InstanceRow GenerateAllBlankCheckInstance(string instanceName, EfuseFinalInstanceRow ecidItem)
        {
            var row = new InstanceRow { TestName = instanceName, VbtType = ".NET", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = ecidItem.EfusePatternRow.PayloadList.Any() ? ecidItem.GetTimeSet() : "";
            row.AcCategory = ecidItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = ecidItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_read", "efuse", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenerateAllBlankCheckInstance(instanceName, ecidItem);
            }

            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;

            function.SetParamValue("patternName", ecidItem.GetPatSetNameForArgument());
            if (!string.IsNullOrEmpty(ecidItem.EfusePatternRow.ReadWritePin))
            {
                function.SetParamValue("pinName", ecidItem.EfusePatternRow.ReadWritePin);
            }

            function.SetParamValue("bankName", ecidItem.BankName.Equals(BankType.Cfg) ? "Config" : ecidItem.BankName);
            function.SetParamValue("earlyfuse", ecidItem.ExtraType == EfuseExtraType.Early || ecidItem.ExtraType == EfuseExtraType.Deid ? "TRUE" : "FALSE");
            function.SetParamValue("isEcid", ecidItem.BankName.Equals(BankType.Ecid, StringComparison.CurrentCultureIgnoreCase) ? "TRUE" : "FALSE");
            function.SetParamValue("blankCheckCurrentStage", ecidItem.TestName.Contains(EFuseConst.BlankCheck) ? "TRUE" : "FALSE");
            function.SetParamValue("rvOnly", ecidItem.TestName.EndsWith("_RV", StringComparison.CurrentCultureIgnoreCase) ? "TRUE" : "FALSE");
            function.SetParamValue("printDecode", "TRUE");
            function.SetParamValue("PrintDspWave", "TRUE");
            function.SetParamValue("initPatternSet", "");
            function.SetParamValue("writeReadVerify", "");
            function.SetParamValue("fullPrintForReadWriteVerify", "");

            row.Args = function.ArgList;
            return row;
        }

        internal override InstanceRow GenerateInitInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = ".NET", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = string.IsNullOrEmpty(bankItem.InitPatName) ? "" : bankItem.GetTimeSet(bankItem.InitPatName);
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.CSharpFuncNameFuncTestMain, "efuse", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenerateInitInstanceRow(bankItem);
            }

            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;

            function.SetParamValue("Patterns", bankItem.GetInitPatSetNameForArgument(true));
            function.SetParamValue("DigSource", GetTestAutoSwitchJtagTdi(bankItem));
            function.SetParamValue("ResultMode", "0");
            function.SetParamValue("RelayMode", "1");

            row.Args = function.ArgList;
            return row;
        }

        internal override InstanceRow GenerateCrcInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow
            {
                TestName = bankItem.TestName,
                VbtType = ".NET",
                DcCategory = GetEfuseDcCategory(),
                DcSelector = "Typ",
                TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : ""
            };
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.CSharpFuncNameFuncTestMain, "efuse", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenerateCrcInstanceRow(bankItem);
            }

            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("Patterns", bankItem.GetPatSetNameForArgument());
            function.SetParamValue("ResultMode", "0");
            function.SetParamValue("RelayMode", "1");
            function.SetParamValue("DigSource", GetTestAutoSwitchJtagTdi(bankItem));

            row.Args = function.ArgList;
            return row;
        }

        internal override InstanceRow GenerateCompareWrInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = ".NET" };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_Read", "efuse", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenerateCompareWrInstanceRow(bankItem);
            }

            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;

            function.SetParamValue("patternName", bankItem.GetPatSetNameForArgument());
            if (!string.IsNullOrEmpty(bankItem.EfusePatternRow.ReadWritePin))
            {
                function.SetParamValue("pinName", bankItem.EfusePatternRow.ReadWritePin);
            }

            function.SetParamValue("bankName", bankItem.BankName.Equals(BankType.Cfg) ? "Config" : bankItem.BankName);
            function.SetParamValue("earlyfuse", bankItem.ExtraType == EfuseExtraType.Early || bankItem.ExtraType == EfuseExtraType.Deid ? "TRUE" : "FALSE");
            function.SetParamValue("isEcid", bankItem.BankName.Equals(BankType.Ecid, StringComparison.CurrentCultureIgnoreCase) ? "TRUE" : "FALSE");
            function.SetParamValue("blankCheckCurrentStage", bankItem.TestName.Contains(EFuseConst.BlankCheck) ? "TRUE" : "FALSE");
            function.SetParamValue("rvOnly", bankItem.TestName.EndsWith("_RV", StringComparison.CurrentCultureIgnoreCase) ? "TRUE" : "FALSE");
            function.SetParamValue("printDecode", "TRUE");
            function.SetParamValue("PrintDspWave", "TRUE");
            function.SetParamValue("initPatternSet", "");
            function.SetParamValue("writeReadVerify", "");
            function.SetParamValue("fullPrintForReadWriteVerify", "");
            row.Args = function.ArgList;
            return row;
        }

        internal override InstanceRow GenerateSyntaxInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = ".NET" };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_SyntaxCheck", "efuse", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenerateSyntaxInstanceRow(bankItem);
            }

            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;
            return row;
        }

        internal override InstanceRow GenerateWriteInstanceRow(EfuseFinalInstanceRow bankItem, List<string> powerPin)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = ".NET", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_write", "efuse", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenerateWriteInstanceRow(bankItem, powerPin);
            }

            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("patternName", bankItem.GetPatSetNameForArgument());
            if (!string.IsNullOrEmpty(bankItem.EfusePatternRow.ReadWritePin))
            {
                function.SetParamValue("pinName", bankItem.EfusePatternRow.ReadWritePin);
            }

            function.SetParamValue("bankName", bankItem.BankName.Equals(BankType.Cfg) ? "Config" : bankItem.BankName);
            function.SetParamValue("earlyFuse", bankItem.ExtraType == EfuseExtraType.Early || bankItem.ExtraType == EfuseExtraType.Deid ? "TRUE" : "FALSE");
            function.SetParamValue("isEcid", bankItem.BankName.Equals(BankType.Ecid) ? "TRUE" : "FALSE");

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
                function.SetParamValue("powerPin", pins);
                function.SetParamValue("powerPinVoltage", "1.8");
            }
            function.SetParamValue("printDecode", "TRUE");
            function.SetParamValue("printDspWave", "TRUE");
            function.SetParamValue("rvOnly", bankItem.EfusePatternRow.PatternType.IsDvrv ? "TRUE" : "FALSE");
            function.SetParamValue("initPatternSet", bankItem.GetInitPatSetNameForArgument());
            row.Args = function.ArgList;
            return row;
        }

        internal override InstanceRow GenerateVer1InstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = ".NET", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            Function vbtFunctionBase = TestProgram.VbtFunctionLib.GetFunctionByName("FuncTestMain", "efuse", true);
            if (!vbtFunctionBase.IsFound || vbtFunctionBase.Type == "VBT")
            {
                return base.GenerateVer1InstanceRow(bankItem);
            }

            row.VbtName = vbtFunctionBase.FullFunctionName;
            row.ArgList = vbtFunctionBase.Parameters;
            vbtFunctionBase.SetParamValue("patterns", bankItem.GetPatSetNameForArgument());
            row.Args = vbtFunctionBase.ArgList;
            return row;
        }

        internal override InstanceRow GenerateUdrUfpInstanceRow(EfuseFinalInstanceRow bankItem, List<string> powerPin)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = ".NET", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("FuncTestMain", "efuse", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenerateUdrUfpInstanceRow(bankItem, powerPin);
            }

            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("patterns", bankItem.GetPatSetNameForArgument());
            row.Args = function.ArgList;
            return row;
        }

        internal override InstanceRow GenerateUdrUfrInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = ".NET", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = bankItem.GenerateLevel();

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("FuncTestMain", "efuse", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenerateUdrUfrInstanceRow(bankItem);
            }

            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("patterns", bankItem.GetPatSetNameForArgument());
            row.Args = function.ArgList;

            return row;
        }

        internal override InstanceRow GenerateReadInstanceRow(EfuseFinalInstanceRow bankItem)
        {
            var row = new InstanceRow { TestName = bankItem.TestName, VbtType = ".NET", DcCategory = GetEfuseDcCategory() };
            row.DcSelector = GetEfuseDcSelector(row.TestName);
            row.TimeSets = bankItem.EfusePatternRow.PayloadList.Any() ? bankItem.GetTimeSet() : "";
            row.AcCategory = bankItem.GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";

            row.PinLevels = bankItem.GenerateLevel();
            Function function;
            bool isJtag = false;
            if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) && !bankItem.EfusePatternRow.PatternType.IsDvrv &&
                (bankItem.EfusePatternRow.PatternType.TestMode.Equals(EfuseTestMode.JtagRead) || bankItem.EfusePatternRow.PatternType.TestMode.Equals(EfuseTestMode.JtagReadDap))
                && !bankItem.BankName.Equals(BankType.Mon))
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_TapRead", "efuse", true);
                if (!function.IsFound || function.Type == "VBT")
                {
                    return base.GenerateReadInstanceRow(bankItem);
                }

                isJtag = true;
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName("Bank_read", "efuse", true);
                if (!function.IsFound || function.Type == "VBT")
                {
                    return base.GenerateReadInstanceRow(bankItem);
                }
            }

            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            function.SetParamValue("patternName", bankItem.GetPatSetNameForArgument());
            if (!string.IsNullOrEmpty(bankItem.EfusePatternRow.ReadWritePin))
            {
                function.SetParamValue("pinName", bankItem.EfusePatternRow.ReadWritePin);
            }

            SetBankNameParamValue(function, bankItem);

            function.SetParamValue("isEcid", bankItem.BankName.Equals(BankType.Ecid, StringComparison.CurrentCultureIgnoreCase) ? "TRUE" : "FALSE");
            function.SetParamValue("earlyFuse", bankItem.ExtraType == EfuseExtraType.Early || bankItem.ExtraType == EfuseExtraType.Deid ? "TRUE" : "FALSE");
            function.SetParamValue("blankCheckCurrentStage", bankItem.TestName.Contains(EFuseConst.BlankCheck) ? "TRUE" : "FALSE");
            function.SetParamValue("printdecode", !isJtag ? "TRUE" : "FALSE");
            function.SetParamValue("PrintDspWave", !isJtag ? "TRUE" : "FALSE");

            if (function.FullFunctionName.Contains("Bank_TapRead"))
            {
                function.SetParamValue("writeReadVerify", "FALSE");
            }

            function.SetParamValue("RvOnly", bankItem.EfusePatternRow.PatternType.IsDvrv ? "TRUE" : "FALSE");
            row.Args = function.ArgList;
            return row;
        }

        private void SetBankNameParamValue(Function function, EfuseFinalInstanceRow bankItem)
        {
            if (bankItem.TestName.Contains(EFuseConst.Ver2))
            {
                if (bankItem.BankName.Equals(BankType.UdrE))
                {
                    function.SetParamValue("bankName", "CMP_E");
                }
                else if (bankItem.BankName.Equals(BankType.UdrP))
                {
                    function.SetParamValue("bankName", "CMP_P");
                }
                else if (bankItem.BankName.Equals(BankType.UdrP0))
                {
                    function.SetParamValue("bankName", "CMP_P0");
                }
                else if (bankItem.BankName.Equals(BankType.UdrP1))
                {
                    function.SetParamValue("bankName", "CMP_P1");
                }
            }
            else
            {
                function.SetParamValue("bankName", bankItem.BankName.Equals(BankType.Cfg) ? "Config" : bankItem.BankName);
            }
        }

        internal string GetDcSelector(string selectorName)
        {
            string selector = "Typ";
            switch (selectorName)
            {
                case "HV":
                    selector = "Max";
                    break;
                case "LV":
                    selector = "Min";
                    break;
                case "NV":
                    selector = "Typ";
                    break;
            }
            return selector;
        }
    }
}
