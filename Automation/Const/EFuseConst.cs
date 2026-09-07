using System;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.EFuse.Enums;

namespace Automation.Const
{
    public static class EFuseConst
    {
        public const string End = "end";
        public const string Efuse = "Efuse";
        public const string TestInstSheet = "TestInst_eFuse";
        public const string Deid = "Deid";
        public const string NonDeid = "nonDeid";
        public const string Early = "Early";
        public const string FlowSheetName = "Flow_efuse";
        public const string Partial = "Partial";
        public const string BlankCheck = "BlankCheck";
        public const string Read = "Read";
        public const string CompareWr = "CompareWR";
        public const string SyntaxCheck = "SyntaxCheck";
        public const string BkmFt = "BKM_FT";
        public const string IedaBkmSet = "IEDA_Registry_BKM_SetFuse";
        public const string Usi = "USI";
        public const string Ufp = "UFP";
        public const string Ufr = "UFR";
        public const string UsoRead = "USO_Read";
        public const string Ver1 = "VER1";
        public const string SwitchFlag = "SwitchFlagForPseudoFuse";
        public const string Write = "WRITE";
        public const string Check = "Check";
        public const string ShowEcid = "ShowECIDData";

        public const string Comma = ",";
        public const string VddBin = "vddbin";
        public const string Decimal = "Decimal";
        public const string Na = "N/A";
        public const string Default = "Default";
        public const string Real = "Real";
        public const string Alias = "alias";
        public const string Io = "I/O";
        public const string Reserve = "Reserve";
        public const string Error = "Error";

        public const string EcidSorting = "ECID_Sorting";
        public const string EcidFlowSheet = "Flow_efuse_ECID";
        public const string EcidDeidFlowSheet = "Flow_efuse_ECID_Deid";

        public const string BitDefTableFileName = "EFUSE_BitDef_Table";
        public const string ConfigTableFileName = "Config_table";
        public const string ConfigTableSvmFileName = "Config_table_SVM";
        public const string EvsAllBc = "EVS_All_Blank_Check";
        public const string BankEnable = "Enable";
        public const string BankFail = "Fail";
        public const string BankRead = "BankRead";
        public const string PreWrite = "PreWrite";
        public const string PseudoFuse = "PseudoFuse";
        public const string PseudoFuseWriteItem = "PseudoFuse_WriteToFile";
        public const string PseudoFuseReadItem = "PseudoFuse_ReadFromFile";
        public const string WriteDssc = "WriteDSSC";
        public const string JtagRead = "JTAGRead";
        public const string BkmCp = "BKM_CP";
        public const string IedaBkmGet = "IEDA_Registry_BKM_GetFuse";
        public const string UfrBlank = "UFR_Blank";
        public const string UsoBlank = "USO_BlankCheck";
        public const string Ver2 = "VER2";
        public const string Ver2Read = "VER2_Read";
        public const string SiteVarEq = "site-var=";
        public const string SiteVarGr = "site-var>";
        public const string ChkVar1 = "Chk_Var 1";
        public const string ChkVar0 = "Chk_Var 0";
        public const string GoodBinRetestDetection = "GoodBin_Retest_Detection";
        public const string BypReTestCheck = "BypReTestCheck";

        public const string BinCut = "bincut";
        public const string VddError = "No Binning Equation. Binary=";
        public const string BkmName = "BKM_Info";

        public static bool BankIsUdr(string name)
        {
            return name.Equals(BankType.UdrE) || name.Equals(BankType.UdrP) || name.Equals(BankType.UdrP0) ||
                   name.Equals(BankType.UdrP1) || name.Equals(BankType.CmpE) || name.Equals(BankType.CmpP);
        }

        private class ValidateObject
        {
            public string Name { get; set; }
            public string SecName { get; set; }
        }

        private static readonly (
            Func<ValidateObject, string> Selector,
            string[] Keywords,
            string BankType
        )[] _bankNameRules =
        {
            (vo => vo.Name, new[] { "CONFIG", "CFG" }, BankType.Cfg),
            (vo => vo.Name, new[] { "UDRM0", "UDR_M0" }, BankType.UdrM0),
            (vo => vo.Name, new[] { "UDRM1", "UDR_M1" }, BankType.UdrM1),
            (vo => vo.Name, new[] { "UDRM", "UDR_M" }, BankType.UdrM),
            (vo => vo.Name, new[] { "UDRE", "UDR_E" }, BankType.UdrE),
            (vo => vo.Name, new[] { "UDRP0", "UDR_P0" }, BankType.UdrP0),
            (vo => vo.Name, new[] { "UDRP1", "UDR_P1" }, BankType.UdrP1),
            (vo => vo.Name, new[] { "UDRP", "UDR_P" }, BankType.UdrP),
            (vo => vo.Name, new[] { "UFA", "USO", "USI", "SMR" }, BankType.UdrE),
            (vo => vo.Name, new[] { "ECID", "ECD" }, BankType.Ecid),
            (vo => vo.Name, new[] { "MONITOR", "MON", "SEF" }, BankType.Mon),
            (vo => vo.Name, new[] { "CMP_E" }, BankType.CmpE),
            (vo => vo.Name, new[] { "CMP_P" }, BankType.CmpP),
            (vo => vo.Name, new[] { "SW0" }, BankType.Sw0),
            (vo => vo.Name, new[] { "SW1" }, BankType.Sw1),
            (vo => vo.Name, new[] { "SW2" }, BankType.Sw2),
            (vo => vo.Name, new[] { "SW3" }, BankType.Sw3),
            (vo => vo.Name, new[] { "SW4" }, BankType.Sw4),
            (vo => vo.Name, new[] { "SW5" }, BankType.Sw5),
            (vo => vo.SecName, new[] { "EFMO", "EFAP" }, BankType.Mon),
            (vo => vo.SecName, new[] { "EFUI" }, BankType.Uid),
        };

        public static string GetBankName(string name, string secName = null)
        {
            name = Regex.Replace(name, "bank_", "", RegexOptions.IgnoreCase).ToUpper();
            ValidateObject validateObject = new ValidateObject { Name = name, SecName = secName };
            foreach (
                (
                    Func<ValidateObject, string> Selector,
                    string[] Keywords,
                    string BankType
                ) rule in _bankNameRules
            )
            {

                string target = rule.Selector(validateObject);

                if (!string.IsNullOrEmpty(target) && rule.Keywords.Any(target.Contains))
                {
                    return rule.BankType;
                }
            }
            return BankType.Unknow;
        }

        public static string ConvertToExtraName(EfuseExtraType extraType)
        {
            if (extraType == EfuseExtraType.Early)
            {
                return Early;
            }

            if (extraType == EfuseExtraType.Deid)
            {
                return Deid;
            }

            if (extraType == EfuseExtraType.NonDeid)
            {
                return NonDeid;
            }

            return "";
        }

        public static string ConvertToFullBankName(string bankName)
        {
            if (string.IsNullOrEmpty(bankName))
            {
                return bankName;
            }

            string fullBankName = bankName;
            switch (bankName)
            {
                case BankType.Cfg:
                    fullBankName = "Config";
                    break;
                case BankType.UdrE:
                    fullBankName = "UDRE";
                    break;
                case BankType.UdrE0:
                    fullBankName = "UDRE0";
                    break;
                case BankType.UdrE1:
                    fullBankName = "UDRE1";
                    break;
                case BankType.UdrP:
                    fullBankName = "UDRP";
                    break;
                case BankType.UdrP0:
                    fullBankName = "UDRP0";
                    break;
                case BankType.UdrP1:
                    fullBankName = "UDRP1";
                    break;
                case BankType.Mon:
                    fullBankName = "Monitor";
                    break;
            }
            return fullBankName;
        }

        public static string RemoveHlnAndMrName(string testName)
        {
            if (string.IsNullOrEmpty(testName))
            {
                return testName;
            }

            return testName.Replace("_MR0", "").Replace("_MR1", "").Replace("_HV", "").Replace("_LV", "").Replace("_NV", "");
        }
    }
}
