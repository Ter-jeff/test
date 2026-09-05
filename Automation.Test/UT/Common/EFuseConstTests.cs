using Automation.Const;
using Automation.GenerateIgxl.EFuse.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class EFuseConstTests
    {
        [DataTestMethod]
        [DataRow("UDR_E", true, DisplayName = "01_ShouldReturnTrue_ForUDR_E")]
        [DataRow("CMP_P", true, DisplayName = "02_ShouldReturnTrue_ForCMP_P")]
        [DataRow("UNKNOWN", false, DisplayName = "03_ShouldReturnFalse_ForUnknownBank")]
        public void BankIsUdr_ShouldReturnExpectedResult(string name, bool expected)
        {
            bool result = EFuseConst.BankIsUdr(name);
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("bank_cfg", "CFG", DisplayName = "01_GetBankName_CFG")]
        [DataRow("bank_config", "CFG", DisplayName = "02_GetBankName_CONFIG")]
        [DataRow("bank_udr_m0", "UDR_M0", DisplayName = "03_GetBankName_UDR_M0")]
        [DataRow("bank_udrm0", "UDR_M0", DisplayName = "04_GetBankName_UDRM0")]
        [DataRow("bank_udr_m1", "UDR_M1", DisplayName = "05_GetBankName_UDR_M1")]
        [DataRow("bank_udrm1", "UDR_M1", DisplayName = "06_GetBankName_UDRM1")]
        [DataRow("bank_udr_m", "UDR_M", DisplayName = "07_GetBankName_UDR_M")]
        [DataRow("bank_udrm", "UDR_M", DisplayName = "08_GetBankName_UDRM")]
        [DataRow("bank_udr_e", "UDR_E", DisplayName = "09_GetBankName_UDR_E")]
        [DataRow("bank_udre", "UDR_E", DisplayName = "10_GetBankName_UDRE")]
        [DataRow("bank_udr_p0", "UDR_P0", DisplayName = "11_GetBankName_UDR_P0")]
        [DataRow("bank_udrp0", "UDR_P0", DisplayName = "12_GetBankName_UDRP0")]
        [DataRow("bank_udr_p1", "UDR_P1", DisplayName = "13_GetBankName_UDR_P1")]
        [DataRow("bank_udrp1", "UDR_P1", DisplayName = "14_GetBankName_UDRP1")]
        [DataRow("bank_udr_p", "UDR_P", DisplayName = "15_GetBankName_UDR_P")]
        [DataRow("bank_udrp", "UDR_P", DisplayName = "16_GetBankName_UDRP")]
        [DataRow("bank_ufa", "UDR_E", DisplayName = "17_GetBankName_UFA_ReturnsUDR_E")]
        [DataRow("bank_uso", "UDR_E", DisplayName = "18_GetBankName_USO_ReturnsUDR_E")]
        [DataRow("bank_usi", "UDR_E", DisplayName = "19_GetBankName_USI_ReturnsUDR_E")]
        [DataRow("bank_smr", "UDR_E", DisplayName = "20_GetBankName_SMR_ReturnsUDR_E")]
        [DataRow("bank_ecid", "ECID", DisplayName = "21_GetBankName_ECID")]
        [DataRow("bank_ecd", "ECID", DisplayName = "22_GetBankName_ECD")]
        [DataRow("bank_monitor", "MON", DisplayName = "23_GetBankName_MONITOR")]
        [DataRow("bank_mon", "MON", DisplayName = "24_GetBankName_MON")]
        [DataRow("bank_sef", "MON", DisplayName = "25_GetBankName_SEF_ReturnsMON")]
        [DataRow("bank_cmp_e", "CMP_E", DisplayName = "26_GetBankName_CMP_E")]
        [DataRow("bank_cmp_p", "CMP_P", DisplayName = "27_GetBankName_CMP_P")]
        [DataRow("bank_sw0", "SW0", DisplayName = "28_GetBankName_SW0")]
        [DataRow("bank_sw1", "SW1", DisplayName = "29_GetBankName_SW1")]
        [DataRow("bank_sw2", "SW2", DisplayName = "30_GetBankName_SW2")]
        [DataRow("bank_sw3", "SW3", DisplayName = "31_GetBankName_SW3")]
        [DataRow("bank_sw4", "SW4", DisplayName = "32_GetBankName_SW4")]
        [DataRow("bank_sw5", "SW5", DisplayName = "33_GetBankName_SW5")]
        [DataRow("bank_xxx", "MON", "EFMO", DisplayName = "34_GetBankName_SecNameEFMO_ReturnsMON")]
        [DataRow("bank_xxx", "MON", "EFAP", DisplayName = "35_GetBankName_SecNameEFAP_ReturnsMON")]
        [DataRow("bank_xxx", "UID", "EFUI", DisplayName = "36_GetBankName_SecNameEFUI_ReturnsUID")]
        [DataRow("bank_xxx", "Unknow", "XXX", DisplayName = "37_GetBankName_Unknown_Default")]
        public void GetBankName_ShouldReturnExpectedBank(string name, string expected, string secName = "")
        {
            string result = EFuseConst.GetBankName(name, secName);
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow(EfuseExtraType.Early, EFuseConst.Early, DisplayName = "01_ConvertToExtraName_Early")]
        [DataRow(EfuseExtraType.Deid, EFuseConst.Deid, DisplayName = "02_ConvertToExtraName_Deid")]
        [DataRow(EfuseExtraType.NonDeid, EFuseConst.NonDeid, DisplayName = "03_ConvertToExtraName_NonDeid")]
        [DataRow((EfuseExtraType)999, "", DisplayName = "04_ConvertToExtraName_Unknown")]
        public void ConvertToExtraName_ShouldMapEnumToString(EfuseExtraType efuseExtraType, string expected)
        {
            string result = EFuseConst.ConvertToExtraName(efuseExtraType);
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow(BankType.Cfg, "Config", DisplayName = "01_ConvertToFullBankName_CFG")]
        [DataRow(BankType.UdrE, "UDRE", DisplayName = "02_ConvertToFullBankName_UDR_E")]
        [DataRow(BankType.UdrE0, "UDRE0", DisplayName = "03_ConvertToFullBankName_UDR_E0")]
        [DataRow(BankType.UdrE1, "UDRE1", DisplayName = "04_ConvertToFullBankName_UDR_E1")]
        [DataRow(BankType.UdrP, "UDRP", DisplayName = "05_ConvertToFullBankName_UDR_P")]
        [DataRow(BankType.UdrP0, "UDRP0", DisplayName = "06_ConvertToFullBankName_UDR_P0")]
        [DataRow(BankType.UdrP1, "UDRP1", DisplayName = "07_ConvertToFullBankName_UDR_P1")]
        [DataRow(BankType.Mon, "Monitor", DisplayName = "08_ConvertToFullBankName_MON")]
        [DataRow(BankType.Ecid, BankType.Ecid, DisplayName = "09_ConvertToFullBankName_ECID")]
        [DataRow("", "", DisplayName = "10_ConvertToFullBankName_Empty")]
        [DataRow(null, null, DisplayName = "11_ConvertToFullBankName_Null")]
        public void ConvertToFullBankName_ShouldReturnExpectedName(string input, string expected)
        {
            string result = EFuseConst.ConvertToFullBankName(input);
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("Fuse_MR0_HV", "Fuse", DisplayName = "01_RemoveHlnAndMrName_MR0HV")]
        [DataRow("Fuse_MR1_LV", "Fuse", DisplayName = "02_RemoveHlnAndMrName_MR1LV")]
        [DataRow("Fuse_NV", "Fuse", DisplayName = "03_RemoveHlnAndMrName_NV")]
        [DataRow("Fuse", "Fuse", DisplayName = "04_RemoveHlnAndMrName_NoSuffix")]
        [DataRow("", "", DisplayName = "05_RemoveHlnAndMrName_Empty")]
        [DataRow(null, null, DisplayName = "06_RemoveHlnAndMrName_Null")]
        public void RemoveHlnAndMrName_ShouldStripSuffixes(string input, string expected)
        {
            string result = EFuseConst.RemoveHlnAndMrName(input);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void BankIsUdrTest()
        {
            string[] types = [BankType.CmpE, BankType.UdrE, BankType.UdrP, BankType.UdrP0, BankType.UdrP1, BankType.CmpP, BankType.Unknow];
            foreach (string type in types)
            {
                bool result = EFuseConst.BankIsUdr(type);
                if (type == BankType.UdrE || type == BankType.UdrP || type == BankType.UdrP0 ||
                    type == BankType.UdrP1 || type == BankType.CmpE || type == BankType.CmpP)
                {
                    Assert.IsTrue(result);
                }
                else
                {
                    Assert.IsFalse(result);
                }
            }
        }

        [TestMethod]
        public void GetBankNameTest()
        {
            string[] types =
            [
                "config",
                "cfg",
                "udrm0",
                "udr_m0",
                "udrm",
                "udr_m",
                "udrm1",
                "udr_m1",
                "udre",
                "udr_e",
                "udrp",
                "udr_p",
                "udrp0",
                "udr_p0",
                "udrp1",
                "udr_p1",
                "ufa",
                "uso",
                "usi",
                "smr",
                "ecid",
                "ecd",
                "monitor",
                "mon",
                "sef",
                "cmp_e",
                "cmp_p",
                "sw0",
                "sw1",
                "sw2",
                "sw3",
                "sw4",
                "sw5"
            ];

            string[] secondName = ["EFMO", "EFAP", "EFUI"];

            foreach (string type in types)
            {
                Assert.AreNotEqual(BankType.Unknow, EFuseConst.GetBankName("bank_" + type));
            }

            foreach (string name in secondName)
            {
                string result = EFuseConst.GetBankName("bank_", name);
                Assert.AreEqual(name == "EFUI" ? BankType.Uid : BankType.Mon, result);
            }

            Assert.AreEqual(BankType.Unknow, EFuseConst.GetBankName("bank_"));
        }

        [DataRow(EfuseExtraType.Early, EFuseConst.Early)]
        [DataRow(EfuseExtraType.Deid, EFuseConst.Deid)]
        [DataRow(EfuseExtraType.NonDeid, EFuseConst.NonDeid)]
        [DataRow(EfuseExtraType.Normal, "")]
        [TestMethod]
        public void ConvertToExtraNameTest(EfuseExtraType efuseExtraType, string expected)
        {
            Assert.AreEqual(expected, EFuseConst.ConvertToExtraName(efuseExtraType));
        }

        [DataRow("", "")]
        [DataRow(BankType.Cfg, "Config")]
        [DataRow(BankType.UdrE, "UDRE")]
        [DataRow(BankType.UdrE0, "UDRE0")]
        [DataRow(BankType.UdrE1, "UDRE1")]
        [DataRow(BankType.UdrP, "UDRP")]
        [DataRow(BankType.UdrP1, "UDRP1")]
        [DataRow(BankType.UdrP0, "UDRP0")]
        [DataRow(BankType.Mon, "Monitor")]
        [DataRow("123", "123")]
        [TestMethod]
        public void ConvertToFullBankNameTest(string bankName, string expected)
        {
            Assert.AreEqual(expected, EFuseConst.ConvertToFullBankName(bankName));
        }

        [DataRow("", "")]
        [DataRow("AA_MR0", "AA")]
        [DataRow("EE_MR1", "EE")]
        [DataRow("BB_HV", "BB")]
        [DataRow("CC_LV", "CC")]
        [DataRow("DD_NV", "DD")]
        [TestMethod]
        public void RemoveHlnAndMrNameTest(string input, string expected)
        {
            Assert.AreEqual(expected, EFuseConst.RemoveHlnAndMrName(input));
        }
    }
}
