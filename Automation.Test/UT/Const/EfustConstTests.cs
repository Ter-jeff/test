using Automation.Const;
using Automation.GenerateIgxl.EFuse.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Const
{
    [TestClass]
    public class EfustConstTests
    {
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
