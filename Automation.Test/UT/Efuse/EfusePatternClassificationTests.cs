using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.Static;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Efuse.Input;

using PatternData = TestPlanLib.Basic.PatternData;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfusePatternClassificationTests : FunctionTestBase
    {
        [DataTestMethod]
        [DataRow("1_2_3_4_5_6_7_8_9_ME0", "", BankType.UdrE, DisplayName = "01_1_2_3_4_5_6_7_8_9_ME0_UDR_E")]
        [DataRow("1_2_3_4_5_6_7_8_9_MP0", "", BankType.UdrP, true, DisplayName = "02_1_2_3_4_5_6_7_8_9_MP0_UDR_P_True")]
        [DataRow("1_2_3_4_5_6_7_8_9_MPX", "a_b_c_d_e_f_g_h_i_j_P0", BankType.UdrP0, DisplayName = "03_1_2_3_4_5_6_7_8_9_MPX_P0_UDR_P0")]
        [DataRow("1_2_3_4_5_6_7_8_9_MPX", "a_b_c_d_e_f_g_h_i_j_P1", BankType.UdrP1, DisplayName = "04_1_2_3_4_5_6_7_8_9_MPX_P1_UDR_P1")]
        [DataRow("1_2_3_4_5_6_7_8_9_MPX", "a_b_c_d_e_f_g_h_i_MPX0", BankType.UdrP0, DisplayName = "05_1_2_3_4_5_6_7_8_9_MPX_MPX0_UDR_P0")]
        [DataRow("1_2_3_4_5_6_7_8_9_MPX", "a_b_c_d_e_f_g_h_i_MPX1", BankType.UdrP1, DisplayName = "06_1_2_3_4_5_6_7_8_9_MPX_MPX1_UDR_P1")]
        [DataRow("1_2_3_4_5_6_7_8_9_MP", "", BankType.UdrP, DisplayName = "07_1_2_3_4_5_6_7_8_9_MP_UDR_P")]
        [DataRow("1_2_3_4_5_6_7_8_9_MMX", "a_b_c_d_e_f_g_h_i_j_M0", BankType.UdrM0, DisplayName = "08_1_2_3_4_5_6_7_8_9_MMX_M0_UDR_M0")]
        [DataRow("1_2_3_4_5_6_7_8_9_MMX", "a_b_c_d_e_f_g_h_i_j_M1", BankType.UdrM1, DisplayName = "09_1_2_3_4_5_6_7_8_9_MMX_M1_UDR_M1")]
        [DataRow("1_2_3_4_5_6_7_8_9_MMX", "a_b_c_d_e_f_g_h_i_MMX0", BankType.UdrM0, DisplayName = "10_1_2_3_4_5_6_7_8_9_MMX_MMX0_UDR_M0")]
        [DataRow("1_2_3_4_5_6_7_8_9_MMX", "a_b_c_d_e_f_g_h_i_MMX1", BankType.UdrM1, DisplayName = "11_1_2_3_4_5_6_7_8_9_MMX_MMX1_UDR_M1")]
        [DataRow("1_2_3_4_5_6_7_8_9_MM", "", BankType.UdrM, DisplayName = "12_1_2_3_4_5_6_7_8_9_MM_UDR_M")]
        [DataRow("1_2_3_4_5", "", BankType.Unknow, DisplayName = "13_1_2_3_4_5_UNKNOW")]
        public void GetDetailBankNameTest(string payloadJoined, string init, string expected, bool onlyUdrP = false)
        {
            string[] payload = payloadJoined.Split('_');
            // Arrange
            List<string> initListAsList = string.IsNullOrEmpty(init) ? [] : [init];
            var sut = new EfusePatternClassification(onlyUdrP, []);

            // Act
            string result = sut.GetDetailBankName(initListAsList, payload);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("1_2_3_4_5_6_7_8_9_10_11_CRCMON", false, EfuseTestMode.Crc, false, DvRvType.Not, DisplayName = "01_CRCMON_CRC")]
        [DataRow("1_2_3_4_5_6_7_8_9_10_11_ZEROS", true, EfuseTestMode.Unknow, true, DvRvType.FlatCheck, DisplayName = "02_ZEROS_Unknow")]
        [DataRow("1_2_3_4_5_6_CPX_PRG_9_10_11_CP1B", true, EfuseTestMode.Unknow, false, DvRvType.FlatCheck, DisplayName = "03_CP1B_Unknow")]
        [DataRow("1_2_3_4_5_6_CPX_PRG_9_10_11_CP1EBC", true, EfuseTestMode.Unknow, true, DvRvType.FlatCheck, DisplayName = "04_CP1EBC_Unknow")]
        [DataRow("1_2_3_4_5_6_CPX_PRG_9_10_11_CPE_blank", true, EfuseTestMode.Unknow, true, DvRvType.FlatCheck, DisplayName = "05_CPE_blank_Unknow")]
        [DataRow("1_2_3_4_5_6_CPX_PRG_9_10_11_CPE_FLDSSC", true, EfuseTestMode.Unknow, false, DvRvType.Rv, DisplayName = "06_CPE_FLDSSC_Unknow")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_9_10_CPE_DV", true, EfuseTestMode.DvWrite, true, DvRvType.Dv, DisplayName = "07_CPE_DV_DVWrite")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_9_10_CPE_RV", true, EfuseTestMode.RvWrite, true, DvRvType.Rv, DisplayName = "08_CPE_RV_RVWrite")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_9_10_11_DVE", true, EfuseTestMode.DvWrite, true, DvRvType.Dv, DisplayName = "09_DVE_DVWrite")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_9_10_11_RVE", true, EfuseTestMode.RvWrite, true, DvRvType.Rv, DisplayName = "10_RVE_RVWrite")]
        [DataRow("1_2_3_4_5_6_7_8_9_10_11_12", true, EfuseTestMode.Unknow, false, DvRvType.Not, DisplayName = "11_Unknown_Unknow")]
        [DataRow("1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_FLDSSC1", false, EfuseTestMode.MarginRead, true, DvRvType.Not, DisplayName = "12_FLDSSC1_MarginRead")]
        [DataRow("1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_FLDSSC", false, EfuseTestMode.MarginReadFull, true, DvRvType.Not, DisplayName = "13_FLDSSC_MarginRead_Full")]
        [DataRow("1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_DAP", false, EfuseTestMode.ReadDap, true, DvRvType.Not, DisplayName = "14_DAP_Read_Dap")]
        [DataRow("1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_FLDSSC1", true, EfuseTestMode.RvRead, true, DvRvType.Rv, DisplayName = "15_FLDSSC1_RV_Read")]
        [DataRow("1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_FLDSSC", true, EfuseTestMode.RvRead, true, DvRvType.Rv, DisplayName = "16_FLDSSC_RV_Read")]
        [DataRow("1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_DAP", true, EfuseTestMode.ReadDap, true, DvRvType.Not, DisplayName = "17_DAP_Read_Dap")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_10_11_CPE_FLDSSC1", false, EfuseTestMode.WriteDssc, true, DvRvType.Not, DisplayName = "18_FLDSSC1_WriteDSSC")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_10_11_CPE_Q1DSSC", false, EfuseTestMode.WriteDsscQ, true, DvRvType.Not, DisplayName = "19_Q1DSSC_WriteDSSC_Q")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_10_11_CPE_H1DSSC", false, EfuseTestMode.WriteDsscH, true, DvRvType.Not, DisplayName = "20_H1DSSC_WriteDSSC_H")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_10_11_HF_FLDSSC1", false, EfuseTestMode.WriteDsscHf, false, DvRvType.Not, DisplayName = "21_HF_FLDSSC1_WriteDSSC_HF")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_10_11_CPE_FLDSSC1", true, EfuseTestMode.RvWrite, true, DvRvType.Rv, DisplayName = "22_FLDSSC1_RV_Write")]
        [DataRow("1_2_3_4_5_6_7_JTG_SNS_10_11_PART_ZERO", true, EfuseTestMode.Unknow, false, DvRvType.Not, DisplayName = "23_PART_ZERO_Unknow")]
        [DataRow("1_2_3_4_5_6_UFA_JTG_SNS_10_11_12", true, EfuseTestMode.Ufr, false, DvRvType.Not, DisplayName = "24_UFA_UFR")]
        [DataRow("1_2_3_4_5_6_7_JTG_SNS_10_11_12_DAP", true, EfuseTestMode.JtagReadDap, false, DvRvType.Not, DisplayName = "25_DAP_JTAGRead_Dap")]
        [DataRow("1_2_3_4_5_6_SEF_JTG_SNS_10_11_APB", true, EfuseTestMode.ApbRead, false, DvRvType.Not, DisplayName = "26_APB_APB_Read")]
        [DataRow("1_2_3_4_5_6_7_JTG_SNS_10_11_12", true, EfuseTestMode.JtagRead, false, DvRvType.Not, DisplayName = "27_JTAGRead_Unknow")]
        [DataRow("1_2_3_4_5_6_7_JTG_PRG_10_11_PART_ZERO", true, EfuseTestMode.Unknow, false, DvRvType.Not, DisplayName = "28_PART_ZERO_Unknow")]
        [DataRow("1_2_3_4_5_6_UFA_JTG_PRG_10_11_12", true, EfuseTestMode.Ufp, false, DvRvType.Not, DisplayName = "29_UFP_UFP")]
        [DataRow("1_2_3_4_5_6_SEF_JTG_PRG_10_11_APB", true, EfuseTestMode.ApbWrite, false, DvRvType.Not, DisplayName = "30_APB_APB_Write")]
        [DataRow("1_2_3_4_5_6_7_JTG_PRG_10_11_12", true, EfuseTestMode.JtagWrite, false, DvRvType.Not, DisplayName = "31_JTAGWrite_Unknow")]
        [DataRow("1_2_3_4_5_6_7_JTG_UNS_10_11_UDRVER2", true, EfuseTestMode.Ver2, false, DvRvType.Not, DisplayName = "32_UDRVER2_VER2")]
        [DataRow("1_2_3_4_5_6_7_JTG_UNS_10_11_UDRVER1", true, EfuseTestMode.Ver1, false, DvRvType.Not, DisplayName = "33_UDRVER1_VER1")]
        [DataRow("1_2_3_4_5_6_USO_JTG_UNS_10_11_12", true, EfuseTestMode.Uso, false, DvRvType.Not, DisplayName = "34_USO_USO")]
        [DataRow("1_2_3_4_5_6_USI_JTG_UNS_10_11_12", true, EfuseTestMode.Usi, false, DvRvType.Not, DisplayName = "35_USI_USI")]
        [DataRow("1_2_3_4_5_6_CPX_PRG_9_10_11_blank", true, EfuseTestMode.Unknow, false, DvRvType.FlatCheck, DisplayName = "36_blank_PatCnt12_Unknown_NotEarly")]
        [DataRow("1_2_3_4_5_6_7_JTG_UNS_10_11_UDRVER2_13", true, EfuseTestMode.Ver2, false, DvRvType.Not, DisplayName = "37_UDRVER2_VER2_PatCnt13")]
        [DataRow("1_2_3_4_5_6_7_JTG_UNS_10_11_UDRVER1_13", true, EfuseTestMode.Ver1, false, DvRvType.Not, DisplayName = "38_UDRVER1_VER1_PatCnt13")]
        [DataRow("1_2_3_4_5_6_7_DAA_PRG_9_10_CPE_ZEROS", true, EfuseTestMode.Unknow, true, DvRvType.FlatCheck, DisplayName = "39_DAA_PRG_CPE_ZEROS_Unknown_FlatCheck")]
        public void GetPatternTypeTest(string payload, bool isCfg, EfuseTestMode efuseTestMode, bool expectedIsEarly, DvRvType dvRvType)
        {

            LocalSpecs.Options.Device = EnumDevice.AP;
            string[] arr = payload.Split('_');
            var sut = new EfusePatternClassification(false, []);

            EfusePatternType result = sut.GetPatternType(arr, isCfg);

            Assert.AreEqual(efuseTestMode, result.TestMode, $"Payload: {payload}");
            Assert.AreEqual(expectedIsEarly, result.IsEarly, $"Payload: {payload}");
            Assert.AreEqual(dvRvType, result.DvrvType, $"Payload: {payload}");
        }

        [DataTestMethod]
        [DataRow("1_2_3_4_5_6_7_JTG_UNS_10_11_fullrd", true, EfuseTestMode.MarginReadFull, false, DvRvType.Not, DisplayName = "01_JTG_UNS_fullrd_MarginRead_Full")]
        [DataRow("1_2_3_4_5_6_7_JTG_UNS_10_11_fullwr", true, EfuseTestMode.WriteDssc, false, DvRvType.Not, DisplayName = "02_JTG_UNS_fullwr_WriteDSSC")]
        public void GetPatternTypeTest_RF(string payload, bool isCfg, EfuseTestMode efuseTestMode, bool expectedIsEarly, DvRvType dvRvType)
        {
            LocalSpecs.Options.Device = EnumDevice.RF;
            string[] arr = payload.Split('_');
            var sut = new EfusePatternClassification(false, []);

            EfusePatternType result = sut.GetPatternType(arr, isCfg);

            Assert.AreEqual(efuseTestMode, result.TestMode, $"Payload: {payload}");
            Assert.AreEqual(expectedIsEarly, result.IsEarly, $"Payload: {payload}");
            Assert.AreEqual(dvRvType, result.DvrvType, $"Payload: {payload}");
        }

        [TestMethod]
        public void ClassificationFromScgh_ShouldReturnRows_WhenValidDataProvided()
        {
            // Arrange
            var sut = new EfusePatternClassification(false, []);

            var scghData = new ScghData
            {
                HardIpSheetRowList =
                [
                    new("Sheet1")
                    {
                        InitList = ["INIT_PATTERN_3_4_5_6_8_9"],
                        PayloadList = ["PAYLOAD_PATTERN_3_4_5_6_8_9"],
                        RowNum = 1
                    }
                ]
            };

            var patternDic = new Dictionary<string, PatternData>
            {
                {
                    "DUMMY_UFA_DAA_PRG_XX_BANK_MP0_CP1",
                    new PatternData
                    {
                        IsExist = true,
                        PatternVersion = "DUMMY_UFA_DAA_PRG_XX_BANK_MP0_CP1_V1"
                    }
                }
            };

            var patInK = new HashSet<string>
            {
                Path.Combine("C:\\patterns", "DUMMY_UFA_DAA_PRG_XX_BANK_MP0_CP1_V1.atp.gz")
            };

            var efuseBitDefTables = new List<BitDefTable>
            {
                new()
                {
                    BlockName = "BANK_MP0",
                    AccessMode = "ReadWrite",
                    BitMode = "8"
                }
            };

            // Act
            List<EfusePatternRow> result = sut.ClassificationFromScgh(scghData, patternDic, patInK, efuseBitDefTables);

            // Assert
            Assert.AreEqual(1, result.Count);
            EfusePatternRow row = result.First();
            Assert.AreEqual("Sheet1", row.SheetName);
            Assert.AreEqual(1, row.RowNum);
            CollectionAssert.Contains(row.PatList, "INIT_PATTERN_3_4_5_6_8_9");
            CollectionAssert.Contains(row.PatList, "PAYLOAD_PATTERN_3_4_5_6_8_9");
            Assert.AreEqual(BankType.Unknow, row.BankName);
        }

        [TestMethod]
        public void ClassificationFromInstanceSheet_ShouldSkipRow_WhenNoInitOrPayload()
        {
            // Arrange
            var sut = new EfusePatternClassification(false, []);

            var instanceSheet = new BinCutInstanceSheet("")
            {
                Rows =
                [
                    new()
                    {
                        InitList = ["I1_2_3_4_5_6_7_8_9_10_11_12"],
                        PayloadList = ["1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_FLDSSC1"],
                        SheetName = "EmptySheet",
                        RowNum = 10
                    }
                ]
            };

            var patternDic = new Dictionary<string, PatternData>();
            var patInK = new HashSet<string>();
            var efuseBitDefTables = new List<BitDefTable>();

            // Act
            List<EfusePatternRow> result = sut.ClassificationFromInstanceSheet(instanceSheet, patternDic, patInK, efuseBitDefTables);

            // Assert
            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void GetDetailBankName_ShouldReturn_UdrP0_WhenInitContainsP0()
        {
            // Arrange
            var sut = new EfusePatternClassification(false, []);
            var initList = new List<string> { "DUMMY_INIT_P0" };
            string[] payload = ["A", "B", "C", "D", "E", "F", "G", "H", "BANK", "MP0"];

            // Act
            string bank = sut.GetDetailBankName(initList, payload);

            // Assert
            Assert.AreEqual(BankType.UdrP0, bank);
        }
    }
}
