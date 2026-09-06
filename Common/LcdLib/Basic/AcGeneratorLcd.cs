using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenAc.AcGenerator.Business;
using Automation.GenerateIgxl.Basic.Business.GenAc.AcInput.BassData;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.Singleton;
using Automation.Utility.PatternListManager;

using CommonLib.Extension;

using TestPlanLib.Basic;

namespace LcdLib.Basic
{
    [ExcludeFromCodeCoverage]
    public class AcGeneratorLcd : AcGenerator
    {
        #region Field
        public const string Bscan = "BScan";
        #endregion

        #region Constructor
        public AcGeneratorLcd(AcInputSheet acInputSheet)
            : base(acInputSheet)
        {
            AcInputSheet = acInputSheet;
        }
        #endregion

        #region Member Function

        protected override List<string> InitialAcCatList()
        {
            List<string> categoryList =
            [
                Common,
                Scan,
                Mbist,
                Bscan,
                Jtag
            ];
            return categoryList;
        }

        protected static string GetTimesetcategory(string sheetName)
        {
            string blockName = "TBD";
            //ex: TIMESET_OSPA0_A_AN_SI_1   
            string[] toks = sheetName.Split(['_'], StringSplitOptions.RemoveEmptyEntries);
            //Scan/Mbist/Other/HardIp/Common
            string subBlock = HardIpConstData.AcCommonDefault;
            if (toks.Length >= 5)
            {
                //Sub block part
                if (toks[3] == "SC")
                {
                    subBlock = Scan;
                }
                else if (toks[3] == "BI")
                {
                    subBlock = Mbist;
                }
                else if (toks[3] == "JT")
                {
                    subBlock = Jtag;
                }
                else if (toks[3] == "IO")
                {
                    subBlock = Bscan;
                }

                blockName = subBlock;
            }
            return blockName;
        }

        protected override BlockType GetTimesetBlockType(string sheetName)
        {
            BlockType blockType = BlockType.HardIp;
            string[] toks = sheetName.Split(['_'], StringSplitOptions.RemoveEmptyEntries);
            if (toks.Length >= 5)
            {
                //Sub block part
                if (toks[3] == "SC")
                {
                    blockType = BlockType.Scan;
                }
                else if (toks[3] == "BI")
                {
                    blockType = BlockType.Mbist;
                }
                else if (toks[3] == "JT")
                {
                    blockType = BlockType.Common;
                }
                else if (toks[3] == "IO")
                {
                    blockType = BlockType.BScan;
                }
            }
            return blockType;
        }

        /*
         * Override to gen OTP Ac spec
         */
        protected override void GenEfuseACspec(List<PatternData> patternDatas)
        {
            int otpIndex = 0;
            var lHardIp2OtpDict = new Dictionary<string, string>();
            //Collect all timeSetSheets that are used by OTP pattern
            List<string> lOTpTSetSheets = GetOtpPatternTimeSet(patternDatas);
            foreach (string otpTSet in lOTpTSetSheets) //all timeSetSheets that are used by efuse pattern
            {
                string otp = "Otp";
                if (!AcTSetCategoryMapSingleton.Instance().Contains(otpTSet))
                {
                    // error 
                    continue;
                }
                string hardipCate = AcTSetCategoryMapSingleton.Instance().GetCategory(otpTSet);
                if (!lHardIp2OtpDict.TryGetValue(hardipCate, out string? value))
                {
                    otp = otpIndex == 0 ? otp : otp + "_" + otpIndex;
                    CopyAcCategory(hardipCate, otp);
                    lHardIp2OtpDict.Add(hardipCate, otp);
                    //!! add otp timeset to category!!
                    AcTSetCategoryMapSingleton.Instance().SetRow(otpTSet, BlockType.Otp, otp);
                    otpIndex++;
                }
                else // already create an efuse category
                {
                    //!! use the exist otp cat
                    AcTSetCategoryMapSingleton.Instance().SetRow(otpTSet, BlockType.Otp, value);
                }
            }
        }
        #endregion

        public static List<string> GetOtpPatternTimeSet(List<PatternData> patternDatas)
        {
            var otpTimeSetSheetList = new List<string>();
            PatternListEfuseTimeSet.ReadCfg();
            foreach (PatternData pattern in patternDatas)
            {
                if (JudgeOtpPattern(pattern.PatternName))
                {
                    if (pattern.Use.EqualsIgnoreCase("USE"))
                    {
                        if (pattern.FileVersion.EqualsIgnoreCase("NA") ||
                            pattern.TimeSetVersion.EqualsIgnoreCase("NA"))
                        {
                            continue;
                        }
                        if (!otpTimeSetSheetList.Contains(pattern.TimeSetVersion))
                        {
                            otpTimeSetSheetList.Add(pattern.TimeSetVersion);
                        }
                    }
                }
            }

            return otpTimeSetSheetList;
        }

        private static bool JudgeOtpPattern(string pattern)
        {
            const string otp = "OTP";
            string subName = pattern.Split('_').Last();

            if (subName.ContainsIgnoreCase(otp))
            {
                return true;
            }

            return false;
        }
    }
}
