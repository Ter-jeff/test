using System.Collections.Generic;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpPatternTests : FunctionTestBase
    {
        private static HardIpPattern CreatePattern(bool isDefaultNoBinout, bool hasIgnoreKey, bool hasEnableKey)
        {
            HardIpPattern pat = new HardIpPattern
            {
                Pattern = new PatternClass("dummy")
                {
                    IsDefaultNobinout = isDefaultNoBinout
                },

                MiscInfoDict = []
            };

            if (hasIgnoreKey)
            {
                pat.MiscInfoDict.Add(HardIpConstData.IgnorePatBinOut, "");
            }

            if (hasEnableKey)
            {
                pat.MiscInfoDict.Add(HardIpConstData.EnablePattBinout, "");
            }

            return pat;
        }

        [TestMethod]
        public void Should_ReturnTrue_When_IgnoreKeyExists()
        {
            HardIpPattern pat = CreatePattern(isDefaultNoBinout: false, hasIgnoreKey: true, hasEnableKey: false);

            Assert.IsTrue(pat.IsIgnorePatBinOut());
        }

        [TestMethod]
        public void Should_ReturnTrue_When_DefaultNoBinout_And_NoEnable()
        {
            HardIpPattern pat = CreatePattern(isDefaultNoBinout: true, hasIgnoreKey: false, hasEnableKey: false);

            Assert.IsTrue(pat.IsIgnorePatBinOut());
        }

        [TestMethod]
        public void Should_ReturnFalse_When_NotDefault_And_NoIgnore()
        {
            HardIpPattern pat = CreatePattern(isDefaultNoBinout: false, hasIgnoreKey: false, hasEnableKey: false);

            Assert.IsFalse(pat.IsIgnorePatBinOut());
        }

        [TestMethod]
        public void Should_ReturnFalse_When_DefaultButEnableExists()
        {
            HardIpPattern pat = CreatePattern(isDefaultNoBinout: true, hasIgnoreKey: false, hasEnableKey: true);

            Assert.IsFalse(pat.IsIgnorePatBinOut());
        }

        [TestMethod]
        public void Should_ReturnTrue_When_BothIgnoreAndEnableExist()
        {
            HardIpPattern pat = CreatePattern(isDefaultNoBinout: false, hasIgnoreKey: true, hasEnableKey: true);

            Assert.IsTrue(pat.IsIgnorePatBinOut());
        }

        [TestMethod]
        public void GetFreq_ShouldConvertMHz()
        {
            string result = HardIpPattern.GetFreq("100MHz");
            Assert.AreEqual("100000000", result);
        }

        [TestMethod]
        public void GetFreq_ShouldHandleDecimal()
        {
            string result = HardIpPattern.GetFreq("1.5MHz");

            Assert.AreEqual("1500000", result);
        }

        [TestMethod]
        public void GetFreq_ShouldHandleLowercaseUnit()
        {
            string result = HardIpPattern.GetFreq("10khz");

            Assert.AreEqual("10000", result);
        }

        [TestMethod]
        public void GetFreq_ShouldHandleInvalidInput()
        {
            string result = HardIpPattern.GetFreq("XYZ");

            Assert.AreNotEqual(null, result);
        }

        [TestMethod]
        public void GetFreq_ShouldHandleGHz()
        {
            string result = HardIpPattern.GetFreq("1GHz");

            Assert.AreEqual("1000000000", result);
        }

        [TestMethod]
        public void GetDigCapName_ShouldExtractName()
        {
            var pat = new HardIpPattern
            {
                MiscInfo = "MeasCapName:MyCap"
            };

            string result = pat.GetDigCapNameByMiscInfo();

            Assert.AreEqual("MyCap", result);
        }

        [TestMethod]
        public void GetDigCapName_ShouldReturnEmpty_WhenNotFound()
        {
            var pat = new HardIpPattern
            {
                MiscInfo = "SomethingElse:123"
            };

            Assert.AreEqual("", pat.GetDigCapNameByMiscInfo());
        }

        [TestMethod]
        public void CapBitsInTp_ShouldSumBits()
        {
            var pat = new HardIpPattern
            {
                MeasPins =
                [
                    new() { MeasType = MeasType.MeasC, CapBit = "2" },
                    new() { MeasType = MeasType.MeasC, CapBit = "3" }
                ]
            };

            int result = pat.CapBitsInTp();

            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void CapBitsInTp_ShouldIgnoreNonMeasC()
        {
            var pat = new HardIpPattern
            {
                MeasPins =
                [
                    new() { MeasType = MeasType.MeasI, CapBit = "10" }
                ]
            };

            Assert.AreEqual(0, pat.CapBitsInTp());
        }

        [TestMethod]
        public void CapBitsInTp_ShouldSumOnlyValidMeasC()
        {
            HardIpPattern pat = new HardIpPattern();

            pat.MeasPins.Add(new MeasPin { MeasType = "MeasC", CapBit = "2" });
            pat.MeasPins.Add(new MeasPin { MeasType = "MeasV", CapBit = "100" });
            pat.MeasPins.Add(new MeasPin { MeasType = "MeasC", CapBit = "" });

            int result = pat.CapBitsInTp();

            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void PatternIndexFlag_ShouldReturnEmpty_WhenAllZero()
        {
            HardIpPattern pat = new HardIpPattern
            {
                DupIndex = 0,
                ConditionIndex = 0
            };

            Assert.AreEqual("", pat.PatternIndexFlag);
        }

        [TestMethod]
        public void PatternIndexFlag_ShouldReturnDupOnly()
        {
            HardIpPattern pat = new HardIpPattern
            {
                DupIndex = 1
            };

            Assert.AreEqual("_1", pat.PatternIndexFlag);
        }

        [TestMethod]
        public void PatternIndexFlag_ShouldReturnCondOnly()
        {
            HardIpPattern pat = new HardIpPattern
            {
                ConditionIndex = 2
            };

            Assert.AreEqual("_2", pat.PatternIndexFlag);
        }

        [TestMethod]
        public void PatternIndexFlag_ShouldReturnBoth()
        {
            HardIpPattern pat = new HardIpPattern
            {
                DupIndex = 1,
                ConditionIndex = 2
            };

            Assert.AreEqual("_1_2", pat.PatternIndexFlag);
        }

        [TestMethod]
        public void CopyCtor_ShouldDeepCopy()
        {
            HardIpPattern src = new HardIpPattern();
            src.MeasPins.Add(new MeasPin { MeasType = MeasType.MeasC, CapBit = "2" });
            src.SweepCodes.Add("A", []);
            src.SkipList.Add("X");

            HardIpPattern dst = new HardIpPattern(src);

            Assert.AreNotSame(src, dst);
            Assert.AreNotSame(src.MeasPins, dst.MeasPins);
            Assert.AreNotSame(src.SweepCodes, dst.SweepCodes);
            Assert.AreNotSame(src.SkipList, dst.SkipList);
        }

        [TestMethod]
        public void CopyCtor_ShouldDeepCopyPattern()
        {
            HardIpPattern src = new HardIpPattern
            {
                Pattern = new PatternClass("A+B")
            };

            HardIpPattern dst = new HardIpPattern(src);

            Assert.AreNotSame(src.Pattern, dst.Pattern);
            Assert.AreEqual(src.Pattern.RealPatternName, dst.Pattern.RealPatternName);
        }

        [TestMethod]
        public void CopyCtor_ShouldCopySweepCodes()
        {
            HardIpPattern src = new HardIpPattern();

            List<SweepCode> list = [new SweepCode { SweepInfo = "1,2,1" }];

            src.SweepCodes.Add("1", list);

            HardIpPattern dst = new HardIpPattern(src);

            Assert.AreNotSame(src.SweepCodes, dst.SweepCodes);
            Assert.AreNotSame(src.SweepCodes["1"], dst.SweepCodes["1"]);
        }

        [TestMethod]
        public void CopyCtor_ShouldDeepCopyMeasPin()
        {
            HardIpPattern src = new HardIpPattern();

            MeasPin pin = new MeasPin
            {
                CapBit = "2"
            };
            src.MeasPins.Add(pin);

            HardIpPattern dst = new HardIpPattern(src);

            Assert.AreNotSame(src.MeasPins, dst.MeasPins);
            Assert.AreNotSame(src.MeasPins[0], dst.MeasPins[0]);
        }

        [TestMethod]
        public void CopyMethod_ShouldCopyBasicFields()
        {
            HardIpPattern src = new HardIpPattern
            {
                SheetName = "A",
                SubBlock = "B"
            };

            HardIpPattern dst = new HardIpPattern();
            dst.Copy(src);

            Assert.AreEqual("A", dst.SheetName);
            Assert.AreEqual("B", dst.SubBlock);
        }

        [TestMethod]
        public void CopyMethod_ShouldNotShareReference()
        {
            HardIpPattern src = new HardIpPattern();
            src.MeasPins.Add(new MeasPin { CapBit = "1" });

            HardIpPattern dst = new HardIpPattern();
            dst.Copy(src);

            Assert.AreNotSame(src.MeasPins, dst.MeasPins);
        }

        [TestMethod]
        public void GetTimingsByAc_ShouldHandleEmpty()
        {
            HardIpPattern pat = new HardIpPattern
            {
                ForceCondition = new ForceClass
                {
                    ForceCondition = ""
                }
            };

            List<Timing> result = pat.GetTimingsByAc();

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetTimingsByAc_ShouldIgnoreInvalid()
        {
            HardIpPattern pat = new HardIpPattern
            {
                ForceCondition = new ForceClass
                {
                    ForceCondition = "AC:INVALID"
                }
            };

            List<Timing> result = pat.GetTimingsByAc();

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void ProcessSweepData_ShouldReturnOriginal_WhenNoSweep()
        {
            HardIpPattern pat = new HardIpPattern
            {
                RegisterAssignment = "A=B"
            };

            bool flag = false;

            string result = pat.ProcessSweepData(ref flag);

            Assert.AreEqual("A=B", result);
            Assert.IsFalse(flag);
        }

        [TestMethod]
        public void ProcessSweepData_ShouldHandleBracketSweep()
        {
            HardIpPattern pat = new HardIpPattern
            {
                RegisterAssignment = "REG=[1,2,3]"
            };

            bool flag = false;

            string result = pat.ProcessSweepData(ref flag);

            Assert.IsTrue(flag);
            Assert.IsTrue(result.Contains("sweep"));
        }

        [TestMethod]
        public void ProcessSweepData_ShouldHandleNestSweep()
        {
            HardIpPattern pat = new HardIpPattern
            {
                RegisterAssignment = "REG=nestsweep(1,2,3)"
            };

            bool flag = false;

            pat.ProcessSweepData(ref flag);

            Assert.IsTrue(flag);
        }

        [TestMethod]
        public void ProcessSweepData_ShouldHandleBurst()
        {
            HardIpPattern parent = new HardIpPattern
            {
                Pattern = new PatternClass("A,B")
            };

            HardIpPattern child = new HardIpPattern
            {
                RegisterAssignment = "A=B"
            };

            parent.BurstPatterns.Add(child);

            bool flag = false;

            string result = parent.ProcessSweepData(ref flag);

            Assert.IsTrue(result.Contains("A=B"));
        }

        [TestMethod]
        public void ProcessSweepData_ShouldSkipInvalidAssignment()
        {
            HardIpPattern pat = new HardIpPattern
            {
                RegisterAssignment = "INVALID_NO_EQUAL"
            };

            bool flag = false;

            string result = pat.ProcessSweepData(ref flag);

            Assert.IsTrue(string.IsNullOrEmpty(result) || result == "INVALID_NO_EQUAL");
            Assert.IsFalse(flag);
        }

        [TestMethod]
        public void ProcessSweepData_ShouldHandleMultipleAssignments()
        {
            HardIpPattern pat = new HardIpPattern
            {
                RegisterAssignment = "A=B;C=D"
            };

            bool flag = false;

            string result = pat.ProcessSweepData(ref flag);

            Assert.AreEqual("A=B;C=D", result);
        }

        [TestMethod]
        public void ProcessSweepData_ShouldHandleEmpty()
        {
            HardIpPattern pat = new HardIpPattern
            {
                RegisterAssignment = ""
            };

            bool flag = false;

            string result = pat.ProcessSweepData(ref flag);

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void ProcessSweepData_ShouldMergeSameSrcCodeIndex()
        {
            HardIpPattern pat = new HardIpPattern
            {
                RegisterAssignment = "REG1=[1,2];REG1=[3,4]"
            };

            bool flag = false;

            pat.ProcessSweepData(ref flag);

            Assert.IsTrue(flag);
            Assert.IsTrue(pat.SweepCodes.Count <= 1);
        }

        [TestMethod]
        public void ProcessSweepData_ShouldGenerateSweepFlag_ForBracketSweep()
        {
            HardIpPattern pat = new HardIpPattern
            {
                RegisterAssignment = "REG=[1,2,3]"
            };

            bool flag = false;

            string result = pat.ProcessSweepData(ref flag);

            Assert.IsTrue(flag);
            Assert.IsTrue(result.Contains('1'));
        }

        [TestMethod]
        public void ProcessSweepData_ShouldHandleBurstWithSweep()
        {
            HardIpPattern parent = new HardIpPattern
            {
                Pattern = new PatternClass("A,B")
            };

            HardIpPattern child = new HardIpPattern
            {
                RegisterAssignment = "REG=[1,2,3]"
            };

            parent.BurstPatterns.Add(child);

            bool flag = false;

            parent.ProcessSweepData(ref flag);

            Assert.IsTrue(child.RegisterAssignment.Contains("sweep"));
        }

        [TestMethod]
        public void MiscInfo_ShouldParseKeyOnly()
        {
            HardIpPattern pat = new HardIpPattern
            {
                MiscInfo = "Ignore_Patt_BinOut"
            };

            Assert.IsTrue(pat.MiscInfoDict.ContainsKey("Ignore_Patt_BinOut"));
        }

        [TestMethod]
        public void MiscInfo_ShouldParseKeyValue()
        {
            HardIpPattern pat = new HardIpPattern
            {
                MiscInfo = "AAA:123"
            };

            Assert.AreEqual("123", pat.MiscInfoDict["AAA"]);
        }

        [TestMethod]
        public void PatternClass_ShouldBeMultiple()
        {
            PatternClass p = new PatternClass("A,B");

            Assert.IsTrue(p.IsMultiple());
        }

        [TestMethod]
        public void PatternClass_ShouldGetLastPayload()
        {
            PatternClass p = new PatternClass("A+B");

            Assert.AreEqual("b", p.GetLastPayload());
        }

        [TestMethod]
        public void PatternClass_ShouldNotBeMultiple()
        {
            PatternClass p = new PatternClass("A");

            Assert.IsFalse(p.IsMultiple());
        }

        [TestMethod]
        public void ForceClass_ShouldExtractLevel()
        {
            ForceClass fc = new ForceClass
            {
                ForceCondition = "Level:LVL1"
            };

            Assert.AreEqual("LVL1", fc.GetLevelSetting());
        }

        [TestMethod]
        public void ForceClass_ShouldHandleEmptyAcSelector()
        {
            ForceClass fc = new ForceClass
            {
                ForceCondition = "DC:XXX"
            };

            Assert.AreEqual("", fc.GetAcSelector());
        }

        [TestMethod]
        public void SweepVData_ShouldReturnPlusOperand()
        {
            SweepVData data = new SweepVData("PIN,1,2,1");

            Assert.AreEqual("+", data.Operand);
        }

        [TestMethod]
        public void SweepVData_ShouldReturnLessComparator()
        {
            SweepVData data = new SweepVData("PIN,1,2,1");

            Assert.AreEqual("<", data.Comparator);
        }

        [TestMethod]
        public void SweepVData_ShouldFailCheckStep()
        {
            SweepVData data = new SweepVData("PIN,1,2,-1");

            Assert.IsFalse(data.CheckStep);
        }
    }
}
