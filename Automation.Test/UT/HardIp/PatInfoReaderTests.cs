using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class PatInfoReaderTests : FunctionTestBase
    {
        private PatInfoReader _reader = null!;

        [TestInitialize]
        public void Setup()
        {
            _reader = new PatInfoReader();

            _reader.PinGroupList.Clear();
            _reader.PinGroupList["GRP_A"] = ["PIN1", "PIN2"];
            _reader.PinGroupList["GRP_B"] = ["GRP_A", "PIN3"];
        }

        [TestMethod]
        public void DecompGroups_ShouldReturnFlattenedPinList()
        {
            // Act
            List<string> result = _reader.DecompGroups("GRP_B");

            // Assert
            CollectionAssert.AreEquivalent(new List<string> { "PIN1", "PIN2", "PIN3" }, result);
        }

        [TestMethod]
        public void ConvertToNetName_ShouldReturnPinName_WhenNoNetFound()
        {
            // Act
            string result = _reader.ConvertToNetName("PIN_A");

            // Assert
            Assert.AreEqual("PIN_A", result);
        }

        [DataTestMethod]
        [DataRow("v", "PIN_V", "VMEAS", "MeasVStr", DisplayName = "01_v")]
        [DataRow("i", "PIN_I", "IMEAS", "MeasIStr", DisplayName = "02_i")]
        [DataRow("f", "PIN_F", "FMEAS", "MeasFStr", DisplayName = "03_f")]
        [DataRow("vdiff", "PIN_A", "VMEAS_DIFF", "MeasVdiffStr", DisplayName = "04_vdiff")]
        [DataRow("vdiff2", "PIN_B", "VMEAS_DIFF2", "MeasVdiff2Str", DisplayName = "05_vdiff2")]
        [DataRow("idiff", "PIN_C", "IMEAS_DIFF", "MeasIdiffStr", DisplayName = "06_idiff")]
        [DataRow("fdiff", "PIN_D", "FMEAS_DIFF", "MeasFdiffStr", DisplayName = "07_fdiff")]
        [DataRow("vocm", "PIN_E", "VOCM", "MeasVocmStr", DisplayName = "08_vocm")]
        [DataRow("r1", "PIN_F", "R1MEAS", "MeasR1Str", DisplayName = "09_r1")]
        [DataRow("r2", "PIN_G", "R2MEAS", "MeasR2Str", DisplayName = "10_r2")]
        [DataRow("dutycycle", "PIN_H", "DUTY", "MeasDutyCycleStr", DisplayName = "11_dutycycle")]
        [DataRow("vdm", "PIN_I", "VDM", "MeasVdmStr", DisplayName = "12_vdm")]
        public void ConvertHardIpInfoNewToOld_ShouldHandleSpecialSeqsCorrectly(string seq, string pin, string measName, string expectedProperty)
        {
            // Arrange
            var oldInfo = new HardIpInfo();
            var newInfo = new HardIpInfoNew
            {
                MeasSeq = seq,
                MeasPin = pin,
                MeasName = measName,
                CalcStoreName = "CalcX"
            };

            // Act
            HardIpInfo result = _reader.ConvertHardIpInfoNewToOld(oldInfo, newInfo);

            // Assert
            Assert.AreEqual(seq, result.MeasSeqStr);
            Assert.AreEqual(measName, result.MeasName);

            var accessors = new Dictionary<string, Func<HardIpInfo, string>>
            {
                [nameof(HardIpInfo.MeasVStr)] = i => i.MeasVStr,
                [nameof(HardIpInfo.MeasIStr)] = i => i.MeasIStr,
                [nameof(HardIpInfo.MeasFStr)] = i => i.MeasFStr,
                [nameof(HardIpInfo.MeasVdiffStr)] = i => i.MeasVdiffStr,
                [nameof(HardIpInfo.MeasVdiff2Str)] = i => i.MeasVdiff2Str,
                [nameof(HardIpInfo.MeasIdiffStr)] = i => i.MeasIdiffStr,
                [nameof(HardIpInfo.MeasFdiffStr)] = i => i.MeasFdiffStr,
                [nameof(HardIpInfo.MeasVocmStr)] = i => i.MeasVocmStr,
                [nameof(HardIpInfo.MeasR1Str)] = i => i.MeasR1Str,
                [nameof(HardIpInfo.MeasR2Str)] = i => i.MeasR2Str,
                [nameof(HardIpInfo.MeasDutyCycleStr)] = i => i.MeasDutyCycleStr,
                [nameof(HardIpInfo.MeasVdmStr)] = i => i.MeasVdmStr
            };
            Assert.IsTrue(accessors.ContainsKey(expectedProperty), $"Property {expectedProperty} not found on HardIpReference.");

            string value = accessors[expectedProperty](result);
            Assert.IsFalse(string.IsNullOrEmpty(value), $"Expected {expectedProperty} to be updated.");
            Assert.IsTrue(value.Contains(pin), $"{expectedProperty} should contain '{pin}'.");

            Assert.IsTrue(result.MeaSet.ContainsKey(seq), $"MeaSet should contain key '{seq}'.");
            string mappedPin = result.MeaSet[seq][0].Keys.First();
            Assert.AreEqual(pin, mappedPin, $"MeaSet pin mismatch for '{seq}'.");
        }

        [TestMethod]
        public void PatInfoReader()
        {
            string subName = "PatInfoReader";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);
            _ = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "RF", "HardIPInfo.log");
            var patInfoReader = new PatInfoReader();
            List<HardIpInfo> result = patInfoReader.ExtractHardIpInfos(file);

            string json = JsonConvert.SerializeObject(result.GetRange(0, 20), Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void ExtractHardIpInfo_Should_ParseSimpleHardIpReference()
        {
            // Arrange
            var lines = new List<string>
            {
                "======",
                "test_inst: MyPayload",
                "tset: TSET_MAIN",
                "measseq: V,I,F",
                "measname: VMEAS|IMEAS|FMEAS",
                "f: PIN_A",
                "v: PIN_B",
                "i: PIN_C",
                "======"
            };

            // Act
            List<HardIpInfo> result = _reader.ExtractHardIpInfos(lines);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(2, result.Count);

            HardIpInfo refItem = result[0];
            Assert.AreEqual("mypayload", refItem.Payload);
            Assert.AreEqual("tset_main", refItem.TimeSet);
            Assert.AreEqual("V,I,F".ToUpper(), refItem.MeasSeqStr);
            Assert.AreEqual("VMEAS|IMEAS|FMEAS".ToUpper(), refItem.MeasName);

            Assert.IsTrue(refItem.MeaSet.ContainsKey("f"));
            Assert.AreEqual("PIN_A", refItem.MeaSet["f"][0].Keys.First());
        }

        [TestMethod]
        public void ExtractHardIpInfo_Should_HandleNewSeqFields()
        {
            // Arrange
            var lines = new List<string>
            {
                "======",
                "test_inst: pattern1",
                "newseq: V|I",
                "newseqmeaspin: PIN_X|PIN_Y",
                "newseqmeasname: VOUT|IOUT",
                "newseqcalcstorename: Calc_01",
                "======"
            };

            // Act
            List<HardIpInfo> result = _reader.ExtractHardIpInfos(lines);

            // Assert
            Assert.AreEqual(2, result.Count);
            HardIpInfo refItem = result[0];
            Assert.AreNotEqual(null, refItem.NewInfo);
            Assert.AreEqual("V|I", refItem.NewInfo.MeasSeq);
            Assert.AreEqual("PIN_X|PIN_Y", refItem.NewInfo.MeasPin);
            Assert.AreEqual("VOUT|IOUT", refItem.NewInfo.MeasName);
            Assert.AreEqual("Calc_01", refItem.NewInfo.CalcStoreName);
        }

        [TestMethod]
        public void ExtractHardIpInfo_Should_IgnoreDuplicatePatterns()
        {
            // Arrange
            var lines = new List<string>
            {
                "======",
                "test_inst: PATTERN_DUP",
                "tset: TSET_A",
                "======",
                "======",
                "test_inst: PATTERN_DUP",
                "tset: TSET_B",
                "======"
            };

            // Act
            List<HardIpInfo> result = _reader.ExtractHardIpInfos(lines);

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("pattern_dup", result[0].Payload);
            Assert.AreEqual("tset_a", result[0].TimeSet);
        }

        [TestMethod]
        public void ExtractHardIpInfo_Should_MapKeyToPropertyCorrectly()
        {
            // Arrange
            string subName = "ExtractHardIpInfos";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var lines = new List<string>
            {
                "======",
                "test_inst: PATTERN_TEST",
                "tset: TSET_A",
                "time_domain: AST",
                "test_inst:PAT1",
                "vm_vector:VEC123",
                "tset:TSET_X",
                "subr:SUBR_1",
                "call_subrs:CSUBR",
                "call_subrs_cnt:5",
                "call_subrs_cnt_diff:DIFF",
                "capbit:12",
                "capbitstr:CBSTR",
                "capbitname:CBNAME",
                "cappinname:CPIN",
                "sendbit:SBIT",
                "sendbitstr:SBSTR",
                "sendbitname:SBNAME",
                "sendpinname:SPIN",
                "digsrcsignalname:SIGNAL",
                "xpins:XP1",
                "unusediopins:UIPIN",
                "measseq:VIF",
                "measname:NAME1",
                "version:1.0=",
                "ssncorename:CORE_X=",
                "efuse_prog_time:100",
                "======"
            };

            // Act
            List<HardIpInfo> result = _reader.ExtractHardIpInfos(lines);

            // Assert
            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
