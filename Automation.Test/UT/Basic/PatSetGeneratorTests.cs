using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenPatSet.Business;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Basic;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class PatSetGeneratorTests : FunctionTestBase
    {
        private PatSetGenerator _generator = null!;

        [TestInitialize]
        public void Setup()
        {
            LocalSpecs.Options.IsIgnorePatternCheck = false;
            LocalSpecs.Options.InstrumentType = EnumInstrument.UFlex;
            LocalSpecs.Options.SearchValidPatternRev = true;

            _generator = new PatSetGenerator(Path.Combine("C:\\", "Pattern"));
        }

        [TestMethod]
        public void GenerateFlow_Should_GenerateSimplePatSet()
        {
            // Arrange
            var patternList = new List<PatternData>
            {
                new() { PatternName = "PAT1", FileVersion = "PAT1_V1", Use = "USE" }
            };
            var fileList = new Dictionary<string, string>
            {
                { "PAT1_V1.PAT.GZ", Path.Combine("C:\\Pattern", "PAT1_V1.PAT.GZ") }
            };

            // Act
            _generator.GenerateFlow(patternList, [], fileList);

            // Assert
            Assert.AreNotEqual(null, _generator.PatSetSheetAll);
            Assert.AreEqual(1, _generator.PatSetSheetAll.Rows.Count);
            Assert.IsTrue(_generator.PatSetSheetAll.Rows[0].PatSetRows[0].File.Contains("PAT1_V1"));
        }

        [TestMethod]
        public void GenerateFlow_Should_HandleSubroutineReference()
        {
            // Arrange
            var patternList = new List<PatternData>
            {
                new() { PatternName = "PAT_SUBR", FileVersion = "PAT_SUBR_V1", Use = "USE" }
            };

            var hardIpReferences = new List<HardIpInfo>
            {
                new()
                {
                    Payload = "PAT_SUBR",
                    Subroutine =   "SUB_A,SUB_B" ,
                    VmVector = "VM_1"
                }
            };

            var fileList = new Dictionary<string, string>
            {
                { "PAT_SUBR_V1", Path.Combine("C:\\Pattern", "PAT_SUBR_V1.PAT.GZ") }
            };

            // Act
            _generator.GenerateFlow(patternList, hardIpReferences, fileList);

            // Assert
            Assert.AreEqual(1, _generator.PatSetSheetAll.Rows.Count);
            Assert.AreEqual("PAT_SUBR", _generator.PatSetSheetAll.Rows[0].PatSetName);
        }

        [TestMethod]
        public void GenerateFlow_Should_ReportMissingPatternFile()
        {
            // Arrange
            ErrorReportManager.ClearErrors();

            var patternList = new List<PatternData>
            {
                new() { PatternName = "PAT_MISS", FileVersion = "PAT_MISS_V1", Use = "USE" }
            };

            var fileList = new Dictionary<string, string>();

            // Act
            _generator.GenerateFlow(patternList, [], fileList);

            // Assert
            List<Error> errors = ErrorReportManager.GetErrorList();
            Assert.IsTrue(errors.Any(e => e.Message.Contains("doesn't exist in pattern folder")));
        }

        [TestMethod]
        public void GenerateFlow_Should_AcceptDebugUseType()
        {
            // Arrange
            var patternList = new List<PatternData>
            {
                new() { PatternName = "PAT_DEBUG", FileVersion = "PAT_DEBUG_V1", Use = "DEBUG" }
            };
            var fileList = new Dictionary<string, string>
            {
                { "PAT_DEBUG_V1.PAT.GZ", Path.Combine("C:\\Pattern", "PAT_DEBUG_V1.PAT.GZ") }
            };

            // Act
            _generator.GenerateFlow(patternList, [], fileList);

            // Assert
            Assert.AreEqual(1, _generator.PatSetSheetAll.Rows.Count);
            Assert.IsTrue(_generator.PatSetSheetAll.Rows[0].PatSetName.Contains("PAT_DEBUG"));
        }
    }
}
