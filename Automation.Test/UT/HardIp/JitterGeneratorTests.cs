using System.Collections.Generic;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.InputObject;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{

    [TestClass]
    public class JitterGeneratorTests
    {
        private JitterGenerator _generator = null!;
        private Dictionary<string, HardIpSheet> _planDic = null!;

        [TestInitialize]
        public void Setup()
        {
            _generator = new JitterGenerator();
            _planDic = [];
        }

        [TestMethod]
        public void GenJitterSheet_NoMeasD_ReturnsNull()
        {
            // Arrange: Add a row with a different MeasType
            var sheet = new HardIpSheet { Rows = [] };
            var pattern = new HardIpPattern();
            pattern.MeasPins.Add(new MeasPin { PinName = "Pin1", MeasType = "NotMeasD" });
            sheet.Rows.Add(pattern);
            _planDic.Add("Sheet1", sheet);

            // Act
            JitterSheet? result = _generator.GenJitterSheet(_planDic);

            // Assert
            Assert.AreEqual(null, result, "Should return null if no pins have MeasType.MeasD");
        }

        [TestMethod]
        public void GenJitterSheet_ValidMeasD_CreatesRowsWithDefaultValues()
        {
            // Arrange
            var sheet = new HardIpSheet { Rows = [] };
            var pattern = new HardIpPattern();
            pattern.MeasPins.Add(new MeasPin { PinName = "ClockPin", MeasType = MeasType.MeasD });
            sheet.Rows.Add(pattern);
            _planDic.Add("Sheet1", sheet);

            // Act
            JitterSheet? result = _generator.GenJitterSheet(_planDic, "MyCustomJitter");

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual("MyCustomJitter", result!.Name);
            Assert.AreEqual(1, result.Rows.Count);

            JitterRow row = result.Rows[0];
            Assert.AreEqual("ClockPin", row.PinOrGroup);
            Assert.AreEqual("Duty_Jitter", row.JitterSet);
            Assert.AreEqual("Disabled", row.Mode);
        }

        [TestMethod]
        public void GenJitterSheet_DuplicatePins_AreDistinctInOutput()
        {
            // Arrange: Multiple patterns/sheets referencing the same pin
            var sheet1 = new HardIpSheet { Rows = [] };
            sheet1.Rows.Add(CreatePatternWithPin("PinA"));

            var sheet2 = new HardIpSheet { Rows = [] };
            // Duplicate
            sheet2.Rows.Add(CreatePatternWithPin("PinA"));

            _planDic.Add("S1", sheet1);
            _planDic.Add("S2", sheet2);

            // Act
            JitterSheet? result = _generator.GenJitterSheet(_planDic);

            // Assert
            Assert.AreEqual(1, result!.Rows.Count, "Duplicate pins should be filtered out by Distinct().");
        }

        private static HardIpPattern CreatePatternWithPin(string name)
        {
            var p = new HardIpPattern();
            p.MeasPins.Add(new MeasPin { PinName = name, MeasType = MeasType.MeasD });
            return p;
        }
    }

}
