using System.Collections.Generic;

using Automation.GenerateIgxl.PostAction.InstCommon;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class InstCommonOutputTests : FunctionTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            LocalSpecs.TarFolder = OutputPath;
            PinMapSheet pinMapSheet = new PinMapSheet("");
            PinGroup pin = new PinGroup("Pins_1p8v", PinMapConst.TypeIo);
            pin.AddPin(new Pin("TX_P", "I/O"));
            pinMapSheet.AddGroup(pin);
            TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(FolderStructure.DirPinMap, pinMapSheet);
        }

        [TestMethod]
        public void SetPpmuClampPara_ValidPinName_CalculatesClampCorrecty()
        {
            // Arrange
            var row = new InstanceRow { VbtName = "Set_PPMU_Clamp", TestName = "MyTest" };
            var rows = new List<InstanceRow> { row };

            var service = new InstCommonOutput();

            // Act
            service.SetPpmuClampPara(rows);

            // Assert
            // 1.8 * 1.2 = 2.16
            Assert.AreEqual("Pins_1p8v", row.Args[0]);
            Assert.AreEqual("2.16", row.Args[1]);
        }

        [TestMethod]
        public void SetPpmuClampPara_ContiTest_ShouldSkip()
        {
            // Arrange
            var row = new InstanceRow { VbtName = "Set_PPMU_Clamp", TestName = "Conti_Check" };
            row.Args.Add("OldValue");

            // Act
            new InstCommonOutput().SetPpmuClampPara([row]);

            // Assert
            Assert.AreEqual("OldValue", row.Args[0], "Args should not be cleared for 'conti' tests.");
        }
    }
}
