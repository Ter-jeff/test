using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase.MultiRow
{
    [TestClass]
    public class InstanceRowsTests
    {
        [TestMethod]
        public void AddHeader_CreatesValidHeaderRowWithBlockArg()
        {
            // Arrange
            var collection = new InstanceRows();
            string blockName = "FunctionalTest";

            // Act
            collection.AddHeader(blockName);

            // Assert
            Assert.AreEqual(1, collection.Count);
            InstanceRow header = collection[0];
            Assert.AreEqual("FunctionalTest_Header_1", header.TestName);
            Assert.AreEqual("VBT", header.VbtType);
            Assert.AreEqual("Print_Header", header.VbtName);
            Assert.AreEqual("PrintInfo", header.ArgList);
            Assert.AreEqual(1, header.Args.Count);
            Assert.AreEqual("FunctionalTest", header.Args[0]);
        }

        [TestMethod]
        public void AddFooter_CreatesValidFooterRowWithBlockArg()
        {
            // Arrange
            var collection = new InstanceRows();
            string blockName = "PowerTest";

            // Act
            collection.AddFooter(blockName);

            // Assert
            Assert.AreEqual(1, collection.Count);
            InstanceRow footer = collection[0];
            Assert.AreEqual("PowerTest_Footer_1", footer.TestName);
            Assert.AreEqual("VBT", footer.VbtType);
            Assert.AreEqual("Print_Footer", footer.VbtName);
            Assert.AreEqual("PrintInfo", footer.ArgList);
            Assert.AreEqual(1, footer.Args.Count);
            Assert.AreEqual("PowerTest", footer.Args[0]);
        }

        [TestMethod]
        public void AddHeaderFooter_StripsFlowPrefixAndAddsBothRows()
        {
            // Arrange
            var collection = new InstanceRows();
            string sheetName = "Flow_MainExecution";

            // Act
            collection.AddHeaderFooter(sheetName);

            // Assert
            Assert.AreEqual(2, collection.Count);

            InstanceRow header = collection[0];
            Assert.AreEqual("MainExecution_Header_1", header.TestName);
            Assert.AreEqual("MainExecution", header.Args[0]);

            InstanceRow footer = collection[1];
            Assert.AreEqual("MainExecution_Footer_1", footer.TestName);
            Assert.AreEqual("MainExecution", footer.Args[0]);
        }

        [TestMethod]
        public void AddHeaderFooter_HandlesSheetNameWithoutFlowPrefix()
        {
            // Arrange
            var collection = new InstanceRows();
            string sheetName = "NoPrefixSheet";

            // Act
            collection.AddHeaderFooter(sheetName);

            // Assert
            Assert.AreEqual(2, collection.Count);
            Assert.AreEqual("NoPrefixSheet_Header_1", collection[0].TestName);
            Assert.AreEqual("NoPrefixSheet_Footer_1", collection[1].TestName);
        }
    }
}
