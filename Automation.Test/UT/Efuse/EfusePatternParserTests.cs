using System;
using System.IO;
using System.IO.Compression;
using System.Text;

using Automation.GenerateIgxl.EFuse.Business;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfusePatternParserTests : FunctionTestBase
    {
        private string _tempFilePath = null!;

        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(_tempFilePath))
            {
                File.Delete(_tempFilePath);
            }
        }

        private string CreateGzipFile(string content)
        {
            _tempFilePath = Path.Combine(Path.GetTempPath(), $"pattern_{Guid.NewGuid():N}.gz");
            using (var fs = new FileStream(_tempFilePath, FileMode.Create))
            using (var gzip = new GZipStream(fs, CompressionMode.Compress))
            using (var writer = new StreamWriter(gzip, Encoding.UTF8))
            {
                writer.Write(content);
            }
            return _tempFilePath;
        }

        [TestMethod]
        public void WorkFlow_WriteMode_Extracts_DigSrc_Pin_And_Counts_Send()
        {
            // Arrange
            string patternContent = @"
                instruments = {
                    moduleA(PIN_A):DigSrc|
                }
                PIN_A):DigSrc = Send)
                PIN_A):DigSrc = Send)
            ";

            string file = CreateGzipFile(patternContent);
            var parser = new EfusePatternParser(file, isRead: false, isWrite: true);

            // Act
            parser.WorkFlow();

            // Assert
            Assert.AreEqual("PIN_A", parser.ReadWritePin);
            Assert.AreEqual(2, parser.SendCount);
            Assert.AreEqual(0, parser.StoreCount);
        }

        [TestMethod]
        public void WorkFlow_ReadMode_Extracts_DigCap_Pin_And_Counts_Store()
        {
            // Arrange
            string patternContent = @"
                instruments = {
                    moduleB(PIN_B):DigCap|
                }
                PIN_B):DigCap = Store)
                PIN_B):DigCap = Store)
                PIN_B):DigCap = Store)
            ";

            string file = CreateGzipFile(patternContent);
            var parser = new EfusePatternParser(file, isRead: true, isWrite: false);

            // Act
            parser.WorkFlow();

            // Assert
            Assert.AreEqual("PIN_B", parser.ReadWritePin);
            Assert.AreEqual(3, parser.StoreCount);
            Assert.AreEqual(0, parser.SendCount);
        }

        [TestMethod]
        public void WorkFlow_NoMatch_Returns_EmptyPin_And_ZeroCounts()
        {
            // Arrange
            string patternContent = @"
                instruments = {
                    moduleC(PIN_X):OtherType;
                }
            ";

            string file = CreateGzipFile(patternContent);
            var parser = new EfusePatternParser(file, isRead: true, isWrite: false);

            // Act
            parser.WorkFlow();

            // Assert
            Assert.AreEqual("", parser.ReadWritePin);
            Assert.AreEqual(0, parser.SendCount);
            Assert.AreEqual(0, parser.StoreCount);
        }

        [TestMethod]
        public void WorkFlow_Ignores_Commented_Lines()
        {
            // Arrange
            string patternContent = @"
                // instruments = { moduleD(PIN_D):DigSrc| }
                instruments = { (moduleD(PIN_D):DigSrc|) }
                //PIN_D):DigSrc = Send)
                PIN_D):DigSrc = Send)
            ";

            string file = CreateGzipFile(patternContent);
            var parser = new EfusePatternParser(file, isRead: false, isWrite: true);

            // Act
            parser.WorkFlow();

            // Assert
            Assert.AreEqual("PIN_D", parser.ReadWritePin);
            Assert.AreEqual(1, parser.SendCount);
        }

        [TestMethod]
        public void WorkFlow_Stops_At_GlobalSubr_Block()
        {
            // Arrange
            string patternContent = @"
                instruments = {
                    moduleE(PIN_E):DigCap|
                }
                PIN_E):DigCap = Store)
                global subr Something()
                PIN_E):DigCap = Store)
            ";

            string file = CreateGzipFile(patternContent);
            var parser = new EfusePatternParser(file, isRead: true, isWrite: false);

            // Act
            parser.WorkFlow();

            // Assert
            Assert.AreEqual("PIN_E", parser.ReadWritePin);
            Assert.AreEqual(1, parser.StoreCount);
        }
    }
}
