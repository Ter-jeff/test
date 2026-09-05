using System.Collections.Generic;

using CommonLib.ErrorReport.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.ErrorReport
{
    [TestClass]
    public class ErrorInfoTests
    {
        [TestMethod]
        public void DefaultConstructor_CommentsIsEmptyList()
        {
            // Arrange & Act
            var errorInfo = new ErrorInfo();

            // Assert
            Assert.AreEqual(0, errorInfo.Comments.Count);
        }

        [TestMethod]
        public void DefaultConstructor_PatternIsNull()
        {
            // Arrange & Act
            var errorInfo = new ErrorInfo();

            // Assert
            Assert.IsNull(errorInfo.Pattern);
        }

        [TestMethod]
        public void Pattern_SetAndGet_ReturnsAssignedValue()
        {
            // Arrange
            var errorInfo = new ErrorInfo
            {
                // Act
                Pattern = "pattern_001"
            };

            // Assert
            Assert.AreEqual("pattern_001", errorInfo.Pattern);
        }

        [TestMethod]
        public void Comments_SetAndGet_ReturnsAssignedList()
        {
            // Arrange
            var errorInfo = new ErrorInfo();
            var comments = new List<string> { "first", "second" };

            // Act
            errorInfo.Comments = comments;

            // Assert
            CollectionAssert.AreEqual(comments, errorInfo.Comments);
        }

        [TestMethod]
        public void Comments_SetToNull_ReturnsNull()
        {
            // Arrange
            var errorInfo = new ErrorInfo
            {
                // Act
                Comments = null
            };

            // Assert
            Assert.AreEqual(null, errorInfo.Comments);
        }
    }
}
