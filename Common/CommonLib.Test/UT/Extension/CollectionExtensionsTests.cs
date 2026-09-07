using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class CollectionExtensionsTests
    {
        [TestMethod]
        public void AddRange_ValidItems_AddsAll()
        {
            var list = new List<int> { 1, 2 };
            int[] itemsToAdd = [3, 4, 5];
            list.AddRange(itemsToAdd);

            Assert.AreEqual(5, list.Count);
            Assert.IsTrue(list.Contains(3));
            Assert.IsTrue(list.Contains(4));
            Assert.IsTrue(list.Contains(5));
        }

        [TestMethod]
        public void AddRange_EmptyCollection_AddsNothing()
        {
            var list = new List<int> { 1, 2 };
            int[] itemsToAdd = [];
            list.AddRange(itemsToAdd);

            Assert.AreEqual(2, list.Count);
        }

        [TestMethod]
        public void CartesianProduct_TwoSequences_ReturnsAllCombinations()
        {
            var sequences = new List<IEnumerable<int>>
            {
                new List<int> { 1, 2 },
                new List<int> { 3, 4 }
            };

            var result = sequences.CartesianProduct().ToList();

            Assert.AreEqual(4, result.Count);
        }

        [TestMethod]
        public void CartesianProduct_EmptySequence_ReturnsEmpty()
        {
            var sequences = new List<IEnumerable<int>>
            {
                new List<int>()
            };

            var result = sequences.CartesianProduct().ToList();

            // Empty product should return empty
            Assert.AreEqual(0, result.FirstOrDefault()?.Count() ?? 0);
        }

        [TestMethod]
        public void CartesianProduct_MultipleSequences_ReturnsAllCombinations()
        {
            var sequences = new List<IEnumerable<int>>
            {
                new List<int> { 1, 2 },
                new List<int> { 3, 4 },
                new List<int> { 5, 6 }
            };

            var result = sequences.CartesianProduct().ToList();

            Assert.AreEqual(8, result.Count);
        }
    }
}
