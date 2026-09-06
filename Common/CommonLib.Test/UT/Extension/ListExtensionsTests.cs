using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class ListExtensionsTests
    {
        [TestMethod]
        public void ChunkBy_GroupsByKey_ReturnsCorrectGroups()
        {
            KeyValuePair<string, int>[] items =
            [
                new KeyValuePair<string, int>("A", 1),
                new KeyValuePair<string, int>("A", 2),
                new KeyValuePair<string, int>("B", 3),
                new KeyValuePair<string, int>("B", 4),
                new KeyValuePair<string, int>("A", 5)
            ];

            var groups = items.ChunkBy(x => x.Key).ToList();

            Assert.AreEqual(3, groups.Count);
            Assert.AreEqual(2, groups[0].Count());
            Assert.AreEqual(2, groups[1].Count());
            Assert.AreEqual(1, groups[2].Count());
        }

        [TestMethod]
        public void ChunkBy_SingleGroup_ReturnsOneGroup()
        {
            KeyValuePair<string, int>[] items =
            [
                new KeyValuePair<string, int>("A", 1),
                new KeyValuePair<string, int>("A", 2),
                new KeyValuePair<string, int>("A", 3)
            ];

            var groups = items.ChunkBy(x => x.Key).ToList();

            Assert.AreEqual(1, groups.Count);
            Assert.AreEqual(3, groups[0].Count());
        }

        [TestMethod]
        public void ChunkBy_EmptySequence_ReturnsNoGroups()
        {
            KeyValuePair<string, int>[] items = [];

            var groups = items.ChunkBy(x => x.Key).ToList();

            Assert.AreEqual(0, groups.Count);
        }

        [TestMethod]
        public void ChunkBy_IntKey_ReturnsCorrectGroups()
        {
            int[] items = [1, 1, 2, 2, 3, 1, 1];

            var groups = items.ChunkBy(x => x).ToList();

            Assert.AreEqual(4, groups.Count);
            Assert.AreEqual(1, groups[0].Key);
            Assert.AreEqual(2, groups[0].Count());
            Assert.AreEqual(2, groups[1].Key);
            Assert.AreEqual(2, groups[1].Count());
        }

        [TestMethod]
        public void Test1()
        {
            KeyValuePair<string, string>[] list =
                [
                    new KeyValuePair<string, string>("A", "We"),
                    new KeyValuePair<string, string>("A", "think"),
                    new KeyValuePair<string, string>("A", "that"),
                    new KeyValuePair<string, string>("B", "Linq"),
                    new KeyValuePair<string, string>("C", "is"),
                    new KeyValuePair<string, string>("A", "really"),
                    new KeyValuePair<string, string>("B", "cool"),
                    new KeyValuePair<string, string>("B", "!")
                ];
            IEnumerable<IGrouping<string, KeyValuePair<string, string>>> query = list.ChunkBy(p => p.Key);
            foreach (IGrouping<string, KeyValuePair<string, string>> item in query)
            {
                Trace.WriteLine("Group key = " + item.Key);
                foreach (KeyValuePair<string, string> inner in item)
                {
                    Trace.WriteLine("\t" + inner.Value);
                }
            }
        }
    }
}
