using System.Collections.Generic;
using System.Diagnostics;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test
{
    [TestClass]
    public class ChunkByTests
    {
        [TestMethod]
        public void Test1()
        {
            KeyValuePair<string, string>[] list = new[]
                {
                    new KeyValuePair<string, string>("A", "We"),
                    new KeyValuePair<string, string>("A", "think"),
                    new KeyValuePair<string, string>("A", "that"),
                    new KeyValuePair<string, string>("B", "Linq"),
                    new KeyValuePair<string, string>("C", "is"),
                    new KeyValuePair<string, string>("A", "really"),
                    new KeyValuePair<string, string>("B", "cool"),
                    new KeyValuePair<string, string>("B", "!")
                };
            IEnumerable<System.Linq.IGrouping<string, KeyValuePair<string, string>>> query = list.ChunkBy(p => p.Key);
            foreach (System.Linq.IGrouping<string, KeyValuePair<string, string>> item in query)
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
