using System.Collections.Generic;
using System.Diagnostics;

using BinCutScriptLib.Base.Line;

namespace BinCutScriptLib.Extension
{
    [DebuggerDisplay("{Line.Line}")]
    public class Node
    {
        public BinCutLineBase Line { get; set; } = new();
        public List<Node> Children { get; set; } = [];
    }
}
