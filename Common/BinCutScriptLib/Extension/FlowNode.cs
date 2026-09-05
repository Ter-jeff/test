using System.Diagnostics;

namespace BinCutScriptLib.Extension
{
    [DebuggerDisplay("{FlowName}")]
    public class FlowNode : Node
    {
        public string FlowName { get; internal set; } = string.Empty;
    }
}
