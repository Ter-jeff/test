using System.Collections.Generic;

using CommonLib.Extension;

namespace TestPlanLib.Basic
{
    public class VariableInitTable
    {
        public List<VariableRow> Rows = [];
    }

    public class VariableRow
    {
        public string Opcode = "";
        public string Parameter = "";
        public Dictionary<string, string> JobValues = new(StringExtensions.IgnoreCase);

        public string Comment { get; internal set; } = "";
    }
}
