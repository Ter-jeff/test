using System.Collections.Generic;

using CommonLib.Extension;

namespace IgxlLib.IgxlBase.MultiRow
{
    public class InstanceRows : List<InstanceRow>
    {
        public InstanceRows()
        {
        }

        public InstanceRows(IEnumerable<InstanceRow> instanceRows) : base(instanceRows)
        {
        }

        public void AddFooter(string block)
        {
            var footer = new InstanceRow
            {
                TestName = block + "_Footer_1",
                VbtType = "VBT",
                ArgList = "PrintInfo",
                VbtName = "Print_Footer"
            };
            footer.Args.Add(block);
            Add(footer);
        }

        public void AddHeader(string block)
        {
            var header = new InstanceRow
            {
                TestName = block + "_Header_1",
                VbtType = "VBT",
                ArgList = "PrintInfo",
                VbtName = "Print_Header"
            };
            header.Args.Add(block);
            Add(header);
        }

        public void AddHeaderFooter(string sheetName)
        {
            string block = sheetName.ReplaceStartsWith("Flow_", "");

            AddHeader(block);

            AddFooter(block);
        }
    }
}
