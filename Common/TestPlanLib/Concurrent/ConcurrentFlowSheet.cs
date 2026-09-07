using System.Collections.Generic;

using CommonReaderLib;

namespace TestPlanLib.Concurrent
{
    public class ConcurrentFlowSheet : MySheet
    {
        public List<ConcurrentFlowSheetRow> Rows = [];

        public int SequenceNameColNumber = -1;
        public int SubflowColStart = -1;
        public int SubflowColEnd = -1;

        public ConcurrentFlowSheet(string sheetName)
        {
            SheetName = sheetName;
        }
    }

    public class ConcurrentFlowSheetRow : MyRow
    {
        public string SequenceName = "";
        public List<string> Subflows = [];
    }
}
