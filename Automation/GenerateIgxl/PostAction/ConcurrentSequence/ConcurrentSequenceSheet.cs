using System.Collections.Generic;

namespace Automation.GenerateIgxl.PostAction.ConcurrentSequence
{
    public class ConcurrentSequenceSheet
    {
        public List<ConcurrentSequenceRow> Rows { get; set; } = new List<ConcurrentSequenceRow>();

        public ConcurrentSequenceSheet(string sheetName)
        {
            Rows = new List<ConcurrentSequenceRow>();
        }

        public void AddRow(ConcurrentSequenceRow concurrentSequenceRow)
        {
            Rows.Add(concurrentSequenceRow);
        }
    }
}
