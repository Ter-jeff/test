using System.Collections.Generic;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class TestPlanSequence
    {
        public int StartRow { get; }
        public int EndRow { get; }
        public int SeqIndex { get; set; }
        public int MeasType { get; set; }
        public List<string> ForceCondition { set; get; }

        public TestPlanSequence(int startRow, int endRow, int seqIndex)
        {
            StartRow = startRow;
            EndRow = endRow;
            SeqIndex = seqIndex;
            ForceCondition = new List<string>();
        }

        public TestPlanSequence(TestPlanSequence other)
        {
            if (other == null)
            {
                return;
            }

            StartRow = other.StartRow;
            EndRow = other.EndRow;

            SeqIndex = other.SeqIndex;
            MeasType = other.MeasType;

            ForceCondition = other.ForceCondition != null ? new List<string>(other.ForceCondition) : new List<string>();
        }

        public TestPlanSequence Copy()
        {
            return new TestPlanSequence(this);
        }
    }
}
