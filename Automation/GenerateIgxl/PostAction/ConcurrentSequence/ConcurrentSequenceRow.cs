using System.Collections.Generic;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.PostAction.ConcurrentSequence
{
    public class ConcurrentSequenceRow : IgxlRow
    {
        public string Disable { get; set; }
        public string SequenceName { get; set; }
        public Dictionary<string, List<FlowStep>> FlowSteps = new Dictionary<string, List<FlowStep>>();

        public ConcurrentSequenceRow()
        {
        }
    }

    public class FlowStep
    {
        public string FlowStepItem { get; set; }
        public string BackgroundSubStep { get; set; }
    }
}
