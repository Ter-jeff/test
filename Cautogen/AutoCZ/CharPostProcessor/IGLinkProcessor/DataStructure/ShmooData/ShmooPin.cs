using System;

namespace Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.ShmooData
{
    public class ShmooPin : IEquatable<ShmooPin>
    {
        /* Property */
        public string SweepPinName { get; set; }
        public string SweepType { get; set; }
        public string StartPoint { get; set; }
        public string StopPoint { get; set; }
        public string StepSize { get; set; }
        public string ShmooType { get; set; }
        public string PortName { get; set; }

        /* Constructor */
        public ShmooPin(string sweepPin, string start = "", string stop = "", string size = "", string shmooType = "")
        {
            if (sweepPin.Contains(':'))
            {
                SweepPinName = sweepPin.Split(':')[1];
                SweepType = sweepPin.Split(':')[0];
            }

            StartPoint = start;
            StopPoint = stop;
            StepSize = size;
            ShmooType = shmooType;
            PortName = "";
        }

        /* Member function */
        public bool Equals(ShmooPin other)
        {
            return SweepPinName == other.SweepPinName && SweepType == other.SweepType && StartPoint == other.StartPoint && StopPoint == other.StopPoint && StepSize == other.StepSize;
        }
    }
}
