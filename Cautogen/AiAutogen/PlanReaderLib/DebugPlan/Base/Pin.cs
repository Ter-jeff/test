using System;

using CommonLib.Extension;

namespace DebugPlanReaderLib.DebugPlan
{
    public class Pin
    {
        public string Name { get; set; }
        public string Start { get; set; }
        public int IndexStart { get; set; }
        public string Stop { get; set; }
        public int IndexStop { get; set; }
        public string Step { get; set; }
        public string Type { get; set; }
        public int IndexStep { get; set; }

        public bool IsSearch
        {
            get { return Start != Stop; }
        }

        public bool IsForce
        {
            get
            {

                if (!string.IsNullOrEmpty(Start) && !string.IsNullOrEmpty(Stop))
                {
                    double startValue;
                    double stopValue;
                    var start = Start.ConvertNumber();
                    var stop = Stop.ConvertNumber();

                    if (double.TryParse(start, out startValue) && double.TryParse(stop, out stopValue))
                    {
                        ForceValue = start;
                        return startValue == stopValue;
                    }
                }
                return false;
            }
        }

        public string ForceValue
        {
            get;
            private set;
        }

        public string ShmooName
        {
            get { return "Shmoo_" + Name; }
        }

        public decimal Steps
        {
            get
            {
                decimal startValue;
                var start = Start.ConvertNumber();
                if (!decimal.TryParse(start, out startValue))
                {
                    return 0;
                }

                decimal stopValue;
                var stop = Stop.ConvertNumber();
                if (!decimal.TryParse(stop, out stopValue))
                {
                    return 0;
                }

                decimal stepValue;
                var step = Step.ConvertNumber();
                if (!decimal.TryParse(step, out stepValue))
                {
                    return 0;
                }

                if (stepValue == 0)
                    return 0;

                var value = Math.Abs((stopValue - startValue) / stepValue);
                return value;
            }
        }
    }
}
