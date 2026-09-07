namespace Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan
{
    public class Pin
    {
        public string PinCondition;
        public int PlanIndex;
        public string PinName { get; set; }
        public string PinType { get; set; }
        public string Start { get; set; }
        public string Stop { get; set; }
        public string Step { get; set; }
        public string Order { get; set; }

        public Pin()
        {
            Start = "";
            Stop = "";
            Step = "";
            PinType = "";
            PinCondition = "";
            PlanIndex = -1;
            Order = "";
        }
    }
}
