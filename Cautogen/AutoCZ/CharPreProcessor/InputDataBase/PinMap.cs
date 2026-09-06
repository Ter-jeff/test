namespace Cautogen.AutoCZ.CharPreProcessor.InputDataBase
{
    public class PinInfo
    {
        public string PinName { get; private set; }
        public string PinType { get; private set; }

        public PinInfo(string pinName, string pinType)
        {
            PinName = pinName;
            PinType = pinType;
        }
    }
}
