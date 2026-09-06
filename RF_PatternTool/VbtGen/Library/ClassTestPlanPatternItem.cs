namespace Automation.Library
{
    public class PinItem : ICloneable
    {
        public string PinName { get; set; }

        public Dictionary<string, object> PlanData = new Dictionary<string, object>();

        public PinItem(string pin)
        {
            PinName = pin;
        }

        public object Clone()
        {
            PinItem cloneItem = new PinItem(PinName);

            foreach (KeyValuePair<string, object> pair in PlanData)
            {
                cloneItem.PlanData.Add(pair.Key, pair.Value);
            }
            return cloneItem;

        }
    }
}
