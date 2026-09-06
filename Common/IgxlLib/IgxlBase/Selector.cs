namespace IgxlLib.IgxlBase
{
    public class Selector
    {
        public string? SelectorName { get; set; }
        public string? SelectorValue { get; set; }

        public Selector()
        {
        }

        public Selector(string name, string value)
        {
            SelectorName = name;
            SelectorValue = value;
        }
    }
}
