
namespace Automation.GenerateIgxl.PostAction.SelSram
{
    public class SelSramData
    {
        public string Category { get; }

        public string Pattern { get; }

        public string Type { get; }

        public SelSramData(string category, string pattern, string type)
        {
            Category = category;
            Pattern = pattern;
            Type = type;
        }
    }
}
