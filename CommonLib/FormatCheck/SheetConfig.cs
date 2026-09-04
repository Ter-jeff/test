namespace CommonLib.FormatCheck
{
    public class SheetConfig
    {
        public string SheetName { get; set; } = string.Empty;
        public string FirstHeaderName { get; set; } = string.Empty;
        public string HeaderName { get; set; } = string.Empty;
        public bool Optional { get; set; } = false;
        public EnumColumn Type { get; set; } = EnumColumn.None;
    }
}
