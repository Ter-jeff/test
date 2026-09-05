namespace FileDiffLib
{
    public class FileDiff
    {
        public string ExpectedFile { get; internal set; } = string.Empty;
        public string OutputFile { get; internal set; } = string.Empty;
        public bool IsSame { get; set; }
    }
}
