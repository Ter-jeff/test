namespace IgxlLib.IgxlSheets
{
    public interface IIgxlSheet
    {
        string IgxlSheetName { get; }
        string Name { get; set; }
        string SheetType { get; }

        void Write(string fileName, string version = "");
    }
}
