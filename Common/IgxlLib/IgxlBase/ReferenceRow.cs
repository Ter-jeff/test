using System.Linq;

namespace IgxlLib.IgxlBase
{
    public class ReferenceRow : IgxlRow, IIgxlRow
    {
        public string FilePath { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;

        public string[] Print()
        {
            string[] arr = [.. Enumerable.Repeat("", 3)];
            if (!string.IsNullOrEmpty(FilePath))
            {
                arr[0] = ColumnA ?? "";
                arr[1] = FilePath;
                arr[2] = Comment;
            }
            else
            {
                arr = ["\t"];
            }
            return arr;
        }
    }
}
