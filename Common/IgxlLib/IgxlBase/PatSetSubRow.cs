using System.Linq;

namespace IgxlLib.IgxlBase
{
    public class PatSetSubRow : IgxlRow, IIgxlRow
    {
        public string PatternFileName { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;

        public PatSetSubRow()
        {
        }

        public PatSetSubRow(PatSetSubRow patSetSubRow) : base(patSetSubRow)
        {
            if (patSetSubRow == null)
            {
                return;
            }

            PatternFileName = patSetSubRow.PatternFileName;
            Comment = patSetSubRow.Comment;
        }

        public PatSetSubRow Copy()
        {
            return new PatSetSubRow(this);
        }

        public string[] Print()
        {
            string[] arr = [.. Enumerable.Repeat("", 3)];
            if (!string.IsNullOrEmpty(PatternFileName))
            {
                arr[0] = ColumnA ?? "";
                arr[1] = PatternFileName;
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
