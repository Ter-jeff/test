namespace TagDiff.Core.Common
{
    public class Location(int row, int col)
    {
        public int Row { set; get; } = row;
        public int Col { set; get; } = col;
        public bool IsUsed { set; get; }
    }
}
