using System.Text.RegularExpressions;

namespace ScghLib.Reader
{
    internal static partial class ProdCharSheetRowHelpers
    {

        [GeneratedRegex(@"\w+_(?<chiplet>[A-z]\d+$)")]
        public static partial Regex MyRegex();
    }
}
