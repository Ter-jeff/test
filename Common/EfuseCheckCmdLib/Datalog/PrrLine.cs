using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Datalog;

namespace EfuseCheckCmdLib.Datalog
{
    public partial class PrrLine : LineBase
    {
        [GeneratedRegex(@"(?<=\[Site\s+\d+\]\s).*(?=\s+Hex)")]
        private static partial Regex MyRegex();

        public PrrLine(string line, int lineNo)
        {
            Line = line;
            LineNo = lineNo;
        }

        public PrrRow ToRow()
        {
            int siteNum = GetSite();
            string[] arr = Line.Split(' ');
            Match match = MyRegex().Match(Line);

            return new PrrRow
            {
                Site = siteNum,
                Type = match.Success ? match.Value : arr[4],
                Prr = arr.Last().Trim('.').Replace("'", ""),
                Line = this,
            };
        }
    }
}
