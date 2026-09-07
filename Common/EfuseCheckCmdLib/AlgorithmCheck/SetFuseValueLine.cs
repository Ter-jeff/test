
using System.Text.RegularExpressions;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    internal partial class SetFuseValueLine : XLine
    {
        private const string FusePattern1 = @"\[Site (?<site>\d+)\] Set fuse value in cache '(?<name>\w+)' = 'SiteGeneric`1 \{(?<value>.*?)\}'";
        private const string FusePattern2 = @"\[Site (?<site>\d+)\] Set fuse value in cache '(?<name>\w+)' =\s*(?<value>\d+)\s*";
        public static readonly Regex RegexFuse1 = Fuse1Regex();
        public static readonly Regex RegexFuse2 = Fuse2Regex();

        [GeneratedRegex(FusePattern1, RegexOptions.IgnoreCase & RegexOptions.Compiled)]
        private static partial Regex Fuse1Regex();

        [GeneratedRegex(FusePattern2, RegexOptions.IgnoreCase & RegexOptions.Compiled)]
        private static partial Regex Fuse2Regex();

        //[INFO][Site 0] config Fuse SetWriteVariable_SiteAware mtr_bts_ta000_c3 = 1048448
        //[INFO][Site 0] Setting value '1048448' for fuse 'mtr_bts_ta000_c3' in bank 'config' starts.
        //[INFO][Site 0] Set fuse value in cache 'mtr_bts_ta000_c3' = 'SiteGeneric`1 { [0] = 1048448, , ,  }'.
        //[INFO][Site 0] Fuse 'mtr_bts_ta000_c3' is set to value '1048448'.
        //[INFO][Site 0] Fuse value locked for 'mtr_bts_ta000_c3'.
        public SetFuseValueRow ConvertSetFuseValueRow()
        {
            var row = new SetFuseValueRow { Site = GetSite() };
            string name = "";
            string value = "";
            if (Line.Contains('{'))
            {
                Match match = RegexFuse1.Match(Line);
                if (match.Success)
                {
                    name = match.Groups["name"].ToString();
                    value = match.Groups["value"].ToString();
                }
            }
            else
            {
                Match match = RegexFuse2.Match(Line);
                if (match.Success)
                {
                    name = match.Groups["name"].ToString();
                    value = match.Groups["value"].ToString();
                }
            }
            row.Name = name;
            row.Value = value;

            return row;
        }
    }
}
