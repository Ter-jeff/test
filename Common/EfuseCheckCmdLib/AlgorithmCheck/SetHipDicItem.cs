using System.Text.RegularExpressions;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    public class SetHipDicItem
    {
        public string Site = "";
        public string Key = "";
        public string Value = "";
        public string Format = "";
        public string TestName = "";
        public int Number;
        public SetHipDicItem(string line)
        {
            string regHip = @"Site\((?<site>\d+)\).*Key_Name:= (?<key>\w+).*Value:= (?<value>[\d\.\-]+)";

            Site = Regex.Match(line, regHip, RegexOptions.IgnoreCase).Groups["site"].Value;
            Key = Regex.Match(line, regHip, RegexOptions.IgnoreCase).Groups["key"].Value;
            Value = Regex.Match(line, regHip, RegexOptions.IgnoreCase).Groups["value"].Value;
        }
    }
}
