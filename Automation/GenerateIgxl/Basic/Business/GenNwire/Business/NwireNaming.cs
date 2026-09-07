using System.Text.RegularExpressions;

using Automation.Singleton;

namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Business
{
    public static class NwireNaming
    {
        private static readonly Regex _regex = new Regex("^F");

        public static string GetBinTableName()
        {
            return _regex.Replace(NwireSingleton.NwireFlag, "Bin_DC");
        }
    }
}
