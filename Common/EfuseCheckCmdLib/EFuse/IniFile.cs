using System.Globalization;
using System.IO;
using System.Reflection;

using IniParser;
using IniParser.Model;

namespace EfuseCheckCmdLib.EFuse
{
    public class IniFile(string? iniPath = null)
    {
        public string Path = new FileInfo(iniPath ?? _exe + ".ini").FullName.ToString(CultureInfo.InvariantCulture);
        private static readonly string _exe = Assembly.GetExecutingAssembly().GetName().Name!;
        private static readonly FileIniDataParser _parser = new();

        public string Read(string key, string? section = null)
        {
            IniData data = _parser.ReadFile(Path);
            return data[section ?? _exe][key] ?? string.Empty;
        }
    }
}
