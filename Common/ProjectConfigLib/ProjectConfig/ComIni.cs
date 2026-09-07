using IniParser;
using IniParser.Model;

namespace ProjectConfigLib.ProjectConfig
{
    public class ComIni
    {
        private static readonly FileIniDataParser _parser = new();

        public static string IniRead(string key, string section, string pPath)
        {
            IniData data = _parser.ReadFile(pPath);
            return data[section][key] ?? string.Empty;
        }

        public static void IniWrite(string key, string value, string section, string pPath)
        {
            IniData data = _parser.ReadFile(pPath);
            if (!data.Sections.ContainsSection(section))
            {
                data.Sections.AddSection(section);
            }
            data[section][key] = value;
            _parser.WriteFile(pPath, data);
        }

        public static void IniDeleteKey(string key, string section, string pPath)
        {
            IniData data = _parser.ReadFile(pPath);
            if (data.Sections.ContainsSection(section))
            {
                data[section].RemoveKey(key);
            }
            _parser.WriteFile(pPath, data);
        }

        public static void DeleteSection(string section, string pPath)
        {
            IniData data = _parser.ReadFile(pPath);
            data.Sections.RemoveSection(section);
            _parser.WriteFile(pPath, data);
        }

        public static bool IniKeyExists(string key, string section, string pPath)
        {
            IniData data = _parser.ReadFile(pPath);
            return data.Sections.ContainsSection(section) && data[section].ContainsKey(key);
        }
    }
}
