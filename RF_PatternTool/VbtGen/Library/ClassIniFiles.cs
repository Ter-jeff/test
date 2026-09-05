using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace Automation.Library
{
    public class IniFile
    {
        public string Path;
        private static string _eXE = Assembly.GetExecutingAssembly().GetName().Name;

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string defaultString, StringBuilder retVal, int size, string filePath);

        public IniFile(string iniPath = null)
        {
            Path = new FileInfo(iniPath ?? _eXE + ".ini").FullName.ToString(CultureInfo.InvariantCulture);
        }

        public string Read(string key, string section = null)
        {
            var retVal = new StringBuilder(255);
            GetPrivateProfileString(section ?? _eXE, key, "", retVal, 255, Path);
            return retVal.ToString();
        }
    }
}
