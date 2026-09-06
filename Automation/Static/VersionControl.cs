using System.Diagnostics;
using System.Reflection;

using CommonLib.Static;

namespace Automation.Static
{
    public class VersionControl
    {
        public static string ToolVersion = "V" + AssemblyProvider.Current.GetFileVersion(FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion);

        public static string Timestamp
        {
            get { return "_" + ToolVersion + "_" + TimeContext.Now.ToString("yyyy-MM-dd HHmmss"); }
        }
    }
}
