using System.IO;
using System.Reflection;

using CommonLib.Static;

namespace Cautogen.Utility
{
    public class VersionControl
    {
        #region Field
        public static string ToolVersion = "V" + Assembly.GetExecutingAssembly().GetName().Version;
        #endregion

        #region Property
        public static string Vtimestamp
        {
            get { return "_" + ToolVersion + "_" + TimeContext.Now.ToString("yyyy-MM-dd HHmmss"); }
        }

        public static string AddTimeStamp(string file)
        {
            return Path.GetFileNameWithoutExtension(file) + "_Real" + Vtimestamp + Path.GetExtension(file);
        }

        #endregion
    }
}
