using System.IO;

using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.PreCheck.PreChecks.Basic
{
    public class BasicDirectoryCheck : PreCheckBase
    {
        private readonly string _neededCheckDir;

        public BasicDirectoryCheck(string needCheckDir) : base(null, null)
        {
            _neededCheckDir = needCheckDir;
        }

        protected internal override bool CheckExist()
        {
            bool result = true;
            if (!Directory.Exists(_neededCheckDir) && !LocalSpecs.Options.IsIgnorePatternCheck)
            {
                ErrorReportManager.AddError(BasicErrorType.E_FormatError_10, "directory", 1, 0, "The directory is not exist : " + _neededCheckDir, new string[] { _neededCheckDir });
                result = false;
            }
            return result;
        }

        protected internal override bool CheckHeaders()
        {
            return true;
        }

        protected internal override bool CheckFormat()
        {
            return true;
        }

        protected internal override bool CheckBusiness()
        {
            return true;
        }
    }
}
