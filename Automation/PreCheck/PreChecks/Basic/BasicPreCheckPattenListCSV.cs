using System.IO;

namespace Automation.PreCheck.PreChecks.Basic
{
    public class BasicPreCheckPattenListCsv : BasicPreCheckBase
    {
        private readonly string _path;

        public BasicPreCheckPattenListCsv(string path) : base(null, null)
        {
            _path = path;
        }

        protected internal override bool CheckExist()
        {
            return File.Exists(_path);
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
