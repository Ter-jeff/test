using System.Collections.Generic;

namespace Automation.Singleton
{
    public class InstanceExceptionList
    {
        public List<InstanceExceptionEntry> ExceptionList { get; } = new List<InstanceExceptionEntry>();

        public void AddRow(InstanceExceptionEntry entry)
        {
            ExceptionList.Add(entry);
        }
    }

    public class InstanceExceptionEntry
    {
        public string TestName;
        public string DcCategory;
    }
}
