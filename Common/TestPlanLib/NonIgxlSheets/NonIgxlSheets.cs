using System.Collections.Generic;
using System.IO;

namespace TestPlanLib.NonIgxlSheets
{
    public class NonIgxlSheets
    {
        #region Property
        public List<string> SheetList { get; } = [];

        #endregion

        #region Member Function
        public void Add(string dir, string fileName)
        {
            string fileFullName = Path.Combine(dir, fileName);
            SheetList.Add(fileFullName);
        }

        public void Clear()
        {
            SheetList.Clear();
        }
        #endregion
    }
}
