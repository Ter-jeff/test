using System.Collections.Generic;
using System.Linq;

namespace Automation.GenerateIgxl.Basic.Business.GenAc.AcInput.BassData
{
    public class AcInputSheet
    {
        #region Field

        #region Const Field
        public const string HeaderStart = "DTACSpecSheet";
        public const string HeaderName = "Name";
        public const string HeaderTyp = "Typ";
        #endregion

        private readonly List<AcInputRow> _acInputData = new List<AcInputRow>();

        #endregion

        #region Property
        public List<AcInputRow> AcInputData
        {
            get { return _acInputData.ToList(); }
        }
        #endregion

        #region Member Function
        public void AddRow(AcInputRow acInputRow)
        {
            _acInputData.Add(acInputRow);
        }
        #endregion
    }
}
