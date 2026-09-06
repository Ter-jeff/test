using System.Collections.Generic;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Utility;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class PreCheckCtrl
    {
        /* properties */
        private readonly List<IPreCheck> _preCheckerList;
        private readonly List<string> _dupCategory;
        private readonly Dictionary<string, List<Characterization>> _payload1UseCount;

        /* constructor */
        public PreCheckCtrl(List<IPreCheck> preCheckerList)
        {
            _preCheckerList = preCheckerList;
            _dupCategory = new List<string>();
            _payload1UseCount = new Dictionary<string, List<Characterization>>();
        }

        /* methods*/
        public void WorkFlow()
        {
            // run each check for all plan sheet
            foreach (KeyValuePair<string, List<Characterization>> planSheet in CharPlan.CharPlanSheetDict)
            {
                LogHelper.Info("Checking sheet: " + planSheet.Key + "...");
                MessageWriter.WriteMessage("Checking sheet: " + planSheet.Key + "...", EnumMessageLevel.Info);

                foreach (IPreCheck checker in _preCheckerList)
                {
                    checker.Check(planSheet.Value, planSheet.Key);
                    checker.UpdateErrorMessages();
                }

                _dupCategory.AddRange(planSheet.Value.Select(row => row.Category).Distinct().ToList());

                foreach (Characterization charRow in planSheet.Value)
                {
                    _UpdatePayload1UseCount(charRow);
                }
            }

            /* cross sheets check */
            // check test instance num for each payload1
            foreach (string payload1 in _payload1UseCount.Keys)
            {

                if (_payload1UseCount[payload1].Count <= 200)
                {
                    continue;
                }

                Characterization item = _payload1UseCount[payload1][0];
                const string outString = "Too many instance names";
                ErrorManager.AddWarning(ErrorType.TooManyInstanceName, item.SheetName, item.RowNum, item.ColNum("payload1"), "Use", outString, payload1);
                foreach (int col in item.ColNum("payload1"))
                {
                    ErrorReportManager.AddError(CharErrorType.W_TooManyInstanceName_01, item.SheetName, item.RowNum, col, [],
                        new ErrorInfo() { Comments = new List<string>() { payload1 } });
                }
            }

            // check pin list
            PinListReader.PinListCheck();
        }

        private void _UpdatePayload1UseCount(Characterization charRow)
        {
            string payload1 = UtilityFunction.RemoveDateCode(charRow.Payload1);
            if (payload1 == "")
            {
                return;
            }

            if (_payload1UseCount.TryGetValue(payload1, out List<Characterization> value))
            {
                value.Add(charRow);
            }
            else
            {
                _payload1UseCount[payload1] = new List<Characterization> { charRow };
            }
        }
    }
}
