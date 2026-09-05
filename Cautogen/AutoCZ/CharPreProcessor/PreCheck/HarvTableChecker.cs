using System.Collections.Generic;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class HarvTableChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization item in charList)
            {
                if (!string.IsNullOrEmpty(item.HarvFstp) && !HarvReader.HarvItems.Contains(item.HarvFstp))
                {
                    //ErrorMessages.Add( new ErrorMessage()
                    //{
                    //    ErrorLevel=ErrorLevel.Error,
                    //    ErrorType=ErrorType.NotInHarvTable,
                    //    SheetName=item.SheetName,
                    //    RowNum=item.RowNum,
                    //    ColList  = item.ColNum(item.HarvFstp),
                    //    Message = item.HarvFstp + " in " + item.SheetName + " Row: " + item.RowNum +" is not defined in Harv_Mapping_Table",
                    //});
                    ErrorManager.AddWarning(ErrorType.NotInHarvTable, item.SheetName, item.RowNum,
                                item.ColNum("harv_fstp"), "use", item.HarvFstp + " in " + item.SheetName + " Row: " + item.RowNum + " is not defined in Harv_Mapping_Table");
                    foreach (int col in item.ColNum("harv_fstp"))
                    {
                        ErrorReportManager.AddError(CharErrorType.W_NotInHarvTable_01, item.SheetName, item.RowNum, col, [item.HarvFstp, item.SheetName, $"{item.RowNum}"]);
                    }
                }
            }
        }
    }
}
