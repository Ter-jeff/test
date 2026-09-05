using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Writer;

using CommonLib.Enums;

namespace Cautogen.AutoCZ.CharPreProcessor.ErrorReport
{
    public class ErrorReportCtrl : ExcelWriterBase
    {
        public ErrorReportCtrl(string charPlanPath, string excelFilePath, IEnumerable<IExcelSheetWriter> sheetWriterList)
            : base(excelFilePath)
        {
            File.Copy(charPlanPath, excelFilePath, true);
            SheetWriterList.AddRange(sheetWriterList);
        }

        protected override void _Logging()
        {
            if (ErrorManager.ErrorListDict.Count == 0)
            {
                MessageWriter.WriteMessage("No error be found ...", EnumMessageLevel.Info);
            }
            else
            {
                MessageWriter.WriteMessage("Generating ErrorReport Sheet ...", EnumMessageLevel.Info);

                // print error summary 
                foreach (ErrorType errorType in ErrorManager.ErrorListDict.Keys)
                {
                    MessageWriter.WriteMessage(errorType + " : " + ErrorManager.ErrorListDict[errorType].ToList().Count,
                        EnumMessageLevel.Error);
                }
            }
        }
    }
}
