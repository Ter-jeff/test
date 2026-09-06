using System;
using System.Collections.Generic;
using System.IO;

using OfficeOpenXml;

namespace Cautogen.common.ReaderWriter.Writer
{
    public class ExcelWriterBase : IWriter
    {
        protected List<IExcelSheetWriter> SheetWriterList = new List<IExcelSheetWriter>();
        public List<Action> CallbackFuncs = new List<Action>();

        public ExcelWriterBase(string excelFilePath)
        {
            ExcelFilePath = excelFilePath;
        }

        public string ExcelFilePath { get; }

        public void Write()
        {
            var fileInfo = new FileInfo(Path.GetFullPath(ExcelFilePath));
            using (var excel = new ExcelPackage(fileInfo))
            {
                foreach (IExcelSheetWriter shWriter in SheetWriterList)
                {
                    shWriter.Write(excel.Workbook);
                }

                excel.Save();
            }

            foreach (Action callbackFunc in CallbackFuncs)
            {
                callbackFunc();
            }

            _Logging();
        }

        protected virtual void _Logging()
        {

        }
    }
}
