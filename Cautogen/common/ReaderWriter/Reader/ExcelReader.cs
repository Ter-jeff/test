using System;
using System.IO;

using OfficeOpenXml;

namespace Cautogen.common.ReaderWriter.Reader
{
    public class ExcelReader : ReaderBase
    {
        protected ExcelWorkbook Workbook;

        /* constructor */
        public ExcelReader(string filePath, Action callbackFunc = null) : base(filePath, callbackFunc)
        {
        }

        /* methods */
        public override void Read()
        {
            if (!IsFileExist())
            {
                return;
            }

            try
            {
                Workbook = new ExcelPackage(new FileInfo(FilePath)).Workbook;

                foreach (ExcelWorksheet sh in Workbook.Worksheets)
                {
                    _ReadPreProcess(sh);
                }
                foreach (ExcelWorksheet sh in Workbook.Worksheets)
                {
                    _Read(sh);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        protected virtual void _Read(ExcelWorksheet sh)
        {
            throw new NotImplementedException();
        }

        protected virtual void _ReadPreProcess(ExcelWorksheet sh)
        {
        }
    }
}
