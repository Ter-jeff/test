using System;
using System.IO;

using Automation.GenerateIgxl.EFuse;
using Automation.Static;

using CommonLib.Enums;

using EfuseCheckCmdLib.EFuse.EFuseApp;

using LogLib.Static;

using OfficeOpenXml;

namespace EfuseCheckCmdLib
{
    internal class EFuseMainScript : EFuseMain
    {
        public void NonIgxlSheetProcess(bool isFromCheckScript = false)
        {
            try
            {
                Response.Report("Reading Efuse files...", EnumMessageLevel.General, 10);
                var efuseInputManager = new EfuseInputManager();
                if (isFromCheckScript)
                {
                    efuseInputManager.ReadForCheckScript();
                }
                else
                {
                    efuseInputManager.ReadAndPreCheck();
                }
                var efuseGenerateBdf = new EfuseGenerateBdf(EFuseInputData, isFromCheckScript);
                efuseGenerateBdf.WorkFlow();
            }
            catch (Exception e)
            {
                Response.Report("Efuse NonIgxlSheetProcess Failed: " + e.StackTrace, EnumMessageLevel.Error, 100);
            }
        }

        public static void CheckProberHexCode(string waferIdFile)
        {
            var inputExcel = new ExcelPackage(new FileInfo(waferIdFile));
            EpWorkbook.TestPlanWorkbook = inputExcel.Workbook;
            ExcelWorksheet ecidSheet = EpWorkbook.TestPlanWorkbook.Worksheets["ECID"];
            if (ecidSheet != null)
            {
                int x = ecidSheet.Dimension.Rows;
                for (int row = 2; row <= x; row++)
                {
                    string proberHex = EfuseScriptUtility.GetProberHexcode(ecidSheet.Cells[row, 1].Text, ecidSheet.Cells[row, 2].Text, ecidSheet.Cells[row, 3].Text, ecidSheet.Cells[row, 4].Text, true);

                    ecidSheet.Cells[row, 5].Value = proberHex;
                }
            }

            inputExcel.Save();

        }
    }
}
