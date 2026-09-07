using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.InputChecker
{
    public class EfuseArraySizeChecker
    {
        public void WorkFlow(EfuseArraySizeSheet sheet)
        {
            CheckOfBitColumn(sheet);
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                CheckOfBitCountIsDivideBy16(sheet);
            }
        }

        private void CheckOfBitCountIsDivideBy16(EfuseArraySizeSheet sheet)
        {
            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                EfuseArraySizeRow row = sheet.Rows[i];
                if (Regex.IsMatch(row.OfBits.Trim(), @"^\d+$") && int.Parse(row.OfBits) % 16 != 0)
                {
                    sheet.AddError(EFuseErrorType.E_MismatchBit_01, sheet.SheetName, row.RowNum, sheet.OfBitsColNumber, "This bank count should be mod by 16 due to the ECC");
                    ErrorReportManager.AddError(EFuseErrorType.E_MismatchBit_01, sheet.SheetName, row.RowNum, sheet.OfBitsColNumber, []);
                }
            }
        }

        private void CheckOfBitColumn(EfuseArraySizeSheet sheet)
        {
            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                EfuseArraySizeRow row = sheet.Rows[i];
                if (string.IsNullOrEmpty(row.OfBits))
                {
                    sheet.AddError(EFuseErrorType.E_MismatchBit_02, sheet.SheetName, row.RowNum, sheet.OfBitsColNumber, "This column cannot be empty");
                    ErrorReportManager.AddError(EFuseErrorType.E_MismatchBit_02, sheet.SheetName, row.RowNum, sheet.OfBitsColNumber, []);
                }
                else
                {
                    if (!Regex.IsMatch(row.OfBits.Trim(), @"^\d+$"))
                    {
                        sheet.AddError(EFuseErrorType.E_MismatchBit_03, sheet.SheetName, row.RowNum, sheet.OfBitsColNumber, "This column must be a number");
                        ErrorReportManager.AddError(EFuseErrorType.E_MismatchBit_03, sheet.SheetName, row.RowNum, sheet.OfBitsColNumber, []);
                    }
                }
            }
        }
    }
}
