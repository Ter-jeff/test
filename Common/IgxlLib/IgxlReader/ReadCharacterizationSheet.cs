using System;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadCharacterizationSheet : IgxlSheetReader<CharSheet>
    {
        private const int StartRowIndex = 4;
        private const int StartColumnIndex = 2;

        public override EnumSheetType SupportedType => EnumSheetType.DTCharacterizationSheet;

        public override CharSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var charSheet = new CharSheet(excelWorksheet);
            int maxRowCount = excelWorksheet.Dimension.End.Row;

            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                CharSetup charSetup = GetCharSetup(excelWorksheet, i);

                CharSetup? existingSetup = charSheet.Rows.FirstOrDefault(x =>
                    x.SetupName.EqualsIgnoreCase(charSetup.SetupName) &&
                    x.TestMethod.EqualsIgnoreCase(charSetup.TestMethod));

                if (existingSetup == null)
                {
                    existingSetup = charSetup;
                    charSheet.AddRow(charSetup);
                }

                existingSetup.CharSteps.Add(GetCharStep(excelWorksheet, i));
            }

            return charSheet;
        }

        public override CharSheet ReadSheet(Stream stream, string sheetName)
        {
            throw new NotImplementedException();
        }

        private CharSetup GetCharSetup(ExcelWorksheet excelWorksheet, int row)
        {
            return new CharSetup
            {
                SheetName = excelWorksheet.Name,
                RowNum = row,
                ColumnA = GetCellText(excelWorksheet, row, 1),
                SetupName = GetCellText(excelWorksheet, row, StartColumnIndex),
                TestMethod = GetCellText(excelWorksheet, row, StartColumnIndex + 1),
                CharSteps = []
            };
        }

        private CharStep GetCharStep(ExcelWorksheet excelWorksheet, int row)
        {
            int index = 4;
            string setupName = GetCellText(excelWorksheet, row, StartColumnIndex);
            string stepName = GetCellText(excelWorksheet, row, index++);

            var charStep = new CharStep(setupName, stepName)
            {
                VoltageType = "",
                SetupName = setupName,
                Mode = GetCellText(excelWorksheet, row, index++),
                ParameterType = GetCellText(excelWorksheet, row, index++),
                ParameterName = GetCellText(excelWorksheet, row, index++),
                RangeCalcField = GetCellText(excelWorksheet, row, index++),
                RangeFrom = GetCellText(excelWorksheet, row, index++),
                RangeTo = GetCellText(excelWorksheet, row, index++),
                RangeSteps = GetCellText(excelWorksheet, row, index++),
                RangeStepSize = GetCellText(excelWorksheet, row, index++),
                PerformTest = GetCellText(excelWorksheet, row, index++),
                TestLimitLow = GetCellText(excelWorksheet, row, index++),
                TestLimitHigh = GetCellText(excelWorksheet, row, index++),
                AlgorithmName = GetCellText(excelWorksheet, row, index++),
                AlgorithmArguments = GetCellText(excelWorksheet, row, index++),
                AlgorithmResultsCheck = GetCellText(excelWorksheet, row, index++),
                AlgorithmTransition = GetCellText(excelWorksheet, row, index++),
                ApplyToPins = GetCellText(excelWorksheet, row, index++),
                ApplyToPinExecMode = GetCellText(excelWorksheet, row, index++),
                ApplyToTimeSets = GetCellText(excelWorksheet, row, index++),
                DeviceMarginContexts = GetCellText(excelWorksheet, row, index++),
                DeviceMarginPatterns = GetCellText(excelWorksheet, row, index++),
                DeviceMarginInstances = GetCellText(excelWorksheet, row, index++),
                AdjustBackoff = GetCellText(excelWorksheet, row, index++),
                AdjustSpecName = GetCellText(excelWorksheet, row, index++),
                AdjustFromSetup = GetCellText(excelWorksheet, row, index++),
                AxisExecutionOrder = GetCellText(excelWorksheet, row, index++),
                Function = GetCellText(excelWorksheet, row, index++),
                Arguments = GetCellText(excelWorksheet, row, index++),
                PreSetup = GetCellText(excelWorksheet, row, index++),
                PreSetupArguments = GetCellText(excelWorksheet, row, index++),
                PreStep = GetCellText(excelWorksheet, row, index++),
                PreStepArguments = GetCellText(excelWorksheet, row, index++),
                PrePoint = GetCellText(excelWorksheet, row, index++),
                PrePointArguments = GetCellText(excelWorksheet, row, index++),
                PostPoint = GetCellText(excelWorksheet, row, index++),
                PostPointArguments = GetCellText(excelWorksheet, row, index++),
                PostStep = GetCellText(excelWorksheet, row, index++),
                PostStepArguments = GetCellText(excelWorksheet, row, index++),
                PostSetup = GetCellText(excelWorksheet, row, index++),
                PostSetupArguments = GetCellText(excelWorksheet, row, index++),
                OutputFormat = GetCellText(excelWorksheet, row, index++),
                OutputTextFile = GetCellText(excelWorksheet, row, index++),
                OutputSheet = GetCellText(excelWorksheet, row, index++),
                OutputSuspendDatalog = GetCellText(excelWorksheet, row, index++),
                OutputDestinationsTextFile = GetCellText(excelWorksheet, row, index++),
                OutputDestinationsSheet = GetCellText(excelWorksheet, row, index++),
                OutputDestinationsDatalog = GetCellText(excelWorksheet, row, index++),
                OutputDestinationsImmediateWin = GetCellText(excelWorksheet, row, index++),
                OutputDestinationsOutputWin = GetCellText(excelWorksheet, row, index++),
                Comment = GetCellText(excelWorksheet, row, index++)
            };

            return charStep;
        }
    }
}
