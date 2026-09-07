using System.IO;
using System.Linq;
using System.Text;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class CharSheet : IgxlSheet<CharSetup>
    {
        public override string SheetType => "DTCharacterizationSheet";
        public override string IgxlSheetName => IgxlSheetNames.Characterization;

        public CharSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public CharSheet(string sheetName) : base(sheetName)
        {
        }

        public override void AddRow(CharSetup charSetup)
        {
            base.AddRow(charSetup);
        }

        public override void Write(string fileName, string version = "")
        {
            Write2P5(fileName, version);
        }

        protected override void WriteSheet(string fileName, string version, SheetInfo sheetInfo)
        {
            if (Rows.Count == 0)
            {
                return;
            }

            using var sw = new StreamWriter(fileName, false);
            var sb = new StringBuilder();
            int maxCount = GetMaxCount(sheetInfo);
            int setupNameIndex = GetIndexFrom(sheetInfo, "Setup Name");
            int testMethodIndex = GetIndexFrom(sheetInfo, "Test Method");
            int stepNameIndex = GetIndexFrom(sheetInfo, "Step Name");
            int modeIndex = GetIndexFrom(sheetInfo, "Mode");
            int parameterTypeIndex = GetIndexFrom(sheetInfo, "Parameter", "Type");
            int parameterNameIndex = GetIndexFrom(sheetInfo, "Parameter", "Name");
            int rangeCalcFieldIndex = GetIndexFrom(sheetInfo, "Range", "Calc Field");
            int rangeFromIndex = GetIndexFrom(sheetInfo, "Range", "From");
            int rangeToIndex = GetIndexFrom(sheetInfo, "Range", "To");
            int rangeStepsIndex = GetIndexFrom(sheetInfo, "Range", "Steps");
            int rangeStepSizeIndex = GetIndexFrom(sheetInfo, "Range", "Step Size");
            int performTestIndex = GetIndexFrom(sheetInfo, "Perform Test");
            int testLimitsLowIndex = GetIndexFrom(sheetInfo, "Test Limits", "Low");
            int testLimitsHighIndex = GetIndexFrom(sheetInfo, "Test Limits", "High");
            int algorithmNameIndex = GetIndexFrom(sheetInfo, "Algorithm", "Name");
            int algorithmArgumentsIndex = GetIndexFrom(sheetInfo, "Algorithm", "Arguments");
            int algorithmResultsCheckIndex = GetIndexFrom(sheetInfo, "Algorithm", "Results Check");
            int algorithmTransitionIndex = GetIndexFrom(sheetInfo, "Algorithm", "Transition");
            int applyToPinsIndex = GetIndexFrom(sheetInfo, "Apply To", "Pins");
            int applyToPinExecModeIndex = GetIndexFrom(sheetInfo, "Apply To", "Pin Exec Mode");
            int applyToTimeSetsIndex = GetIndexFrom(sheetInfo, "Apply To", "Time Sets");
            int deviceMarginContextsIndex = GetIndexFrom(sheetInfo, "Device Margin", "Contexts");
            int deviceMarginPatternsIndex = GetIndexFrom(sheetInfo, "Device Margin", "Patterns");
            int deviceMarginInstancesIndex = GetIndexFrom(sheetInfo, "Device Margin", "Instances");
            int adjustBackoffIndex = GetIndexFrom(sheetInfo, "Adjust", "Backoff");
            int adjustSpecNameIndex = GetIndexFrom(sheetInfo, "Adjust", "Spec Name");
            int adjustFromSetupIndex = GetIndexFrom(sheetInfo, "Adjust", "From Setup");
            int axisExecutionOrderIndex = GetIndexFrom(sheetInfo, "Shmoo", "Axis Execution Order");
            int functionIndex = GetIndexFrom(sheetInfo, "Function");
            int argumentsIndex = GetIndexFrom(sheetInfo, "Arguments");
            int interposeFunctionsIndex = GetIndexFrom(sheetInfo, "Interpose Functions");
            int preSetupIndex = GetIndexFrom(sheetInfo, "Interpose Functions", "Pre Setup");
            int preStepIndex = GetIndexFrom(sheetInfo, "Interpose Functions", "Pre Step");
            int prePointIndex = GetIndexFrom(sheetInfo, "Interpose Functions", "Pre Point");
            int postPointIndex = GetIndexFrom(sheetInfo, "Interpose Functions", "Post Point");
            int postStepIndex = GetIndexFrom(sheetInfo, "Interpose Functions", "Post Step");
            int postSetupIndex = GetIndexFrom(sheetInfo, "Interpose Functions", "Post Setup");
            int outputFormatIndex = GetIndexFrom(sheetInfo, "Output", "Format");
            int outputTextFileIndex = GetIndexFrom(sheetInfo, "Output", "Text File");
            int outputSheetIndex = GetIndexFrom(sheetInfo, "Output", "Sheet");
            int outputSuspendDatalogIndex = GetIndexFrom(sheetInfo, "Output", "Suspend Datalog");
            int outputDestinationsTextFileIndex = GetIndexFrom(sheetInfo, "Output Destinations", "Text File");
            int outputDestinationsSheetIndex = GetIndexFrom(sheetInfo, "Output Destinations", "Sheet");
            int outputDestinationsDatalogIndex = GetIndexFrom(sheetInfo, "Output Destinations", "Datalog");
            int outputDestinationsImmediateWinIndex = GetIndexFrom(sheetInfo, "Output Destinations", "Immediate Win");
            int outputDestinationsOutputWinIndex = GetIndexFrom(sheetInfo, "Output Destinations", "Output Win");
            int commentIndex = GetIndexFrom(sheetInfo, "Comment");

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                SetColumns(sheetInfo, i, arr);

                WriteHeader(arr, sb);
            }
            #endregion

            #region data
            for (int index = 0; index < Rows.Count; index++)
            {
                CharSetup row = Rows[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                foreach (CharStep charStep in row.CharSteps)
                {
                    if (!string.IsNullOrEmpty(row.SetupName))
                    {
                        arr[0] = row.ColumnA ?? "";
                        arr[setupNameIndex] = row.SetupName;
                        arr[testMethodIndex] = row.TestMethod;
                        arr[stepNameIndex] = charStep.StepName;
                        arr[modeIndex] = charStep.Mode;
                        arr[parameterTypeIndex] = charStep.ParameterType;
                        arr[parameterNameIndex] = charStep.ParameterName;
                        arr[rangeCalcFieldIndex] = charStep.RangeCalcField;
                        arr[rangeFromIndex] = charStep.RangeFrom;
                        arr[rangeToIndex] = charStep.RangeTo;
                        arr[rangeStepsIndex] = charStep.RangeSteps;
                        arr[rangeStepSizeIndex] = charStep.RangeStepSize;
                        arr[performTestIndex] = charStep.PerformTest;
                        arr[testLimitsLowIndex] = charStep.TestLimitLow;
                        arr[testLimitsHighIndex] = charStep.TestLimitHigh;
                        arr[algorithmNameIndex] = charStep.AlgorithmName;
                        arr[algorithmArgumentsIndex] = charStep.AlgorithmArguments;
                        arr[algorithmResultsCheckIndex] = charStep.AlgorithmResultsCheck;
                        arr[algorithmTransitionIndex] = charStep.AlgorithmTransition;
                        arr[applyToPinsIndex] = charStep.ApplyToPins;
                        arr[applyToPinExecModeIndex] = charStep.ApplyToPinExecMode;
                        arr[applyToTimeSetsIndex] = charStep.ApplyToTimeSets;
                        arr[deviceMarginContextsIndex] = charStep.DeviceMarginContexts;
                        arr[deviceMarginPatternsIndex] = charStep.DeviceMarginPatterns;
                        arr[deviceMarginInstancesIndex] = charStep.DeviceMarginInstances;
                        arr[adjustBackoffIndex] = charStep.AdjustBackoff;
                        arr[adjustSpecNameIndex] = charStep.AdjustSpecName;
                        arr[adjustFromSetupIndex] = charStep.AdjustFromSetup;
                        if (axisExecutionOrderIndex != -1)
                        {
                            arr[axisExecutionOrderIndex] = charStep.AxisExecutionOrder;
                        }

                        arr[functionIndex] = charStep.Function;
                        arr[argumentsIndex] = charStep.Arguments;
                        arr[interposeFunctionsIndex] = charStep.InterposeFunctions;
                        arr[preSetupIndex] = charStep.PreSetup;
                        arr[preSetupIndex + 1] = charStep.PreSetupArguments;
                        arr[preStepIndex] = charStep.PreStep;
                        arr[preStepIndex + 1] = charStep.PreStepArguments;
                        arr[prePointIndex] = charStep.PrePoint;
                        arr[prePointIndex + 1] = charStep.PrePointArguments;
                        arr[postPointIndex] = charStep.PostPoint;
                        arr[postPointIndex + 1] = charStep.PostPointArguments;
                        arr[postStepIndex] = charStep.PostStep;
                        arr[postStepIndex + 1] = charStep.PostStepArguments;
                        arr[postSetupIndex] = charStep.PostSetup;
                        arr[postSetupIndex + 1] = charStep.PostSetupArguments;
                        arr[outputFormatIndex] = charStep.OutputFormat;
                        arr[outputTextFileIndex] = charStep.OutputTextFile;
                        arr[outputSheetIndex] = charStep.OutputSheet;
                        if (outputSuspendDatalogIndex != -1)
                        {
                            arr[outputSuspendDatalogIndex] = charStep.OutputSuspendDatalog;
                        }

                        arr[outputDestinationsTextFileIndex] = charStep.OutputDestinationsTextFile;
                        arr[outputDestinationsSheetIndex] = charStep.OutputDestinationsSheet;
                        arr[outputDestinationsDatalogIndex] = charStep.OutputDestinationsDatalog;
                        arr[outputDestinationsImmediateWinIndex] = charStep.OutputDestinationsImmediateWin;
                        arr[outputDestinationsOutputWinIndex] = charStep.OutputDestinationsOutputWin;
                        arr[commentIndex] = charStep.Comment;
                    }
                    else
                    {
                        arr = ["\t"];
                    }
                    sb.AppendLine(string.Join("\t", arr));
                }
            }
            #endregion

            sw.Write(sb.ToString());
        }
    }
}
