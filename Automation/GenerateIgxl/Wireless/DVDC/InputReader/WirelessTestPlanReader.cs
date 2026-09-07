using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.GenerateIgxl.Wireless.DVDC.InputObject;
using Automation.GenerateIgxl.Wireless.DVDC.WirelessConst;

using CommonLib.Extension;

using LogLib.Utility;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.Wireless.DVDC.InputReader
{
    public class WirelessTestPlanReader : TestPlanReader
    {
        public const string ConHeaderOtpName = WirelessConstData.TrimFuseName;
        public const string ConHeaderTrimTarget = WirelessConstData.TrimTarget;
        public const string ConHeaderTrimMeas = "TrimMeas";
        public const string ConHeaderTrimTargetCalc = WirelessConstData.TrimCalcEqn;
        public const string ConHeaderTrimType = WirelessConstData.TrimType;
        //public const string ConHeaderRFInstrumentSetup = WirelessConstData.RFInstrumentSetup;
        public const string ConHeaderRfCalc = WirelessConstData.RfCalc;
        public const string ConHeaderRfInterPose = WirelessConstData.RfInterpose;
        public const string ConHeaderReadBackSequence = WirelessConstData.FwReadBackSequence;
        public const string ConHeaderTrimDataWidth = "Trim_DataWidth";
        public const string ConHeaderTrimCodeType = @"Trim\s*Code\s*Type";
        public const string ConHeaderTrimCodeStart = @"Trim\s*Code\s*Start";
        public const string ConHeaderTrimCodeStop = @"Trim\s*Code\s*Stop";
        public const string ConHeaderTrimCodeStopBit = @"Trim\s*Code\s*Stop\s*Bit";
        public const string ConHeaderTrimStartBit = @"Trim\s*Start\s*Bit";
        public const string ConHeaderTrimStart = @"Trim\s*Start";
        public const string ConHeaderTrimFormat = @"Trim\s*Format";
        public const string ConHeaderTrimTable = @"Trim\s*Table";
        public const string ConHeaderTrimDefaultCode = @"Trim\s*Default\s*Code";
        public const string ConHeaderTrimCodeStepSize = @"Trim\s*Code\s*StepSize";
        public const string ConHeaderRfInstrumentSetup = @"RF\s*Instrument\s*Setup";
        public const string ConHeaderAnalogInstrumentSetup = @"Analog\s*Instrument\s*Setup";
        public const string ConHeaderBbInstrumentSetup = @"BB\s*Instrument\s*Setup";
        public const string ConHeaderBbDeviceSetup = @"BB\s*Device\s*Setup";
        public const string ConHeaderBbCalc = @"BB\s*Calc";

        internal int _oTpName;
        internal int _trimTarget;
        internal int _trimMeas;
        internal int _trimCalcEqn;
        internal int _trimType;
        internal int _rfSetup;
        internal int _rfCalc;
        internal int _rfInterpose;
        internal int _bbSetup;
        internal int _bbDeviceSetup;
        internal int _bbCalc;

        private readonly int _startColumnNumber = 1;
        private readonly int _startRowNumber = 1;
        internal int _endColNumber = 1;

        public override TestPlanSheet ReadSheet(ExcelWorksheet sheet, bool ignoreDescription = true)
        {
            var planSheet = new TestPlanSheet();
            ExcelWorksheet = sheet;
            HeaderOrder = EpplusExtensions.GetHeaderOrder(sheet);
            StartRow = 2;
            EndRow = sheet.Dimension.End.Row;
            planSheet.SheetName = sheet.Name;
            WirelessPatternRow patRow = null;

            //CheckUnrecognizeHeader
            List<string> knownHeaders =
            [
                .. HardIpPattern.KnownHeaders,
                .. new List<string>
                {
                    ConHeaderOtpName,
                    ConHeaderTrimTarget,
                    ConHeaderTrimMeas,
                    ConHeaderTrimTargetCalc,
                    ConHeaderTrimType,
                    ConHeaderRfInstrumentSetup,
                    ConHeaderRfCalc,
                    ConHeaderReadBackSequence,
                    ConHeaderTrimDataWidth,
                    ConHeaderTrimCodeType,
                    ConHeaderTrimCodeStart,
                    ConHeaderTrimCodeStop,
                    ConHeaderTrimCodeStopBit,
                    ConHeaderTrimStartBit,
                    ConHeaderTrimStart,
                    ConHeaderTrimFormat,
                    ConHeaderTrimTable,
                    ConHeaderTrimDefaultCode,
                    ConHeaderTrimCodeStepSize,
                    ConHeaderRfInstrumentSetup,
                    ConHeaderAnalogInstrumentSetup
                },
            ];
            CheckUnrecognizeHeader(knownHeaders);

            //Initial herder index
            GetHeaderIndexs();
            GetDimensions();
            GetWirelessHeaderIndexs();
            SetHardipSheetsHeaderIdx(planSheet);

            //find the forceCondition used by all the patterns in the sheet
            planSheet.ForceIndex = ForceIndex;
            planSheet.ForceStr = GetSheetForceConditions();

            #region read testPlan patterns
            for (int j = StartRow; j <= EndRow; j++)
            {
                string patternName = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, j, PatternIndex)).Trim(',').Trim(';');
                if (patternName != "")
                {
                    patRow = new WirelessPatternRow(patternName);
                    ReadPatternRow(sheet, j, patRow, patternName);
                    planSheet.PatternRows.Add(patRow);
                }
                else
                {
                    CheckRunPattern(sheet, j);
                    string measStr = EpplusExtensions.GetCellValue(sheet, j, MeasIndex);
                    string testName = EpplusExtensions.GetCellValue(sheet, j, TestNameIndex);
                    //If meas column not blank, but not belong to any pattern, flag error
                    if (patRow == null && !string.IsNullOrEmpty(measStr))
                    {
                        continue;
                    }
                    //Will ignore the rows that Pattern column, Force Condition column, RegisterAssignment and Meas column are blank
                    if (string.IsNullOrEmpty(measStr) && string.IsNullOrEmpty(testName))
                    {
                        string registerAssignment = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, j, RegisterIndex));
                        //read RegisterAssignment
                        patRow.RegisterAssignment = UpdatePatternItem(patRow.RegisterAssignment, registerAssignment);

                        string forceStr = EpplusExtensions.GetCellValue(sheet, j, ForceIndex);
                        //Read post pattern force condition
                        patRow.PostPatForceCondition = UpdatePatternItem(patRow.PostPatForceCondition, forceStr);
                        continue;
                    }

                    ReadNonPatternRowNew(sheet, ref j, patRow);
                }
            }
            #endregion

            //DividePatternRow(planSheet.PatternRows.OfType<PatternRow>().ToList());

            return planSheet;
        }

        protected override void ReadNonPatternRowNew(ExcelWorksheet sheet, ref int rowindex, PatternRow patRow)
        {
            try
            {
                //find whether the force condition column is merged or not
                bool isMerged = sheet.MergedCells[rowindex, ForceIndex] != null;

                var patChildRow = new PatSubChildRow { IsMerged = isMerged };
                if (!isMerged)
                {
                    patChildRow.TpRows.Add(ReadTestPlanRow(rowindex));
                }
                else
                {
                    for (int mergedRow = rowindex;
                        mergedRow <= new ExcelAddress(sheet.MergedCells[rowindex, ForceIndex]).End.Row;
                        mergedRow++)
                    {
                        patChildRow.TpRows.Add(ReadTestPlanRow(mergedRow));
                    }
                    rowindex = new ExcelAddress(sheet.MergedCells[rowindex, ForceIndex]).End.Row;
                }
                patRow.PatChildRows.Add(patChildRow);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        protected override TestPlanRow ReadTestPlanRow(int rowNum)
        {
            TestPlanRow testPlanRow = base.ReadTestPlanRow(rowNum);

            if (_rfCalc != 0)
            {
                testPlanRow.InterposeFunc = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _rfCalc));
            }

            if (_rfInterpose != 0)
            {
                testPlanRow.RfInterpose = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _rfInterpose));
            }

            if (_rfSetup != 0)
            {
                if (ExcelWorksheet.MergedCells[rowNum, _rfSetup] != null)
                {
                    testPlanRow.MergeRowNumForMeas = new ExcelAddress(ExcelWorksheet.MergedCells[rowNum, _rfSetup]).Start.Row;
                    testPlanRow.RfIntrumentSetup =
                        EpplusExtensions.GetCellValue(ExcelWorksheet, testPlanRow.MergeRowNumForMeas, _rfSetup)
                            .Replace("\n", "")
                            .Replace("\t", "");
                }
                else
                {
                    testPlanRow.RfIntrumentSetup = string.IsNullOrEmpty(testPlanRow.RfIntrumentSetup) ?
                          SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _rfSetup)) :
                        testPlanRow.RfIntrumentSetup + ";" + SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _rfSetup));
                }
            }
            if (_bbSetup != 0)
            {
                testPlanRow.RfIntrumentSetup = string.IsNullOrEmpty(testPlanRow.RfIntrumentSetup) ?
                      SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _bbSetup)) :
                     testPlanRow.RfIntrumentSetup + ";" + SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _bbSetup));
            }
            if (_bbDeviceSetup != 0)
            {
                testPlanRow.RfIntrumentSetup = string.IsNullOrEmpty(testPlanRow.RfIntrumentSetup) ?
                      SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _bbDeviceSetup)) :
                 testPlanRow.RfIntrumentSetup + ";" + SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _bbDeviceSetup));
            }
            if (_bbCalc != 0)
            {
                testPlanRow.RfIntrumentSetup = string.IsNullOrEmpty(testPlanRow.RfIntrumentSetup) ?
                      SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _bbCalc)) :
                  testPlanRow.RfIntrumentSetup + ";" + SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, _bbCalc));
            }

            //20181214 add by Kimi
            if (testPlanRow.Meas == "" && testPlanRow.TestName != "" && testPlanRow.ForceCondition == "" && testPlanRow.MiscInfo == "" && testPlanRow.RegisterAssignment == "")
            {
                testPlanRow.Meas = MeasType.MeasLimit;
            }

            return testPlanRow;
        }

        internal void GetWirelessHeaderIndexs()
        {
            for (int i = _startColumnNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusExtensions.GetCellValue(ExcelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderOtpName, StringComparison.OrdinalIgnoreCase) || lStrHeader.Equals("TrimRegName", StringComparison.OrdinalIgnoreCase))
                {
                    _oTpName = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderTrimTarget, StringComparison.OrdinalIgnoreCase) || lStrHeader.Equals(ConHeaderTrimTarget.Replace("_", ""), StringComparison.OrdinalIgnoreCase))
                {
                    _trimTarget = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderTrimMeas, StringComparison.OrdinalIgnoreCase))
                {
                    _trimMeas = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderTrimTargetCalc, StringComparison.OrdinalIgnoreCase) || lStrHeader.Equals("Trim_Calc_Eqn", StringComparison.OrdinalIgnoreCase))
                {
                    _trimCalcEqn = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderTrimType, StringComparison.OrdinalIgnoreCase) || lStrHeader.Equals(ConHeaderTrimType.Replace("_", ""), StringComparison.OrdinalIgnoreCase))
                {
                    _trimType = i;
                    continue;
                }
                if (Regex.IsMatch(lStrHeader.Replace("_", ""), ConHeaderRfInstrumentSetup, RegexOptions.IgnoreCase))
                {
                    _rfSetup = i;
                    continue;
                }
                if (Regex.IsMatch(lStrHeader.Replace("_", ""), ConHeaderRfCalc, RegexOptions.IgnoreCase))
                {
                    _rfCalc = i;
                    continue;
                }
                if (Regex.IsMatch(lStrHeader.Replace("_", ""), ConHeaderRfInterPose, RegexOptions.IgnoreCase))
                {
                    _rfInterpose = i;
                    continue;
                }
                if (Regex.IsMatch(lStrHeader.Replace("_", ""), ConHeaderBbInstrumentSetup, RegexOptions.IgnoreCase))
                {
                    _bbSetup = i;
                    continue;
                }
                if (Regex.IsMatch(lStrHeader.Replace("_", ""), ConHeaderBbDeviceSetup, RegexOptions.IgnoreCase))
                {
                    _bbDeviceSetup = i;
                    continue;
                }
                if (Regex.IsMatch(lStrHeader.Replace("_", ""), ConHeaderBbCalc, RegexOptions.IgnoreCase))
                {
                    _bbCalc = i;
                    continue;
                }

            }
        }

        internal bool GetDimensions()
        {
            if (ExcelWorksheet.Dimension != null)
            {
                _endColNumber = ExcelWorksheet.Dimension.End.Column;
                return true;
            }
            return false;
        }

        protected void ReadPatternRow(ExcelWorksheet sheet, int rowIndex, WirelessPatternRow patRow, string patternName)
        {
            base.ReadPatternRow(sheet, rowIndex, patRow, patternName);
            if (_rfInterpose != 0)
            {
                patRow.RfInterPose = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowIndex, _rfInterpose));
            }

            patRow.WirelessData.RegisterAssignment = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, RegisterIndex));
            patRow.WirelessData.TrimFuseName = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, _oTpName));
            patRow.WirelessData.TrimTarget = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, _trimTarget));
            patRow.WirelessData.TrimMeas = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, _trimMeas));
            patRow.WirelessData.TrimCalcEqn = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, _trimCalcEqn));
            patRow.WirelessData.TrimType = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, _trimType));
        }
    }
}
