using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;

namespace CommonLib.ErrorReport.Interop
{
    internal class ReportHardIp : ErrorReportInterop
    {
        public ReportHardIp(List<Error> errorList)
            : base(errorList)
        {
        }

        public override void Write(Workbook workbook)
        {
            if (ErrorList.Count == 0)
            {
                return;
            }

            WriteReport(workbook);
            //WriteCorrectItems(workbook);
        }

        //private void WriteCorrectItems(Workbook planWorkbook)
        //{
        //    try
        //    {
        //        Dictionary<string, int> linkIndex = new Dictionary<string, int>();

        //        if (HardIpDataMain.PatInfoData != null && HardIpDataMain.PatInfoData.PatInfoErrorList.Count > 0)
        //        {
        //            #region Delete sheet if alreadey exists
        //            foreach (Worksheet sheet in planWorkbook.Worksheets)
        //            {
        //                if (sheet.Name == "CorrectItems")
        //                {
        //                    planWorkbook.Application.DisplayAlerts = false;
        //                    sheet.Delete();
        //                    break;
        //                }
        //            }
        //            #endregion

        //            #region Generate Correct items sheet
        //            Worksheet wSheet_correct = planWorkbook.Worksheets.Add(planWorkbook.Worksheets[3], Type.Missing, Type.Missing, Type.Missing);
        //            wSheet_correct.Name = "CorrectItems";
        //            int rowIndex = 2;
        //            wSheet_correct.Cells[1, 1].Value = "ATE Test Number";
        //            wSheet_correct.Cells[1, 2].Value = "Test Item";
        //            wSheet_correct.Cells[1, 3].Value = "Step";
        //            wSheet_correct.Cells[1, 4].Value = "Description";
        //            wSheet_correct.Cells[1, 5].Value = "PatternData";
        //            wSheet_correct.Cells[1, 6].Value = "Meas seq";
        //            wSheet_correct.Cells[1, 7].Value = "Pass Fail Only";
        //            wSheet_correct.Cells[1, 8].Value = "Meas";
        //            wSheet_correct.Cells[1, 9].Value = "Send Bit";
        //            wSheet_correct.Cells[1, 10].Value = "Send Bit Str";
        //            wSheet_correct.Cells[1, 11].Value = "Cap Bit";
        //            wSheet_correct.Cells[1, 12].Value = "Cap Bit Str";
        //            wSheet_correct.get_Range((Range)wSheet_correct.Cells[1, 1], (Range)wSheet_correct.Cells[1, 12]).Font.Bold = true;

        //            Hashtable describe = new Hashtable();
        //            describe.Add("F", "frequency");
        //            describe.Add("V", "voltage");
        //            describe.Add("I", "current");
        //            foreach (HardIpReference patInfo in HardIpDataMain.PatInfoData.PatInfoErrorList)
        //            {
        //                if (!linkIndex.ContainsKey(patInfo.Payload))
        //                    linkIndex.Add(patInfo.Payload, rowIndex);
        //                else
        //                {
        //                    continue;
        //                    //MessageWriter.WriteMessage("Duplicate pattern " + patInfo.Payload, MessageLevel.Error);
        //                }
        //                //HardIpPattern errorPattern = SearchInfo.GetPattern(patInfo.Payload);
        //                wSheet_correct.Cells[rowIndex, 4].Value = "Run the pattern provided";
        //                wSheet_correct.Cells[rowIndex, 5].Value = patInfo.Payload;
        //                wSheet_correct.Cells[rowIndex, 6].Value = patInfo.MeasSeqStr;
        //                //wSheet_correct.Cells[rowIndex, 7].Value = errorPattern.PassOrFail;
        //                wSheet_correct.Cells[rowIndex, 9].Value = patInfo.SendBit;
        //                wSheet_correct.Cells[rowIndex, 10].Value = patInfo.SendBitStr;
        //                wSheet_correct.Cells[rowIndex, 11].Value = patInfo.CapBit;
        //                wSheet_correct.Cells[rowIndex, 12].Value = patInfo.CapBitStr;
        //                rowIndex++;
        //                foreach (HardIpSeqInfo info in patInfo.SeqInfo)
        //                {
        //                    List<string> pinList = info.PinList;
        //                    pinList.Sort();
        //                    foreach (string pin in pinList)
        //                    {
        //                        wSheet_correct.Cells[rowIndex, 4].Value = "Measure the " + describe[info.SeqName] + " for " + pin;
        //                        wSheet_correct.Cells[rowIndex, 8].Value = "Meas" + info.SeqName + " Pin = " + pin;
        //                        rowIndex++;
        //                    }
        //                }
        //            }

        //            wSheet_correct.Columns[1].AutoFit();
        //            wSheet_correct.Columns[2].AutoFit();
        //            wSheet_correct.Columns[3].AutoFit();
        //            wSheet_correct.Columns[4].AutoFit();
        //            wSheet_correct.Columns[5].AutoFit();
        //            wSheet_correct.Columns[6].AutoFit();
        //            wSheet_correct.Columns[7].AutoFit();
        //            wSheet_correct.Columns[8].AutoFit();
        //            wSheet_correct.Columns[9].AutoFit();
        //            wSheet_correct.Columns[10].AutoFit();
        //            wSheet_correct.Columns[11].AutoFit();
        //            wSheet_correct.Columns[12].AutoFit();
        //            #endregion

        //            #region Add link for correct items to error location
        //            foreach (ErrorRow error in ErrorList)
        //            {
        //                if (error.ErrorType.Equals(HardIpErrorType.WrongTotalMeasCount) ||
        //                    error.ErrorType.Equals(HardIpErrorType.WrongMeasCountInPatInfo) ||
        //                    error.ErrorType.Equals(HardIpErrorType.WrongRegisterAssignment) ||
        //                    error.ErrorType.Equals(HardIpErrorType.WrongMeasPinInPatInfo) ||
        //                    error.ErrorType.Equals(HardIpErrorType.MissingPinName) ||
        //                    error.ErrorType.Equals(HardIpErrorType.WrongMeasSequence) ||
        //                    error.ErrorType.Equals(HardIpErrorType.WrongLimit))
        //                {
        //                    if (error.Comments.Count == 0)
        //                    {
        //                        error.Comments.Add("NA");
        //                    }
        //                    if (linkIndex.ContainsKey(error.Comments[0]))
        //                    {
        //                        HardIpPattern errorPattern = error.ErrorPattern;
        //                        if (errorPattern == null || errorPattern.Pattern.GetLastPayload() == "")
        //                            continue;
        //                        Worksheet errorSheet = planWorkbook.Worksheets[error.SheetName];
        //                        Range correctRange = wSheet_correct.Range[(Range)wSheet_correct.Cells[linkIndex[error.Comments[0]], 1], (Range)wSheet_correct.Cells[linkIndex[error.Comments[0]], 12]];
        //                        correctRange.Interior.Pattern = XlPattern.xlPatternSolid;
        //                        correctRange.Interior.Color = Color.Lime;
        //                        Range linkRange = errorSheet.Cells[errorPattern.RowNum, errorPattern.ColumnNum];
        //                        errorSheet.Hyperlinks.Add(linkRange, "#" + "'" + wSheet_correct.Name + "'" + "!" + correctRange.Address, Type.Missing, Type.Missing, Type.Missing);

        //                        Range errorRange = linkRange;
        //                        Range linkRange_correct = wSheet_correct.Cells[linkIndex[error.Comments[0]], 5];
        //                        wSheet_correct.Hyperlinks.Add(linkRange_correct, "#" + "'" + errorSheet.Name + "'" + "!" + errorRange.Address, Type.Missing, Type.Missing, Type.Missing);
        //                    }
        //                }
        //            }
        //            planWorkbook.Worksheets["SummaryReport"].Activate();
        //            #endregion
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        throw new Exception("Write HardIP error report failed. " + e.StackTrace);
        //    }
        //}
    }
}
