using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.common.ReaderWriter.Writer;

using CommonLib.Extension;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Cautogen.AutoCZ.CharPreProcessor.ErrorReport
{
    internal class CorrectItemSheet : IExcelSheetWriter
    {
        public void Write(ExcelWorkbook wb)
        {
            if (UtilityMain.UtilityData.PatInfoErrorList.Count == 0)
            {
                return;
            }

            // delete sheet if alreadey exists
            foreach (ExcelWorksheet sheet in wb.Worksheets.Where(sheet => sheet.Name == "CorrectItems"))
            {
                wb.Worksheets.Delete(sheet);
                break;
            }

            var linkIndex = new Dictionary<string, int>();

            #region Generate Correct items sheet

            ExcelWorksheet wSheetCorrect = wb.Worksheets.Add("CorrectItems");
            wb.Worksheets.MoveAfter("CorrectItems", "ErrorReport");
            int rowIndex = 2;
            wSheetCorrect.Cells[1, 1].Value = "ATE Test Number";
            wSheetCorrect.Cells[1, 2].Value = "Test Item";
            wSheetCorrect.Cells[1, 3].Value = "Step";
            wSheetCorrect.Cells[1, 4].Value = "Description";
            wSheetCorrect.Cells[1, 5].Value = "PatternData";
            wSheetCorrect.Cells[1, 6].Value = "Meas seq";
            wSheetCorrect.Cells[1, 7].Value = "Pass Fail Only";
            wSheetCorrect.Cells[1, 8].Value = "Meas";
            wSheetCorrect.Cells[1, 9].Value = "Send Bit";
            wSheetCorrect.Cells[1, 10].Value = "Send Bit Str";
            wSheetCorrect.Cells[1, 11].Value = "Cap Bit";
            wSheetCorrect.Cells[1, 12].Value = "Cap Bit Str";
            wSheetCorrect.Cells[1, 1, 1, 12].Style.Font.Bold = true;

            var describe = new Hashtable { { "F", "frequency" }, { "V", "voltage" }, { "I", "current" } };

            foreach (HardIpReference patInfo in UtilityMain.UtilityData.PatInfoErrorList)
            {
                if (!linkIndex.ContainsKey(patInfo.Payload))
                {
                    linkIndex.Add(patInfo.Payload, rowIndex);
                }
                else
                {
                    continue;
                }

                wSheetCorrect.Cells[rowIndex, 4].Value = "Run the pattern provided";
                wSheetCorrect.Cells[rowIndex, 5].Value = patInfo.Payload;
                wSheetCorrect.Cells[rowIndex, 6].Value = patInfo.MeasSeqStr;
                wSheetCorrect.Cells[rowIndex, 9].Value = patInfo.SendBit;
                wSheetCorrect.Cells[rowIndex, 10].Value = patInfo.SendBitStr;
                wSheetCorrect.Cells[rowIndex, 11].Value = patInfo.CapBit;
                wSheetCorrect.Cells[rowIndex, 12].Value = patInfo.CapBitStr;
                rowIndex++;
                foreach (HardIpSeqInfo info in patInfo.SeqInfo)
                {
                    List<string> pinList = info.PinList.Split(',').ToList();
                    pinList.Sort();
                    foreach (string pin in pinList)
                    {
                        wSheetCorrect.Cells[rowIndex, 4].Value = "Measure the " + describe[info.SeqName] + " for " + pin;
                        wSheetCorrect.Cells[rowIndex, 8].Value = "Meas" + info.SeqName + " Pin = " + pin;
                        rowIndex++;
                    }
                }
            }

            wSheetCorrect.Column(1).TryAutoFit();
            wSheetCorrect.Column(2).TryAutoFit();
            wSheetCorrect.Column(3).TryAutoFit();
            wSheetCorrect.Column(4).TryAutoFit();
            wSheetCorrect.Column(5).TryAutoFit();
            wSheetCorrect.Column(6).TryAutoFit();
            wSheetCorrect.Column(7).TryAutoFit();
            wSheetCorrect.Column(8).TryAutoFit();
            wSheetCorrect.Column(9).TryAutoFit();
            wSheetCorrect.Column(10).TryAutoFit();
            wSheetCorrect.Column(11).TryAutoFit();
            wSheetCorrect.Column(12).TryAutoFit();

            #endregion

            #region Add link for correct items to error location

            foreach (ErrorType errorType in ErrorManager.ErrorListDict.Keys)
            {
                foreach (ErrorMessage error in ErrorManager.ErrorListDict[errorType])
                {
                    if (errorType != ErrorType.WrongMeasCount && errorType != ErrorType.WrongMeasPin)
                    {
                        continue;
                    }

                    if (!linkIndex.ContainsKey(error.CommentsList[0]))
                    {
                        continue;
                    }

                    InputDataBase.CharPlan.Characterization charItem = UtilityMain.UtilityFunction.GetCharItem(error.CommentsList[0]);
                    if (charItem.Payload1 == "")
                    {
                        continue;
                    }

                    ExcelWorksheet errorSheet = wb.Worksheets[error.SheetName];
                    ExcelRange correctRange =
                        wSheetCorrect.Cells[linkIndex[error.CommentsList[0]], 1, linkIndex[error.CommentsList[0]], 12];
                    correctRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    correctRange.Style.Fill.BackgroundColor.SetColor(Color.Lime);

                    ExcelRange linkRange = errorSheet.Cells[charItem.RowNum, error.ColList[0]];
                    linkRange.Hyperlink = new ExcelHyperLink(
                        "'" + wSheetCorrect.Name + "'" + "!" + correctRange.Address,
                        linkRange.Value.ToString());


                    ExcelRange errorRange = linkRange;
                    ExcelRange linkRangeCorrect = wSheetCorrect.Cells[linkIndex[error.CommentsList[0]], 5];
                    linkRangeCorrect.Hyperlink =
                        new ExcelHyperLink("'" + errorSheet.Name + "'" + "!" + errorRange.Address,
                            linkRangeCorrect.Value.ToString());
                }
            }
            #endregion
        }
    }
}
