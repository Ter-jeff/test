using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace RfLib.Dvdc.GenTemplate.TestPlanFormat
{
    public class Template
    {
        protected const string Cp1LoHeader = "CP1 Lo Limit (H,L,N)";
        protected const string Cp1HiHeader = "CP1 Hi Limit (H,L,N)";
        protected const string Cp2LoHeader = "CP2 Lo Limit (H,L,N)";
        protected const string Cp2HiHeader = "CP2 Hi Limit (H,L,N)";
        protected const string Ft1LoHeader = "FT1 Lo Limit (H,L,N)";
        protected const string Ft1HiHeader = "FT1 Hi Limit (H,L,N)";
        protected const string Ft2LoHeader = "FT2 Lo Limit (H,L,N)";
        protected const string Ft2HiHeader = "FT2 Hi Limit (H,L,N)";
        protected const string Ft3LoHeader = "FT3 Lo Limit (H,L,N)";
        protected const string Ft3HiHeader = "FT3 Hi Limit (H,L,N)";
        protected const string QaLoHeader = "QA Lo Limit (H,L,N)";
        protected const string QaHiHeader = "QA Hi Limit (H,L,N)";
        protected const string HtolLoHeader = "HTOL Lo Limit (H,L,N)";
        protected const string HtolHiHeader = "HTOL Hi Limit (H,L,N)";
        protected const string CommentHeader = "Comment";

        public HeaderIndex Ttr { set; get; } = new HeaderIndex() { Header = "TTR", Index = 1 };
        public HeaderIndex TestItem { set; get; } = new HeaderIndex() { Header = "Test Item", Index = 2 };
        public HeaderIndex Step { set; get; } = new HeaderIndex() { Header = "Step", Index = 3 };
        public HeaderIndex Description { set; get; } = new HeaderIndex() { Header = "Description", Index = 4 };
        public HeaderIndex Pattern { set; get; } = new HeaderIndex() { Header = "Pattern", Index = 5 };
        public HeaderIndex TestName { set; get; } = new HeaderIndex() { Header = "TestName", Index = 6 };
        public HeaderIndex ForceCondition { set; get; } = new HeaderIndex() { Header = "Force Condition", Index = 7 };
        public HeaderIndex RegisterAssignment { set; get; } = new HeaderIndex() { Header = "Register Assignment", Index = 8 };
        public HeaderIndex MiscInfo { set; get; } = new HeaderIndex() { Header = "Misc Info", Index = 9 };
        public HeaderIndex Meas { set; get; } = new HeaderIndex() { Header = "Meas", Index = 10 };
        public HeaderIndex LoLimit { set; get; } = new HeaderIndex() { Header = Cp1LoHeader, Index = 11 };
        public HeaderIndex HiLimit { set; get; } = new HeaderIndex() { Header = Cp1HiHeader, Index = 12 };

        public virtual List<string> GetHeaders()
        {
            return [Ttr.Header,
                TestItem.Header,
                Step.Header,
                Description.Header,
                Pattern.Header,
                TestName.Header,
                ForceCondition.Header,
                RegisterAssignment.Header,
                MiscInfo.Header,
                Meas.Header,
                Cp1LoHeader,
                Cp1HiHeader,
                Cp2LoHeader,
                Cp2HiHeader,
                Ft1LoHeader,
                Ft1HiHeader,
                Ft2LoHeader,
                Ft2HiHeader,
                Ft3LoHeader,
                Ft3HiHeader,
                QaLoHeader,
                QaHiHeader,
                HtolLoHeader,
                HtolHiHeader,
                CommentHeader];
        }

        public virtual List<string> GetTrimUse()
        {
            return [];
        }

        public virtual void WriteTemplateHeader(ExcelWorksheet excelWorksheet)
        {
            List<string> headers = GetHeaders();
            for (int headerIndex = 1; headerIndex <= headers.Count; headerIndex++)
            {
                excelWorksheet.Cells[1, headerIndex].Value = headers[headerIndex - 1];
                excelWorksheet.Cells[1, headerIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                excelWorksheet.Cells[1, headerIndex].Style.Fill.BackgroundColor.SetColor(Color.RoyalBlue);
                excelWorksheet.Cells[1, headerIndex].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelWorksheet.Cells[1, headerIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            }
        }

        public virtual void WriteTemplateContent(ExcelWorksheet excelWorksheet, TemplateRow templateRow, int rowIndex)
        {
            excelWorksheet.Cells[rowIndex, TestItem.Index].Value = templateRow.TestItem;
            excelWorksheet.Cells[rowIndex, Description.Index].Value = templateRow.Description;
            if (templateRow.TestItem != 0)
            {
                excelWorksheet.Cells[rowIndex, Description.Index].Style.Fill.PatternType = ExcelFillStyle.Solid;
                excelWorksheet.Cells[rowIndex, Description.Index].Style.Fill.BackgroundColor.SetColor(Color.LightSalmon);
            }

            excelWorksheet.Cells[rowIndex, Ttr.Index].Value = templateRow.TTR;
            excelWorksheet.Cells[rowIndex, Step.Index].Value = templateRow.Step;
            excelWorksheet.Cells[rowIndex, Pattern.Index].Value = templateRow.Pattern;
            excelWorksheet.Cells[rowIndex, TestName.Index].Value = templateRow.TestName;
            excelWorksheet.Cells[rowIndex, ForceCondition.Index].Value = templateRow.ForceCondition;
            excelWorksheet.Cells[rowIndex, RegisterAssignment.Index].Value = templateRow.RegisterAssignment;
            int i = 0;
            foreach (string limitKey in templateRow.LoLimit.Keys)
            {
                excelWorksheet.Cells[rowIndex, LoLimit.Index + (i * 2)].Value = templateRow.LoLimit[limitKey];
                excelWorksheet.Cells[rowIndex, HiLimit.Index + (i * 2)].Value = templateRow.HiLimit[limitKey];
                i++;
            }

            if (!string.IsNullOrWhiteSpace(templateRow.Pattern))
            {
                excelWorksheet.Cells[rowIndex, Pattern.Index].Style.Fill.PatternType = ExcelFillStyle.Solid;
                excelWorksheet.Cells[rowIndex, Pattern.Index].Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
            }

            //wSheet.Cells[rowIndex, MiscInfo.Index].Value = template.MiscInfo;
            SetMiscInfoDifferentColor(excelWorksheet, rowIndex, MiscInfo.Index, templateRow.MiscInfo);
            excelWorksheet.Cells[rowIndex, Meas.Index].Value = templateRow.Meas;

        }

        protected static void SetMiscInfoDifferentColor(ExcelWorksheet excelWorksheet, int row, int col, string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return;
            }

            List<string> miscValue = [.. str.Split([';'], StringSplitOptions.RemoveEmptyEntries)];
            excelWorksheet.Cells[row, col].Style.WrapText = true;
            excelWorksheet.Cells[row, col].IsRichText = true;
            int i = 0;
            foreach (string misc in miscValue)
            {
                if (misc.Split(':').Length <= 1)
                {
                    ExcelRichText rtDir2 = excelWorksheet.Cells[row, col].RichText.Add(misc + ";");
                    rtDir2.Color = Color.Black;
                    i++;
                    continue;
                }

                List<string> miscinfo = [.. misc.Split(':')];
                string key = miscinfo[0];
                miscinfo.RemoveAt(0);
                string value = string.Join(":", miscinfo);
                if (string.IsNullOrEmpty(value))
                {
                    ExcelRichText rtDir1 = excelWorksheet.Cells[row, col].RichText.Add(key + ": ;");
                    rtDir1.Color = Color.Red;
                }
                else
                {
                    ExcelRichText rtDir2 = excelWorksheet.Cells[row, col].RichText.Add(key + ":" + value + ";");
                    rtDir2.Color = Color.Black;
                }
                if (i < miscValue.Count - 1)
                {
                    excelWorksheet.Cells[row, col].RichText.Add("\r\n");
                }
                i++;
            }
        }
    }

    [DebuggerDisplay("{Description}")]
    public class TemplateRow
    {
        public int TestItem { get; set; }
        public string TTR { get; set; }
        public string Step { get; set; }
        public string Description { get; set; }
        public string Pattern { get; set; }
        public string TestName { get; set; }
        public string RegisterAssignment { get; set; }
        public string MiscInfo { get; set; }
        public string Meas { get; set; }
        public Dictionary<string, string> LoLimit { get; set; }
        public Dictionary<string, string> HiLimit { get; set; }
        public string ForceCondition { get; set; }
        public string Seqindex { get; set; }
        public int Rowindex { get; set; }

        public TemplateRow(int testItem, string step)
        {
            TTR = "";
            TestItem = testItem;
            TestName = "";
            Step = step;
            Description = "";
            Pattern = "";
            TestName = "";
            RegisterAssignment = "";
            MiscInfo = "";
            Meas = "";
            LoLimit = [];
            HiLimit = [];
            ForceCondition = "";
            Seqindex = "";
            Rowindex = -1;
        }

    }

    public class HeaderIndex
    {
        public string Header = "";
        public int Index;
    }
}
