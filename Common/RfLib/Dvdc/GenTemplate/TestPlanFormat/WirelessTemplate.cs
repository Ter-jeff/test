using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Utility;

using Newtonsoft.Json;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using Color = System.Drawing.Color;

namespace RfLib.Dvdc.GenTemplate.TestPlanFormat
{
    public class WirelessTemplate : Template
    {
        private const string TestNameHeader = "TestName";
        private const string TrimRegNameHeader = "TrimFuseName";
        private const string TrimTargetHeader = "TrimTarget";
        private const string TrimMeasHeader = "TrimMeas";
        private const string TrimTypeHeader = "TrimType";
        private const string BestCodeCalcFuncHeader = "BestCodeCalcFunc";
        private const string RfInstrumentSetupHeader = "RF Instrument Setup";
        private const string RfTestTypeHeader = "RF Test Type";
        private const string RfInterposeHeader = "RF Interpose";
        public new HeaderIndex TestName { set; get; }
        public HeaderIndex TrimRegName { set; get; }
        public HeaderIndex TrimTarget { set; get; }
        public HeaderIndex TrimMeas { set; get; }
        public HeaderIndex TrimType { set; get; }
        public HeaderIndex BestCodeCalcFunc { set; get; }
        public HeaderIndex RfInstrumentSetup { set; get; }
        public HeaderIndex RfTestType { set; get; }
        public HeaderIndex RfInterpose { set; get; }

        public static List<string> RfUse { get; private set; } = [RfInstrumentSetupHeader, RfTestTypeHeader, RfInterposeHeader];

        public WirelessTemplate()
        {
            TestName = new HeaderIndex() { Header = "TestName", Index = 6 };
            TrimRegName = new HeaderIndex() { Header = "TrimFuseName", Index = 7 };
            TrimTarget = new HeaderIndex() { Header = "TrimTarget", Index = 8 };
            TrimMeas = new HeaderIndex() { Header = "TrimMeas", Index = 9 };
            TrimType = new HeaderIndex() { Header = "TrimType", Index = 10 };
            BestCodeCalcFunc = new HeaderIndex() { Header = "BestCodeCalcFunc", Index = 11 };
            RfInstrumentSetup = new HeaderIndex() { Header = "RF Instrument Setup", Index = 12 };
            RfTestType = new HeaderIndex() { Header = "RF Test Type", Index = 13 };
            RfInterpose = new HeaderIndex() { Header = "RF Interpose", Index = 14 };
            ForceCondition = new HeaderIndex() { Header = "Force Condition", Index = 15 };
            RegisterAssignment = new HeaderIndex() { Header = "Register Assignment", Index = 16 };
            MiscInfo = new HeaderIndex() { Header = "Misc Info", Index = 17 };
            Meas = new HeaderIndex() { Header = "Meas", Index = 18 };
            LoLimit = new HeaderIndex() { Header = Cp1LoHeader, Index = 19 };
            HiLimit = new HeaderIndex() { Header = Cp1HiHeader, Index = 20 };
        }

        public override List<string> GetHeaders()
        {
            return [Ttr.Header,
                TestItem.Header,
                Step.Header,
                Description.Header,
                Pattern.Header,
                TestNameHeader,
                TrimRegNameHeader,
                TrimTargetHeader,
                TrimMeasHeader,
                TrimTypeHeader,
                BestCodeCalcFuncHeader,
                RfInstrumentSetupHeader,
                RfTestTypeHeader,
                RfInterposeHeader,
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

        public override List<string> GetTrimUse()
        {
            return
            [
                TestNameHeader,
                TrimRegNameHeader,
                TrimTargetHeader,
                TrimMeasHeader,
                TrimTypeHeader,
                BestCodeCalcFuncHeader
            ];
        }

        public List<int> GetNonUseIndex()
        {
            var ignoreList = new List<int>();
            if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                ignoreList.Add(RfTestType.Index);
                ignoreList.Add(RfInstrumentSetup.Index);
                //IgnoreList.Add(RfInterpose.Index);
            }
            return ignoreList;
        }

        public override void WriteTemplateHeader(ExcelWorksheet excelWorksheet)
        {
            List<string> headers = GetHeaders();
            for (int headerIndex = 1; headerIndex <= headers.Count; headerIndex++)
            {
                excelWorksheet.Cells[1, headerIndex].Value = headers[headerIndex - 1];
                excelWorksheet.Cells[1, headerIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ////if (GetNonUse().Contains(headers[headerIndex-1])&& ProjectConfigLocalSpecs.Device == DeviceEnum.LCD)
                ////    wSheet.Cells[1, headerIndex].Style.Fill.BackgroundColor.SetColor(Color.Black);
                //else
                if (RfUse.Contains(headers[headerIndex - 1]))
                {
                    excelWorksheet.Cells[1, headerIndex].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }
                else if (GetTrimUse().Contains(headers[headerIndex - 1]))
                {
                    excelWorksheet.Cells[1, headerIndex].Style.Fill.BackgroundColor.SetColor(Color.DarkOrange);
                }
                else
                {
                    excelWorksheet.Cells[1, headerIndex].Style.Fill.BackgroundColor.SetColor(Color.RoyalBlue);
                }

                excelWorksheet.Cells[1, headerIndex].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelWorksheet.Cells[1, headerIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            }
        }

        public override void WriteTemplateContent(ExcelWorksheet excelWorksheet, TemplateRow templateRow, int rowIndex)
        {
            excelWorksheet.Cells[rowIndex, TestItem.Index].Value = templateRow.TestItem;
            excelWorksheet.Cells[rowIndex, Description.Index].Value = templateRow.Description;
            if (templateRow.TestItem != 0)
            {
                excelWorksheet.Cells[rowIndex, Description.Index].Style.Fill.PatternType = ExcelFillStyle.Solid;
                excelWorksheet.Cells[rowIndex, Description.Index].Style.Fill.BackgroundColor.SetColor(Color.LightSalmon);
            }

            excelWorksheet.Cells[rowIndex, Step.Index].Value = templateRow.Step;
            excelWorksheet.Cells[rowIndex, Pattern.Index].Value = templateRow.Pattern;
            try
            {
                string parent = JsonConvert.SerializeObject(templateRow);
                WirelessTemplateRow newtemplate = JsonConvert.DeserializeObject<WirelessTemplateRow>(parent)!;
                if (!string.IsNullOrEmpty(newtemplate.TestName) ||
                    newtemplate.TrimRegName.Count > 0 ||
                    !string.IsNullOrEmpty(newtemplate.TrimTarget) ||
                    !string.IsNullOrEmpty(newtemplate.TrimMeas) ||
                    !string.IsNullOrEmpty(newtemplate.TrimType) ||
                    !string.IsNullOrEmpty(newtemplate.BestCode) ||
                    !string.IsNullOrEmpty(newtemplate.InstrumentSetup) ||
                    !string.IsNullOrEmpty(newtemplate.PostCalc) ||
                    !string.IsNullOrEmpty(newtemplate.Interpose))
                {

                    excelWorksheet.Cells[rowIndex, TestName.Index].Value = newtemplate.TestName;
                    excelWorksheet.Cells[rowIndex, TrimRegName.Index].Value = string.Join(";", newtemplate.TrimRegName.Distinct());
                    excelWorksheet.Cells[rowIndex, TrimTarget.Index].Value = newtemplate.TrimTarget;
                    excelWorksheet.Cells[rowIndex, TrimMeas.Index].Value = newtemplate.TrimMeas;
                    excelWorksheet.Cells[rowIndex, TrimType.Index].Value = newtemplate.TrimType;
                    excelWorksheet.Cells[rowIndex, BestCodeCalcFunc.Index].Value = newtemplate.BestCode;
                    excelWorksheet.Cells[rowIndex, RfInstrumentSetup.Index].Value = newtemplate.InstrumentSetup;
                    excelWorksheet.Cells[rowIndex, RfTestType.Index].Value = newtemplate.PostCalc;
                    excelWorksheet.Cells[rowIndex, RfInterpose.Index].Value = newtemplate.Interpose;
                    //if(newtemplate.InstrumentType.Count != 0)
                    //{
                    //    var instType = wSheet.Cells[rowIndex,RfInstrumentType.Index].DataValidation.AddListDataValidation();
                    //    foreach(var type in newtemplate.InstrumentType)
                    //    {
                    //        instType.Formula.Values.Add(type);
                    //    }
                    //}

                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                throw;
            }
            excelWorksheet.Cells[rowIndex, ForceCondition.Index].Value = templateRow.ForceCondition;
            excelWorksheet.Cells[rowIndex, RegisterAssignment.Index].Value = templateRow.RegisterAssignment;
            if (!string.IsNullOrWhiteSpace(templateRow.Pattern))
            {

                excelWorksheet.Cells[rowIndex, Pattern.Index].Style.Fill.PatternType = ExcelFillStyle.Solid;
                excelWorksheet.Cells[rowIndex, Pattern.Index].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                excelWorksheet.Cells[rowIndex, Pattern.Index].Style.Font.Color.SetColor(Color.White);
                excelWorksheet.Cells[rowIndex, Pattern.Index].Style.Font.Bold = true;

            }
            //wSheet.Cells[rowIndex, MiscInfo.Index].Value = template.MiscInfo;
            SetMiscInfoDifferentColor(excelWorksheet, rowIndex, MiscInfo.Index, templateRow.MiscInfo);

            excelWorksheet.Cells[rowIndex, Meas.Index].Value = templateRow.Meas;
            //wSheet.Cells[rowIndex, LoLimit.Index].Value = template.LoLimit;
            //wSheet.Cells[rowIndex, HiLimit.Index].Value = template.HiLimit;
            int i = 0;
            foreach (string limitKey in templateRow.LoLimit.Keys)
            {
                excelWorksheet.Cells[rowIndex, LoLimit.Index + (i * 2)].Value = templateRow.LoLimit[limitKey];
                excelWorksheet.Cells[rowIndex, HiLimit.Index + (i * 2)].Value = templateRow.HiLimit[limitKey];
                i++;
            }
            if (LocalSpecs.Options.Device == EnumDevice.LCD && !string.IsNullOrWhiteSpace(templateRow.Pattern))
            {
                excelWorksheet.Cells[rowIndex, RfInterpose.Index].Style.Fill.PatternType = ExcelFillStyle.Solid;
                excelWorksheet.Cells[rowIndex, RfInterpose.Index].Style.Fill.BackgroundColor.SetColor(Color.Khaki);
            }
            if (LocalSpecs.Options.Device == EnumDevice.LCD && string.IsNullOrWhiteSpace(templateRow.Pattern))
            {
                if (!string.IsNullOrWhiteSpace(templateRow.TestName) ||
                    (!string.IsNullOrWhiteSpace(templateRow.Meas) && !templateRow.Meas.EqualsIgnoreCase("MeasN")))
                {
                    excelWorksheet.Cells[rowIndex, LoLimit.Index].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    excelWorksheet.Cells[rowIndex, LoLimit.Index].Style.Fill.BackgroundColor.SetColor(Color.Khaki);
                    excelWorksheet.Cells[rowIndex, HiLimit.Index].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    excelWorksheet.Cells[rowIndex, HiLimit.Index].Style.Fill.BackgroundColor.SetColor(Color.Khaki);
                }
            }
        }
    }

    public class WirelessTemplateRow : TemplateRow
    {
        public List<string> TrimRegName = [];
        public string TrimTarget { get; set; }
        public string TrimMeas { get; set; }
        public string TrimType { get; set; }
        public string BestCode { get; set; }
        public string InstrumentSetup { get; set; }
        public string DeviceSetup { get; set; }
        public string PostCalc { get; set; }
        public string Interpose { get; set; }

        public WirelessTemplateRow(int testItem, string step, string description = "", string testname = "")
            : base(testItem, step)
        {
            TestItem = testItem;
            Step = step;
            TrimRegName = [];
            TrimTarget = "";
            TrimMeas = "";
            TrimType = "";
            BestCode = "";
            InstrumentSetup = "";
            PostCalc = "";
            Interpose = "";
            DeviceSetup = "";

            Description = description;
            TestName = testname.ToUpper();
        }

    }
}
