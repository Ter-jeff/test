using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using LogLib.Static;

using OfficeOpenXml;

namespace Automation.PreCheck.PreChecks.Basic
{
    public class BasicPreCheckPowerInfo : BasicPreCheckBase
    {
        private readonly Dictionary<int, string> _titleIndexDic = new Dictionary<int, string>();
        private int _pinNameColIdx;
        private int _powerSequenceColIdx;

        public BasicPreCheckPowerInfo(ExcelWorkbook excelWorkbook, string sheetName) : base(excelWorkbook, sheetName)
        {
            FirstHeader = "(PinName|Chiplet).*";
        }

        protected internal override bool CheckHeaders()
        {
            bool result = CheckColumnTitle();

            return result;
        }

        protected internal override bool CheckFormat()
        {
            bool result = CheckPowerInfoData();

            return result;
        }

        protected internal override bool CheckBusiness()
        {
            bool result = CheckContainsPowerPinInTs();
            if (!CheckIfoldStageMatch())
            {
                result = false;
            }

            return result;
        }

        private bool CheckColumnTitle()
        {
            bool result = true;
            for (int i = StartColumn; i <= ExcelWorksheet.Dimension.Columns; i++)
            {
                string content = ExcelWorksheet.GetCellValue(StartRow, i);
                if (content.ContainsIgnoreCase("PinName"))
                {
                    _titleIndexDic.Add(i, "PinName");
                    _pinNameColIdx = i;
                }
                else if (content.ContainsIgnoreCase("PowerSequence"))
                {
                    _titleIndexDic.Add(i, "PowerSequence");
                    _powerSequenceColIdx = i;
                }
                else if (content.ContainsIgnoreCase("PowerUpSequence"))
                {
                    _titleIndexDic.Add(i, "PowerUpSequence");
                    _powerSequenceColIdx = i;
                }
            }
            if (!_titleIndexDic.ContainsValue("PowerSequence") && !_titleIndexDic.ContainsValue("PowerUpSequence"))
            {
                ErrorReportManager.AddError(BasicErrorType.E_MissingBlock_05, ExcelWorksheet.Name, StartRow, 1, "Omit column \"PowerSequence\"");
                result = false;
            }
            return result;
        }

        private bool CheckContainsPowerPinInTs()
        {
            bool result = true;

            var multiTestSettingSheetsSingleton = MultiTestSettingSheetsSingleton.Instance();
            if (multiTestSettingSheetsSingleton.PowerPinList != null && multiTestSettingSheetsSingleton.PowerPinList.Any())
            {
                List<string> powerPins = multiTestSettingSheetsSingleton.PowerPinList;
                List<ProtocolAwarePin> nWirePin = NwireSingleton.Instance().SettingInfo.NwirePins;
                for (int i = StartRow + 1; i <= ExcelWorksheet.Dimension.Rows; i++)
                {
                    string content = ExcelWorksheet.GetCellValue(i, _pinNameColIdx);
                    string pinSequence = ExcelWorksheet.GetCellValue(i, _powerSequenceColIdx + 1);
                    if (content == string.Empty)
                    {
                        break;
                    }

                    if (pinSequence.Contains(":"))
                    {
                        continue;
                    }

                    if (!powerPins.Contains(content) && !nWirePin.Any(x => content.Split(':').Contains(x.OutClk)))
                    {
                        Response.Report("Pin '" + content + "' does not exist in TestSetting.", EnumMessageLevel.Error);
                        ErrorReportManager.AddError(BasicErrorType.E_MissingBlock_06, ExcelWorksheet.Name, i, _pinNameColIdx, "Pin '" + content + "' does not exist in TestSetting.", new string[] { content });
                        result = false;
                    }
                }
            }

            return result;
        }

        private bool CheckIfoldStageMatch()
        {
            bool result = true;
            Dictionary<int, string> ifoldCols = FindIfoldJobTitleIndex();
            if (ifoldCols.Count == 1 && ifoldCols.ElementAt(0).Value == "")
            {
                return true;
            }

            List<string> nonUsedJobs = LocalSpecs.AllJobs;
            string message;
            foreach (KeyValuePair<int, string> ifoldCol in ifoldCols)
            {
                string ifoldJob = "";
                if (!string.IsNullOrEmpty(ifoldCol.Value))
                {
                    ifoldJob = ifoldCol.Value.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries).Last();
                }

                if (!LocalSpecs.AllJobs.Exists(p => p.Equals(ifoldJob, StringComparison.OrdinalIgnoreCase)))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_08, ExcelWorksheet.Name, StartRow, ifoldCol.Key, $"Job: {ifoldCol.Value} does not define in FlowMain(Test Plan)/jobMapping(Setting file) sheet !!!", new string[] { ifoldCol.Value });
                    result = false;

                }
                else
                {
                    nonUsedJobs.Remove(ifoldJob);
                }
            }
            if (nonUsedJobs.Count > 0)
            {
                message = $"Jobs: {string.Join(",", nonUsedJobs)} need to define in {ExcelWorksheet.Name} !!!";
                ErrorReportManager.AddError(BasicErrorType.E_MissingPin_09, ExcelWorksheet.Name, StartRow, ifoldCols.Last().Key, $"Jobs: {string.Join(",", nonUsedJobs)} need to define in {ExcelWorksheet.Name} !!!", new string[] { string.Join(",", nonUsedJobs), ExcelWorksheet.Name });
                result = false;
            }

            return result;
        }

        private bool CheckPowerInfoData()
        {
            bool result = true;
            for (int i = StartRow + 1; i <= ExcelWorksheet.Dimension.Rows; i++)
            {
                string content = ExcelWorksheet.GetCellValue(i, StartColumn);
                if (content == string.Empty)
                {
                    break;
                }

                for (int j = StartColumn + 1; j <= StopColumn; j++)
                {
                    string value = ExcelWorksheet.GetCellValue(i, j);
                    if (value == string.Empty)
                    {
                        if (_titleIndexDic.ContainsKey(j))
                        {
                            if (_titleIndexDic[j].ContainsIgnoreCase("ifold"))
                            {
                                continue;
                            }

                            ErrorReportManager.AddError(BasicErrorType.E_FormatError_14, ExcelWorksheet.Name, i, 1, "Omit \"" + _titleIndexDic[j] + "\" value for pin: " + content, new string[] { _titleIndexDic[j], content });
                            result = false;
                        }
                        continue;
                    }

                    if (_titleIndexDic.ContainsKey(j) && _titleIndexDic[j].Equals("ifold", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (!value.Equals("N/A") && !value.Equals("NA") && !double.TryParse(value, out _))
                        {
                            if (_titleIndexDic.TryGetValue(j, out string value1))
                            {
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_17, ExcelWorksheet.Name, i, j, "Non numerical \"" + value1 + "\" value for pin: " + content, new string[] { value1, content });
                                result = false;
                            }
                            else
                            {
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_18, ExcelWorksheet.Name, i, j, "Non numerical unknown type value for pin: " + content, new string[] { content });
                                result = false;
                            }
                        }
                    }
                }
            }

            return result;
        }

        private Dictionary<int, string> FindIfoldJobTitleIndex()
        {
            var result = new Dictionary<int, string>();
            var titles = new Dictionary<int, string>();
            for (int i = StartColumn + 1; i <= ExcelWorksheet.Dimension.Columns; i++)
            {
                string content = ExcelWorksheet.GetCellValue(StartRow, i);
                titles.Add(i, content);
            }
            var items = titles.Where(x => x.Value.ContainsIgnoreCase("ifold")).ToList();
            if (!items.Any())
            {
                result.Add(0, "");
            }

            foreach (KeyValuePair<int, string> entry in items)
            {
                string job = entry.Value.Contains(":") ? entry.Value.Split(':').Last() : "";
                result.Add(entry.Key, job);
            }
            return result;
        }
    }
}
