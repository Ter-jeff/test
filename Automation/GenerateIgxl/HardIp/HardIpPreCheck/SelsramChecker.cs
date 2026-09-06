using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using OfficeOpenXml;

using TestPlanLib.BinCut;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class SelsramChecker : HardIpPrecheckBase
    {
        private string _selsramSheetName = "";
        internal SelsrmMappingSheet _selsrmMappingSheet;

        public SelsramChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        public override void Check()
        {
            bool isCheckSelsram = Pattern.MiscInfo.Split(';').Any(x => x.Equals("CheckSelsram", StringComparison.OrdinalIgnoreCase));
            if (!isCheckSelsram)
            {
                return;
            }

            CheckTable();

            CheckItem(Pattern);
        }

        private void CheckTable()
        {
            GetTable();

            if (_selsrmMappingSheet != null)
            {
                CheckHeader();
            }
        }

        private void GetTable()
        {
            ExcelWorksheet sheet = EpWorkbook.TestPlanWorkbook.Worksheets["SELSRM_Mapping_Table"] ?? EpWorkbook.TestPlanWorkbook.Worksheets["SELSRAM_Mapping_Table"];
            if (sheet != null)
            {
                _selsramSheetName = sheet.Name;
                var selsrmMappingSheetReader = new SelsrmMappingSheetReader();
                _selsrmMappingSheet = selsrmMappingSheetReader.ReadSheet(sheet);
            }
            else
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_SelsramMappingTableError_01,
                    _selsramSheetName,
                    0,
                    0,
                    []
                );
            }
        }

        internal void CheckHeader()
        {
            if (!_selsrmMappingSheet.HeaderIndex.ContainsKey("DigSrc_Assignment"))
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_SelsramMappingTableError_02,
                    _selsramSheetName,
                    0,
                    0,
                    []
                );
            }
        }

        internal void CheckItem(HardIpPattern pattern)
        {
            string block = pattern.SheetName.Replace("HARDIP_", "");
            IEnumerable<string> registerAssignmentList = pattern.RegisterAssignment.Split(';').Select(x => x.Split('=')[0].ToUpper());
            List<string> patternNames = pattern.Pattern.InstancePatternName;
            bool findSelsramSetting = false;

            foreach (string patternName in patternNames)
            {
                IEnumerable<SelsrmMappingTableRow> selsramList = _selsrmMappingSheet.Rows.Where(x => Regex.IsMatch(block, x.Block.Replace("*", ".*"), RegexOptions.IgnoreCase) && Regex.IsMatch(patternName, x.Pattern.Replace("*", ".*"), RegexOptions.IgnoreCase) && !(x.LogicPins.Trim().ToUpper() == "PRESERVED" || x.SramPins.Trim().ToUpper() == "PRESERVED"));
                if (selsramList.Any())
                {
                    findSelsramSetting = true;
                    //TODO
                    foreach (SelsrmMappingTableRow selsram in selsramList)
                    {
                        string tableDsa = selsram.DigSrcAssignment.ToUpper();
                        if (string.IsNullOrWhiteSpace(tableDsa) || tableDsa.Trim() == "N/A" || tableDsa.Trim() == "NA")
                        {
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_SelsramDigSrcAssignmentNotDefineInTable_01,
                                _selsramSheetName,
                                selsram.RowNum,
                                _selsrmMappingSheet.HeaderIndex["DigSrc_Assignment"],
                                [selsram.Block, selsram.Pattern, selsram.Alpha, _selsramSheetName]
                            );
                            continue;
                        }
                        if (!registerAssignmentList.Any(x => x.Equals(tableDsa)))
                        {
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_SelsramDigSrcAssignmentNotDefineInInstance_01,
                                pattern.SheetName,
                                pattern.RowNum,
                                HardIpSheet.PlanHeaderIdx["registerIndex"],
                                [tableDsa]
                            );
                        }
                    }
                    break;
                }
            }

            if (!findSelsramSetting)
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_CanNotGetSelsramSetting_01,
                    pattern.SheetName,
                    pattern.RowNum,
                    HardIpSheet.PlanHeaderIdx["patternIndex"],
                    [block, pattern.Pattern.GetPatternName(), _selsramSheetName]
                );
            }
        }
    }
}
