using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Enums;

using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

using FlowSheet = IgxlLib.IgxlSheets.SubFlowSheet;

namespace Automation.Utility.TpUpdate.HardIPBinoutTPUpdate
{
    public class HardIpBinOutTpUpdateMain
    {
        private readonly HashSet<string> _currentJobs = new HashSet<string>();

        private readonly string _binOutReportPath;
        private readonly Dictionary<string, FlowSheet> _subFlowSheets;
        private readonly List<BinTableSheet> _binTableSheets;

        private readonly Regex _regBinOutStatus = new Regex(@"(?<job>\w+)_BinOutStatus", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _regTpBinOut = new Regex(@"(?<job>\w+)_TPBinOut$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private const string ConBinOutUpdateConfigSheet = "BinOutUpdate_Config";

        public HardIpBinOutTpUpdateMain(string binOutReportPath, Dictionary<string, FlowSheet> subFlowSheets, List<BinTableSheet> binTableSheets)
        {
            _binOutReportPath = binOutReportPath;
            _subFlowSheets = subFlowSheets;
            _binTableSheets = binTableSheets;
        }

        public void WorkFlow()
        {
            try
            {
                #region Load And Update EnableWords with Flow Sheet

                var ep = new ExcelPackage(new FileInfo(_binOutReportPath));

                #region SearchJob

                foreach (ExcelWorksheet worksheet in ep.Workbook.Worksheets)
                {
                    if (_regBinOutStatus.IsMatch(worksheet.Name))
                    {
                        _currentJobs.Add(_regBinOutStatus.Match(worksheet.Name).Groups["job"].Value);
                    }
                    else if (_regTpBinOut.IsMatch(worksheet.Name))
                    {
                        _currentJobs.Add(_regTpBinOut.Match(worksheet.Name).Groups["job"].Value);
                    }
                }
                #endregion

                ExcelWorksheet binOutExcelSheet = ep.Workbook.Worksheets.ToList().Find(p =>
                    p.Name.Equals("BinOutStatus_HIP_List_Merge", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.Equals("LogBinOutHIP_Merge", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.Equals("TPBinOutHIPAll", StringComparison.OrdinalIgnoreCase)
                    );

                if (binOutExcelSheet == null)
                {
                    Response.Report("There are no \"TPBinOutHIPAll\" sheet...", EnumMessageLevel.Error);
                    return;
                }

                var binOutReader = new BinOutStatusHipListReader(_currentJobs);
                BinOutStatusHipListSheet binOutSheet = binOutReader.ReadSheet(binOutExcelSheet);

                var controller = new HardIpBinOutUpdateController(true);
                BinOutUpdateConfigSheet binoutUpdateConfig = new BinOutUpdateConfigReader().ReadSheet(ep.Workbook.Worksheets[ConBinOutUpdateConfigSheet]);

                // Default set to overwrite limits from test plan
                bool isOverWriteLimits = binoutUpdateConfig == null || binoutUpdateConfig.EnableOverWriteTestLimits;
                controller.ReferenceStatusAndLimits(binOutSheet, isOverWriteLimits);

                // TODO: this is for update the limit and status for all NV/LV/HV
                var binOutHipItems = binOutSheet.Rows.GroupBy(p => p.Testinstance + "#" + p.Type).ToDictionary(p => p.Key, p => p.ToList());

                binOutHipItems = binOutHipItems.Where(p => p.Value.Any(q => q.BinOutEnableDictionary.Any(k => !string.IsNullOrEmpty(k.Value)) ||
                                                                            q.UpdatedLoLimitDic.Any(k => !string.IsNullOrEmpty(k.Value)) ||
                                                                            q.UpdatedHiLimitDic.Any(k => !string.IsNullOrEmpty(k.Value)))).ToDictionary(p => p.Key, p => p.Value);
                if (!binOutHipItems.Any())
                {
                    Response.Report("There are no items need to be updated in \"TPBinOutHIPAll\"...", EnumMessageLevel.Error);
                    return;
                }

                #endregion

                var binOutByBlocks = binOutHipItems.GroupBy(p => p.Key.Split('#')[0].Split('_')[0]).ToDictionary(p => p.Key, p => p.ToList());

                List<string> allHipFlows = controller.AccessFileList(binOutByBlocks.Keys.ToList(), _subFlowSheets);
                List<FlowSheet> hipFlowSheets = ReadHipFlows(allHipFlows);

                Dictionary<string, List<string>> binTable = controller.GetHardipBintableRows(_binTableSheets);
                List<FlowSheet> results = controller.UpdateFlowRowsMainByMultiJobs(hipFlowSheets, binOutByBlocks, binTable, _currentJobs);

                WriteResult(results, allHipFlows);
            }
            catch (Exception ex)
            {
                Response.Report("Error occurred while binOut updating", EnumMessageLevel.Error);
                Response.Report(ex.ToString(), EnumMessageLevel.Error);
            }
        }

        private List<FlowSheet> ReadHipFlows(List<string> sheets)
        {
            Dictionary<string, FlowSheet> flows = _subFlowSheets;
            var hardipFlows = new List<FlowSheet>();

            foreach (string sheet in sheets)
            {
                hardipFlows.Add(flows[sheet]);
            }

            return hardipFlows;
        }

        private void WriteResult(List<FlowSheet> sheets, List<string> allHipSheets)
        {
            foreach (FlowSheet sheet in sheets)
            {
                string targetHipSheet = allHipSheets.FirstOrDefault(p => Path.GetFileNameWithoutExtension(p).Equals(sheet.Name, StringComparison.OrdinalIgnoreCase));
                if (targetHipSheet != null)
                {
                    Response.Report($"Update/Write sheet: {targetHipSheet + ".txt"}, version: 3.0", EnumMessageLevel.Info);
                    sheet.Write(targetHipSheet + ".txt", "3.0");
                }
            }

            Response.Report("Update/Write sheets done for Hardip BinOut TP Update.", EnumMessageLevel.Info);
        }
    }
}
