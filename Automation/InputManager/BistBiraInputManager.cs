using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.BistBira.Base;
using Automation.GenerateIgxl.BistBira.BistInputLib;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Reader.ConfigFile.NamingRule.Business;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using LogLib.Static;

using OfficeOpenXml;

using ScghLib.Base;
using ScghLib.Enums;
using ScghLib.Reader;

using TestPlanLib.Basic;
using TestPlanLib.Mbist;
using TestPlanLib.Static;

namespace Automation.InputManager
{
    public class BistBiraInputManager : InputManagerBase<BistBiraInputData>
    {
        private static readonly Regex _regex1 = new Regex(@"\w+_(?<chiplet>[A-z]\d+$)", RegexOptions.IgnoreCase);

        public BistBiraInputManager(ExcelWorkbook excelWorkbook, ParaData paraData) : base(excelWorkbook)
        {
            ParaData = paraData;
        }

        public override BistBiraInputData Read()
        {
            var result = new BistBiraInputData
            {
                Naming = new BistNaming(new MbistConfig()),
            };
            if (EpWorkbook.EquationVoltages != null)
            {
                result.EquationVoltage = new EquationVoltageReader().ReadSheet(EpWorkbook.EquationVoltages.Worksheets["EquationVoltages"]).First();
            }

            Response.Report("Mbist Initialization... ", percentage: 10);

            var configReader = new MbistConfigReader();
            result.Config = configReader.ReadConfig();

            if (!File.Exists(LocalSpecs.PatternListCsvFileName))
            {
                result.PatternDic = null;
                Response.Report("PatterList.csv does not exist!", EnumMessageLevel.Warning, 20);
            }
            else
            {
                result.PatternDic = AcTSetCategoryMapSingleton.Instance().PatternList;
            }

            result.PerformanceModeFilter = new PerformanceModeFilter();

            result.VoltageConverter = new VoltageConverter(MultiTestSettingSheetsSingleton.Instance());

            var mbistSheets = new List<MbistSheet>();
            if (TestPlanStatic.MainFlowSheet != null && TestPlanStatic.MainFlowSheet.Rows != null)
            {
                List<FlowSequenceNew> mbistFlows = TestPlanStatic.MainFlowSheet.Rows.FirstOrDefault().SequencesNew.FindAll(x => x.Module.ToUpper().Equals("MBIST"));
                mbistSheets = AddSourceSheetFromFlowMain(mbistFlows);
            }
            else
            {
                foreach (MbistSheetInfo mbistSheetInfo in NeededSheets.MbistDomainOrder)
                {
                    string[] sheets = mbistSheetInfo.Sheets.Split(',');
                    bool isMultiChipLet = IsMultiChipLet(sheets);
                    foreach (string sheet in sheets)
                    {
                        mbistSheets.Add(new MbistSheet { SheetName = sheet, IsMultipleSheet = true, IsMultiChipLet = isMultiChipLet });
                    }
                }
            }

            foreach (ExcelWorksheet sheet in EpWorkbook.ScghWorkbook.Worksheets)
            {
                if (!sheet.Name.Contains("@"))
                {
                    continue;
                }

                if (sheet.Name.Split('@')[0].Equals(NeededSheets.MbistScghGpu, StringComparison.OrdinalIgnoreCase) ||
                    sheet.Name.Split('@')[0].Equals(NeededSheets.MbistScghCpu, StringComparison.OrdinalIgnoreCase) ||
                    sheet.Name.Split('@')[0].Equals(NeededSheets.MbistScghSoc, StringComparison.OrdinalIgnoreCase))
                {
                    mbistSheets.Add(new MbistSheet { SheetName = sheet.Name, Jobs = sheet.Name.Split('@').Last() });
                }
            }

            var prodFlowSheets = new List<BistProdFlowSheet>();
            foreach (MbistSheet mbistSheet in mbistSheets)
            {
                string sheetName = mbistSheet.SheetName;
                ExcelWorksheet worksheet = EpWorkbook.ScghWorkbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    Response.Report($"Warning: Cannot find WorkSheet:{sheetName} in Scgh File", EnumMessageLevel.Warning, 70);
                    continue;
                }
                Response.Report($"Generating for sheet {sheetName} ...", percentage: 50);

                var sheetReader = new BistProdFlowReader(mbistSheet);
                prodFlowSheets.Add(sheetReader.ReadSheet(worksheet));
            }

            result.ProdFlowSheets = prodFlowSheets;

            return result;
        }

        public List<MbistSheet> AddSourceSheetFromFlowMain(List<FlowSequenceNew> mbistFlows)
        {
            var scgSheetList = new List<MbistSheet>();
            if (mbistFlows.Any())
            {
                scgSheetList.Clear();
            }
            foreach (FlowSequenceNew flow in mbistFlows)
            {
                string sheet = flow.SheetName;
                MbistPatSetType mbistPatSetType = MbistPatSetType.Single;
                MbistBinTableType mbistBinTableType = MbistBinTableType.Single;
                bool mbistLoop = false;
                SetPatSetsOption(flow.OptionDict, ref mbistPatSetType, ref mbistBinTableType);
                SetBinTableOption(flow.OptionDict, ref mbistBinTableType);
                SetMbistLoopOption(flow.OptionDict, ref mbistLoop);

                scgSheetList.Add(
                    new MbistSheet
                    {
                        SheetName = sheet,
                        IsMultipleSheet = true,
                        MbistPatSetType = mbistPatSetType,
                        IsCof = flow.SubFlowName.ToUpper().StartsWith("COF"),
                        MbistBinTableType = mbistBinTableType,
                        MbistLoop = mbistLoop
                    });
            }
            return scgSheetList;
        }

        private void SetPatSetsOption(Dictionary<string, string> flowMainOption, ref MbistPatSetType mbistPatSetType, ref MbistBinTableType mbistBinTableType)
        {
            if (flowMainOption.TryGetValue("PatSets", out string type))
            {
                switch (type.ToUpper())
                {
                    case "SINGLE":
                        mbistPatSetType = MbistPatSetType.Single;
                        mbistBinTableType = MbistBinTableType.Single;
                        break;
                    case "BURSTNO":
                        mbistPatSetType = MbistPatSetType.BurstNo;
                        mbistBinTableType = MbistBinTableType.Burst;
                        break;
                    case "BURSTYES":
                        mbistPatSetType = MbistPatSetType.BurstYes;
                        mbistBinTableType = MbistBinTableType.Burst;
                        break;
                }
            }
        }

        private void SetBinTableOption(Dictionary<string, string> flowMainOption, ref MbistBinTableType mbistBinTableType)
        {
            if (flowMainOption.TryGetValue("BinTable", out string type))
            {
                switch (type.ToUpper())
                {
                    case "SINGLE":
                        mbistBinTableType = MbistBinTableType.Single;
                        break;
                    case "BURST":
                        mbistBinTableType = MbistBinTableType.Burst;
                        break;
                }
            }
        }

        private void SetMbistLoopOption(Dictionary<string, string> flowMainOption, ref bool mbistLoop)
        {
            if (flowMainOption.ContainsKey("MbistLoop"))
            {
                mbistLoop = true;
            }
        }

        internal bool IsMultiChipLet(string[] sheets)
        {
            var chipLetList = new List<string>();
            bool isMultiChipLet = false;
            IEnumerable<string> chiplets = sheets.Select(sheet => _regex1.Match(sheet).Groups["chiplet"].ToString());
            foreach (string chiplet in chiplets)
            {
                if (chipLetList.Contains(chiplet))
                {
                    isMultiChipLet = true;
                }
                else
                {
                    chipLetList.Add(chiplet);
                }
            }
            return isMultiChipLet;
        }
    }
}
