using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.PostAction.InstCommon
{
    public class InstCommonMain
    {
        private static readonly Regex _regex = new Regex("_Footer_|_Header_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex1 = new Regex("init|nWire", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public void WorkFlowWithoutHeader(string commonSheet)
        {
            if (string.IsNullOrEmpty(commonSheet) || SettingStatic.BasicConfigWorkbook.Worksheets[commonSheet] == null)
            {
                return;
            }

            WorkFlow(commonSheet, new List<InstanceRow>());
        }

        public void WorkFlowWithHeader(string commonSheet)
        {
            if (string.IsNullOrEmpty(commonSheet) || SettingStatic.BasicConfigWorkbook.Worksheets[commonSheet] == null)
            {
                return;
            }

            List<InstanceRow> instanceRows = AddHeaderForFlowAndInstance();
            WorkFlow(commonSheet, instanceRows);
        }

        private void WorkFlow(string commonSheet, List<InstanceRow> instanceRows)
        {
            var output = new InstCommonOutput();
            ExcelWorksheet instanceCommonWorksheet = SettingStatic.BasicConfigWorkbook.Worksheets[commonSheet];
            InstanceSheet instanceSheet = new ReadInstanceSheet().ReadSheet(instanceCommonWorksheet);
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                ReplaceCsCommonInstanceFromConfig(commonSheet, ref output);
            }
            else
            {
                output.InstanceRows.AddRange(instanceSheet.Rows);
            }
            output.InstanceRows.AddRange(instanceRows);

            AddStoreReferenceClockName(ref output);
            AddInitializeModulesFromInstance(ref output);
            output.AddDataToLocalSpec();
        }

        private List<InstanceRow> AddHeaderForFlowAndInstance()
        {
            List<InstanceRow> instanceRows = CollectFooterHeaderInstanceRows();

            bool flag = TestProgram.IgxlWorkBk.TryGetTestInstCommon(out _);
            List<InstanceRow> all = CollectAllInstanceRows();

            return DistributeFooterHeaderInstanceRows(instanceRows, all, flag);
        }

        private List<InstanceRow> CollectFooterHeaderInstanceRows()
        {
            var instanceRows = new List<InstanceRow>();
            foreach (SubFlowSheet subFlowSheet in TestProgram.IgxlWorkBk.SubFlowSheets.Values)
            {
                if (!_regex1.IsMatch(subFlowSheet.Name))
                {
                    #region insert Footer/Header flow to each subflowsheet except init and nWire, temperary put here for short-term solution
                    string blockName = subFlowSheet.Name.Replace("Flow_", "");
                    int headerCnt = subFlowSheet.Rows.Count(p => p.Parameter != null && p.Parameter.ContainsIgnoreCase("_HEADER"));
                    int footerCnt = subFlowSheet.Rows.Count(p => p.Parameter != null && p.Parameter.ContainsIgnoreCase("_FOOTER"));

                    int headerIndex = subFlowSheet.Rows.FindIndex(p => p.Opcode.Equals("print", StringComparison.CurrentCultureIgnoreCase));
                    int footerIndex = subFlowSheet.Rows.FindLastIndex(p => p.Opcode.Equals("print", StringComparison.CurrentCultureIgnoreCase));
                    int printCount = subFlowSheet.Rows.Count(p => p.Opcode.Equals("print", StringComparison.CurrentCultureIgnoreCase));

                    if (footerCnt == 0 && headerIndex != -1 && footerIndex != -1 && printCount >= 2)
                    {
                        subFlowSheet.Rows.Insert(footerIndex + 1, GenFooterHeaderFlow(blockName, "Footer"));
                    }

                    if (headerCnt == 0 && headerIndex != -1 && footerIndex != -1 && printCount >= 2)
                    {
                        subFlowSheet.Rows.Insert(headerIndex, GenFooterHeaderFlow(blockName, "Header"));
                    }
                    #endregion

                    List<FlowRow> flowRows = subFlowSheet.Rows.FindAll(p => p.Parameter != null && _regex.IsMatch(p.Parameter));
                    if (flowRows.Count != 0)
                    {
                        instanceRows.AddRange(flowRows.Select(row => GenFooterHeaderInst(row.Parameter)));
                    }
                }
            }

            return instanceRows;
        }

        private List<InstanceRow> CollectAllInstanceRows()
        {
            var all = new List<InstanceRow>();
            foreach (KeyValuePair<string, InstanceSheet> keyValuePair in TestProgram.IgxlWorkBk.InsSheets)
            {
                for (int index = 0; index < keyValuePair.Value.Rows.Count; index++)
                {
                    InstanceRow row = keyValuePair.Value.Rows[index];
                    row.RowNum = index + 4;
                }
                all.AddRange(keyValuePair.Value.Rows);
            }

            return all;
        }

        private List<InstanceRow> DistributeFooterHeaderInstanceRows(List<InstanceRow> instanceRows, List<InstanceRow> all, bool flag)
        {
            var instanceRowsInCommon = new List<InstanceRow>();
            foreach (InstanceRow instanceRow in instanceRows)
            {
                int headerIndex = instanceRow.TestName.ToUpper().IndexOf("_HEADER", StringComparison.Ordinal);
                int footerIndex = instanceRow.TestName.ToUpper().IndexOf("_FOOTER", StringComparison.Ordinal);
                string sheetName = "";
                if (headerIndex > 0)
                {
                    sheetName = instanceRow.TestName.Substring(0, headerIndex);
                }

                if (footerIndex > 0)
                {
                    sheetName = instanceRow.TestName.Substring(0, footerIndex);
                }

                sheetName = "TestInst_" + sheetName;
                string sheet = TestProgram.IgxlWorkBk.InsSheets.Keys.ToList().FirstOrDefault(x => Path.GetFileName(x).ToUpper().Equals(sheetName.ToUpper()));
                if (sheet == null)
                {
                    sheetName = sheetName.Replace("TestInst_HARDIP", "Inst_H");
                    sheet = TestProgram.IgxlWorkBk.InsSheets.Keys.ToList().FirstOrDefault(x => Path.GetFileName(x).ToUpper().Equals(sheetName.ToUpper()));
                }
                if (!all.Exists(x => x.TestName.Equals(instanceRow.TestName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    if (sheet != null && TestProgram.IgxlWorkBk.InsSheets.ContainsKey(sheet))
                    {
                        TestProgram.IgxlWorkBk.InsSheets[sheet].Rows.Insert(0, instanceRow);
                    }
                    else
                    {
                        if (flag)
                        {
                            instanceRowsInCommon.Add(instanceRow);
                        }
                    }
                }
            }

            return instanceRowsInCommon;
        }

        public string GenFooterHeaderName(string block, string footerHeader)
        {
            return block + '_' + footerHeader + "_1";
        }

        public FlowRow GenFooterHeaderFlow(string block, string footerHeader, string tn = "")
        {
            var outputRow = new FlowRow
            {
                Opcode = OpCode.Test,
                Parameter = GenFooterHeaderName(block, footerHeader),
                TNum = tn
            };
            return outputRow;
        }

        public InstanceRow GenFooterHeaderInst(string instName)
        {

            var outputRow = new InstanceRow { TestName = instName, VbtType = "VBT", ArgList = "PrintInfo" };
            if (instName.IndexOf("_Footer_", StringComparison.CurrentCultureIgnoreCase) >= 0)
            {
                outputRow.VbtName = "Print_Footer";
                outputRow.Args.Add(instName.Substring(0, instName.IndexOf("_Footer_", StringComparison.CurrentCultureIgnoreCase)));
            }
            else if (instName.IndexOf("_Header_", StringComparison.CurrentCultureIgnoreCase) >= 0)
            {
                outputRow.VbtName = "Print_Header";
                outputRow.Args.Add(instName.Substring(0, instName.IndexOf("_Header_", StringComparison.CurrentCultureIgnoreCase)));
            }
            else
            {
                outputRow.VbtName = "";
            }
            return outputRow;
        }

        private void ReplaceCsCommonInstanceFromConfig(string commonSheet, ref InstCommonOutput instCommonSheet)
        {
            if (string.IsNullOrEmpty(commonSheet) || SettingStatic.BasicConfigWorkbook.Worksheets[$"{commonSheet}_Cs"] == null)
            {
                return;
            }

            ExcelWorksheet instanceCommonCsWorksheet = SettingStatic.BasicConfigWorkbook.Worksheets[$"{commonSheet}_Cs"];
            InstanceSheet csCommonInstanceSheet = new ReadInstanceSheet().ReadSheet(instanceCommonCsWorksheet);
            var tempAllInstanceList = new List<InstanceRow>();
            var vbtWhiteList = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "GetBarCode", "EEPROM_Write_Hw", "Read_EEPROM_ProberTemp_Compare" };
            tempAllInstanceList.AddRange(instCommonSheet.InstanceRows);
            tempAllInstanceList.AddRange(csCommonInstanceSheet.Rows);
            tempAllInstanceList.ForEach(x => x.VbtName = RenewFunctionName(x.VbtName));
            var newInstanceList = new List<InstanceRow>();
            IEnumerable<IGrouping<string, InstanceRow>> groupByInstanceName = tempAllInstanceList.GroupBy(x => x.TestName.ToUpper());
            foreach (IGrouping<string, InstanceRow> group in groupByInstanceName)
            {
                InstanceRow csInstance = group.FirstOrDefault(x => string.Equals(x.VbtType, ".NET", StringComparison.OrdinalIgnoreCase));
                if (csInstance != null)
                {
                    if (TestProgram.VbtFunctionLib.VbtLib.Any(x => string.Equals(x.FullFunctionName, csInstance.VbtName, StringComparison.OrdinalIgnoreCase) && string.Equals(x.Type, ".NET", StringComparison.OrdinalIgnoreCase)))
                    {
                        newInstanceList.Add(csInstance);
                    }
                }
                else if (vbtWhiteList.Contains(group.FirstOrDefault()?.VbtName))
                {
                    newInstanceList.Add(group.FirstOrDefault());
                }
            }
            instCommonSheet.InstanceRows = newInstanceList;
        }

        private string RenewFunctionName(string functionName)
        {
            return TestProgram.VbtFunctionLib.GetFunctionByName(functionName.Split('.').LastOrDefault(), "").FullFunctionName;
        }

        private void AddStoreRefClockFlow()
        {
            SubFlowSheet initFlowSheet = TestProgram.IgxlWorkBk.SubFlowSheets.Values.ToList().Find(x => x.Name.Equals("Flow_Table_Main_Init_Flows", StringComparison.OrdinalIgnoreCase));
            if (initFlowSheet == null)
            {
                return;
            }

            var flowRow = new FlowRow();
            int index = initFlowSheet.Rows.FindLastIndex(p => p.Opcode.Equals("return", StringComparison.CurrentCultureIgnoreCase));
            if (index != -1)
            {
                flowRow.Opcode = "Test";
                flowRow.Parameter = "StoreReferenceClockName";
                initFlowSheet.Rows.Insert(index, flowRow);
            }
        }

        private void AddStoreRefClockInstance(ref InstCommonOutput instCommonSheet)
        {
            var instRow = new InstanceRow();
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("StoreReferenceClockName", "basic", true);
            function.SetParamValue("clockName", "Refclk");
            instRow.TestName = "StoreReferenceClockName";
            instRow.VbtName = function.FullFunctionName;
            instRow.VbtType = function.Type;
            instRow.ArgList = function.Parameters;
            instRow.Args = function.ArgList;
            instCommonSheet.InstanceRows.Add(instRow);
        }

        private void AddStoreReferenceClockName(ref InstCommonOutput instCommonSheet)
        {
            bool needGen = false;
            if (TestProgram.IgxlWorkBk.PortMapSheets.Values.Any())
            {
                IEnumerable<PortRow> portRows = TestProgram.IgxlWorkBk.PortMapSheets.Values.SelectMany(x => x.Rows).SelectMany(x => x.PortRows);
                if (portRows.Any(x => x.FunctionName.Equals("Refclk", StringComparison.OrdinalIgnoreCase)))
                {
                    needGen = true;
                }
            }

            if (!needGen)
            {
                return;
            }

            if (!instCommonSheet.InstanceRows.Any(x => x.TestName.Equals("StoreReferenceClockName", StringComparison.OrdinalIgnoreCase))
                && !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder)
                && TestProgram.VbtFunctionLib.VbtLib.Any(x => x.FunctionName.Equals("StoreReferenceClockName", StringComparison.OrdinalIgnoreCase))
                && TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlex))
            {
                AddStoreRefClockInstance(ref instCommonSheet);
                AddStoreRefClockFlow();
            }
        }

        private void AddInitializeModulesFromInstance(ref InstCommonOutput instCommonSheet)
        {
            var instanceRow = new InstanceRow();
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.ConInitializeModulesFromInstance, "", true);
            if (function.Type == ".NET")
            {
                instanceRow.TestName = FuncNameConst.ConInitializeModulesFromInstance;
                instanceRow.VbtType = function.Type;
                instanceRow.VbtName = function.FullFunctionName;
                instanceRow.ArgList = function.Parameters;
                if (TestProgram.IgxlWorkBk.PinMapPair.Value.TryGetGroup("All_DiffPairs", out PinGroup target))
                {
                    function.SetParamValue("hasDefaultIODifferential", target != null ? "TRUE" : "FALSE");
                    instanceRow.Args = function.ArgList;
                    instCommonSheet.InstanceRows.Add(instanceRow);
                }
            }
        }
    }
}
