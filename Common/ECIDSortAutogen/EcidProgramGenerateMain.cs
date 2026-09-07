using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

namespace ECIDSortAutogen
{
    public partial class EcidProgramGenerateMain
    {
        [GeneratedRegex(@"\w+JTAGRead", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        [GeneratedRegex("^ECID_JTAGRead", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex("SyntaxCheck", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex("ECID_Check", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex("ShowECIDData", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex4();
        [GeneratedRegex("ShowECIDData", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex5();

        public static void GenSyntaxCheckFlow(SubFlowSheet subFlowSheet)
        {
            var jtagReadRow = subFlowSheet.Rows.Where(p => MyRegex1().IsMatch(p.Parameter)).ToList();
            var syntaxCheckRow = subFlowSheet.Rows.Where(p => MyRegex2().IsMatch(p.Parameter)).ToList();
            var syntacCheckBinTableRow = subFlowSheet.Rows.Where(p => p.Opcode.EqualsIgnoreCase("bintable") && MyRegex3().IsMatch(p.Parameter)).ToList();
            if (jtagReadRow.Count != 0)
            {
                if (syntaxCheckRow.Count != 0)
                {
                    if (subFlowSheet.Rows.IndexOf(jtagReadRow.LastOrDefault()!) > subFlowSheet.Rows.IndexOf(syntaxCheckRow.LastOrDefault()!)) //If Syntax_Check before JTAG_Read, inserting Syntax_Check after JTAG_Read.
                    {
                        var insertSyntaxCheckRow = new FlowRow
                        {
                            Job = "FT",
                            Parameter = syntaxCheckRow.LastOrDefault()!.Parameter,
                            Opcode = syntaxCheckRow.LastOrDefault()!.Opcode,
                            FailAction = syntaxCheckRow.LastOrDefault()!.FailAction,
                            Enable = "ECIDSort_Enable",
                        };

                        var insertSyntaxCheckBinTableRow = new FlowRow
                        {
                            Job = "FT",
                            Parameter = syntacCheckBinTableRow.LastOrDefault()!.Parameter,
                            Enable = "ECIDSort_Enable",
                            Opcode = syntacCheckBinTableRow.LastOrDefault()!.Opcode,
                        };

                        subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(jtagReadRow.LastOrDefault()!) + 2, insertSyntaxCheckRow);
                        subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(jtagReadRow.LastOrDefault()!) + 3, insertSyntaxCheckBinTableRow);

                    }
                }
                else
                {
                    string errMsg = $"The Test Item: {"ECID_SyntaxCheck"} can not found in {subFlowSheet.Name} table.";
                    Response.Report(errMsg, EnumMessageLevel.Error, 0);
                }
            }
            else
            {
                var jtagReadFlowRow = new FlowRow
                {
                    Job = "FT",
                    Parameter = "ECID_JTAGRead",
                    Enable = "ECIDSort_Enable",
                    Opcode = "Test"
                };
                var jtagReadBinTableRow = new FlowRow
                {
                    Job = "FT",
                    Parameter = "Bin_EFUSE_ECID_JTAG",
                    Enable = "BinTable&&ECIDSort_Enable",
                    Opcode = "nop"
                };

                foreach (FlowRow row in subFlowSheet.Rows)
                {
                    if (MyRegex4().IsMatch(row.Parameter))
                    {
                        subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(row), jtagReadFlowRow);
                        subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(row), jtagReadBinTableRow);
                        break;
                    }
                }
                if (syntaxCheckRow.Count == 0)
                {
                    string errMsg = $"The Test Item: {"ECID_SyntaxCheck"} can not found in {subFlowSheet.Name} table.";
                    Response.Report(errMsg, EnumMessageLevel.Error, 0);
                }
                else if (subFlowSheet.Rows.IndexOf(syntaxCheckRow.LastOrDefault()!) < subFlowSheet.Rows.IndexOf(jtagReadFlowRow))
                {
                    var insertSyntaxCheckRow = new FlowRow
                    {
                        Job = "FT",
                        Parameter = syntaxCheckRow.LastOrDefault()!.Parameter,
                        Opcode = syntaxCheckRow.LastOrDefault()!.Opcode,
                        FailAction = syntaxCheckRow.LastOrDefault()!.FailAction,
                        Enable = "ECIDSort_Enable",
                    };

                    var insertSyntaxCheckBinTableRow = new FlowRow
                    {
                        Job = "FT",
                        Parameter = syntacCheckBinTableRow.LastOrDefault()!.Parameter,
                        Enable = "ECIDSort_Enable",
                        Opcode = syntacCheckBinTableRow.LastOrDefault()!.Opcode,
                    };

                    subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(jtagReadFlowRow) + 2, insertSyntaxCheckRow);
                    subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(jtagReadFlowRow) + 3, insertSyntaxCheckBinTableRow);
                }
            }
        }

        public static void GenEnableToMainInitEnableWd(SubFlowSheet subFlowSheet, string para, string op, string columnA)
        {
            if (subFlowSheet == null)
            {
                return;
            }

            var initFlowRow = new FlowRow { Enable = para, Opcode = "nop", ColumnA = columnA };
            var flowRow = new FlowRow
            {
                Parameter = para,
                Opcode = op,
                ColumnA = columnA
            };

            string stopPrint = "Flow_Table_Main_Init_EnableWd Stop";
            IEnumerable<FlowRow> printStopRow = subFlowSheet.Rows.Where(p => p.Opcode.EqualsIgnoreCase("print") && Regex.IsMatch(p.Parameter, stopPrint)).Select(p => p);
            if (printStopRow.Any())
            {
                subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(printStopRow.FirstOrDefault()!), initFlowRow);
                subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(printStopRow.FirstOrDefault()!), flowRow);
            }
            else
            {
                FlowRow returnRow = subFlowSheet.Rows.Where(p => p.Opcode.EqualsIgnoreCase("return")).Select(p => p).FirstOrDefault()!;
                subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(returnRow), initFlowRow);
                subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(returnRow), flowRow);
            }
        }

        public static void GenEcidSortingFlow(SubFlowSheet subFlowSheet)
        {
            foreach (FlowRow row in subFlowSheet.Rows)
            {
                if (MyRegex5().IsMatch(row.Parameter))
                {
                    var flowRow = new FlowRow
                    {
                        Parameter = "ECID_Sorting",
                        Enable = "ECIDSort_Enable",
                        Opcode = "Test",
                        Job = "FT",
                    };
                    subFlowSheet.InsertRow(subFlowSheet.Rows.IndexOf(row), flowRow);
                    break;
                }
            }
        }

        public static void GenEcidSortingInst(InstanceSheet instanceSheet)
        {
            foreach (InstanceRow row in instanceSheet.Rows)
            {
                if (_regex.IsMatch(row.TestName))
                {
                    var instanceRow = new InstanceRow
                    {
                        TestName = "ECID_Sorting",
                        VbtType = "VBT",
                        VbtName = "BinSorting_Compare_FT_ECID_S"
                    };
                    instanceSheet.AddRow(instanceRow);
                    break;
                }
            }
        }
    }
}
