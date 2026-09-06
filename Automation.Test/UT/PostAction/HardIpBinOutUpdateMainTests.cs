using System.Collections.Generic;
using System.IO;

using Automation.Utility.TpUpdate.HardIPBinoutTPUpdate;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OpCode = IgxlLib.IgxlConst.OpCode;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class HardIpBinOutUpdateMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void HardIpBinOutTpUpdateMainTest()
        {
            string subName = "BinOut";

            string inputPath = Path.Combine(InputPath, "PostAction", subName);
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string binOutReportFileName = Path.Combine(inputPath, "Komodo_A0_BinOut_V06A_X_HardIP_1.xlsx");

            var subFlowSheets = new Dictionary<string, SubFlowSheet>
            {
                {
                    Path.Combine(outputPath, "Flow_WALKINGZ_3"),
                    new SubFlowSheet("Flow_WALKINGZ_3")
                    {
                        Rows = new FlowRows
                        {
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode = OpCode.Test,
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_P_25C_FT1",
                                FailAction = "F_WALKINGZ_3",
                                Job = "CP2",
                                TName = "CONTINUITYNEG",
                                Env = "Env"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode = OpCode.UseLimit,
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_P_25C_FT1",
                                FailAction = "F_WALKINGZ_3",
                                Job = "CP2",
                                TName = "CONTINUITYNEG",
                                Env = "Env"
                            },

                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode = OpCode.Test,
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_N_25C_CP1",
                                FailAction = "F_WALKINGZ_3",
                                Job = "CP2",
                                TName = "CONTINUITYNEG",
                                Env = "Env"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode = OpCode.Test,
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_N_25C_CP1",
                                FailAction = "F_WALKINGZ_3",
                                Job = "CP2",
                                TName = "CONTINUITYNEG",
                                Env = "Env"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode = OpCode.UseLimit,
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_N_25C_CP1",
                                FailAction = "F_WALKINGZ_3",
                                Job = "CP2",
                                TName = "CONTINUITYNEG",
                                Env = "Env"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode = OpCode.UseLimit,
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_N_25C_CP1",
                                FailAction = "F_WALKINGZ_3",
                                Job = "CP2",
                                TName = "CONTINUITYNEG",
                                Env = "Env"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode =OpCode.BinTable,
                                Parameter = "BinName3",
                                Job = "CP2",
                                Part = "Part",
                                Env = "Env"
                            },

                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode = OpCode.Test,
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_N_25C_CP1",
                                FailAction = "F_WALKINGZ_3",
                                Job = "CP2",
                                TName = "CONTINUITYNEG",
                                Env = "Env"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode = OpCode.UseLimit,
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_N_25C_CP1",
                                FailAction = "F_WALKINGZ_3",
                                Job = "CP2",
                                TName = "CONTINUITYNEG",
                                Env = "Env"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_3",
                                Opcode =OpCode.BinTable,
                                Parameter = "BinName3",
                                Job = "CP2",
                                Part = "Part",
                                Env = "Env"
                            }
                        }
                    }
                },
                {
                    Path.Combine(outputPath, "Flow_WALKINGZ_2"),
                    new SubFlowSheet("Flow_WALKINGZ_2")
                    {
                        Rows = new FlowRows
                        {
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_2",
                                Opcode = "test",
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_N_25C_CP1",
                                FailAction = "F_WALKINGZ_2",
                                Job = "CP1",
                                TName = "CONTINUITYNEG"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_2",
                                Opcode =OpCode.BinTable,
                                Parameter = "BinName2",
                                Job = "CP1",
                                Part = "Part"
                            }
                        }
                    }
                },
                {
                    Path.Combine(outputPath, "Flow_WALKINGZ_1"),
                    new SubFlowSheet("Flow_WALKINGZ_1")
                    {
                        Rows = new FlowRows
                        {
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_1",
                                Opcode = "test",
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_N_25C_CP1",
                                FailAction = "F_WALKINGZ_1",
                                Job = "CP1",
                                TName = "CONTINUITYNEG"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_1",
                                Opcode =OpCode.BinTable,
                                Parameter = "BinName1",
                                Env = "Env"
                            }
                        }
                    }
                },
                {
                    Path.Combine(outputPath, "Flow_WALKINGZ_0"),
                    new SubFlowSheet("Flow_WALKINGZ_0")
                    {
                        Rows = new FlowRows
                        {
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_0",
                                Opcode = "test",
                                Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_N_25C_CP1",
                                FailAction = "F_WALKINGZ_0",
                                TName = "CONTINUITYNEG"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_WALKINGZ_0",
                                Opcode =OpCode.BinTable,
                                Parameter = "BinName0",
                                Env = "Env",
                                Job = "CP1"
                            }
                        }
                    }
                },
                {
                    Path.Combine(outputPath, "Flow_DC"),
                    new SubFlowSheet("Flow_DC")
                    {
                        Rows = new FlowRows
                        {
                            new FlowRow
                            {
                                SheetName = "Flow_DC",
                                Opcode = "test",
                                Parameter = "DC_CONTINUITY_PPMU_CONTINUITY_NEG_CP1"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_DC",
                                Opcode = "use-limit",
                                Parameter = "DC_CONTINUITY_PPMU_CONTINUITY_NEG_CP1",
                                TName=  "CONTINUITYNEG"
                            },
                            new FlowRow
                            {
                                SheetName = "Flow_DC",
                                Opcode =OpCode.BinTable,
                                Parameter = "BinName0"
                            }
                        }
                    }
                }
            };
            var binTableSheets = new List<BinTableSheet>
            {
                new( "Bin_Table_HardIP")
                {
                    Rows =
                    [
                        new() {Name="BinName0", ItemList="F_1,F_WALKINGZ_0"},
                        new() {Name="BinName1", ItemList="F_1,F_WALKINGZ_1"},
                        new() {Name="BinName2", ItemList="F_1,F_WALKINGZ_2"},
                        new() {Name="BinName3", ItemList="F_1,F_WALKINGZ_3"}
                    ]
                }
            };

            new HardIpBinOutTpUpdateMain(binOutReportFileName, subFlowSheets, binTableSheets).WorkFlow();
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
