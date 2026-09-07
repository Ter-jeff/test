using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;

using TagDiff.Core.Const;
using TagDiff.Core.Output;
using TagDiff.Core.Static;

using TagDiffCore.Regs;

namespace TagDiff.Core.Utility
{
    public static class ClassifyCompareReport
    {
        public static void SetCategory(EnumSheetType enumSheetType, string sheetName, ref CompareReport compareReport)
        {
            string[] tokens = sheetName.Split('_');
            if (TagDiffStatic.IsNotTracked(sheetName))
            {
                UpdateResult(ref compareReport, Category.NotTracked, sheetName);
                return;
            }

            if (tokens.Any(t => t.EqualsIgnoreCase("Char") || t.EqualsIgnoreCase("CZ") || t.EqualsIgnoreCase("LVCC")))
            {
                UpdateResult(ref compareReport, Category.Cz, SetBlock(enumSheetType, sheetName));
                return;
            }

            if (sheetName.ContainsIgnoreCase("_HardIp"))
            {
                UpdateResult(ref compareReport, Category.Hardip, SetBlock(enumSheetType, sheetName));
                return;
            }

            switch (enumSheetType)
            {
                case EnumSheetType.DTFlowtableSheet:
                    SetCategoryForFlow(sheetName, ref compareReport);
                    break;
                case EnumSheetType.DTTestInstancesSheet:
                    HandleTestInstances(sheetName, compareReport);
                    break;
                case EnumSheetType.DTTimesetBasicSheet:
                    UpdateResult(ref compareReport, Category.Basic, Block.TimeSets);
                    break;
                case EnumSheetType.DTPinMap:
                    UpdateResult(ref compareReport, Category.Basic, Block.PinMap);
                    break;
                case EnumSheetType.DTChanMap:
                    UpdateResult(ref compareReport, Category.Basic, Block.ChannelMap);
                    break;
                case EnumSheetType.DTPortMapSheet:
                    UpdateResult(ref compareReport, Category.Basic, Block.PortMap);
                    break;
                case EnumSheetType.DTACSpecSheet:
                    UpdateResult(ref compareReport, Category.Basic, Block.AcSpecs);
                    break;
                case EnumSheetType.DTDCSpecSheet:
                    UpdateResult(ref compareReport, Category.Basic, Block.DcSpecs);
                    break;
                case EnumSheetType.DTBintablesSheet:
                    UpdateResult(ref compareReport, Category.Basic, Block.BinTable);
                    break;
                case EnumSheetType.DTCharacterizationSheet:
                    string charBlock = sheetName.EndsWithIgnoreCase("_2D") ? Block.DevChar2D : Block.DevChar1D;
                    UpdateResult(ref compareReport, Category.Basic, charBlock);
                    break;
                case EnumSheetType.DTGlobalSpecSheet:
                    UpdateResult(ref compareReport, Category.Basic, Block.GlobalSpec);
                    break;
                case EnumSheetType.DTLevelSheet:
                    UpdateResult(ref compareReport, Category.Basic, Block.LevelSheets);
                    break;
                case EnumSheetType.DTUnknown:
                    HandleUnknown(sheetName, compareReport);
                    break;
                case EnumSheetType.DTJobListSheet:
                    TagDiffStatic.AddNotTracked(sheetName);
                    UpdateResult(ref compareReport, Category.NotTracked, sheetName);
                    break;
                case EnumSheetType.DTReferencesSheet:
                    UpdateResult(ref compareReport, Category.Basic, sheetName);
                    break;
                case EnumSheetType.DTPatternSetSheet:
                case EnumSheetType.DTPatternSubroutineSheet:
                    UpdateResult(ref compareReport, Category.Other, Block.PatternSet);
                    break;
                default:
                    HandleNameBasedFallback(sheetName, compareReport);
                    break;
            }
        }

        private static void HandleUnknown(string sheetName, CompareReport compareReport)
        {
            if (sheetName.ContainsIgnoreCase("EFUSE_BitDef_Table"))
            {
                compareReport.Category = Category.Efuse;
                compareReport.Block = sheetName;
            }
            else if (sheetName.StartsWithIgnoreCase("Reg_Assign"))
            {
                compareReport.Category = Category.RegAssign;
                compareReport.Block = Block.RegAssign;
            }
            else
            {
                compareReport.Category = Category.Other;
                compareReport.Block = sheetName;
            }
        }

        private static void HandleNameBasedFallback(string sheetName, CompareReport compareReport)
        {
            string[] atpgPrefixes = ["SOC", "CPU", "GPU", "Harvest", "NonBinCut"];
            foreach (string prefix in atpgPrefixes)
            {
                if (sheetName.StartsWithIgnoreCase(prefix))
                {
                    compareReport.Category = Category.Atpg;
                    return;
                }
            }

            compareReport.Category = Category.Other;
            compareReport.Block = sheetName;
        }

        private static void HandleTestInstances(string sheetName, CompareReport compareReport)
        {
            string str = sheetName.Trim();
            if (str.StartsWithIgnoreCase("Inst_"))
            {
                str = str[5..];
            }
            else if (str.StartsWithIgnoreCase("TestInst_"))
            {
                str = str[9..];
            }

            compareReport.Category = GetCategoryFromInstanceName(str);
            compareReport.Block = Block.TestInst;
        }

        private static string GetCategoryFromInstanceName(string name)
        {
            // ATPG Group
            if (name.ContainsIgnoreCase("ClockCheck") || name.ContainsIgnoreCase("_BI_") || name.ContainsIgnoreCase("SocScan") || name.ContainsIgnoreCase("SocMbist") || name.ContainsIgnoreCase("GfxScan") || name.ContainsIgnoreCase("GfxMbist") || name.ContainsIgnoreCase("CpuScan") || name.ContainsIgnoreCase("CpuMbist") || name.ContainsIgnoreCase("Harvest") || name.ContainsIgnoreCase("Non_BinCut"))
            {
                return Category.Atpg;
            }

            // DC Group
            if (name.ContainsIgnoreCase("EVS") || name.ContainsIgnoreCase("DC_Conti") || name.ContainsIgnoreCase("DctestIds") || name.ContainsIgnoreCase("DCTEST_IDS") || name.ContainsIgnoreCase("DCTEST_GPIO"))
            {
                return Category.Dc;
            }

            // EFUSE Group
            if (name.ContainsIgnoreCase("eFuse"))
            {
                return Category.Efuse;
            }

            // RTOS Group
            if (name.ContainsIgnoreCase("SPIROM") || name.ContainsIgnoreCase("RTOS"))
            {
                return Category.Rtos;
            }

            // Single matches
            if (Reg.RegexHardip().IsMatch(name))
            {
                return Category.Hardip;
            }

            if (name.ContainsIgnoreCase("Vddbinning"))
            {
                return Category.BinCut;
            }

            return Category.Basic;
        }

        private static void UpdateResult(ref CompareReport compareReport, string category, string block)
        {
            compareReport.Category = category;
            compareReport.Block = block;
        }

        private static string SetBlock(EnumSheetType enumSheetType, string sheetName)
        {
            if (enumSheetType == EnumSheetType.DTFlowtableSheet)
            {
                return Block.Flow;
            }
            else if (enumSheetType == EnumSheetType.DTTestInstancesSheet)
            {
                return Block.TestInst;
            }
            else
            {
                return sheetName;
            }
        }

        public static void SetCategoryForFlow(string sheetName, ref CompareReport compareReport)
        {
            if (compareReport == null)
            {
                return;
            }

            compareReport.Block = Block.Flow;
            if (sheetName.EqualsIgnoreCase("Flow_SOF_SA"))
            {
                compareReport.Category = Category.Atpg;
                return;
            }

            var mappings = new[]
            {
                new { Prefix = "Flow_ClockCheck", Category = Category.Dc },
                new { Prefix = "Flow_EVS", Category = Category.Dc },
                new { Prefix = "Flow_HARDIP", Category = Category.Hardip },
                new { Prefix = "Flow_VddBinning", Category = Category.BinCut },
                new { Prefix = "Flow_PostBinCut", Category = Category.BinCut },
                new { Prefix = "Flow_SOC_", Category = Category.Atpg },
                new { Prefix = "Flow_CPU_", Category = Category.Atpg },
                new { Prefix = "Flow_GPU_", Category = Category.Atpg },
                new { Prefix = "Flow_S_BI_", Category = Category.Atpg },
                new { Prefix = "Flow_C_BI_", Category = Category.Atpg },
                new { Prefix = "Flow_G_BI_", Category = Category.Atpg },
                new { Prefix = "Flow_efuse_", Category = Category.Efuse },
                new { Prefix = "Flow_DcTest", Category = Category.Dc },
                new { Prefix = "Flow_DC_", Category = Category.Dc },
                new { Prefix = "Flow_Rtos_", Category = Category.Rtos },
                new { Prefix = "Main_Flow", Category = Category.Basic }
            };

            foreach (var m in mappings)
            {
                if (sheetName.StartsWithIgnoreCase(m.Prefix))
                {
                    compareReport.Category = m.Category;
                    if (m.Prefix.EqualsIgnoreCase("Main_Flow"))
                    {
                        compareReport.Block = Block.MainFlow;
                    }
                    return;
                }
            }

            if (sheetName.StartsWithIgnoreCase("Flow_M") && sheetName.EndsWithIgnoreCase("_BV"))
            {
                compareReport.Category = Category.BinCut;
                return;
            }

            compareReport.Category = Category.Basic;
        }
    }
}
