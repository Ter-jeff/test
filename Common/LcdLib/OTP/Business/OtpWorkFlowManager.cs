using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Singleton;
using Automation.Static;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LcdLib.Const;
using LcdLib.InputManager.Data;
using LcdLib.OTP.Reader;

using TestPlanLib.Basic;
using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace LcdLib.OTP.Business
{
    public partial class OtpWorkFlowManager(OtpInputData otpInputData)
    {
        [GeneratedRegex("^Flow_", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("^Flow_", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();

        protected readonly OtpInputData OtpInputData = otpInputData;

        public virtual void WorkFlow()
        {
            const string otpSheetName = OtpConst.FlowSheetName;
            const string otpBlankSheetName = OtpConst.FlowBlankCheckSheetName;
            var flagList = new List<string>();
            var otpFlowSheets = new List<SubFlowSheet>();
            var otpInstanceSheet = new InstanceSheet(OtpConst.TestInstSheet);

            otpFlowSheets.Add(GenerateFlow(otpSheetName, OtpInputData.OtpPatternRows, ref flagList));
            otpFlowSheets.Add(GenerateBlankCheckFlow(otpBlankSheetName, OtpInputData.OtpPatternRows, ref flagList));

            otpInstanceSheet.Rows.AddRange(GenerateOtpInstanceRow(OtpInputData.OtpPatternRows));
            otpInstanceSheet.AddHeaderFooter(otpSheetName);

            foreach (SubFlowSheet flow in otpFlowSheets)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirEfuse, flow);
            }

            otpInstanceSheet.RemoveDuplicateInstance();
            GenOtpInstanceSheet(otpInstanceSheet);

            GenBinTable(flagList);
        }

        private static SubFlowSheet GenerateBlankCheckFlow(string sheetName, IEnumerable<OtpPatternRow> otpPatternRows, ref List<string> flagList)
        {
            var flowSheet = new SubFlowSheet(sheetName);
            //flowSheet.FlowRows.Add(NwireSingleton.Instance().SettingInfo.GetNwireCall(flowSheet.Name));
            string headerFooterName = MyRegex().Replace(sheetName, "");
            flowSheet.AddHeaderRow(Combination.CombineByUnderLine([OtpConst.Otp, "Blank", "Check"]));
            flowSheet.AddPrintStartRow(headerFooterName);

            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Test,
                Parameter = "Relay_On_Desc_Bit_Check",
            });

            var readItems = otpPatternRows.Where(x => x.OtpReadWrite.Equals(EnumOtpReadWrite.Read)).OrderBy(x => x.RowNum)
                .ToList();

            # region OTP BIT CHECK

            OtpPatternRow? blankCheckItem = readItems.FirstOrDefault(x => x.OtpReadType.Equals(EnumOtpReadType.Multishot4Byte));
            if (blankCheckItem != null)
            {
                //BlankCheck
                string flag = Combination.CombineByUnderLine(["F", blankCheckItem.Block, blankCheckItem.Mode, blankCheckItem.Item, blankCheckItem.PayloadList.First(), "N", "Flag"]);
                flowSheet.AddTestAndBinTableRow(Combination.CombineByUnderLine([OtpConst.OtpBlankCheck, OtpConst.OtpInstancePostfixMultiShot4Byte]), "", "", "", flag, false);
                flowSheet.AddRow(new FlowRow
                {
                    Opcode = OpCode.BinTable,
                    Enable = "HardIPBin",
                    Parameter = flag.Replace("F_", "Bin_")
                });
                flagList.Add(flag);
            }

            # endregion

            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.If,
                Parameter = OtpConst.VarRunTrim + "=0",
            });

            # region OTP Read

            foreach (OtpPatternRow item in readItems)
            {
                if (item.OtpWriteType.Equals(EnumOtpReadType.None))
                {
                    continue;
                }

                string flag =
                    Combination.CombineByUnderLine(["F", item.Block, item.Mode, item.Item, item.PayloadList.First(), "N", "Flag"]);

                if (item.OtpReadType.Equals(EnumOtpReadType.Multishot4Byte))
                {
                    flowSheet.AddTestAndBinTableRow(Combination.CombineByUnderLine([OtpConst.OtpInstanceReadAll, item.OtpReadType.ToString()]), "", "", "", flag, false);
                    flowSheet.AddRow(new FlowRow
                    {
                        Opcode = OpCode.BinTable,
                        Enable = "HardIPBin",
                        Parameter = flag.Replace("F_", "Bin_")
                    });
                    flagList.Add(flag);
                }
                else
                {
                    flowSheet.AddNopTestAndBinTableRow(Combination.CombineByUnderLine([OtpConst.OtpInstanceReadAll, item.OtpReadType.ToString()]), "", "", "", flag, false);
                    flowSheet.AddRow(new FlowRow
                    {
                        Opcode = OpCode.BinTable,
                        Enable = "HardIPBin",
                        Parameter = flag.Replace("F_", "Bin_")
                    });
                    flagList.Add(flag);
                }
            }
            #endregion

            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.EndIf,
            });

            flowSheet.AddPrintEndRow(headerFooterName);
            flowSheet.AddFooterRow(Combination.CombineByUnderLine([OtpConst.Otp, "Blank", "Check"]));
            flowSheet.AddReturnRow();
            return flowSheet;
        }

        private static SubFlowSheet GenerateFlow(string sheetName, List<OtpPatternRow> otpPatternRows, ref List<string> flagList)
        {
            var flowSheet = new SubFlowSheet(sheetName);
            //flowSheet.FlowRows.Add(NwireSingleton.Instance().SettingInfo.GetNwireCall(flowSheet.Name));
            string headerFooterName = MyRegex1().Replace(sheetName, "");
            flowSheet.AddHeaderRow(OtpConst.Otp);
            flowSheet.AddPrintStartRow(headerFooterName);

            //SW CALC
            flowSheet.AddRow(new FlowRow { Opcode = OpCode.Test, Parameter = OtpConst.OtpInstanceSwCrcCalc });

            //Check
            string flag = Combination.CombineByUnderLine(["F", OtpConst.OtpInstanceCheckDefaultReal]);
            flowSheet.AddTestAndBinTableRow(OtpConst.OtpInstanceCheckDefaultReal, "", "", "", flag, false);
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.BinTable,
                Enable = "HardIPBin",
                Parameter = flag.Replace("F_", "Bin_")
            });
            flagList.Add(flag);

            //OTP
            var functionalItems = otpPatternRows.Where(x => x.OtpReadWrite.Equals(EnumOtpReadWrite.None) && x.BankName.Equals(EnumOtpBankName.None)).OrderBy(x => x.RowNum)
                .ToList();
            foreach (OtpPatternRow item in functionalItems)
            {
                flag =
                    Combination.CombineByUnderLine(["F", item.Block, item.Mode, item.Item, item.PayloadList.First(), "N", "Flag"]);
                string parameter = Combination.CombineByUnderLine([item.Block, item.Mode, item.Item, item.PayloadList.First(), "NV"]).ToUpper();
                flowSheet.AddRow(new FlowRow { Opcode = OpCode.Test, Enable = "HardIP_NV", Parameter = parameter, FailAction = flag });
                flagList.Add(flag);
            }

            //OTP write
            var writeItems = otpPatternRows.Where(x => x.OtpReadWrite.Equals(EnumOtpReadWrite.Write)).OrderBy(x => x.RowNum)
                .ToList();
            foreach (OtpPatternRow item in writeItems)
            {
                if (item.OtpWriteType.Equals(EnumOtpWriteType.None))
                {
                    continue;
                }

                flag =
                    Combination.CombineByUnderLine(["F", item.Block, item.Mode, item.Item, item.PayloadList.First(), "N", "Flag"]);
                if (item.OtpWriteType.Equals(EnumOtpWriteType.Multishot4Byte) && item.BankName.Equals(EnumOtpBankName.All))
                {
                    flowSheet.AddRow(new FlowRow  //OTP_Burn_ALL_MuliShot_MULTISHOT4BYTE
                    {
                        Opcode = OpCode.Test,
                        Parameter = Combination.CombineByUnderLine([OtpConst.OtpInstanceWritePrefix, nameof(EnumOtpBankName.All).ToUpper(), OtpConst.OtpInstancePostfixMultiShot4Byte]),
                        FailAction = flag
                    });
                }
                else
                {
                    flowSheet.AddRow(new FlowRow
                    {
                        Opcode = OpCode.Nop,
                        Parameter = Combination.CombineByUnderLine([OtpConst.OtpInstanceWritePrefix, item.BankName.ToString().ToUpper(), item.OtpWriteType.ToString()]),
                        FailAction = flag
                    });
                }
            }

            //Relay off
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Test,
                Parameter = "Relay_Off_OTP",
            });

            foreach (OtpPatternRow item in functionalItems)
            {
                string parameter =
                    Combination.CombineByUnderLine(["Bin", item.Block.ToUpper(), item.Mode.ToUpper(), item.Item.ToUpper(), item.PayloadList.First().ToUpper()]);
                flowSheet.AddRow(new FlowRow { Opcode = OpCode.BinTable, Enable = "HardIPBin", Parameter = parameter });
            }

            flowSheet.AddPrintEndRow(headerFooterName);
            flowSheet.AddFooterRow(OtpConst.Otp);
            flowSheet.AddReturnRow();
            return flowSheet;
        }

        private static List<InstanceRow> GenerateOtpInstanceRow(List<OtpPatternRow> otpPatternRows)
        {
            var instanceRows = new List<InstanceRow>();
            if (otpPatternRows.Count != 0)
            {
                GenerateCheckDefaultRealInstanceRow();
                IEnumerable<OtpPatternRow> readItems = [.. otpPatternRows.Where(x => x.OtpReadWrite.Equals(EnumOtpReadWrite.Read) && !x.OtpReadType.Equals(EnumOtpReadType.None))];
                if (readItems.Any())
                {
                    OtpPatternRow? blankCheckItem = readItems.FirstOrDefault(x => x.OtpReadType.Equals(EnumOtpReadType.Multishot4Byte));
                    if (blankCheckItem != null)
                    {
                        instanceRows.Add(GenerateBlankCheckInstanceRow(blankCheckItem));
                    }

                    foreach (OtpPatternRow readItem in readItems)
                    {
                        instanceRows.Add(GenerateReadInstanceRow(readItem));
                    }
                }

                IEnumerable<OtpPatternRow> writeItems = [.. otpPatternRows.Where(x => x.OtpReadWrite.Equals(EnumOtpReadWrite.Write) && !x.OtpWriteType.Equals(EnumOtpWriteType.None))];
                if (writeItems.Any())
                {
                    foreach (OtpPatternRow writeItem in writeItems)
                    {
                        instanceRows.Add(GenerateWriteInstanceRow(writeItem));
                    }
                }
            }
            return instanceRows;
        }

        private static InstanceRow GenerateWriteInstanceRow(OtpPatternRow otpPatternRow)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(VbtFunctionLibShared.OtpBurn, "");
            var row = new InstanceRow
            {
                TestName = Combination.CombineByUnderLine([OtpConst.OtpInstanceWritePrefix, otpPatternRow.BankName.ToString().ToUpper(), otpPatternRow.OtpWriteType.ToString()]),
                VbtType = "VBT",
                VbtName = function.FunctionName,
                ArgList = function.Parameters,
                Args = function.ArgList,
                DcCategory = GetDcCategory(),
                DcSelector = "Typ",
                TimeSets = otpPatternRow.PayloadList.Count != 0 ? GetTimeSet(otpPatternRow) : ""
            };
            row.AcCategory = GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = GenerateLevel();
            row.Args = function.ArgList;
            return row;
        }

        private static InstanceRow GenerateReadInstanceRow(OtpPatternRow otpPatternRow)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(VbtFunctionLibShared.OtpRealAllSetOtpData, "");
            var row = new InstanceRow
            {
                TestName = Combination.CombineByUnderLine([OtpConst.OtpInstanceReadAll, otpPatternRow.OtpReadType.ToString()]),
                VbtType = "VBT",
                VbtName = function.FunctionName,
                ArgList = function.Parameters,
                DcCategory = GetDcCategory(),
                DcSelector = "Typ",
                TimeSets = otpPatternRow.PayloadList.Count != 0 ? GetTimeSet(otpPatternRow) : ""
            };
            row.AcCategory = GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = GenerateLevel();

            function.SetParamValue("DSSCReadPat", "");
            function.SetParamValue("interposePrePat", "");
            function.SetParamValue("interposePostPat", "");
            function.SetParamValue("DSSCSetupName", "");
            function.SetParamValue("Validating_", "");
            row.Args = function.ArgList;
            return row;
        }

        private static InstanceRow GenerateBlankCheckInstanceRow(OtpPatternRow otpPatternRow)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(VbtFunctionLibShared.OtpBlankBitCheck, "");
            var row = new InstanceRow
            {
                TestName = Combination.CombineByUnderLine([OtpConst.OtpBlankCheck, OtpConst.OtpInstancePostfixMultiShot4Byte]),
                VbtType = "VBT",
                VbtName = function.FunctionName,
                ArgList = function.Parameters,
                DcCategory = GetDcCategory(),
                DcSelector = "Typ",
                TimeSets = otpPatternRow.PayloadList.Count != 0 ? GetTimeSet(otpPatternRow) : ""
            };
            row.AcCategory = GetAcCategory(row.TimeSets);
            row.AcSelector = "Typ";
            row.PinLevels = GenerateLevel();
            row.Args = function.ArgList;
            return row;
        }

        private static void GenerateCheckDefaultRealInstanceRow()
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(VbtFunctionLibShared.OtpCheckDefaultReal, "");
            var row = new InstanceRow
            {
                TestName = OtpConst.OtpInstanceCheckDefaultReal,
                VbtType = "VBT",
                VbtName = function.FunctionName,
                ArgList = function.Parameters
            };
            function.SetParamValue("DebugPrintLog", "-1");
            row.Args = function.ArgList;
        }

        private static void GenOtpInstanceSheet(InstanceSheet instanceSheet)
        {
            if (instanceSheet.Rows.Count != 0)
            {
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirEfuse, instanceSheet);
            }
        }

        private static string GetDcCategory()
        {
            return "OTP_X_X_X";
        }

        private static string GetTimeSet(OtpPatternRow otpPatternRow)
        {
            var timeSets = new List<string>();
            foreach (string patternName in otpPatternRow.PayloadList)
            {
                if (AcTSetCategoryMapSingleton.Instance().PatternList.TryGetValue(patternName.ToLower(), out PatternData? patternData))
                {
                    timeSets.Add(patternData.TimeSetVersion);
                }
            }

            return string.Join(",", timeSets.Distinct().ToList());
        }

        private static string GetAcCategory(string timeSet)
        {
            string acCategory = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSet, BlockType.Otp);
            if (acCategory.EqualsIgnoreCase("TBD"))
            {
                acCategory = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSet);
            }

            return acCategory;
        }

        private static string GenerateLevel()
        {
            return "Levels_OTP";
        }

        private static void GenBinTable(List<string> otpFlags)
        {
            OtpBinTableWriter.GenBinTable(otpFlags);
        }
    }
}
