using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.Binning
{
    public class BinningTableReader
    {
        public static BinningTable Read(string inPath)
        {
            List<string> lines = [.. File.ReadAllLines(inPath)];
            BinningTable errors = Read(lines, inPath);
            return errors;
        }

        public static BinningTable Read(ExcelWorksheet excelWorksheet, string jobStage)
        {
            return Read(excelWorksheet.ConvertToLines(), excelWorksheet.Name, jobStage);
        }

        private static BinningTable Read(List<string> lines, string sheetName, string jobStage = "")
        {
            try
            {
                BinningTable binningTable = ReadSheet(lines, sheetName, jobStage);

                binningTable.Check();

                return binningTable;
            }
            catch (Exception e)
            {
                string msg = "Find exception when reading " + sheetName + "!!! \n ErrMsg: " + e;
                var errMsg = new Exception(msg);
                throw errMsg;
            }
        }

        private static BinningTable ReadSheet(List<string> lines, string sheetName, string jobStage)
        {
            int binnedIdx = -1;
            //init
            var binningTable = new BinningTable
            {
                SheetName = sheetName
            };
            binningTable.TitleList.Clear();
            binningTable.Rows.Clear();
            binningTable.Job = jobStage;
            int index = 0;
            {
                GetTitle(lines, ref binnedIdx, binningTable, ref index);

                index = GetGbTable(lines, binnedIdx, binningTable, index);
            }
            return binningTable;
        }

        private static int GetGbTable(List<string> lines, int binnedIdx, BinningTable binningTable, int index)
        {
            //STEP2. Get Gb table
            //-----------------------------------------------------------
            for (; index < lines.Count; index++)
            {
                //SoC	MS001	 MC601 E1 Voltage	MG101 E1 Voltage	MS001 Evaluate Bin	MC601 CSRAM V	MG002 GSRAM V	CP LVCC	CP LVCC	LV Levels	SOC TD	SOC BIST		SoC	MS001	MC601 Product-CP2GB	MG101 Product-CP2GB	MS001 Product-CP2GB	MC601 CSRAM V	MG002 GSRAM V	CP LVCC	CP LVCC	LV Levels	SOC TD	SOC BIST		SoC	MS001	MC601 Product-FT1GB	MG101 Product-FT1GB	MS001 Product-FT1GB	MC601 CSRAM V	MG002 GSRAM V	FT LVCC	FT LVCC	LV Levels	SOC TD	SOC BIST		SoC	MS001	MC601 Product-FT2GB	MG101 Product-FT2GB	MS001 Product-FT2GB	MC601 CSRAM V	MG002 GSRAM V	FT LVCC	FT LVCC	LV Levels	SOC TD	SOC BIST		SoC	MS001	MC601 Product-FQAGB	MG101 Product-FQAGB	MS001 Product-FQAGB	MC601 CSRAM V	MG002 GSRAM V	FQA LVCC	FQA LVCC	LV Levels	SOC TD	SOC BIST		SOC TD & BIST	
                string line = lines[index];

                //3a. end line
                if (line.Contains("END", StringComparison.OrdinalIgnoreCase)) //End	End	End
                {
                    break;
                }

                //3b. empty, just continue
                if (line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries).Length == 0)
                {
                    continue;
                }

                //3c. parser GB equation table
                string[] spt = line.Split(['\t'], StringSplitOptions.None);
                if (binnedIdx == -1 || binnedIdx >= spt.Length)
                {
                }
                if (spt[binnedIdx].EqualsIgnoreCase("TRUE"))
                {
                    var vLine = spt.Select(token => token.Trim()).ToList();
                    //Fill the cell for SRAM binning table to avoid program crach
                    while (vLine.Count < binningTable.TitleList.Count)
                    {
                        vLine.Add("");
                    }
                    var vddBinTableRow = new BinningRow
                    {
                        Line = line,
                        RowNum = index + 1,
                        RowData = vLine
                    };
                    binningTable.Rows.Add(vddBinTableRow);
                }
                else if (spt[binnedIdx].EqualsIgnoreCase("ATE") ||
                          spt[binnedIdx].EqualsIgnoreCase("FALSE"))
                {
                    var vLine = spt.Select(token => token.Trim()).ToList();
                    var vddBinTableRow = new BinningRow
                    {
                        Line = line,
                        RowNum = index + 1,
                        RowData = vLine,
                        IsOtherRail = true
                    };
                    binningTable.Rows.Add(vddBinTableRow);
                }
            }

            return index;
        }

        private static void GetTitle(List<string> lines, ref int binnedIdx, BinningTable binningTable, ref int index)
        {
            //STEP1. GetTitle and specific field index(eg. CPGB/CP2GB/FT1GB/FT2GB/ATE_FQAGB)
            //-----------------------------------------------------------
            for (; index < lines.Count; index++)
            {
                string line = lines[index];
                //Line#3>   Domain	ID	Mode	EQN	C	M	CPIDSMax	CPVmax	CPVmin	CPGB	CP2GB	FT1GB	FT2GB	FTIDS	SLTGB	ATE_FQAGB	HTOL_RO_GB	SLT_FQA_GB	PMUMax	PMUMin	CPHV	FTHV	QAHV	Comment	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	

                if (line.Contains("REV:", StringComparison.OrdinalIgnoreCase))
                {
                    binningTable.TitleRow1 = line;
                    string[] lineSpt = line.Split(['\t'], StringSplitOptions.None);
                    if (lineSpt[0] == "Rev:")
                    {
                        binningTable.Version = lineSpt[1];
                    }
                }

                if (line.Contains("BASE VOLTAGE", StringComparison.OrdinalIgnoreCase))
                {
                    binningTable.TitleRow2 = line;
                    string[] lineSpt = line.Split(['\t'], StringSplitOptions.None);
                    binningTable.BaseVoltage = lineSpt[1];
                }

                if (line.Contains("STEP SIZE", StringComparison.OrdinalIgnoreCase))
                {
                    string[] lineSpt = line.Split(['\t'], StringSplitOptions.None);
                    int sizeIdx = lineSpt.ToList().FindIndex(x => x.EqualsIgnoreCase("STEP SIZE")) + 1;
                    if (sizeIdx != 0)
                    {
                        binningTable.StepSize = Convert.ToDouble(lineSpt[sizeIdx]);
                    }
                }

                if (line.Contains("DOMAIN", StringComparison.OrdinalIgnoreCase))
                {
                    binningTable.TitleRow3 = line;
                    if (binningTable.StartRowIdx == -1)
                    {
                        binningTable.StartRowIdx = index;
                    }

                    string[] arr = line.Split(['\t'], StringSplitOptions.None);
                    for (int i = 0; i < arr.Length; i++)
                    {
                        if (binningTable.CommentIdx == -1 || i <= binningTable.CommentIdx)
                        {
                            binningTable.TitleList.Add(new Tuple<string, int>(arr[i], i));
                        }

                        binnedIdx = SetIndex(binnedIdx, binningTable, arr, i);
                    }
                    break;
                }
            }
        }

        private static int SetIndex(int binnedIdx, BinningTable binningTable, string[] arr, int i)
        {
            binnedIdx = SetIndexBasic(binnedIdx, binningTable, arr, i);

            if (arr[i].Contains("CPHV", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.CpHvIdx = i;
            }
            else if (arr[i].Contains("FTHV", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.FtHvIdx = i;
            }
            else if (arr[i].Contains("QAHV", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.QaHvIdx = i;
            }
            else if (arr[i].Contains("ALLOW EQUAL", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.AllowEqualIdx = i;
            }
            else if (arr[i].Contains("MONO DELTA", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.MonoDeltaIdx = i;
            }
            else if (arr[i].Contains("MONOTONICITY_OFFSET", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.MonoDeltaIdx = i;
            }
            else if (arr[i].Contains("SRAMTHRESH_PRODUCT", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.SramthreshProductIdx = i;
            }
            else if (arr[i].Contains("SRAMTHRESH_CP1", StringComparison.OrdinalIgnoreCase) || arr[i].Contains("SRAMTHRESH_BINSEARCH", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.SramthreshCp1Idx = i;
            }
            else if (arr[i].Contains("SRAMTHRESH_CP2", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.SramthreshCp2Idx = i;
            }
            else if (arr[i].Contains("INT_MODE_L", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.IntModeLIdx = i;
            }
            else if (arr[i].Contains("INT_MODE_H", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.IntModeHIdx = i;
            }
            else if (arr[i].Contains("INT_MF", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.IntMfIdx = i;
            }
            else if (arr[i].Contains("INT_OFFSET", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.IntOffsetIdx = i;
            }
            else if (arr[i].Contains("INT_SKIPTEST", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.IntSkipTestIdx = i;
            }

            SetIndexOther(binningTable, arr, i);

            return binnedIdx;
        }

        private static void SetIndexOther(BinningTable binningTable, string[] arr, int i)
        {
            if (arr[i].Contains("OFFSET_CP1_TD", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetCp1TdIdx = i;
            }
            else if (arr[i].Contains("OFFSET_CP1_BIST", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetCp1BistIdx = i;
            }
            else if (arr[i].Contains("OFFSET_CP1_FUNC", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetCp1FuncIdx = i;
            }
            else if (arr[i].Contains("OFFSET_CP2_TD", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetCp2TdIdx = i;
            }
            else if (arr[i].Contains("OFFSET_CP2_BIST", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetCp2BistIdx = i;
            }
            else if (arr[i].Contains("OFFSET_CP2_FUNC", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetCp2FuncIdx = i;
            }
            else if (arr[i].Contains("OFFSET_FT1_TD", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetFt1TdIdx = i;
            }
            else if (arr[i].Contains("OFFSET_FT1_BIST", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetFt1BistIdx = i;
            }
            else if (arr[i].Contains("OFFSET_FT1_FUNC", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetFt1FuncIdx = i;
            }
            else if (arr[i].Contains("OFFSET_FT2_TD", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetFt2TdIdx = i;
            }
            else if (arr[i].Contains("OFFSET_FT2_BIST", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetFt2BistIdx = i;
            }
            else if (arr[i].Contains("OFFSET_FT2_FUNC", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetFt2FuncIdx = i;
            }
            else if (arr[i].Contains("OFFSET_QA_TD", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetQaTdIdx = i;
            }
            else if (arr[i].Contains("OFFSET_QA_BIST", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetQaBistIdx = i;
            }
            else if (arr[i].Contains("OFFSET_QA_FUNC", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.OffsetQaFuncIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("Comment"))
            {
                binningTable.CommentIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("BincutCalc_IDSrail"))
            {
                binningTable.BincutCalcIdSrailIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("BincutCalc_IDSmax"))
            {
                binningTable.BincutCalcIdSmax = i;
            }
        }

        private static int SetIndexBasic(int binnedIdx, BinningTable binningTable, string[] arr, int i)
        {
            if (arr[i].Contains("DOMAIN", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.DomainIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("ID"))
            {
                binningTable.IdIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("BINNED"))
            {
                binningTable.BinnedIdx = i;
                binnedIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("MODE"))
            {
                binningTable.ModeIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("EQN"))
            {
                binningTable.EqnIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("MODE_EQN"))
            {
                binningTable.ModeEqnIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("EQN_BIN"))
            {
                binningTable.EqnBinIdx = i;
            }
            else if (arr[i].EqualsIgnoreCase("BINX_IDSMAX") || arr[i].EqualsIgnoreCase("BINX_CPIDSMAX"))
            {
                binningTable.BinXIdsMaxIdx = i;
            }
            else if (arr[i].Trim().EqualsIgnoreCase("C"))
            {
                binningTable.CIdx = i;
            }
            else if (arr[i].Trim().EqualsIgnoreCase("M"))
            {
                binningTable.MIdx = i;
            }
            else if (arr[i].Trim().EqualsIgnoreCase("CPIDSMAX"))
            {
                binningTable.IdsMaxIdx = i;
            }
            else if (arr[i].Trim().EqualsIgnoreCase("FTIDS") || arr[i].Trim().EqualsIgnoreCase("IDSMax_HOT"))
            {
                binningTable.FtMaxIdx = i;
            }
            else if (arr[i].Trim().EqualsIgnoreCase("CPVMAX") || arr[i].Trim().EqualsIgnoreCase("BinningVmax"))
            {
                binningTable.CpVMaxIdx = i;
            }
            else if (arr[i].Trim().EqualsIgnoreCase("CPVMIN") || arr[i].Trim().EqualsIgnoreCase("BinningVmin"))
            {
                binningTable.CpVMinIdx = i;
            }
            SetIndexGb(binningTable, arr, i);

            return binnedIdx;
        }

        private static void SetIndexGb(BinningTable binningTable, string[] arr, int i)
        {
            if (arr[i].Contains("CPGB", StringComparison.OrdinalIgnoreCase) || arr[i].Trim().EqualsIgnoreCase("BinningGB"))
            {
                binningTable.CpGbIdx = i;
            }
            else if (arr[i].Contains("CP2GB", StringComparison.OrdinalIgnoreCase) || arr[i].Contains("CP_GB_HOT", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.Cp2GbIdx = i;
            }
            else if (arr[i].Contains("FT_GB_ROOM", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.Ft1GbIdx = i;
            }
            else if (arr[i].Contains("FT_GB_HOT", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.Ft2GbIdx = i;
            }
            else if (arr[i].Contains("ATE_FQAGB", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.QaGbIdx = i;
            }
            else if (arr[i].Contains("HTOL_RO_GB_ROOM", StringComparison.OrdinalIgnoreCase) || arr[i].Contains("HTOL_T0TX_GB_ROOM", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.HtolGbRoomIdx = i;
            }
            else if (arr[i].Contains("HTOL_RO_GB_HOT", StringComparison.OrdinalIgnoreCase) || arr[i].Contains("HTOL_T0TX_GB_HOT", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.HtolGbHotIdx = i;
            }
            else if (arr[i].Contains("HTOL_RO_GB", StringComparison.OrdinalIgnoreCase))
            {
                binningTable.HtolGbIdx = i;
            }
        }
    }
}
