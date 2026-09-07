using System;

using CommonLib.Datalog;
using CommonLib.Extension;

namespace BinCutScriptLib.Base.Line
{
    public class BinCutLineBase : LineBase
    {
        public const string RailConst = "[INFO]  Starting BinCut test for ";

        public bool IsInstanceLine()
        {
            if (Line.EndsWithIgnoreCase("BV>") ||
                Line.EndsWithIgnoreCase("HBV>") ||
                Line.EndsWithIgnoreCase("BV_Csharp>") ||
                Line.EndsWithIgnoreCase("HBV_Csharp>") ||
                Line.EndsWithIgnoreCase("BV_CS>") ||
                Line.EndsWithIgnoreCase("HBV_CS>"))
            {
                return true;
            }

            return false;
        }

        public bool IsSelsram()
        {
            if (Line.StartsWithIgnoreCase("SELSRAM_") ||
                Line.StartsWithIgnoreCase("SRAM_Vth") ||
                Line.StartsWithIgnoreCase("DSSC_SELSRAM") ||
                Line.StartsWithIgnoreCase("Efuse_Product"))
            {
                return true;
            }

            return false;
        }

        public PowerBinningLine NewPowerBinningLine()
        {
            return new PowerBinningLine() { Line = Line, LineNo = LineNo };
        }

        public AdjustVddBinningLine NewAdjustVddBinningLine()
        {
            return new AdjustVddBinningLine() { Line = Line, LineNo = LineNo };
        }

        public ProductIdentifierLine NewProductIdentifierLine()
        {
            return new ProductIdentifierLine() { Line = Line, LineNo = LineNo };
        }

        public JudgeStoredIdsLine NewJudgeStoredIdsLine()
        {
            return new JudgeStoredIdsLine() { Line = Line, LineNo = LineNo };
        }

        public IdsOffLine NewIdsOffLine()
        {
            return new IdsOffLine() { Line = Line, LineNo = LineNo };
        }

        public IdsOnLine NewIdsOnLine()
        {
            return new IdsOnLine() { Line = Line, LineNo = LineNo };
        }

        public ReadFromDsscLine1 NewReadfromDsscLine()
        {
            return new ReadFromDsscLine1() { Line = Line, LineNo = LineNo };
        }

        public IdsSetWriteDecimalLine NewIdsSetWriteDecimalLine()
        {
            return new IdsSetWriteDecimalLine() { Line = Line, LineNo = LineNo };
        }

        public BvLine NewBvLine()
        {
            return new BvLine() { Line = Line, LineNo = LineNo };
        }

        public BvEqLine NewBvEqLine()
        {
            return new BvEqLine() { Line = Line, LineNo = LineNo };
        }

        public ReadDvfmLine NewReadDvfmLine()
        {
            return new ReadDvfmLine() { Line = Line, LineNo = LineNo };
        }

        public LimitLine NewJudgeResultLine()
        {
            return new LimitLine() { Line = Line, LineNo = LineNo };
        }

        public LimitLine NewPatternResultLine()
        {
            return new LimitLine() { Line = Line, LineNo = LineNo };
        }

        public EnLine NewEnLine()
        {
            return new EnLine() { Line = Line, LineNo = LineNo };
        }

        public OffsetLine1 NewOffsetLine()
        {
            return new OffsetLine1() { Line = Line, LineNo = LineNo };
        }

        public VBinResultWoTestLine1 NewVBinResultWoTestLine()
        {
            return new VBinResultWoTestLine1() { Line = Line, LineNo = LineNo };
        }

        public PatternLine NewPatternLine()
        {
            return new PatternLine() { Line = Line, LineNo = LineNo };
        }

        public HarvestSourceCodetLine NewHarvestSourceCodetLine()
        {
            return new HarvestSourceCodetLine() { Line = Line, LineNo = LineNo };
        }

        public SearchResultWoTestLine1 NewSearchResultWoTestLine()
        {
            return new SearchResultWoTestLine1() { Line = Line, LineNo = LineNo };
        }

        public bool IsPostStartPoint()
        {
            return IsCp2LvStartPoint();
        }

        public bool IsPostStartPointCsharp()
        {
            return IsCp2LvStartPointCsharp();
        }

        public bool IsPostStopPoint()
        {
            return Line.Contains("Site    Sort     Bin") ||
                   Line.StartsWithIgnoreCase("*print: PostBincut end*") ||
                   Line.StartsWithIgnoreCase("Flow PostBincut Stop");
        }

        public bool IsPostStopPointCsharp()
        {
            return Line.Contains("Site    Sort     Bin") ||
                   Line.StartsWithIgnoreCase("*print: PostBincut end*") ||
                   Line.StartsWithIgnoreCase("Flow PostBincut Stop");
        }

        public bool IsHvStopPoint()
        {
            return Line.Contains("Site    Sort     Bin") ||
                   Line.StartsWithIgnoreCase("*print: VddBinning_HVCC end") ||
                   Line.StartsWithIgnoreCase("*print: Flow_VddBinning_PostBinCut start*") ||
                   Line.StartsWithIgnoreCase("Flow VddBinning_OutsideBV Start") ||
                   Line.StartsWithIgnoreCase("Flow OutsideBincut Start") ||
                   Line.StartsWithIgnoreCase("*print: Flow_Vddbinning_HVCC end*") ||
                    Line.StartsWithIgnoreCase("Flow PostBincut Stop");
        }

        public bool IsHvStopPointCs()
        {
            return Line.Contains("Site    Sort     Bin") ||
                   Line.StartsWithIgnoreCase("Flow VddBinning_HVCC Stop ") ||
                    Line.StartsWithIgnoreCase("Flow PostBincut Stop");
        }

        public bool IsPatternLine()
        {
            string[] spt = Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
            bool isHramPattern = false;
            if (spt.Length == 7)
            {
                isHramPattern = int.TryParse(spt[5], out _) && int.TryParse(spt[6], out _);
            }
            if (Line.Contains("N/A               N/A") || isHramPattern)
            {
                return true;
            }
            return false;
        }

        public bool IsPatternLineCsharp()
        {
            string[] spt = Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
            bool isHramPattern = false;
            if (spt.Length == 7)
            {
                isHramPattern = int.TryParse(spt[5], out _) && int.TryParse(spt[6], out _);
            }
            if (Line.Contains("N/A               N/A") || isHramPattern)
            {
                return true;
            }
            return false;
        }

        public bool IsCp2LvStopPoint()
        {
            return Line.Contains("<PrintOutVddBinning>") ||
                   Line.Contains("Power_Binning_Calculation") ||
                   Line.Contains("<Power_Binning_J2>") ||
                   Line.Contains("<PrintOutVddBinning_Csharp>");
        }

        public bool IsCp1LvStopPoint()
        {
            return Line.Contains("<PrintOutVddBinning>") ||
                   Line.Contains("Power_Binning_Calculation") ||
                   Line.Contains("<Power_Binning_J2>") ||
                   Line.Contains("PwrBin Sheet :") ||
                   Line.Contains("PwrBin Harvest_Bin") ||
                   Line.Contains("Site    Sort     Bin") ||
                   Line.Contains("<PrintOutVddBinning_Csharp>");
        }

        public bool IsCp1LvStopPointCsharp()
        {
            return Line.Contains("<PrintOutVddBinning>") ||
                   Line.Contains("print: Vddbinning_PrintOutVddBinning start") ||
                   Line.Contains("Power_Binning_Calculation") ||
                   Line.Contains("<Power_Binning_J2>") ||
                   Line.Contains("PwrBin Sheet :") ||
                   Line.Contains("PwrBin Harvest_Bin") ||
                   Line.Contains("Site    Sort     Bin") ||
                   Line.Contains("<PrintOutVddBinning_Csharp>");
        }

        public bool IsCp1LvStartPoint()
        {
            return Line.Contains("Alg=") ||
                   Line.StartsWithIgnoreCase("BV_") ||
                   Line.StartsWithIgnoreCase("Initial_Voltage") ||
                   Line.EqualsIgnoreCase("****************separated for BinCut step****************") ||
                   Line.Contains("to call instance:");
        }

        public bool IsCp1LvStartPointCsharp()
        {
            if (Line.StartsWithIgnoreCase(RailConst))
            {
                return true;
            }

            return false;
        }

        public bool IsCp2LvStartPoint()
        {
            return Line.Contains("Alg=") ||
                   Line.StartsWithIgnoreCase("Initial_Voltage") ||
                   Line.Contains("to call instance:");
        }

        public bool IsCp2LvStartPointCsharp()
        {
            if (Line.StartsWithIgnoreCase(RailConst))
            {
                return true;
            }

            return false;
        }

        public bool IsCp1LvStepStartPoint()
        {
            return Line.Contains("Alg=") ||
                   Line.StartsWithIgnoreCase("BV_") ||
                   Line.StartsWithIgnoreCase("Initial_Voltage") ||
                   Line.EqualsIgnoreCase("****************separated for BinCut step****************") ||
                   IsPatternLine();
        }

        public bool IsCp1LvStepStopPoint()
        {
            return Line.Contains("Site  Test Name") && Line.Contains("Pin");
        }

        public bool IsCp1LvStepStopPointCsharp()
        {
            return Line.Contains("Flow ") && Line.Contains("BV Stop");
        }

        public bool IsPatternLimit()
        {
            return Line.Contains("Site  Test Name");
            //&& Line.Contains("Pattern"));
        }

        public bool IsDvfmStopPoint()
        {
            return Line.StartsWithIgnoreCase("Initial_Voltage") ||
                   string.IsNullOrEmpty(Line)
                   || Line.StartsWithIgnoreCase("******");
        }

        public bool IsCfgFuseStopPoint()
        {
            return Line.StartsWithIgnoreCase("Initial_Voltage");
            //|| Line.StartsWith("******", StringComparison.CurrentCultureIgnoreCase)
        }

        public bool IsAssignFuseStopPoint()
        {
            return Line.EqualsIgnoreCase("[INFO]  ") ||
                   Line.StartsWithIgnoreCase(" Number       Site  Test Name") ||
                   Line.Length == 0;
        }

        public bool IsHarvestSourceCode()
        {
            return Line.Contains("HarvestSourceCode");
        }

        public bool IsAlgLine()
        {
            string[] arr = Line.Split([','], StringSplitOptions.RemoveEmptyEntries);
            if (arr.Length > 3 && arr[2].StartsWithIgnoreCase("Alg="))
            {
                return true;
            }

            return false;
        }

        public bool IsStepStartCsharp()
        {
            if (Line.StartsWithIgnoreCase(RailConst))
            {
                return true;
            }

            //****Start of calculation for Vx_SOC_MS002**** 
            if (Line.StartsWith("****Start of calculation for"))
            {
                return true;
            }

            return false;
        }

        public bool IsHarvestSourceCodeStart()
        {
            return Line.Contains("Setup Dig Src Test Start");
        }

        public bool IsHarvestSourceCodeStop()
        {
            return Line.Contains("Setup Dig Src Test End");
        }

        public bool IsFstpHeader()
        {
            return Line.Contains("MultiFstp_Header");
        }

        public bool IsFstpFooter()
        {
            return Line.Contains("MultiFstp_Footer");
        }

        public bool IsHarvestBinningFlag()
        {
            return Line.Contains("HarvestBinningFlag");
        }

        public bool IsPreVddHeader()
        {
            return Line.Contains("Vddbinning_Pre start");
        }

        public bool IsPreVddFooter()
        {
            return Line.Contains("Vddbinning_Pre end");
        }

        public bool IsReJudgeHeader()
        {
            return Line.Contains("Vddbinning_ReJudge_stored_IDS start");
        }

        public bool IsVbinResultLine()
        {
            return Line.Contains("Start of Set_VBinResult_without_Test");
        }

        public bool IsEqnNStart()
        {
            return Line.Contains("Equation_N_Start");
        }

        public bool IsPinGroupDisabledLine()
        {
            return Line.StartsWithIgnoreCase("Site") && (Line.Contains("HAC_HARV_X_MASKED") || Line.Contains("disabled PinGroup"));
        }

        public bool IsTotalDisabledPinGroupCountLine()
        {
            return Line.StartsWithIgnoreCase("Total disabled PinGroup");
        }

        public bool IsPinMaskFeatureLine()
        {
            return Line.StartsWith("--(") && Line.Contains("Pin mask Feature");
        }

        public bool IsHarMandHarvFailLine()
        {
            return Line.StartsWith("HRAM Fail Pin:") && Line.Contains("Harv Fail Pin:");
        }

        public bool IsHarvFailLine()
        {
            return Line.StartsWith("Site") && Line.Contains("Pat Fail,") && Line.Contains("did not fail.");
        }

        public bool IsSyncUpLine()
        {
            return Line.StartsWithIgnoreCase("Bincut voltage switch");
        }

        public bool IsNoBinOutLine()
        {
            return Line.StartsWithIgnoreCase("Test Instance") && Line.Contains("judge_PF_func is not used. Please check if any device condition for BinOut control exists in the instance flow table!!!");
        }

        public bool IsTotalFailCycleLine()
        {
            return Line.StartsWithIgnoreCase("Site") &&
                   Line.Contains("total fail cycle =");
        }

        public bool IsHarvFlagSetTrueLine()
        {
            return Line.StartsWithIgnoreCase("Site") &&
                   Line.Contains("Failed in Bincut Bin1EQ1, so harvest flag is assigned true!");
        }

        public bool IsHarvRecoveryLine()
        {
            return Line.StartsWithIgnoreCase("Site") &&
                   Line.Contains("HarvestBinningFlag:") && Line.Contains("overwrites test result of the failed DUT");
        }

        public bool IsSetSearchResultWoTestLine()
        {
            return Line.Contains("_mode is not tested, set VBIN_RESULT of p_mode for Set_E1_Voltage_ForPmode. EQN=");
        }

        public bool IsMfstpBypassBinoutLine()
        {
            return Line.StartsWithIgnoreCase("Site") &&
                   Line.Contains("test failed, but MultiFSTP bypassed BinOut!");
        }
    }
}
