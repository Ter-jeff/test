using System;
using System.Collections.Generic;
using System.Drawing;

using BinCutScriptLib.Base;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Printer.PrintExcel;
using BinCutScriptLib.Static;

using CommonLib.ErrorReport;

using TestPlanLib;

namespace BinCutScriptLib
{
    internal static class PrintReportMain
    {
        internal static void PrintReport(Action<string, Color> richTextBoxAppend, List<List<SiteInfo>> allDiceInfos, string outputFile, List<Alarm> alarms, Job job, CheckManager checkManager, string tempFolder)
        {
            richTextBoxAppend($"Totally {allDiceInfos.Count} dice parse complete.", Color.Blue);
            if (allDiceInfos.Count == 0)
            {
                return;
            }

            var outputExcel = new WriteExcel(outputFile, allDiceInfos, BinCutData.BinningTables);
            richTextBoxAppend("Print Result sheet !!!", Color.Blue);
            outputExcel.WriteResult();

            richTextBoxAppend("Print Histogram sheet !!!", Color.Blue);
            outputExcel.WriteHistogram();

            richTextBoxAppend("Print PassList sheet !!!", Color.Blue);
            outputExcel.WritePassList();
            if (BinCutConfig.FlagUseCofInstance.Equals(true))
            {
                richTextBoxAppend("Print COFSummary sheet !!!", Color.Blue);
                outputExcel.WriteCofSummary();
            }
            if (BinCutConfig.IsDoAll)
            {
                richTextBoxAppend("Print AllSitePatternFailList sheet !!!", Color.Blue);
                outputExcel.WriteAllSitePatternFailList();
                richTextBoxAppend("Print PatternFailList sheet !!!", Color.Blue);
                outputExcel.WritePatternFailList();
            }
            richTextBoxAppend("Print PerformanceList sheet !!!", Color.Blue);
            outputExcel.WritePerformanceList();
            outputExcel.WritePerformanceList1();

            richTextBoxAppend("Print PerTouchdownStepList sheet !!!", Color.Blue);
            outputExcel.WritePerTouchdownStepList();

            if (BinCutData.PowerBinningSheetList.Count != 0)
            {
                richTextBoxAppend("Print PowerBinning sheet !!!", Color.Blue);
                outputExcel.WritePowerBinningList();
            }

            richTextBoxAppend("Print AlarmList sheet !!!", Color.Blue);
            outputExcel.OutputAlarmList(alarms);

            richTextBoxAppend("Print ErrorReport sheet !!!", Color.Blue);
            outputExcel.WriteErrorReport();
            int errorCount = ErrorReportManager.GetErrorList().Count;

            richTextBoxAppend("Print PatternCheck sheet !!!", Color.Blue);
            if (BinCutData.HasBinCutInstance && BinCutData.BinCutPatternReport != null)
            {
                if (BinCutData.HasBinCutInstance)
                {
                    outputExcel.WriteMissExtraCheck(job, BinCutData.BinCutPatternReport, BinCutData.BinCutFlowTables, tempFolder);
                }
            }

            outputExcel.ExcelClose();

            AlgorithmBaseHelpers.PrintCheckedMessage(richTextBoxAppend, allDiceInfos, checkManager, errorCount);
        }
    }
}
