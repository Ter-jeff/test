using System.Collections.Generic;
using System.Linq;

using CommonReaderLib;

using TestPlanLib.BinCut.Flow;

namespace TestPlanLib.BinCut.FlowNew
{
    public class NewBinCutFlowTable : MySheet
    {
        public int StartRowIndex = -1;
        public int BinningDomainIndex = -1;
        public int PerformanceModeIndex = -1;

        //FT_HOT @ 85'C, mV
        public string JobName = "";
        //FT2
        public string FinalJob = "";

        public List<NewBinCutFlowSheetRow> Rows = [];
        public Dictionary<string, int> SubFlowMappingTable = [];

        public NewBinCutFlowTable(NewBinCutFlowTable newBinCutFlowTable) : base(newBinCutFlowTable)
        {
            if (newBinCutFlowTable == null)
            {
                return;
            }

            StartRowIndex = newBinCutFlowTable.StartRowIndex;
            BinningDomainIndex = newBinCutFlowTable.BinningDomainIndex;
            PerformanceModeIndex = newBinCutFlowTable.PerformanceModeIndex;
            JobName = newBinCutFlowTable.JobName;
            FinalJob = newBinCutFlowTable.FinalJob;

            Rows = newBinCutFlowTable.Rows?.Select(row => row.Copy()).ToList() ?? [];
            SubFlowMappingTable = newBinCutFlowTable.SubFlowMappingTable != null ? new Dictionary<string, int>(newBinCutFlowTable.SubFlowMappingTable) : [];
        }

        public NewBinCutFlowTable()
        {
        }

        public NewBinCutFlowTable Copy() => new(this);

        public static void Check()
        {
        }
    }

    public class NewBinCutFlowSheetRow : MyRow
    {
        public EnumBinCutTableType TableType;
        public EnumBinCutTableBinType TableBinType;
        public string BinningDomain { get; set; } = string.Empty;
        public string PerformanceMode { get; set; } = string.Empty;
        public List<string> SubFlows { get; set; } = [];
        public List<(string, bool)> SubFlowsByType { get; set; } = [];
        public string Enable { get; set; } = string.Empty;
        public string Job = "";

        public NewBinCutFlowSheetRow(string sheetName, string job)
        {
            SheetName = sheetName;
            Job = job;
            BinningDomain = "";
            PerformanceMode = "";
        }

        public NewBinCutFlowSheetRow(NewBinCutFlowSheetRow newBinCutFlowSheetRow) : base(newBinCutFlowSheetRow)
        {
            if (newBinCutFlowSheetRow == null)
            {
                return;
            }

            TableType = newBinCutFlowSheetRow.TableType;
            TableBinType = newBinCutFlowSheetRow.TableBinType;
            BinningDomain = newBinCutFlowSheetRow.BinningDomain;
            PerformanceMode = newBinCutFlowSheetRow.PerformanceMode;
            Enable = newBinCutFlowSheetRow.Enable;
            Job = newBinCutFlowSheetRow.Job;

            SubFlows = newBinCutFlowSheetRow.SubFlows != null ? [.. newBinCutFlowSheetRow.SubFlows] : [];

            SubFlowsByType = newBinCutFlowSheetRow.SubFlowsByType != null ? [.. newBinCutFlowSheetRow.SubFlowsByType] : [];
        }

        public NewBinCutFlowSheetRow Copy()
        {
            return new NewBinCutFlowSheetRow(this);
        }
    }
}
