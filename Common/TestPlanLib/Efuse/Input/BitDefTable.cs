using System.Collections.Generic;
using System.Linq;

using CommonReaderLib;

namespace TestPlanLib.Efuse.Input
{
    public class BitDefTable : MySheet
    {
        public int MaxColIdx = -1;

        //bank_XX eFuse Bit Def
        public int BankEfuseBitDefIdx = -1;
        public int MsbBitIdx = -1;
        public int LsbBitIdx = -1;
        public int BitWidthIdx = -1;
        public int Na1Idx = -1;
        public int Na2Idx = -1;
        public int Na3Idx = -1;
        public int Na4Idx = -1;
        public int ProgrammingStageIdx = -1;
        public int LowLimitIdx = -1;
        public int HighLimitIdx = -1;
        public int ResolutionIdx = -1;
        public int AlgorithmIdx = -1;
        public int DescriptionIdx = -1;
        public int DefaultOrRealIdx = -1;
        public int DefaultValueIdx = -1;
        public int UsiMsbBitCycleIdx = -1;
        public int UsiLsbBitCycleIdx = -1;
        public int UsoMsbBitCycleIdx = -1;
        public int UsoLsbBitCycleIdx = -1;

        public List<BitDefRow> Rows = [];
        public string AccessMode = "";
        public string BitMode = "";
        public string Header = "";
        public string BlockName = "";
        public int HeaderRowNum;

        public BitDefTable()
        {
        }

        public BitDefTable(BitDefTable bitDefTable) : base(bitDefTable)
        {
            if (bitDefTable == null)
            {
                return;
            }

            MaxColIdx = bitDefTable.MaxColIdx;
            BankEfuseBitDefIdx = bitDefTable.BankEfuseBitDefIdx;
            MsbBitIdx = bitDefTable.MsbBitIdx;
            LsbBitIdx = bitDefTable.LsbBitIdx;
            BitWidthIdx = bitDefTable.BitWidthIdx;
            Na1Idx = bitDefTable.Na1Idx;
            Na2Idx = bitDefTable.Na2Idx;
            Na3Idx = bitDefTable.Na3Idx;
            Na4Idx = bitDefTable.Na4Idx;
            ProgrammingStageIdx = bitDefTable.ProgrammingStageIdx;
            LowLimitIdx = bitDefTable.LowLimitIdx;
            HighLimitIdx = bitDefTable.HighLimitIdx;
            ResolutionIdx = bitDefTable.ResolutionIdx;
            AlgorithmIdx = bitDefTable.AlgorithmIdx;
            DescriptionIdx = bitDefTable.DescriptionIdx;
            DefaultOrRealIdx = bitDefTable.DefaultOrRealIdx;
            DefaultValueIdx = bitDefTable.DefaultValueIdx;
            UsiMsbBitCycleIdx = bitDefTable.UsiMsbBitCycleIdx;
            UsiLsbBitCycleIdx = bitDefTable.UsiLsbBitCycleIdx;
            UsoMsbBitCycleIdx = bitDefTable.UsoMsbBitCycleIdx;
            UsoLsbBitCycleIdx = bitDefTable.UsoLsbBitCycleIdx;

            AccessMode = bitDefTable.AccessMode;
            BitMode = bitDefTable.BitMode;
            Header = bitDefTable.Header;
            BlockName = bitDefTable.BlockName;
            HeaderRowNum = bitDefTable.HeaderRowNum;

            Rows = bitDefTable.Rows?.Select(row => row.Copy()).ToList() ?? [];
        }

        public BitDefTable Copy()
        {
            return new BitDefTable(this);
        }
    }

    public class BitDefRow : MyRow
    {
        public string Line = "";
        public List<string> RowData = [];

        public BitDefRow()
        { }

        public BitDefRow(BitDefRow bitDefRow) : base(bitDefRow)
        {
            if (bitDefRow == null)
            {
                return;
            }

            Line = bitDefRow.Line;

            RowData = bitDefRow.RowData != null ? [.. bitDefRow.RowData] : [];
        }

        public BitDefRow Copy()
        {
            return new BitDefRow(this);
        }
    }

    public class BitDefBankRange
    {
        public string BankName = "";
        public double Lsb = 0;
        public double Msb = 0;
        public double Width = 0;
    }
}
