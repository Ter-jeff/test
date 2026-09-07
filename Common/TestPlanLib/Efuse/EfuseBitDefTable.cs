using System.Collections.Generic;

namespace TestPlanLib.Efuse
{
    public class EfuseBitDefTable
    {
        //LotId,WaferId...
        public int NameIdx = -1;

        public int MsbBitIdx = -1;
        public int LsbBitIdx = -1;
        public int BitWidthIdx = -1;

        //CP1/CP2..
        public int PrgStageIdx = -1;
        //number and could be N/A
        public int LowLimitIdx = -1;
        //number and could be N/A
        public int HighLimitIdx = -1;
        public int ResolutionIdx = -1;

        public int AlgorithmIdx = -1;
        public int DescriptionIdx = -1;
        //Read/Default/BinCut
        public int DefaultOrRealIdx = -1;
        public int DefaultValueIdx = -1;
        public int DifferentIdx = -1;
        public int HipNameIdx = -1;
        public int HipEquationIdx = -1;

        public int BlockIdx = -1;
        public int AccessModeIdx = -1;
        public int HiplimitlIdx = -1;
        public int HiplimithIdx = -1;

        //ECID eFuse Bit Def	MSB BIT	LSB BIT	Bit Width	N/A	N/A	N/A	N/A	programming stage	Low Limit	High Limit	Resolution	Algorithm	Description	Default or Real	Default Value	Difference
        public List<string> Titles = [];
        public List<List<string>> Rows = [];
        public List<string> BlockList = [];
        public Dictionary<string, string> HipList = [];
        public List<CrcItem> CrcItem = [];
    }
}
