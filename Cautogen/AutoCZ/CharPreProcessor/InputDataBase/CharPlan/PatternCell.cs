namespace Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan
{
    public class PatternCell
    {
        public string Header { get; set; } = "";
        public string Name { get; set; } = "";
        public int ColIndex { get; set; } = -1;
        public string PatternDefine { get; set; } = "";
        public string Retention { get; set; } = "";
        public string PowerRunScenario { get; set; } = "";
        public string DigSrcBitSize { get; set; } = "";
        public string DigSrcSeg { get; set; } = "";
        public string DigSrcPin { get; set; } = "";
        public string DigSrcEq { get; set; } = "";
        public PatternCell()
        {

        }
        public PatternCell(PatternCell item)
        {
            Copy(item);
        }

        public void Copy(PatternCell item)
        {
            Header = item.Header;
            ColIndex = item.ColIndex;
            PatternDefine = item.PatternDefine;
            Retention = item.Retention;
            PowerRunScenario = item.PowerRunScenario;
            DigSrcBitSize = item.DigSrcBitSize;
            DigSrcSeg = item.DigSrcSeg;
            DigSrcPin = item.DigSrcPin;
            DigSrcEq = item.DigSrcEq;
        }
    }
}
