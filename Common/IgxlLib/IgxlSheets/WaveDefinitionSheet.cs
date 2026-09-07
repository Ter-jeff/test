using System.IO;

using IgxlLib.Const;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class WaveDefinitionSheet : IgxlSheet<WaveDefRow>
    {
        public override string SheetType => "DTWaveDefinitionSheet";
        public override string IgxlSheetName => IgxlSheetNames.WaveDefinition;

        public WaveDefinitionSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public WaveDefinitionSheet(string sheetName) : base(sheetName)
        {
        }

        public override void Write(string fileName, string version = "")
        {
            Directory.CreateDirectory(Path.GetDirectoryName(fileName) ?? "");

            using var writer = new StreamWriter(fileName);

            const string firstLine1 = "DTWaveDefinitionSheet,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1";
            const string firstLine2 = "Wave Definitions";
            writer.WriteLine(firstLine1 + '\t' + firstLine2);
            writer.WriteLine();
            writer.WriteLine("	WaveDefName	WaveDefType	WaveDef Component	Repeat Count	Relative Period	Relative Amplitude	Relative Offset	Primitive Parameters	Comment");
            foreach (WaveDefRow waveDefRow in Rows)
            {
                writer.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", waveDefRow.ColumnA, waveDefRow.WaveDefName, waveDefRow.WaveDefType, waveDefRow.WaveDefComponent, waveDefRow.RepeatCount, waveDefRow.RelativePeriod, waveDefRow.RelativeAmplitude, waveDefRow.RelativeOffset, waveDefRow.PrimitiveParameters, waveDefRow.Comment);
            }
        }
    }
}
