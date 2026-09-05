using System.IO;

using IgxlLib.Const;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class MixedSignalSheet : IgxlSheet<MixedSigRow>
    {
        public override string SheetType => "DTMixedSignalTimingSheet";
        public override string IgxlSheetName => IgxlSheetNames.MixedSignal;

        public MixedSignalSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public MixedSignalSheet(string sheetName) : base(sheetName)
        {
        }

        public override void Write(string fileName, string version = "")
        {
            Directory.CreateDirectory(Path.GetDirectoryName(fileName) ?? "");

            using var writer = new StreamWriter(fileName);

            const string firstLine1 = "DTMixedSignalTimingSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1";
            const string firstLine2 = "Mixed Signal Timing";
            writer.WriteLine(firstLine1 + '\t' + firstLine2);
            writer.WriteLine();
            writer.WriteLine("			Resource		Clocking					Instrument	Waveform		MSW	Data for Pre v5.0 Upgrades			");
            writer.WriteLine("\tSet Name	Subset	Type	ID	Fs	N	Fr	M	USR	Data	Definition	Filter	Settings	WaveName	Amplitude	Offset	Old Instrument Data	Comment");
            foreach (MixedSigRow mixedSigRow in Rows)
            {
                writer.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}\t{11}\t{12}\t{13}\t{14}\t{15}\t{16}\t{17}\t{18}\t", mixedSigRow.ColumnA, mixedSigRow.Name, mixedSigRow.Subset, mixedSigRow.Type, mixedSigRow.Id, mixedSigRow.Fs, mixedSigRow.N, mixedSigRow.Fr, mixedSigRow.M, mixedSigRow.Usr, mixedSigRow.Data, mixedSigRow.Definition, mixedSigRow.Filter, mixedSigRow.Settings, mixedSigRow.WaveName, mixedSigRow.Amplitude, mixedSigRow.Offset, mixedSigRow.OldInstData, mixedSigRow.Comment);
            }
        }
    }
}
