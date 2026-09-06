using System.IO;
using System.Linq;
using System.Text;

using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class LevelSheet : IgxlSheet<LevelRow>
    {
        public override string SheetType => "DTLevelSheet";
        public override string IgxlSheetName { get; } = IgxlSheetNames.PinLevel;

        public LevelSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public LevelSheet(string sheetName) : base(sheetName)
        {
        }

        public LevelSheet(LevelSheet levelSheet) : base(levelSheet.Name)
        {
            Rows = [.. levelSheet.Rows.Select(row => row.Copy())];

            IgxlSheetName = levelSheet.IgxlSheetName;
            FilePath = levelSheet.FilePath;
            IgxlSheetContext = levelSheet.IgxlSheetContext;
            SourceInfo = levelSheet.SourceInfo.Copy();
        }

        public LevelSheet Copy()
        {
            return new LevelSheet(this);
        }

        public void AddPowerPinLevel(PowerLevel powerLevel)
        {
            AddRow(new LevelRow(powerLevel.PinName, "Vmain", powerLevel.Vmain, powerLevel.Comment));
            AddRow(new LevelRow(powerLevel.PinName, "valt", powerLevel.Valt, powerLevel.Comment));
            AddRow(new LevelRow(powerLevel.PinName, "iFoldLevel", powerLevel.IFoldLevel, powerLevel.Comment));
            AddRow(new LevelRow(powerLevel.PinName, "tdelay", powerLevel.Tdelay, powerLevel.Comment));
        }

        public void AddPowerPinVSlewLevel(PowerLevel powerLevel)
        {
            AddPowerPinLevel(powerLevel);

            AddRow(new LevelRow(powerLevel.PinName, "VSlewRate", powerLevel.VSlewRate, powerLevel.Comment));
        }

        public void AddDcviPowerPinLevel(DcviPowerLevel dcviPowerLevel)
        {
            var levelRow = new LevelRow(dcviPowerLevel.PinName, "Vps", dcviPowerLevel.Vps, dcviPowerLevel.Comment);

            AddRow(levelRow);

            levelRow = new LevelRow(dcviPowerLevel.PinName, "Isc", dcviPowerLevel.Isc, dcviPowerLevel.Comment);

            AddRow(levelRow);

            levelRow = new LevelRow(dcviPowerLevel.PinName, "tdelay", dcviPowerLevel.TDelay, dcviPowerLevel.Comment);

            AddRow(levelRow);
        }

        public void AddIoPinLevel(IoLevel ioLevel)
        {
            var levelRow = new LevelRow(ioLevel.PinName, "vil", ioLevel.Vil, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "vih", ioLevel.Vih, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "vol", ioLevel.Vol, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "voh", ioLevel.Voh, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "voh_alt1", ioLevel.VohAlt1, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "voh_alt2", ioLevel.VohAlt2, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "iol", ioLevel.Iol, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "ioh", ioLevel.Ioh, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "vt", ioLevel.Vt, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "vcl", ioLevel.Vcl, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "vch", ioLevel.Vch, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "voutLoTyp", ioLevel.VOutLoTyp, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "voutHiTyp", ioLevel.VOutHiTyp, "");
            AddRow(levelRow);
            levelRow = new LevelRow(ioLevel.PinName, "driverMode", ioLevel.DriverMode, "");
            AddRow(levelRow);
        }

        public void AddDiffLevel(DiffLevel diffLevel)
        {
            var row = new LevelRow(diffLevel.PinName, "Vicm", diffLevel.Vicm, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "Vid", diffLevel.Vid, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "dVid0", diffLevel.DVid0, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "dVid1", diffLevel.DVid1, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "dVicm0", diffLevel.DVicm0, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "dVicm1", diffLevel.DVicm1, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "Vod", diffLevel.Vod, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "Vod_Alt1", diffLevel.VodAlt1, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "Vod_Alt2", diffLevel.VodAlt2, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "dVod0", diffLevel.DVod0, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "dVod1", diffLevel.DVod1, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "Iol", diffLevel.Iol, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "Ioh", diffLevel.Ioh, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "VodTyp", diffLevel.VodTyp, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "VocmTyp", diffLevel.VocmTyp, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "Vt", diffLevel.Vt, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "Vcl", diffLevel.Vcl, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "Vch", diffLevel.Vch, "");
            AddRow(row);
            row = new LevelRow(diffLevel.PinName, "DriverMode", diffLevel.DriverMode, "");
            AddRow(row);
        }

        public override void Write(string fileName, string version = "")
        {
            Write2P1(fileName, version);
        }

        protected override void WriteSheet(string fileName, string version, SheetInfo sheetInfo)
        {
            if (Rows.Count == 0)
            {
                return;
            }

            using var sw = new StreamWriter(fileName, false);
            var sb = new StringBuilder();
            int maxCount = GetMaxCount(sheetInfo);
            int pinGroupIndex = GetIndexFrom(sheetInfo, "Pin/Group");
            int seqIndex = GetIndexFrom(sheetInfo, "Seq.");
            int parameterIndex = GetIndexFrom(sheetInfo, "Parameter");
            int valueIndex = GetIndexFrom(sheetInfo, "Value");
            int commentIndex = GetIndexFrom(sheetInfo, "Comment");

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                SetColumns(sheetInfo, i, arr);

                WriteHeader(arr, sb);
            }
            #endregion

            #region data
            for (int index = 0; index < Rows.Count; index++)
            {
                LevelRow row = Rows[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.PinName))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[pinGroupIndex] = row.PinName;
                    arr[seqIndex] = row.Seq;
                    arr[parameterIndex] = row.Parameter ?? "";
                    if (!string.IsNullOrEmpty(row.Value) && !row.Value.ToUpper().EqualsIgnoreCase("HIZ") &&
                        !row.Value.ToUpper().EqualsIgnoreCase("VT"))
                    {
                        arr[valueIndex] = "=" + row.Value.Trim().TrimStart('=');
                    }
                    else
                    {
                        arr[valueIndex] = row.Value ?? "";
                    }

                    arr[commentIndex] = row.Comment ?? "";
                }
                else
                {
                    arr = ["\t"];
                }
                sb.AppendLine(string.Join("\t", arr));
            }
            #endregion

            sw.Write(sb.ToString());
        }
    }
}
