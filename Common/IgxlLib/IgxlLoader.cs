using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;

using IgxlLib.Enums;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using OfficeOpenXml;

namespace IgxlLib
{
    public sealed class IgxlLoader
    {
        public List<SubFlowSheet> FlowSheets { get; } = [];
        public List<InstanceSheet> InstanceSheets { get; } = [];
        public List<BinTableSheet> BinTableSheets { get; } = [];
        public List<PinMapSheet> PinMapSheets { get; } = [];
        public List<ChannelMapSheet> ChannelMapSheets { get; } = [];
        public List<LevelSheet> LevelSheets { get; } = [];
        public List<AcSpecSheet> AcSpecSheets { get; } = [];
        public List<DcSpecSheet> DcSpecSheets { get; } = [];
        public List<TimeSetBasicSheet> TimeSetBasicSheets { get; } = [];
        public List<PatSetSheet> PatSetSheets { get; } = [];
        public List<PatSetSubSheet> PatSetSubSheets { get; set; } = [];
        public List<CharSheet> CharSheets { get; } = [];
        public JobListSheet? JobListSheet { get; internal set; }
        public GlobalSpecSheet? GlobalSpecSheet { get; internal set; }
        public PortMapSheet? PortMapSheet { get; private set; }
        public PatSetSheet? PatSetsAll { get; internal set; }

        private readonly IgxlSheetReaderRegistry _registry;

        public IgxlLoader(string testProgram, List<EnumSheetType> enumSheetTypes)
        {
            _registry = new IgxlSheetReaderRegistry(this);
            ReadSheets(testProgram, enumSheetTypes);
        }

        public IgxlLoader(ExcelWorkbook excelWorkbook, List<EnumSheetType> enumSheetTypes)
        {
            _registry = new IgxlSheetReaderRegistry(this);
            ReadSheets(excelWorkbook, enumSheetTypes);
        }

        private void ReadSheets(ExcelWorkbook excelWorkbook, List<EnumSheetType> enumSheetTypes)
        {
            foreach (ExcelWorksheet worksheet in excelWorkbook.Worksheets)
            {
                EnumSheetType sheetType = IgxlLoaderHelpers.GetIgxlSheetType(worksheet.Cells[1, 1].Value.ToString());
                if (enumSheetTypes.Contains(sheetType))
                {
                    ReadSheet(sheetType, worksheet, worksheet.Name);
                }
            }
        }

        private void ReadSheets(string igxl, List<EnumSheetType> enumSheetTypes)
        {
            using var package = Package.Open(igxl, FileMode.Open, FileAccess.Read);
            foreach (PackagePart part in package.GetParts())
            {
                string file = part.Uri.ToString().Replace("%20", " ").TrimStart('/');
                string sheetName = Path.GetFileNameWithoutExtension(file);
                EnumSheetType sheetType = IgxlLoaderHelpers.GetIgxlSheetType(part.GetStream(FileMode.Open, FileAccess.Read), part.Uri.ToString());
                if (enumSheetTypes.Contains(sheetType))
                {
                    ReadSheet(sheetType, part.GetStream(FileMode.Open, FileAccess.Read), sheetName);
                }
            }
        }

        private void ReadSheet(EnumSheetType enumSheetType, object source, string sheetName)
        {
            (IIgxlSheetReader reader, Action<object, string> assign) = _registry.GetMapping(enumSheetType);

            object result = reader.Read(source, sheetName);

            if (result != null)
            {
                assign(result, sheetName);
            }
        }

        public List<IIgxlSheet> GetAllSheets()
        {
            var allSheets = new List<IIgxlSheet>();
            if (JobListSheet != null)
            {
                allSheets.Add(JobListSheet);
            }

            if (GlobalSpecSheet != null)
            {
                allSheets.Add(GlobalSpecSheet);
            }

            if (PortMapSheet != null)
            {
                allSheets.Add(PortMapSheet);
            }

            if (PatSetsAll != null)
            {
                allSheets.Add(PatSetsAll);
            }

            allSheets.AddRange(FlowSheets);
            allSheets.AddRange(InstanceSheets);
            allSheets.AddRange(BinTableSheets);
            allSheets.AddRange(PinMapSheets);
            allSheets.AddRange(ChannelMapSheets);
            allSheets.AddRange(LevelSheets);
            allSheets.AddRange(AcSpecSheets);
            allSheets.AddRange(DcSpecSheets);
            allSheets.AddRange(TimeSetBasicSheets);
            allSheets.AddRange(PatSetSheets);
            allSheets.AddRange(PatSetSubSheets);
            allSheets.AddRange(CharSheets);

            return allSheets;
        }
    }
}
