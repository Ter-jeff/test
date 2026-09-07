using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager;
using Cautogen.common.IgxlProgramLib.IgxlProgramParser;

using IgxlLib;
using IgxlLib.Enums;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using Ionic.Zip;

using LogLib.Utility;

namespace DebugPlanReaderLib.DebugPlan
{
    public class IgxlDataReader
    {
        private readonly string _jobName;
        public List<AcSpecSheet> AcSpecSheets = new List<AcSpecSheet>();
        public List<DcSpecSheet> DcSpecSheets = new List<DcSpecSheet>();
        public List<SubFlowSheet> FlowSheets = new List<SubFlowSheet>();
        public List<InstanceSheet> InstanceSheets = new List<InstanceSheet>();
        public PatSetSheet PatSetsAll;
        public List<PatSetSheet> PatSetsSheets = new List<PatSetSheet>();
        public PatSetSubSheet PatSetSubSheet = new PatSetSubSheet("Pattern_Subroutine");
        public List<PinMapSheet> PinMapSheets = new List<PinMapSheet>();
        public List<string> TimeSetBasicSheetNames = new List<string>();
        public List<TimeSetBasicSheet> TimeSetBasicSheets = new List<TimeSetBasicSheet>();
        public VbtFunctionLib VbtFunctionLib = new VbtFunctionLib();
        public List<PortMapSheet> PortMapSheets = new List<PortMapSheet>();
        public List<LevelSheet> LevelSheets = new List<LevelSheet>();
        public List<BinTableSheet> BinTableSheets = new List<BinTableSheet>();
        public List<ReferenceSheet> ReferenceSheets = new List<ReferenceSheet>();

        public JobListSheet JobListSheet { get; set; }
        public GlobalSpecSheet GlobalSpecSheet { get; set; }

        public IgxlDataReader(string testProgram)
        {
            using (var zip = new ZipFile(testProgram))
            {
                foreach (ZipEntry entry in zip.Entries)
                {
                    string sheetName = Path.GetFileNameWithoutExtension(entry.FileName);

                    EnumSheetType sheetType;
                    using (var reader = new StreamReader(entry.OpenReader()))
                    {
                        sheetType = IgxlLoaderHelpers.GetIgxlSheetType(reader.ReadLine());
                    }

                    if (sheetType == EnumSheetType.DTTimesetBasicSheet)
                    {
                        TimeSetBasicSheetNames.Add(sheetName);
                    }
                    else if (sheetType == EnumSheetType.DTDCSpecSheet)
                    {
                        DcSpecSheets.Add(new ReadDcSpecSheet().ReadSheet(entry.OpenReader(), sheetName));
                    }
                    else if (sheetType == EnumSheetType.DTACSpecSheet)
                    {
                        AcSpecSheets.Add(new ReadAcSpecSheet().ReadSheet(entry.OpenReader(), sheetName));
                    }
                    else if (sheetType == EnumSheetType.DTPinMap)
                    {
                        PinMapSheets.Add(new ReadPinMapSheet().ReadSheet(entry.OpenReader(), sheetName));
                    }
                    else if (sheetType == EnumSheetType.DTPatternSetSheet)
                    {
                        PatSetsSheets.Add(new ReadPatSetSheet().ReadSheet(entry.OpenReader(), sheetName));
                    }
                    else if (sheetType == EnumSheetType.DTTestInstancesSheet)
                    {
                        InstanceSheets.Add(
                            new ReadInstanceSheet().ReadSheet(entry.OpenReader(), sheetName));
                    }
                }
            }
        }

        public IgxlDataReader(string testProgram, string jobName, bool isCSharp = false, string csharpLibraryFolder = null, string dllOutPutFolder = null)
        {
            //VbtFunctionLib.Read(testProgram, isCSharp, csharpLibraryFolder);
            LogHelper.Info($"Build CS Library");
            LogHelper.Info($"{csharpLibraryFolder}");
            //VbtFunctionLib.BuildCsLibrary(csharpLibraryFolder, dllOutPutFolder);
            BasMain.Parse(csharpLibraryFolder, true);
            VbtFunctionLib = BasMain.VbtFunctionLib;

            var exportDir = Path.Combine(Path.GetDirectoryName(testProgram), "exportProg");

            IgxlManager.ExportWorkBook(testProgram, exportDir);

            var _igxlProgram = new IgxlProgram(jobName);
            _igxlProgram.LoadIgxlForPreProcess(exportDir);
            AcSpecSheets = _igxlProgram.AcSpecSheets;
            DcSpecSheets = _igxlProgram.DcSpecSheets;
            FlowSheets = _igxlProgram.FlowSheets;
            InstanceSheets = _igxlProgram.InstanceSheets;
            PatSetsAll = _igxlProgram.PatSetSheets.FirstOrDefault(p => p.Name == "PatSets_All");
            PatSetsSheets = _igxlProgram.PatSetSheets;
            PatSetSubSheet = _igxlProgram.PatSetSubSheets.FirstOrDefault(p => p.Name == "Pattern_Subroutine");
            PinMapSheets = _igxlProgram.PinMaps;
            TimeSetBasicSheets = _igxlProgram.TimeSetSheets;
            PortMapSheets = _igxlProgram.PortMapSheets;
            LevelSheets = _igxlProgram.LevelSheets;
            BinTableSheets = _igxlProgram.BintableSheets;
            GlobalSpecSheet = _igxlProgram.GlbSpecSheet;
            JobListSheet = _igxlProgram.JoblistSheet;
            ReferenceSheets = _igxlProgram.ReferenceSheets;

            _jobName = jobName;
        }

        public AcSpecSheet CurrentAcSpecSheet
        {
            get
            {
                if (JobListSheet != null)
                {
                    var jobRow = JobListSheet.GetRow(_jobName);
                    if (jobRow != null)
                    {
                        if (AcSpecSheets.Exists(x =>
                                x.Name.Equals(jobRow.AcSpecs, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            return AcSpecSheets.Find(x =>
                                x.Name.Equals(jobRow.AcSpecs, StringComparison.CurrentCultureIgnoreCase));
                        }
                    }
                }

                if (AcSpecSheets.Count > 0)
                {
                    return AcSpecSheets.First();
                }

                return null;
            }
        }

        public PinMapSheet CurrentPinMapSheet
        {
            get
            {
                if (JobListSheet != null)
                {
                    var jobRow = JobListSheet.GetRow(_jobName);
                    if (jobRow != null)
                    {
                        if (PinMapSheets.Exists(x =>
                                x.Name.Equals(jobRow.PinMap, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            return PinMapSheets.Find(x =>
                                x.Name.Equals(jobRow.PinMap, StringComparison.CurrentCultureIgnoreCase));
                        }
                    }
                }

                if (PinMapSheets.Count > 0)
                {
                    return PinMapSheets.First();
                }

                return null;
            }
        }

        public PortMapSheet CurrentPortMapSheet
        {
            get
            {
                if (JobListSheet != null)
                {
                    var jobRow = JobListSheet.GetRow(_jobName);
                    if (jobRow != null)
                    {
                        if (PortMapSheets.Exists(x =>
                                x.Name.Equals(jobRow.PortMap, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            return PortMapSheets.Find(x =>
                                x.Name.Equals(jobRow.PortMap, StringComparison.CurrentCultureIgnoreCase));
                        }
                    }
                }

                if (PortMapSheets.Count > 0)
                {
                    return PortMapSheets.First();
                }

                return null;
            }
        }
    }
}
