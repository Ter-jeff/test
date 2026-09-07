using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenIgxlProj;
using Automation.IgxlPackaging.Contract;
using Automation.Static;

using IgxlLib;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Basic;

namespace Automation.Test.UT.PostAction
{
    /// <summary>
    /// Covers the net8-testable logic that replaced the deleted <c>IgxlProjMain</c> when
    /// IGXL packaging moved into the net48 <c>Automation.IgxlPackaging</c> side-car:
    /// the packager factory, the JSON descriptor builders, and the path helpers.
    /// The actual IGLinkBase object-building now lives in the side-car exe and is out of scope.
    /// </summary>
    [TestClass]
    public class IgxlPackagerTests : TestBase
    {
        private static readonly string[] _jobNames = ["Job1", "Job2"];

        [TestInitialize]
        public void Setup()
        {
            // The descriptor builders call Response.Report for progress; silence it and
            // avoid touching real log sinks.

            LocalSpecs.CurrentProject = "TestProject";
            LocalSpecs.DefaultChannelMap = "DefaultMap";
            // FolderStructure.DirCommon/DirLib/DirMain derive from the target folder.
            LocalSpecs.TarFolder = Path.GetTempPath();

            // Force the "no TestProgram_Def sheet" (plain project) path deterministically,
            // independent of whatever workbook a previously-run test class left behind.
            TestPlanStatic.TestProgramDefSheet = new TestProgramDefSheet();
            TestPlanStatic.SubprogramMappingSheet = new SubprogramMappingSheet();
        }

        // ----------------------------------------------------------------- Factory

        [TestMethod]
        public void Create_ShouldReturnNoOp_WhenSkipIgLinkRequested()
        {
            IIgxlPackager packager = IgxlPackagerFactory.Create(skipIgLink: true);

            Assert.IsInstanceOfType(packager, typeof(NoOpIgxlPackager));
        }

        [TestMethod]
        public void Create_ShouldReturnProcessLauncher_WhenPackagerPathResolves()
        {
            string? originalPackager = Environment.GetEnvironmentVariable("IGXL_PACKAGER_PATH");
            string? originalMono = Environment.GetEnvironmentVariable("MONO_PATH");
            string fakePackager = Path.Combine(Path.GetTempPath(), $"pkg-{Guid.NewGuid():N}.exe");
            string fakeMono = Path.Combine(Path.GetTempPath(), $"mono-{Guid.NewGuid():N}");
            try
            {
                File.WriteAllText(fakePackager, "stub");
                File.WriteAllText(fakeMono, "stub");
                Environment.SetEnvironmentVariable("IGXL_PACKAGER_PATH", fakePackager);
                // MONO_PATH only matters on non-Windows; harmless on Windows.
                Environment.SetEnvironmentVariable("MONO_PATH", fakeMono);

                IIgxlPackager packager = IgxlPackagerFactory.Create();

                Assert.IsInstanceOfType(packager, typeof(IgxlPackagerProcessLauncher));
            }
            finally
            {
                Environment.SetEnvironmentVariable("IGXL_PACKAGER_PATH", originalPackager);
                Environment.SetEnvironmentVariable("MONO_PATH", originalMono);
                File.Delete(fakePackager);
                File.Delete(fakeMono);
            }
        }

        [TestMethod]
        public void Create_ShouldAlwaysReturnAPackager_WithNoOverrides()
        {
            string? originalPackager = Environment.GetEnvironmentVariable("IGXL_PACKAGER_PATH");
            string? originalMono = Environment.GetEnvironmentVariable("MONO_PATH");
            try
            {
                Environment.SetEnvironmentVariable("IGXL_PACKAGER_PATH", null);
                Environment.SetEnvironmentVariable("MONO_PATH", null);

                // Resolution depends on the OS and whether the side-car is bundled next to the
                // test assembly, so we only assert a packager is always produced (never null).
                IIgxlPackager packager = IgxlPackagerFactory.Create();

                Assert.AreNotEqual(null, packager);
            }
            finally
            {
                Environment.SetEnvironmentVariable("IGXL_PACKAGER_PATH", originalPackager);
                Environment.SetEnvironmentVariable("MONO_PATH", originalMono);
            }
        }

        // ----------------------------------------------------------------- Factory helpers (OS-independent)

        [TestMethod]
        public void LocateMono_ShouldReturnExplicitPath_WhenMonoPathSetToExistingFile()
        {
            string? original = Environment.GetEnvironmentVariable("MONO_PATH");
            string fakeMono = Path.Combine(Path.GetTempPath(), $"mono-{Guid.NewGuid():N}");
            try
            {
                File.WriteAllText(fakeMono, "stub");
                Environment.SetEnvironmentVariable("MONO_PATH", fakeMono);

                Assert.AreEqual(fakeMono, IgxlPackagerFactory.LocateMono());
            }
            finally
            {
                Environment.SetEnvironmentVariable("MONO_PATH", original);
                File.Delete(fakeMono);
            }
        }

        [TestMethod]
        public void LocateMono_ShouldFindMonoOnPath_WhenNoExplicitOverride()
        {
            string? originalMono = Environment.GetEnvironmentVariable("MONO_PATH");
            string? originalPath = Environment.GetEnvironmentVariable("PATH");
            string binDir = Path.Combine(Path.GetTempPath(), $"monobin-{Guid.NewGuid():N}");
            try
            {
                Directory.CreateDirectory(binDir);
                string monoOnPath = Path.Combine(binDir, "mono");
                File.WriteAllText(monoOnPath, "stub");

                Environment.SetEnvironmentVariable("MONO_PATH", null);
                // Leading empty entry exercises the string.IsNullOrEmpty(dir) continue.
                Environment.SetEnvironmentVariable("PATH", Path.PathSeparator + binDir);

                Assert.AreEqual(monoOnPath, IgxlPackagerFactory.LocateMono());
            }
            finally
            {
                Environment.SetEnvironmentVariable("MONO_PATH", originalMono);
                Environment.SetEnvironmentVariable("PATH", originalPath);
                Directory.Delete(binDir, true);
            }
        }

        [TestMethod]
        public void LocateMono_ShouldReturnNull_WhenMonoNotFound()
        {
            string? originalMono = Environment.GetEnvironmentVariable("MONO_PATH");
            string? originalPath = Environment.GetEnvironmentVariable("PATH");
            string emptyDir = Path.Combine(Path.GetTempPath(), $"nomono-{Guid.NewGuid():N}");
            try
            {
                Directory.CreateDirectory(emptyDir);
                Environment.SetEnvironmentVariable("MONO_PATH", null);
                Environment.SetEnvironmentVariable("PATH", emptyDir);

                Assert.AreEqual(null, IgxlPackagerFactory.LocateMono());
            }
            finally
            {
                Environment.SetEnvironmentVariable("MONO_PATH", originalMono);
                Environment.SetEnvironmentVariable("PATH", originalPath);
                Directory.Delete(emptyDir, true);
            }
        }

        [TestMethod]
        public void LocateSidecarExe_ShouldReturnExplicitPath_WhenEnvVarPointsAtExistingFile()
        {
            string? original = Environment.GetEnvironmentVariable("IGXL_PACKAGER_PATH");
            string fakeExe = Path.Combine(Path.GetTempPath(), $"pkg-{Guid.NewGuid():N}.exe");
            try
            {
                File.WriteAllText(fakeExe, "stub");
                Environment.SetEnvironmentVariable("IGXL_PACKAGER_PATH", fakeExe);

                Assert.AreEqual(fakeExe, IgxlPackagerFactory.LocateSidecarExe());
            }
            finally
            {
                Environment.SetEnvironmentVariable("IGXL_PACKAGER_PATH", original);
                File.Delete(fakeExe);
            }
        }

        [TestMethod]
        public void CandidateBaseDirectories_ShouldYieldThreeProbeLocations()
        {
            Assert.AreEqual(3, IgxlPackagerFactory.CandidateBaseDirectories().Count());
        }

        // ----------------------------------------------------------------- Path helpers

        [TestMethod]
        public void IsVbFile_ShouldRecognizeVbSourceExtensions()
        {
            Assert.IsTrue(IgxlPackagerProcessLauncher.IsVbFile(Path.Combine("C:", "x", "Module.bas")));
            Assert.IsTrue(IgxlPackagerProcessLauncher.IsVbFile(Path.Combine("C:", "x", "Module.CLS")));
            Assert.IsTrue(IgxlPackagerProcessLauncher.IsVbFile(Path.Combine("C:", "x", "Form.frm")));
            Assert.IsTrue(IgxlPackagerProcessLauncher.IsVbFile(Path.Combine("C:", "x", "Form.FRX")));
            Assert.IsFalse(IgxlPackagerProcessLauncher.IsVbFile(Path.Combine("C:", "x", "Sheet.txt")));
        }

        [TestMethod]
        public void MakeRelative_ShouldStripRefPath_WhenPathIsUnderIt()
        {
            string refPath = Path.Combine("root", "IGLink");
            string full = Path.Combine(refPath, "HardIP", "Flow.txt");

            string result = IgxlPackagerProcessLauncher.MakeRelative(full, refPath);

            Assert.AreEqual(Path.Combine("HardIP", "Flow.txt"), result);
        }

        [TestMethod]
        public void MakeRelative_ShouldReturnInput_WhenRefPathEmptyOrUnrelated()
        {
            Assert.AreEqual(Path.Combine("C:", "x", "Flow.txt"), IgxlPackagerProcessLauncher.MakeRelative(Path.Combine("C:", "x", "Flow.txt"), string.Empty));
            Assert.AreEqual(Path.Combine("C:", "x", "Flow.txt"), IgxlPackagerProcessLauncher.MakeRelative(Path.Combine("C:", "x", "Flow.txt"), Path.Combine("C:", "other")));
        }

        [TestMethod]
        public void StartsWithStaticDir_ShouldMatchLibCommonAndMainRoots()
        {
            Assert.IsTrue(IgxlPackagerProcessLauncher.StartsWithStaticDir(Path.Combine(FolderStructure.DirCommon, "s.txt")));
            Assert.IsTrue(IgxlPackagerProcessLauncher.StartsWithStaticDir(Path.Combine(FolderStructure.DirLib, "s.txt")));
            Assert.IsTrue(IgxlPackagerProcessLauncher.StartsWithStaticDir(Path.Combine(FolderStructure.DirMain, "s.txt")));
            Assert.IsFalse(IgxlPackagerProcessLauncher.StartsWithStaticDir(Path.Combine(Path.GetTempPath(), "elsewhere", "s.txt")));
        }

        // ----------------------------------------------------------------- Job / workbook descriptors

        [TestMethod]
        public void BuildJobDescriptor_ShouldSetChannelMap_WhenProvided()
        {
            DeviceJobDescriptor job = IgxlPackagerProcessLauncher.BuildJobDescriptor("Job1", "Main_Flow_Job1", ["Job1", "Job2"], [new() { Name = "SubA" }], "MapA");

            Assert.AreEqual("Job1", job.Name);
            Assert.AreEqual("Job1", job.DefaultJob);
            Assert.AreEqual("Job1,Job2", job.JobNames);
            Assert.IsTrue(job.MainFlow.Contains("Main_Flow_Job1"));
            CollectionAssert.Contains(job.SubProgramNames, "SubA");
            Assert.IsTrue(job.DefaultChannelMap.Contains("MapA"));
        }

        [TestMethod]
        public void BuildJobDescriptor_ShouldOmitChannelMap_WhenEmpty()
        {
            DeviceJobDescriptor job = IgxlPackagerProcessLauncher.BuildJobDescriptor("Job1", "Main_Flow_Job1", ["Job1"], [], string.Empty);

            Assert.AreEqual(null, job.DefaultChannelMap);
        }

        [TestMethod]
        public void BuildWorkbookDescriptor_ShouldListJobsAndHonorChannelMap()
        {
            var jobs = new List<DeviceJobDescriptor>
            {
                new() { Name = "Job1" },
                new() { Name = "Job2" },
            };

            WorkbookDescriptor withMap = IgxlPackagerProcessLauncher.BuildWorkbookDescriptor(jobs, "MapA", "Job1");
            Assert.AreEqual("Job1", withMap.DefaultJob);
            CollectionAssert.AreEqual(_jobNames, withMap.Jobs.ToArray());
            Assert.IsTrue(withMap.DefaultChannelMap.Contains("MapA"));

            WorkbookDescriptor noMap = IgxlPackagerProcessLauncher.BuildWorkbookDescriptor(jobs, string.Empty, "Job1");
            Assert.AreEqual(null, noMap.DefaultChannelMap);
        }

        // ----------------------------------------------------------------- BuildDescriptor
        [TestMethod]
        public void BuildDescriptor_PlainProject_ShouldSplitCommonCodeSubProgramsAndJobs()
        {
            string igLinkFolder = NewTempFolder();
            IgxlWorkBook workbook = WorkbookWithSingleJob(defaultChannelMapName: "DefaultMap");

            var sources = new List<string>
            {
                Path.Combine(FolderStructure.DirCommon, "CommonSheet.txt"),      // -> common code sheet
                Path.Combine(igLinkFolder, "HardIP", "Flow_HardIP.txt"),         // -> folder sub-program "HardIP"
                Path.Combine(igLinkFolder, "HardIP", "Inst_HardIP.txt"),         // -> appended to "HardIP"
                Path.Combine(igLinkFolder, "VBT_Instrument_Setup.txt"),          // -> skipped
                Path.Combine(FolderStructure.DirLib, "Library.bas"),             // -> common code vb file
            };

            DeviceProjectDescriptor descriptor = IgxlPackagerProcessLauncher.BuildDescriptor(sources, "TestProject", workbook, igLinkFolder);

            Assert.AreNotEqual(null, descriptor);
            Assert.AreEqual("ProjectTemple", descriptor.Name);
            StringAssert.EndsWith(descriptor.OutputProjFile, "TestProject.igxlProj");
            Assert.AreEqual("DefaultMap", descriptor.DefaultChannelMap);

            Assert.IsTrue(descriptor.CommonCode.SheetSources.Any(s => s.Contains("CommonSheet.txt")));
            Assert.IsTrue(descriptor.CommonCode.VbFileSources.Any(s => s.Contains("Library.bas")));

            // subProgramTotal + the "HardIP" folder sub-program.
            Assert.IsTrue(descriptor.SubPrograms.Any(s => s.Name == "TestProject"));
            SubProgramDescriptor? hardIp = descriptor.SubPrograms.FirstOrDefault(s => s.Name == "HardIP");
            Assert.AreNotEqual(null, hardIp);
            Assert.AreEqual(2, hardIp!.SheetSources.Count);

            Assert.AreEqual(1, descriptor.Jobs.Count);
            Assert.AreEqual("Job1", descriptor.Jobs[0].Name);
            Assert.IsTrue(descriptor.Jobs[0].MainFlow.Contains("Main_Flow_Job1"));
            Assert.AreEqual(1, descriptor.WorkBooks.Count);
        }

        [TestMethod]
        public void BuildDescriptor_ShouldClearChannelMap_WhenDefaultNotInWorkbook()
        {
            string igLinkFolder = NewTempFolder();
            // Workbook has NO channel map matching LocalSpecs.DefaultChannelMap.
            IgxlWorkBook workbook = WorkbookWithSingleJob(defaultChannelMapName: null);

            DeviceProjectDescriptor descriptor = IgxlPackagerProcessLauncher.BuildDescriptor([Path.Combine(igLinkFolder, "HardIP", "Flow.txt")], "TestProject", workbook, igLinkFolder);

            Assert.AreEqual(string.Empty, descriptor.DefaultChannelMap);
        }

        [TestMethod]
        public void BuildDescriptor_ShouldProduceNoJobs_WhenNoMainFlows()
        {
            string igLinkFolder = NewTempFolder();
            // no main flows
            var workbook = new IgxlWorkBook();

            DeviceProjectDescriptor descriptor = IgxlPackagerProcessLauncher.BuildDescriptor([Path.Combine(igLinkFolder, "HardIP", "Flow.txt")], "TestProject", workbook, igLinkFolder);

            Assert.AreEqual(0, descriptor.Jobs.Count);
            Assert.AreEqual(0, descriptor.WorkBooks.Count);
            StringAssert.EndsWith(descriptor.SubPrograms[0].MainFlow, "TestProject:");
        }

        [TestMethod]
        public void BuildDescriptor_SubProgram_ShouldRouteSheetsAndSetSubMainFlow()
        {
            string igLinkFolder = NewTempFolder();
            IgxlWorkBook workbook = WorkbookWithSingleJob(defaultChannelMapName: "DefaultMap");

            TestPlanStatic.SubprogramMappingSheet = new SubprogramMappingSheet
            {
                SubprogramSettings = { new SubprogramSetting("Sub1", []) },
            };

            var sources = new List<string>
            {
                Path.Combine(igLinkFolder, "Main_Sub1.txt"),                 // -> Sub1 (Main_ prefix, EndsWith Name)
                Path.Combine(igLinkFolder, "JobList_Sub1.txt"),              // -> JobList_ prefix branch
                Path.Combine(FolderStructure.DirCommon, "CommonSheet.txt"),  // -> common code (static dir)
                Path.Combine(igLinkFolder, "Inst_Generic.txt"),             // -> broadcast to all sub-programs
            };

            DeviceProjectDescriptor descriptor = IgxlPackagerProcessLauncher.BuildDescriptor(sources, "TestProject_Sub", workbook, igLinkFolder);

            SubProgramDescriptor? sub1 = descriptor.SubPrograms.FirstOrDefault(s => s.Name == "Sub1");
            Assert.AreNotEqual(null, sub1);
            Assert.AreEqual("Sub1:Main_Sub1", sub1!.MainFlow);
            Assert.AreEqual("UseMainFlowOnly", sub1.FlowGenMode);
            Assert.IsTrue(sub1.SheetSources.Any(s => s.Contains("Main_Sub1")));
            // broadcast sheet
            Assert.IsTrue(sub1.SheetSources.Any(s => s.Contains("Inst_Generic")));
            Assert.IsTrue(descriptor.CommonCode.SheetSources.Any(s => s.Contains("CommonSheet.txt")));
        }

        // ----------------------------------------------------------------- NoOp packager

        [TestMethod]
        public void NoOpIgxlPackager_GenIgxlProg_ShouldNotThrow()
        {
            var packager = new NoOpIgxlPackager("test reason");

            packager.GenIgxlProg([], Path.GetTempPath(), "TestProject", new IgxlWorkBook(), isUnitTest: true);
        }

        // ----------------------------------------------------------------- helpers

        private static string NewTempFolder()
        {
            string dir = Path.Combine(Path.GetTempPath(), $"igxltest-{Guid.NewGuid():N}");
            Directory.CreateDirectory(dir);
            return dir;
        }

        private static IgxlWorkBook WorkbookWithSingleJob(string? defaultChannelMapName)
        {
            var workbook = new IgxlWorkBook();
            workbook.MainFlowSheets["Main_Flow_Job1"] = new MainFlow("Main_Flow_Job1")
            {
                Name = "Main_Flow_Job1",
                JobNames = ["Job1"],
            };
            if (!string.IsNullOrEmpty(defaultChannelMapName))
            {
                workbook.ChannelMapSheets[defaultChannelMapName] = new ChannelMapSheet(defaultChannelMapName)
                {
                    Name = defaultChannelMapName,
                };
            }
            return workbook;
        }
    }
}
