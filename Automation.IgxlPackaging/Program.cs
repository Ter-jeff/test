using System;
using System.IO;

using Automation.IgxlPackaging.Contract;

using Newtonsoft.Json;

using Teradyne.Oasis.IGLinkBase;

namespace Automation.IgxlPackaging
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.Error.WriteLine("Usage: Automation.IgxlPackaging.exe <descriptor.json>");
                return 2;
            }

            string descriptorPath = args[0];
            if (!File.Exists(descriptorPath))
            {
                Console.Error.WriteLine($"Descriptor file not found: {descriptorPath}");
                return 2;
            }

            try
            {
                string json = File.ReadAllText(descriptorPath);
                PackageRequest request = JsonConvert.DeserializeObject<PackageRequest>(json);
                if (request == null || request.Projects == null)
                {
                    Console.Error.WriteLine("Descriptor parsed as empty.");
                    return 2;
                }

                foreach (DeviceProjectDescriptor project in request.Projects)
                {
                    BuildAndSave(project);
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"IGXL packager failed: {ex.GetType().FullName}: {ex.Message}");
                Console.Error.WriteLine(ex.ToString());
                return 1;
            }
        }

        private static void BuildAndSave(DeviceProjectDescriptor d)
        {
            var deviceProject = new DeviceProject
            {
                Name = d.Name,
                FileName = d.OutputProjFile,
                SaveAsXLS = d.SaveAsXLS,
                SheetOrder = ParseEnum(d.SheetOrder, SheetOrderPreference.Alphabetically),
            };

            var commonCode = new CommonCodeAndSheets();
            foreach (string source in d.CommonCode.SheetSources)
            {
                commonCode.Add(new Sheet { Source = source });
            }
            foreach (string source in d.CommonCode.VbFileSources)
            {
                commonCode.Add(new VBFile { Source = source });
            }

            foreach (SubProgramDescriptor sp in d.SubPrograms)
            {
                var subProgram = new SubProgram
                {
                    Name = sp.Name,
                    JobNames = sp.JobNames ?? string.Empty,
                };
                if (!string.IsNullOrEmpty(sp.MainFlow))
                {
                    subProgram.MainFlow = sp.MainFlow;
                }
                if (sp.GenerateJobListSheet.HasValue)
                {
                    subProgram.GenerateJobListSheet = sp.GenerateJobListSheet.Value;
                }
                if (!string.IsNullOrEmpty(sp.FlowGenMode))
                {
                    subProgram.FlowGenMode = ParseEnum(sp.FlowGenMode, FlowGenerationMode.UseMainFlowOnly);
                }
                foreach (string source in sp.SheetSources)
                {
                    subProgram.Add(new Sheet { Source = source });
                }
                foreach (string source in sp.VbFileSources)
                {
                    subProgram.Add(new VBFile { Source = source });
                }
                deviceProject.SubPrograms.Add(subProgram);
            }

            deviceProject.SubPrograms.CommonCode = commonCode;

            foreach (DeviceJobDescriptor jd in d.Jobs)
            {
                var job = new DeviceJob
                {
                    Name = jd.Name,
                    GenerateJobListSheet = jd.GenerateJobListSheet,
                    GenerateExecIPModule = jd.GenerateExecIPModule,
                    FlowGenMode = ParseEnum(jd.FlowGenMode, FlowGenerationMode.UseMainFlowOnly),
                    MainFlow = jd.MainFlow ?? string.Empty,
                    AppendToFlow = jd.AppendToFlow,
                    PinMap = jd.PinMap ?? string.Empty,
                    JobNames = jd.JobNames ?? string.Empty,
                    DefaultJob = jd.DefaultJob ?? string.Empty,
                    ChannelMapDisplayMode = ParseEnum(jd.ChannelMapDisplayMode, ChanDisplayMode.Signal),
                    IncludeOnlyOnePinMap = jd.IncludeOnlyOnePinMap,
                    IncludeOnlyDefaultChanMap = jd.IncludeOnlyDefaultChanMap,
                };
                if (jd.SubProgramNames != null)
                {
                    job.subprograms.AddRange(jd.SubProgramNames);
                    foreach (string subProgramName in jd.SubProgramNames)
                    {
                        job.AddSubProgram(subProgramName);
                    }
                }
                if (!string.IsNullOrEmpty(jd.DefaultChannelMap))
                {
                    job.DefaultChannelMap = jd.DefaultChannelMap;
                }
                deviceProject.Jobs.Add(job);
            }

            foreach (WorkbookDescriptor wd in d.WorkBooks)
            {
                var workbook = new IGXLWorkbook
                {
                    Name = wd.Name,
                    GenerateJobListSheet = wd.GenerateJobListSheet,
                    ChannelMapDisplayMode = ParseEnum(wd.ChannelMapDisplayMode, ChanDisplayMode.Signal),
                    DefaultJob = wd.DefaultJob ?? string.Empty,
                };
                if (!string.IsNullOrEmpty(wd.DefaultChannelMap))
                {
                    workbook.DefaultChannelMap = wd.DefaultChannelMap;
                }
                if (wd.Jobs != null)
                {
                    foreach (string jobName in wd.Jobs)
                    {
                        workbook.AddJob(jobName);
                    }
                }
                deviceProject.WorkBooks.Add(workbook);
            }

            string outputDir = Path.GetDirectoryName(d.OutputProjFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            Console.Out.WriteLine($"Saving IGLink project: {d.OutputProjFile}");
            DeviceProject.SaveProjectCfg(deviceProject);
            NormalizeRootNamespaceOrder(d.OutputProjFile);
        }

        // .NET Framework's XmlSerializer emits the root xmlns:xsi/xmlns:xsd attributes in
        // an order driven by an internal hashtable that has shifted between framework
        // patch levels. The baselines in Automation.Test/Expected/**/*.igxlProj were
        // captured with xsi declared first. Force that order so byte-level file diffs
        // stay clean across machines / framework revs.
        private static void NormalizeRootNamespaceOrder(string projectPath)
        {
            if (!File.Exists(projectPath))
            {
                return;
            }

            string content = File.ReadAllText(projectPath);
            const string xsiFirst = "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"";
            const string xsdFirst = "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"";
            if (content.IndexOf(xsdFirst, StringComparison.Ordinal) < 0 ||
                content.IndexOf(xsiFirst, StringComparison.Ordinal) >= 0)
            {
                return;
            }
            content = content.Replace(xsdFirst, xsiFirst);
            File.WriteAllText(projectPath, content);
        }

        private static T ParseEnum<T>(string value, T fallback) where T : struct
        {
            if (string.IsNullOrEmpty(value))
            {
                return fallback;
            }
            return Enum.TryParse(value, ignoreCase: true, out T parsed) ? parsed : fallback;
        }
    }
}
