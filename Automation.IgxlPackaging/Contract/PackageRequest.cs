using System.Collections.Generic;

namespace Automation.IgxlPackaging.Contract
{
    /// <summary>
    /// Wire format between the .NET 8 Autogen launcher and the net48 IGLinkBase sidecar.
    /// One file describes one or more device projects to build.
    /// Schema changes here must stay backward-compatible or both sides must be redeployed together.
    /// </summary>
    public class PackageRequest
    {
        public List<DeviceProjectDescriptor> Projects { get; set; } = [];
    }

    public class DeviceProjectDescriptor
    {
        public string Name { get; set; }
        public string OutputProjFile { get; set; }
        public bool SaveAsXLS { get; set; }
        public string SheetOrder { get; set; }
        public string CurrentProject { get; set; }
        public string DefaultChannelMap { get; set; }
        public List<SubProgramDescriptor> SubPrograms { get; set; } = [];
        public CommonCodeDescriptor CommonCode { get; set; } = new CommonCodeDescriptor();
        public List<DeviceJobDescriptor> Jobs { get; set; } = [];
        public List<WorkbookDescriptor> WorkBooks { get; set; } = [];
    }

    public class SubProgramDescriptor
    {
        public string Name { get; set; }
        public string JobNames { get; set; }
        public string MainFlow { get; set; }
        public bool? GenerateJobListSheet { get; set; }
        public string FlowGenMode { get; set; }
        public List<string> SheetSources { get; set; } = [];
        public List<string> VbFileSources { get; set; } = [];
    }

    public class CommonCodeDescriptor
    {
        public List<string> SheetSources { get; set; } = [];
        public List<string> VbFileSources { get; set; } = [];
    }

    public class DeviceJobDescriptor
    {
        public string Name { get; set; }
        public string MainFlow { get; set; }
        public string PinMap { get; set; }
        public string JobNames { get; set; }
        public List<string> SubProgramNames { get; set; } = [];
        public string DefaultChannelMap { get; set; }
        public string DefaultJob { get; set; }
        public string ChannelMapDisplayMode { get; set; }
        public bool GenerateJobListSheet { get; set; }
        public bool GenerateExecIPModule { get; set; }
        public bool AppendToFlow { get; set; }
        public bool IncludeOnlyOnePinMap { get; set; }
        public bool IncludeOnlyDefaultChanMap { get; set; }
        public string FlowGenMode { get; set; }
    }

    public class WorkbookDescriptor
    {
        public string Name { get; set; }
        public string DefaultChannelMap { get; set; }
        public string DefaultJob { get; set; }
        public string ChannelMapDisplayMode { get; set; }
        public bool GenerateJobListSheet { get; set; }
        public List<string> Jobs { get; set; } = [];
    }
}
