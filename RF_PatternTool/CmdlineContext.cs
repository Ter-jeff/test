namespace RF_PatternTool
{
    public class CmdlineContext
    {
        #region Input / Output Path
        public string BenchLogPath { get; set; }
        public string OutputDir { get; set; }
        public string RegisterMapPath { get; set; }
        public string PinMapPath { get; set; }
        public string EfuseBitDefinitionPath { get; set; }
        public string ATPPath { get; set; }

        #endregion

        #region Project Setting
        public string ProjectName { get; set; }
        public string SiliconVersion { get; set; }
        public string PatternFolder { get; set; }

        #endregion

        #region Pattern Generation Setting
        public string PatternType { get; set; }
        public string PreSetupPatterns { get; set; }
        public string TimeStamp { get; set; }
        public string JtagFreq { get; set; }
        public bool IsR16 { get; set; }
        public bool IsFullSweep { get; set; }
        public string AddComment { get; set; }
        public string DebugMode { get; set; }

        #endregion

        #region Test Program Generation Setting
        public string TestPlanShellPath { get; set; }
        public string LibraryPath { get; set; }
        public string ProgramPath { get; set; }
        public string CppDspControl { get; set; }

        #endregion

        #region Generated Runtime Path
        public string PatternInfoPath { get; set; }
        public string ScghPath { get; set; }

        #endregion

        #region Runtime Data
        public HashSet<string> AddrFor64InBin { get; set; } = new HashSet<string>();

        #endregion
    }
}
