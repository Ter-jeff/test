using System.Collections.Generic;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public enum InterposeAssignType
    {
        None = 0,
        InterposePreInit = 1,
        InterposePrePat = 2,
        InterposePostInit = 3,
        InterposePreMeas = 4,
        InterposePostMeas = 5,
        InterposePreRst = 6,
        InterposePostPat = 7,
        InterposePostRst = 8,
        StartOfBodyFArgs = 9,
        EndOfBodyFArgs = 10,
    }

    public class InterposeAssign
    {
        public List<string> InterposeAssignList = new List<string>();
        public string AssignName { get; set; }
        public string BlockName { get; set; }
        public InterposeAssignType Type { get; set; }

        public const string ColSetupName = "SetupName";
        public const string ColInterposes = "Interposes";

        public List<string> Headers { get; } = new List<string> { "SetupName", "Interposes" };
    }
}
