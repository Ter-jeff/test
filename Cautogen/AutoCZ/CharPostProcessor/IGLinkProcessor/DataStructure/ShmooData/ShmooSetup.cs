using System.Collections.Generic;

namespace Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.ShmooData
{
    public class ShmooSetup
    {
        public bool IsUseCmd;
        public string ShmooSetupName;
        public string PlanShmooSetupName;
        public string Timeset;
        public string SearchMethod;
        public string SuspendDatalog;
        public List<ShmooPin> ShmooPins;

        public ShmooSetup()
        {
            ShmooPins = new List<ShmooPin>();
        }
    }
}
