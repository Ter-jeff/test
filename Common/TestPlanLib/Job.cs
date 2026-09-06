using CommonLib.Extension;

using IgxlLib.Enums;

namespace TestPlanLib
{
    public class Job(string jobName)
    {
        public string JobName { get; } = jobName;
        private EnumJob _jobType = EnumJob.None;
        public EnumJob JobType
        {
            get
            {
                if (_jobType != EnumJob.None)
                {
                    return _jobType;
                }

                string jobName = JobName.ToUpper();
                //could be WLFT_QA/FT2_25C_QA/FT2_85C_QA...
                if (jobName.EqualsIgnoreCase("QA") || jobName.EqualsIgnoreCase("FT1_FQA") || jobName.EqualsIgnoreCase("FT2_FQA"))
                {
                    return EnumJob.QA;
                }

                if (jobName.EqualsIgnoreCase("CP1"))
                {
                    return EnumJob.CP1;
                }

                if (jobName.EqualsIgnoreCase("CP2"))
                {
                    return EnumJob.CP2;
                }

                if (jobName.EqualsIgnoreCase("FT1") || jobName.EqualsIgnoreCase("FT2_25C") || jobName.EqualsIgnoreCase("WLFT") ||
                    jobName.EqualsIgnoreCase("WLFT1") || jobName.EqualsIgnoreCase("FT_ROOM") || jobName.EqualsIgnoreCase("RMA_ROOM"))
                {
                    return EnumJob.FT1;
                }

                if (jobName.EqualsIgnoreCase("FT2") || jobName.EqualsIgnoreCase("FT2_85C") || jobName.EqualsIgnoreCase("WLFT2") ||
                    jobName.EqualsIgnoreCase("FT_HOT") || jobName.EqualsIgnoreCase("RMA_HOT"))
                {
                    return EnumJob.FT2;
                }

                return EnumJob.None;
            }
            set { _jobType = value; }
        }
    }
}
