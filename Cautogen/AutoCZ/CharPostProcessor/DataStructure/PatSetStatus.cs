namespace Cautogen.AutoCZ.CharPostProcessor.DataStructure
{
    public class PatSetStatus
    {
        public string Used { get; set; }
        public string Existed { get; set; }
        public string ValidTs { get; set; }
        public string ContainSubr { get; set; }

        public PatSetStatus()
        {
            Used = "NoCheck";//1st check rule
            ValidTs = "NoCheck";//2nd check rule
            Existed = "NoCheck";//3rd check rule
            ContainSubr = "NoCheck"; // 4th check rule
        }
    }
}
