namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class PatChildRow
    {
        public bool IsMerged { get; set; }

        public PatChildRow()
        {
        }

        protected PatChildRow(PatChildRow other)
        {
            if (other == null)
            {
                return;
            }

            IsMerged = other.IsMerged;
        }

        public PatChildRow Copy()
        {
            return new PatChildRow(this);
        }
    }
}
