namespace CommonLib.Datalog
{
    public class LimitRow
    {
        public string Number = "";
        public int Site;
        public string TestName = "";
        public string Pin = "";
        public double Channel;
        public double LowLimit;
        public double Measured;
        public double HighLimit;
        public int Force = -1;
        public int Loc = -1;

        public LimitRow() { }

        public LimitRow(LimitRow limitRow)
        {
            if (limitRow == null)
            {
                return;
            }

            Number = limitRow.Number;
            Site = limitRow.Site;
            TestName = limitRow.TestName;
            Pin = limitRow.Pin;
            Channel = limitRow.Channel;
            LowLimit = limitRow.LowLimit;
            Measured = limitRow.Measured;
            HighLimit = limitRow.HighLimit;
            Force = limitRow.Force;
            Loc = limitRow.Loc;
        }

        public LimitRow Copy()
        {
            return new LimitRow(this);
        }
    }
}
