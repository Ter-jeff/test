namespace TestPlanLib.DataStruct
{
    public class IoLevelsItem
    {
        #region Field
        public bool IsSameDomain;
        public string Level;
        public string Domain;
        public string Vdd;
        public string Vih;
        public string Vil;
        public string Voh;
        public string Vol;
        #endregion

        public IoLevelsItem()
        {
            IsSameDomain = true;
            Level = "";
            Domain = "";
            Vdd = "";
            Vih = "";
            Vil = "";
            Voh = "";
            Vol = "";
        }

        public IoLevelsItem(string level)
        {
            IsSameDomain = true;
            Level = level;
            Domain = "";
            Vdd = "";
            Vih = "";
            Vil = "";
            Voh = "";
            Vol = "";
        }

        public object Clone(string level)
        {
            var ioLevelsItem = new IoLevelsItem
            {
                Level = level,
                Vdd = Vdd,
                IsSameDomain = true,
                Domain = Domain,
                Vih = Vih,
                Vil = Vil,
                Voh = Voh,
                Vol = Vol
            };
            return ioLevelsItem;
        }
    }
}
