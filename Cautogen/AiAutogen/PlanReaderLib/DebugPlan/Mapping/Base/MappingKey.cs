namespace DebugPlanReaderLib.DebugPlan.Mapping.Base
{
    public class MappingKey
    {
        private string _patternSet;
        private string _payload;
        private string _pmode;
        public string PatternSet
        {
            get
            {
                return _patternSet;
            }
            set
            {
                _patternSet = value.ToUpper();
            }
        }
        public string Payload
        {
            get
            {
                return _payload;
            }
            set
            {
                _payload = value.ToUpper();
            }
        }
        public string Pmode
        {
            get
            {
                return _pmode;
            }
            set
            {
                _pmode = value.ToUpper();
            }
        }
        public string Domain
        {
            get
            {
                if (_payload.Split('_').Length >= 3)
                {
                    return _payload.Split('_')[2];
                }
                return "";
            }
        }
        public string Cluster
        {
            get
            {
                if (_payload.Split('_').Length >= 6)
                {
                    return _payload.Split('_')[5];
                }
                return "";
            }
        }
        public string BlockType
        {
            get
            {
                if (_payload.Split('_').Length >= 7)
                {
                    return _payload.Split('_')[6];
                }
                return "";
            }
        }
    }
}
