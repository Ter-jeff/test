using System.Collections.Generic;

namespace Cautogen.common.ReaderWriter.Reader.InputDataBase
{
    public class AdaptiveCooling : Dictionary<string, AdaptiveCoolingData>
    {
    }

    public class AdaptiveCoolingData
    {

        public string TemperatureC { get; set; }
        public string Enable { get; set; }
        public string MinDeltaC { get; set; }
        public string MaxDeltaC { get; set; }
        public string TimeoutSec { get; set; }
    }
}
