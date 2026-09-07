using System.Collections.Generic;

namespace Cautogen.common.ReaderWriter.Reader.InputDataBase
{
    public class DigSrcReg : Dictionary<string, List<DigSrcRegData>>
    {

    }

    public class DigSrcRegData
    {
        public string RegName { get; set; }
        public string RegValue { get; set; }
    }
}
