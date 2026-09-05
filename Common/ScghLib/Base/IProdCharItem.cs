using System.Collections.Generic;

namespace ScghLib.Base
{
    public interface IProdCharItem
    {
        int RowNum { set; get; }
        string PayloadValue { get; }
        string Block { set; get; }
        string Mode { set; get; }
        string Item { get; }
        string Usage { get; }
        string Inits { set; get; }
        string PayLoads { set; get; }
        string Chiplet { get; }

        List<string> GetInitList();

        List<string> GetPayloadList();

        string GetSourceSheet();
    }
}
