using System.Collections.Generic;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    public class SetEfuseItem
    {
        public string ItemName = "";
        public List<EFuseStoreRow> Contents = [];
        public List<string> Headers = ["CaptureStoreName", "FieldName", "FuseEnable"];
    }
}
