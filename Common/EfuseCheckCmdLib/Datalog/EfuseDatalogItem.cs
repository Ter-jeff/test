namespace EfuseCheckCmdLib.Datalog
{
    public class EfuseDatalogItem(string block, string id, string rawData, string value, string order)
    {
        public string Block { get; set; } = block;
        public string Id { get; set; } = id;
        public string RawData { get; set; } = rawData;
        public string Value { get; set; } = value;
        public string Order { get; set; } = order;

        public override string ToString()
        {
            return $"Block: {Block}, Id: {Id}, RawData: {RawData}, Value: {Value}, Order: {Order}";
        }
    }
}
