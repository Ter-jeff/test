using System.Collections.Generic;

namespace RfLib.Dvdc.Base
{
    public class RelayPathItem
    {
        private string _fromPort = "";
        private string _destPort = "";
        public string PathName
        {
            set
            {
                _fromPort = "";
                _destPort = "";
                List<string> pathsgmts = [.. value.Split('-')];
                _fromPort = pathsgmts[0];
                _destPort = string.Join("_to_", pathsgmts.GetRange(1, pathsgmts.Count - 1));
            }
            get { return $"{_fromPort}_to_{_destPort}"; }
        }

        public List<Component> ComponentInfos = [];

        public List<string> GetPorts()
        {
            return [_fromPort, _destPort];
        }
    }
}
