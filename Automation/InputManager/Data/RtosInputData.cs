using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Reader.ConfigFile.RtosCategory;

using ScghLib.Reader;

namespace Automation.InputManager.Data
{
    public class RtosInputData : InputDataBase
    {
        public RtosProdCharSheet ProdSheet { set; get; }
        public RtosConfig NamingRule { set; get; }
        public RtosCategoryConfigReader RtosConfig { set; get; }
    }
}
