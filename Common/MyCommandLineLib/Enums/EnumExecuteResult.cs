using System.ComponentModel;

namespace MyCommandLineLib.Enums
{
    public enum EnumExecuteResult
    {
        [Description("Pass")]
        Pass = 0,
        [Description("Fail")]
        Fail,
        [Description("Exception")]
        Exception,
        [Description("Done")]
        Done,
        [Description("Skip")]
        Skip
    }
}
