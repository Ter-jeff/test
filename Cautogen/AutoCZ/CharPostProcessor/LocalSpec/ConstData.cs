using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace Cautogen.AutoCZ.CharPostProcessor.LocalSpec
{
    [SuppressMessage("Design", "Ter402", Justification = "IGXL project folder structure constants")]
    public class ConstData
    {
        public const string TrunkFolder = @"IGLink";
        public const string Others = @"IGLink\Others";
        public const string Binfolder = @"IGLink\Common\BinTable";
        public const string CzFolder = @"IGLink\Module\Characterization";
        public const string MainFolder = @"IGLink\Module\Main";
        public const string ModuleFolder = @"IGLink\Module";
        public const string BasLibFolder = @"IGLink\Module\Library";
        public const string CommonFolder = @"IGLink\Common";
        public const string GlobalFolder = @"IGLink\Common\Global_Spec";
        public const string AcFolder = @"IGLink\Common\AC_Spec";
        public const string PatSetFolder = @"IGLink\Common\PatSetsAll";
        public const string TimeSetFolder = @"IGLink\Common\Timings";
        public const string XmlFolder = @"IGLink\xml_Files";
        public static List<string> KeptModuleList = new List<string> { "DC_Conti", "JTAG", "eFuse", "IDS", "Library", "Main", "SPIROM" };

        public const string InterPostPrePoint = "freerunclk_set_XY";
        public const string InterPostPrePointArg = "Y,XI0_Diff_Port,";
        public const string InterPostDefaultPort = "XI0_Diff_Port";
        public const string InterPostPostPoint = "freerunclk_stop";
        public const string InterPostPostPointArg = "Clock_Port";
        public const string InterPostPostStep = "PrintShmooInfo";
        public const string InterPostPostStepCSharp = "CoreTestLibrary.Char.FunctionalTestCharMain.PrintShmooInfoMain";
        public const string InterPostPostStepArg = "CorePower";
        //public const string InterPostPrePointArg = "Y,XI0_Diff_Port,XI0_Diff_Freq_VAR";

        public const string SetErrorOpCode = "set-error-bin";
        public const string FlagInitOpCode = "assign-site-var";
        public const string FlagInit = "Function_Result 1";
        public const string FlagCondOpCode = "site-var=";
        public const string FlagCond = "Function_Result 1";
        public const string FailFlagSet = "Function_Result=0";
        public const string FailFlagDummy = "F_Dummy";

        public const string Htol = "Htol";
        public const string Ttr = "Ttr";
        public static string DefaultSigsrcValue = "sgmt_default=0";

        public const string EnableCzMode = "Enable_Charz_mode";
        public const string EnableCzModeCShrap = "Enable_Charz_mode_CS";
        public const string DisableCzMode = "Disable_Charz_mode";
        public const string TpModeOnModule = "TPmode_Char_on";
        public const string TpModeOnModuleCSharp = "IgxlWrapper.CoreTestLibrary.Char.FunctionalTestCharMain.TPmodeCharOn";
        public const string TpModeOffModule = "TPmode_Char_off";
        public const string TpModeOffModuleCSharp = "IgxlWrapper.CoreTestLibrary.Char.FunctionalTestCharMain.TPmodeCharOff";

        public const string DigSrc = "DigSrc";
    }
}
