using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions
{
    public class GeneralFunc
    {
        public static void WriteMessage(string text)
        {
            LocalSpecs.MessageWriter.WriteLine(text);
        }
    }
}
