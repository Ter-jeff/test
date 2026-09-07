using System.Collections.Generic;

using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager.PatternError;

namespace Cautogen.AutoCZ.CharPreProcessor.Utility
{
    public class UtilityMain
    {
        public static UtilityData UtilityData;
        public static UtilityFunction UtilityFunction;

        public static void Reset()
        {
            UtilityData = new UtilityData();
            UtilityFunction = new UtilityFunction();
            ErrorManager.ErrorListDict = new Dictionary<ErrorType, List<ErrorMessage>>();
            PatternErrorCache.Reset();
        }
    }
}
