using System.Collections.Generic;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public interface IPreCheck
    {
        void Check(List<Characterization> charRows, string sheetName);

        void UpdateErrorMessages();
    }
}
