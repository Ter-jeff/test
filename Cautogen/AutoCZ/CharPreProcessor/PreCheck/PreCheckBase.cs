using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class PreCheckBase : IPreCheck
    {
        /* field */
        protected readonly List<ErrorMessage> ErrorMessages = new List<ErrorMessage>();

        /* method */
        public virtual void Check(List<Characterization> charRows, string sheetName)
        {
            throw new NotImplementedException();
        }

        public void UpdateErrorMessages()
        {
            while (ErrorMessages.Count > 0)
            {
                ErrorManager.Add(ErrorMessages[0]);
                ErrorMessages.RemoveAt(0);
            }
        }

        public bool IsUseItem(Characterization item)
        {
            return Regex.IsMatch(item.Use, "^use$", RegexOptions.IgnoreCase);
        }
    }
}
