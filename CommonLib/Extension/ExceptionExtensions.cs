using System;
using System.Collections.Generic;
using System.Linq;

namespace CommonLib.Extension
{
    public static class ExceptionExtensions
    {
        public static string PrintMessages(this Exception ex)
        {
            var lines = new List<string>();
            while (ex != null)
            {
                lines.Add(ex.Message);
                lines.Add("");
                lines.Add(ex.StackTrace);
                ex = ex.InnerException;
            }
            return string.Join(Environment.NewLine, lines.SelectMany(x => x.ToLines()).Distinct().ToList());
        }
    }
}
