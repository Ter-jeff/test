using System.Collections.Generic;

namespace CommonLib.ErrorReport.Base
{
    public class ErrorInfo
    {
        public string? Pattern { get; set; }
        public List<string> Comments { get; set; } = [];
    }
}
