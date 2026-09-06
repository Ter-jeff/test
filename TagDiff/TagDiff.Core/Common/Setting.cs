using System;
using System.Collections.Generic;

namespace TagDiff.Core.Common
{
    [Serializable]
    internal class Setting
    {
        public string Header { get; set; } = string.Empty;
        public int ColNum { get; set; } = -1;
        public List<string> Content { get; set; } = [];
        public bool IsFromBase { get; set; } = false;
    }
}
