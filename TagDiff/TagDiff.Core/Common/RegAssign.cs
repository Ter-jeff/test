using System;
using System.Collections.Generic;

namespace TagDiff.Core.Common
{
    [Serializable]
    internal class RegAssign
    {
        public string SheetName { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string SettingName { get; set; } = string.Empty;
        public List<Setting> RegSettingList { get; set; } = [];
        public string Status { get; set; } = string.Empty;
    }
}
