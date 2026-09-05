using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TestPlanLib.DataStruct
{
    public partial class Headers
    {
        [GeneratedRegex("[ ]{2,}")]
        private static partial Regex MyRegex();

        #region Field
        private readonly Dictionary<string, int> _headerDic = [];
        #endregion

        #region Member Function

        public void AddHeaderItem(string header, int colNum)
        {
            string key = CompareFormatToLower(header);
            _headerDic.TryAdd(key, colNum);
        }

        public int GetHeaderIndex(string header, bool mustExist = true)
        {
            if (_headerDic.TryGetValue(CompareFormatToLower(header), out int index))
            {
                return index;
            }
            return mustExist ? throw new Exception($" Can't find the header:{header} you want. ") : -1;
        }

        public int GetSimilarHeaderIndex(string header, bool mustExist = true)
        {
            string? key = _headerDic.Keys.FirstOrDefault(p => Regex.IsMatch(p, header, RegexOptions.IgnoreCase));
            if (key != null)
            {
                return _headerDic[key];
            }
            return mustExist ? throw new Exception($" Can't find the header:{header} you want. ") : 0;
        }

        public int GetHeaderIndexEvenNonexistence(string header)
        {
            return _headerDic.TryGetValue(CompareFormatToLower(header), out int index) ? index : -1;
        }

        private static string CompareFormatToLower(string src)
        {
            string target = MyRegex().Replace(src, " ");
            target = target.Replace(" ", "");
            target = target.ToLower();
            return target;
        }
        #endregion
    }
}
