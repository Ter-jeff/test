using System;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace TestPlanLib.BinNumberLegacy
{
    public partial class SoftBinRangeData
    {
        [GeneratedRegex(@"\s*")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"\s+")]
        private static partial Regex MyRegex1();

        private int _softBinStart = -1;
        private int _softBinEnd = -1;
        private int _softBinCurrent = -1;
        private bool _isExceed;

        public string Description { get; set; } = string.Empty;
        public string Block { get; set; } = string.Empty;
        public string Condition { get; set; } = string.Empty;
        public string Start { get; set; } = string.Empty;
        public string End { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string Bin { get; set; } = string.Empty;

        public string HardBin { get; set; } = string.Empty;
        public string HardHvBin { get; set; } = string.Empty;
        public string HardLvBin { get; set; } = string.Empty;
        public string HardNvBin { get; set; } = string.Empty;
        public string HardHlvBin { get; set; } = string.Empty;

        public void ResetNumber()
        {
            ConvertNumber();
            _isExceed = false;
        }

        private void ConvertNumber()
        {
            bool convert1 = int.TryParse(Start, out _softBinStart);
            bool convert2 = int.TryParse(End, out _softBinEnd);

            if (!convert1 || !convert2)
            {
                throw new Exception($"The Bin Number range: {Description} is not a numebr, start:{Start} ,end:{End}");
            }

            if (_softBinStart > _softBinEnd)
            {
                throw new Exception(
                    $"Error for Bin Number range of:{Description} ,start number:{Start} can not be larger than end number:{End}!");
            }
            _softBinCurrent = _softBinStart;
        }

        public int GetSoftBinNumber()
        {
            if (_softBinStart == -1)
            {
                ConvertNumber();
                return _softBinCurrent;
            }

            if (_softBinCurrent < _softBinEnd)
            {
                _isExceed = false;
                return ++_softBinCurrent;
            }
            else
            {
                _isExceed = true;
                return _softBinEnd;
            }
        }

        public int GetSoftBinStart()
        {
            if (_softBinStart == -1)
            {
                ConvertNumber();
            }
            return _softBinStart;
        }

        public int GetSoftBinEnd()
        {
            if (_softBinStart == -1)
            {
                ConvertNumber();
            }
            return _softBinEnd;
        }

        public bool CheckExceed()
        {
            return _isExceed;
        }

        public string GetStatus()
        {
            if (State.EqualsIgnoreCase("B"))
            {
                return "Fail";
            }
            if (State.EqualsIgnoreCase("G"))
            {
                return "Pass";
            }
            return State.EqualsIgnoreCase("O") ? "Fail-Stop" : "Fail";
        }

        public bool Match(string inputCondition)
        {
            string[] conditions = MyRegex().Replace(Condition, "").Split(',');
            inputCondition = MyRegex1().Replace(inputCondition, "");
            foreach (string s in conditions)
            {
                if (s.EqualsIgnoreCase(inputCondition))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
