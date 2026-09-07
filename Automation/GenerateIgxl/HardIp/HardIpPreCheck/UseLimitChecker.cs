using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class OppositeLimitChecker : HardIpPrecheckBase
    {
        internal List<int> _oppositeLimitValue; //eg. low limit > hifh limit
        internal List<int> _wrongSyntax;        //eg. 0.5*VDDIO_PCIE+0.1 , 123.25MHz , 0x2080028 , 0b2080028
        internal List<int> _unitMismatch;
        internal List<int> _noUnit;             //warning
        internal List<int> _noLimit;            //warning
        internal List<int> _clearLimit;            //warning to user check
        private int _errorCnt;
        internal string _unit = "";
        private static readonly Regex _regex = new Regex("MeasV", RegexOptions.IgnoreCase);
        private static readonly Regex _regex2 = new Regex("MeasI", RegexOptions.IgnoreCase);
        private static readonly Regex _regex3 = new Regex("MeasF", RegexOptions.IgnoreCase);
        private static readonly Regex _regex4 = new Regex("MeasR", RegexOptions.IgnoreCase);
        private static readonly Regex _regex5 = new Regex(@"^(\d|\.|-)+(\w)*$");
        private static readonly Regex _regex6 = new Regex(@"^(\d|\.|-)+(?<unit>(\w)*)$");
        private static readonly Regex _regex7 = new Regex(@"^(\d|\.|-)+(\w|-|%)*$");
        private static readonly Regex _regex8 = new Regex("MeasV|MeasI|MeasF|MeasR", RegexOptions.IgnoreCase);

        public OppositeLimitChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        internal string GetMeasurePinUnit(MeasPin pin)
        {
            string unit = "";
            if (_regex.IsMatch(pin.MeasType))
            {
                unit = "V";
            }
            else if (_regex2.IsMatch(pin.MeasType))
            {
                unit = "A";
            }
            else if (_regex3.IsMatch(pin.MeasType))
            {
                unit = "Hz";
            }
            else if (_regex4.IsMatch(pin.MeasType))
            {
                unit = "ohms?";
            }

            return unit;
        }

        internal void CheckOneUseLimit(MeasLimit limit, MeasPin pin)
        {
            string low = DataConvertor.ConvertUnits(limit.LoLimit);
            string high = DataConvertor.ConvertUnits(limit.HiLimit);
            bool lowMatch = double.TryParse(low, out double lowValue);
            bool highMatch = double.TryParse(high, out double highValue);

            if (lowMatch && highMatch)
            {
                #region Check Opposite Limit value
                if (highValue < lowValue)
                {
                    if (!_oppositeLimitValue.Contains(limit.LoHeaderIndex))
                    {
                        _oppositeLimitValue.Add(limit.LoHeaderIndex);
                    }

                    if (!_oppositeLimitValue.Contains(limit.HiHeaderIndex))
                    {
                        _oppositeLimitValue.Add(limit.HiHeaderIndex);
                    }
                }
                #endregion
            }

            CheckOneSideUseLimit(limit, "Low", pin.MeasType);
            CheckOneSideUseLimit(limit, "High", pin.MeasType);

        }

        internal void CheckOneSideUseLimit(MeasLimit limit, string hiLoType, string measType)
        {
            string limitstr;
            int headerIndex;
            if (hiLoType == "Low")
            {
                limitstr = limit.LoLimit;
                headerIndex = limit.LoHeaderIndex;
            }
            else
            {
                limitstr = limit.HiLimit;
                headerIndex = limit.HiHeaderIndex;
            }

            if (limitstr.StartsWith("Binning", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string unitStr = "";
            if (_regex5.IsMatch(limitstr))
            {
                unitStr = _regex6.Match(limitstr).Groups["unit"].ToString();
            }

            if (limitstr == "")
            {
                #region no limit
                if (!_noLimit.Contains(headerIndex))
                {
                    _noLimit.Add(headerIndex);
                }

                #endregion
            }
            else if (limitstr.ContainsIgnoreCase("VARIABLE"))
            {
                if (hiLoType == "Low")
                {
                    limit.LoLimit = "";
                }
                else
                {
                    limit.HiLimit = "";
                }

                if (!_clearLimit.Contains(headerIndex))
                {
                    _clearLimit.Add(headerIndex);
                }
            }
            else if (!limitstr.Contains("VDD", StringComparison.OrdinalIgnoreCase) &&
                     !_regex7.IsMatch(limitstr) &&
                     !limitstr.StartsWith("0X", StringComparison.OrdinalIgnoreCase) &&
                     !limitstr.StartsWith("0B", StringComparison.OrdinalIgnoreCase) &&
                     !limitstr.StartsWith("0D", StringComparison.OrdinalIgnoreCase))
            {
                #region Check Un-supported limit value

                if (!_wrongSyntax.Contains(headerIndex))
                {
                    _wrongSyntax.Add(headerIndex);
                }

                #endregion
            }
            else
            {
                #region Check wrong mismatch limit units

                if (unitStr != "" && !Regex.IsMatch(unitStr, _unit + "$", RegexOptions.IgnoreCase))
                {
                    if (!_unitMismatch.Contains(headerIndex))
                    {
                        _unitMismatch.Add(headerIndex);
                    }
                }
                else
                {
                    if (_regex8.IsMatch(measType))
                    {
                        if (unitStr == "" && limitstr != "" && !limitstr.Contains("VDD", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!_noUnit.Contains(headerIndex))
                            {
                                _noUnit.Add(headerIndex);
                            }
                        }
                    }
                }

                #endregion
            }
        }

        public override void Check()
        {
            string sheetName = Pattern.SheetName;
            bool isTrimInstance = Pattern.MiscInfoDict.ContainsKey("trimCategory");
            foreach (MeasPin pin in Pattern.MeasPins)
            {
                _oppositeLimitValue = new List<int>();
                _wrongSyntax = new List<int>();
                _unitMismatch = new List<int>();
                _noUnit = new List<int>();
                _noLimit = new List<int>();
                _clearLimit = new List<int>();
                _unit = GetMeasurePinUnit(pin);

                foreach (MeasLimit limit in pin.MeasLimitsH)
                {
                    CheckOneUseLimit(limit, pin);
                }

                foreach (MeasLimit limit in pin.MeasLimitsL)
                {
                    CheckOneUseLimit(limit, pin);
                }

                foreach (MeasLimit limit in pin.MeasLimitsN)
                {
                    CheckOneUseLimit(limit, pin);
                }

                if (_oppositeLimitValue.Count > 0)
                {
                    foreach (int item in _oppositeLimitValue)
                    {
                        ErrorReportManager.AddError(HardIpErrorType.E_OppositeLimit_01, sheetName, pin.RowNum, item);
                    }
                }

                if (_wrongSyntax.Count > 0)
                {
                    //const string errorMessage = "Unrecognied limit value";
                    foreach (int item in _wrongSyntax)
                    {
                        ErrorReportManager.AddError(HardIpErrorType.E_WrongLimitValue_01, sheetName, pin.RowNum, item);
                    }
                }

                if (_unitMismatch.Count > 0)
                {
                    foreach (int item in _unitMismatch)
                    {
                        ErrorReportManager.AddError(HardIpErrorType.E_MismatchLimitUnit_01, sheetName, pin.RowNum, item);
                    }
                }

                if (_noUnit.Count > 0)
                {
                    if (_errorCnt < 500)
                    {
                        foreach (int item in _noUnit)
                        {
                            ErrorReportManager.AddError(HardIpErrorType.W_MissingLimitUnit_01, sheetName, pin.RowNum, item);
                            _errorCnt++;
                        }

                    }
                }

                if (_noLimit.Count > 0)
                {
                    if (_errorCnt < 500)
                    {
                        foreach (int item in _noLimit)
                        {
                            ErrorReportManager.AddError(HardIpErrorType.W_MissingLimit_01, sheetName, pin.RowNum, item);
                            _errorCnt++;
                        }
                    }
                }

                if (_clearLimit.Count > 0)
                {
                    foreach (int item in _clearLimit)
                    {
                        ErrorReportManager.AddError(HardIpErrorType.W_ClearLimit_01, sheetName, pin.RowNum, item);
                    }
                }

                if (isTrimInstance)
                {
                    if (pin.SequenceIndex != 0 && !pin.MiscInfo.ContainsIgnoreCase(HardIpConstData.IgnoreFlowLimit))
                    {
                        ErrorReportManager.AddError(HardIpErrorType.E_MissingPinName_08, Pattern.SheetName, pin.RowNum, HardIpSheet.PlanHeaderIdx["miscInfoIndex"], [pin.PinName]);
                    }
                }
            }
        }
    }
}
