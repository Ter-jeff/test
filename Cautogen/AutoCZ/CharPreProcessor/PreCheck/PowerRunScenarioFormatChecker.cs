using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class PowerRunScenarioFormatChecker : PreCheckBase
    {
        /* properity and field */
        private readonly List<Tuple<string, string>> _initPlPowerRunTuples = new List<Tuple<string, string>>();

        private readonly Dictionary<string, int> _initPlCounter = new Dictionary<string, int>();

        private Dictionary<string, int> _initPlTypeCounter;

        private static readonly Regex _validatePwrRun = new Regex(@"(^NV$)|(^SWEEP(:[+-]\d+mv)?$)|(^VRS$)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static List<string> ValidateInitPls
        {
            get
            {
                var initPls = new List<string> { "INIT", "PL" };
                for (int i = 1; i <= 10; i++)
                {
                    initPls.Add("INIT" + i.ToString(CultureInfo.CurrentCulture));
                }

                for (int i = 1; i <= 5; i++)
                {
                    initPls.Add("PL" + i.ToString(CultureInfo.CurrentCulture));
                }

                return initPls;
            }
        }

        /* method */
        public override void Check(List<Characterization> charRows, string sheetName)
        {
            if (UtilityMain.UtilityData.InputParam.CharPreCheckForNewTChar)
            {
                CheckForNewTChar(charRows, sheetName);
            }
            else
            {
                CheckForOldTChar(charRows, sheetName);
            }
        }

        private void CheckForNewTChar(List<Characterization> charRows, string sheetName)
        {
            foreach (Characterization row in charRows.Where(row => row.PowerRunScenario != ""))
            {
                // read powerRunScenario into list of tuple
                if (!_ReadPowerRunPass(row))
                {
                    continue;
                }

                // general format check
                _PowerRunScenarioFormatCheck(row);
            }
        }

        private void CheckForOldTChar(List<Characterization> charRows, string sheetName)
        {
            foreach (Characterization row in charRows.Where(row => row.PowerRunScenario != ""))
            {
                // only allow init_X_PL_X for char row having extra patterns
                if (row.ExtraInits.Count > 0 || row.ExtraPLs.Count > 0)
                {
                    if (!Regex.IsMatch(row.PowerRunScenario, "init_(NV|Sweep|VRS)_PL_(NV|Sweep|VRS)", RegexOptions.IgnoreCase))
                    {
                        _AddWarning(row, "PowerRunScenario should be init_(NV|Sweep|VRS)_PL_(NV|Sweep|VRS) due to exist extra inits or payloads");
                    }

                    continue;
                }

                // read powerRunScenario into list of tuple
                if (!_ReadPowerRunPass(row))
                {
                    continue;
                }

                // general format check
                _PowerRunScenarioFormatCheck(row);
            }
        }

        private bool _ReadPowerRunPass(Characterization row)
        {
            // reset previous result
            _initPlPowerRunTuples.Clear();
            _initPlCounter.Clear();
            _initPlTypeCounter = new Dictionary<string, int>
            {
                {"init", 0},
                {"inits", 0},
                {"pl", 0},
                {"pls", 0}
            };

            string[] strArray = row.PowerRunScenario.ToUpper().Split('_');  // upper case

            if (strArray.Length % 2 != 0)
            {
                _AddError(row, "The value is not a pairs");
                return false;
            }

            string initPl = "";

            for (int i = 0; i < strArray.Length; i++)
            {
                if (i % 2 == 0)
                {
                    initPl = strArray[i];
                }
                else
                {
                    // update _initPlPowerRunTupels
                    string pwrRun = strArray[i];
                    _initPlPowerRunTuples.Add(new Tuple<string, string>(initPl, pwrRun));

                    // update _initPlCounter
                    if (_initPlCounter.ContainsKey(initPl))
                    {
                        // each init and pl should only appear once
                        _AddError(row, $"{initPl} should only appear once");
                        //return false;
                    }
                    _initPlCounter[initPl] = 1;

                    // update _initPlTypeCounter
                    if (initPl == "INIT")
                    {
                        _initPlTypeCounter["init"] += 1;
                    }
                    else if (Regex.IsMatch(initPl, @"INIT[\d]{1,2}"))
                    {
                        _initPlTypeCounter["inits"] += 1;
                    }
                    else if (initPl == "PL")
                    {
                        _initPlTypeCounter["pl"] += 1;
                    }
                    else if (Regex.IsMatch(initPl, @"PL[\d]{1}"))
                    {
                        _initPlTypeCounter["pls"] += 1;
                    }
                }
            }
            return true;
        }

        public void _PowerRunScenarioFormatCheck(Characterization row, bool newTChar = true)
        {
            foreach (Tuple<string, string> initPlPwrRun in _initPlPowerRunTuples)
            {
                if (newTChar)
                {
                    if (initPlPwrRun.Item1.ToUpper() != "INIT" && initPlPwrRun.Item1.ToUpper() != "PL")
                    {
                        if (!row.PatternCellList.Select(x => x.Name).Any(x => x.Equals(initPlPwrRun.Item1, StringComparison.OrdinalIgnoreCase)))
                        {
                            _AddError(row, string.Format($"{initPlPwrRun.Item1} pattern is not defined, please check.", initPlPwrRun.Item1));
                        }
                    }
                }
                else
                {
                    if (!ValidateInitPls.Contains(initPlPwrRun.Item1))
                    {
                        _AddError(row, string.Format(" is not allowed", initPlPwrRun.Item1));
                    }
                }
                if (!_validatePwrRun.IsMatch(initPlPwrRun.Item2))
                {
                    _AddError(row, string.Format(" is not allowed", initPlPwrRun.Item2));
                }
            }

            // init and init[\d] does not appear together
            if (_initPlTypeCounter["init"] > 0 && _initPlTypeCounter["inits"] > 0)
            {
                _AddError(row, "All/Single INIT can't set at same time");
            }

            // pl and pl[\d] does not appear together
            if (_initPlTypeCounter["pl"] > 0 && _initPlTypeCounter["pls"] > 0)
            {
                _AddError(row, "All/Single PL can't set at same time");
            }

            // at leate one init and one pl
            if (_initPlTypeCounter["init"] + _initPlTypeCounter["inits"] == 0)
            {
                _AddError(row, "INIT hasn't been set");
            }
            if (_initPlTypeCounter["pl"] + _initPlTypeCounter["pls"] == 0)
            {
                _AddError(row, "PL hasn't been set");
            }
        }

        private void _AddError(Characterization charRow, string msg)
        {
            ErrorMessages.Add(new ErrorMessage
            {
                ErrorLevel = ErrorLevel.Error,
                ErrorType = ErrorType.WrongPowerRunScenario,
                SheetName = charRow.SheetName,
                RowNum = charRow.RowNum,
                Message = $"Wrong PowerRunScenario ({msg}): {charRow.PowerRunScenario}",
            });
            ErrorReportManager.AddError(CharErrorType.E_WrongPowerRunScenario_01, charRow.SheetName, charRow.RowNum, 0, [msg, charRow.PowerRunScenario]);
        }

        private void _AddWarning(Characterization charRow, string message)
        {
            ErrorMessages.Add(new ErrorMessage
            {
                ErrorLevel = ErrorLevel.Warning,
                ErrorType = ErrorType.WrongPowerRunScenario,
                SheetName = charRow.SheetName,
                RowNum = charRow.RowNum,
                Message = "Wrong PowerRunScenario : " + message,
            });
            ErrorReportManager.AddError(CharErrorType.W_WrongPowerRunScenario_02, charRow.SheetName, charRow.RowNum, 0, [message]);
        }
    }
}
