using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class RetentionChecker : PreCheckBase
    {
        private readonly Regex _regexRetentionForNewTChar = new Regex(@"^(?:INIT|PL)\d+:\d+(?:\.\d+)?:(?:[+-]?\d+mV)?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private int _retentionColIdx = -1;
        /* Check retention value and "EXT" */
        public override void Check(List<Characterization> charList, string sheetName)
        {
            ResolveRetentionColIdx(charList);
            foreach (Characterization item in charList)
            {
                CheckRetentionItem(item, sheetName);
            }
        }

        private void ResolveRetentionColIdx(List<Characterization> charList)
        {
            Characterization firstItem = charList.FirstOrDefault();
            if (firstItem == null)
            {
                return;
            }
            if (firstItem.ColumnDic.TryGetValue("retentiontime", out int value))
            {
                _retentionColIdx = value;
            }
            else if (firstItem.ColumnDic.TryGetValue("retentiontime:guardband", out int value1))
            {
                _retentionColIdx = value1;
            }
        }

        private void CheckRetentionItem(Characterization item, string sheetName)
        {
            bool extRet = item.Category != null &&
                (item.Category.Contains("EXTRET") ||
                item.Category.Contains("ERT") ||
                item.Category.Contains("NAPRET") ||
                item.Category.Contains("NRT") ||
                item.Category.Contains("SRT"));
            bool intRet = item.Category != null && item.Category.Contains("INTRET");
            bool disturb = item.Category != null && item.Category.Contains("DISTURB");
            bool hasNoRetention = string.IsNullOrEmpty(item.Retention);
            bool hasNoPowerRunScenario = string.IsNullOrEmpty(item.PowerRunScenario);

            EmitCategoryRetentionWarnings(item, sheetName, extRet, intRet, disturb, hasNoRetention, hasNoPowerRunScenario);

            if (UtilityMain.UtilityData.InputParam.CharPreCheckForNewTChar)
            {
                CheckForNewTChar(item, sheetName);
            }
            else
            {
                CheckForOldTChar(item, sheetName);
            }

            CheckWriteRead(item, sheetName);
        }

        private static void EmitCategoryRetentionWarnings(Characterization item, string sheetName, bool extRet, bool intRet, bool disturb, bool hasNoRetention, bool hasNoPowerRunScenario)
        {
            // for category "EXTRET", must specifiy retention and power-run-scenario
            if (extRet && (hasNoRetention || hasNoPowerRunScenario))
            {
                const string outString = "Missing Retention time or PowerRunScenario for Category 'EXTRET' or 'NAPRET'";
                ErrorManager.AddWarning(ErrorType.WrongRetention, sheetName, item.RowNum,
                    item.ColNum(new List<string> { "retentiontime", "powerrunscenario" }), item.Use, outString);
                foreach (int col in item.ColNum(new List<string> { "retentiontime", "powerrunscenario" }))
                {
                    ErrorReportManager.AddError(CharErrorType.W_WrongRetention_02, sheetName, item.RowNum, col, []);
                }
            }

            // for category "INTRET", must specify power-run-scenario
            if (intRet && hasNoPowerRunScenario)
            {
                const string outString = "Missing PowerRunScenario for Category 'INTRET' ";
                ErrorManager.AddWarning(ErrorType.WrongRetention, sheetName, item.RowNum,
                    item.ColNum(new List<string> { "category", "powerrunscenario" }), item.Use, outString);
                foreach (int col in item.ColNum(new List<string> { "category", "powerrunscenario" }))
                {
                    ErrorReportManager.AddError(CharErrorType.W_WrongRetention_03, sheetName, item.RowNum, col, []);
                }
            }

            // for category "DISTURB", must specify power-run-scenario
            if (disturb && hasNoPowerRunScenario)
            {
                const string outString = "Missing PowerRunScenario for Category  'DISTURB' ";
                ErrorManager.AddWarning(ErrorType.WrongRetention, sheetName, item.RowNum,
                    item.ColNum(new List<string> { "category", "powerrunscenario" }), item.Use, outString);
                foreach (int col in item.ColNum(new List<string> { "category", "powerrunscenario" }))
                {
                    ErrorReportManager.AddError(CharErrorType.W_WrongRetention_04, sheetName, item.RowNum, col, []);
                }
            }

            // if specified retention, the category must be EXTRET
            if (!hasNoRetention && !extRet)
            {
                const string outString = "Wait time exists but Category is not 'EXTRET' ";
                ErrorManager.AddWarning(ErrorType.WrongRetention, sheetName, item.RowNum,
                    item.ColNum(new List<string> { "category", "retentiontime" }), item.Use, outString);
                foreach (int col in item.ColNum(new List<string> { "category", "retentiontime" }))
                {
                    ErrorReportManager.AddError(CharErrorType.W_WrongRetention_05, sheetName, item.RowNum, col, []);
                }
            }
        }
        private void CheckForNewTChar(Characterization item, string sheetName)
        {
            CheckRetentionFormatForNewTChar(item, sheetName);
        }

        private void CheckWriteRead(Characterization item, string sheetName)
        {
            //var patterns = item.PatternCellList.SelectMany(x => x.PatternDefine.Split(','));
            var colPatternPairList = new List<KeyValuePair<int, string>>();
            item.PatternCellList
                .ForEach(x => x.PatternDefine.Split(',').ToList()
                    .ForEach(y => colPatternPairList.Add(new KeyValuePair<int, string>(x.ColIndex, y))));
            var patternLocal = item.PatternCellList.ToDictionary(x => x.ColIndex, x => x.PatternDefine.Split(',').ToList()).ToList();
            var writePatternRegex = new Regex(@"_(\d+)W(\w)_", RegexOptions.IgnoreCase);
            var readPatternRegex = new Regex(@"_(\d+)R(\w)_", RegexOptions.IgnoreCase);

            var writeReadColPatternPairDict = new Dictionary<string, List<KeyValuePair<int, string>>>();

            foreach (KeyValuePair<int, string> pattern in colPatternPairList)
            {
                Match writeMatch = writePatternRegex.Match(pattern.Value);
                if (writeMatch.Success)
                {
                    string key = $"{writeMatch.Groups[1].Value}_{writeMatch.Groups[2].Value}";

                    if (writeReadColPatternPairDict.ContainsKey(key))
                    {
                        writeReadColPatternPairDict[key].Add(pattern);
                    }
                    else
                    {
                        writeReadColPatternPairDict[key] = new List<KeyValuePair<int, string>>() { pattern };
                    }
                }

                Match readMatch = readPatternRegex.Match(pattern.Value);
                if (readMatch.Success)
                {
                    string key = $"{readMatch.Groups[1].Value}_{readMatch.Groups[2].Value}";

                    if (writeReadColPatternPairDict.ContainsKey(key))
                    {
                        if (writeReadColPatternPairDict[key].Count > 0)
                        {
                            writeReadColPatternPairDict[key].RemoveAt(writeReadColPatternPairDict[key].Count - 1);
                        }
                        else
                        {
                            ErrorMessages.Add(new ErrorMessage
                            {
                                ErrorLevel = ErrorLevel.Error,
                                ErrorType = ErrorType.WrongRetention,
                                SheetName = sheetName,
                                RowNum = item.RowNum,
                                Message = $"No matching write pattern for read pattern: {pattern.Value}.",
                                ColList = new List<int> { pattern.Key },
                            });
                            ErrorReportManager.AddError(CharErrorType.E_WrongRetention_06, sheetName, item.RowNum, pattern.Key, [pattern.Value]);
                        }
                    }
                    else
                    {
                        ErrorMessages.Add(new ErrorMessage
                        {
                            ErrorLevel = ErrorLevel.Error,
                            ErrorType = ErrorType.WrongRetention,
                            SheetName = sheetName,
                            RowNum = item.RowNum,
                            Message = $"No matching write pattern for read pattern: {pattern.Value}.",
                            ColList = new List<int> { pattern.Key },
                        });
                        ErrorReportManager.AddError(CharErrorType.E_WrongRetention_06, sheetName, item.RowNum, pattern.Key, [pattern.Value]);
                    }
                }
            }

            foreach (KeyValuePair<string, List<KeyValuePair<int, string>>> writeRead in writeReadColPatternPairDict)
            {
                foreach (KeyValuePair<int, string> colPatternPair in writeRead.Value)
                {
                    ErrorMessages.Add(new ErrorMessage
                    {
                        ErrorLevel = ErrorLevel.Error,
                        ErrorType = ErrorType.WrongRetention,
                        SheetName = sheetName,
                        RowNum = item.RowNum,
                        Message = $"Unmatched write pattern {colPatternPair.Value} no found with read pattern.",
                        ColList = new List<int> { colPatternPair.Key },
                    });
                    ErrorReportManager.AddError(CharErrorType.E_WrongRetention_07, sheetName, item.RowNum, colPatternPair.Key, [colPatternPair.Value]);
                }
            }
        }

        private void CheckForOldTChar(Characterization item, string sheetName)
        {
            // retention segment must be 1, 5 or 15
            int commaount = Regex.Matches(item.Retention, ",").Count;
            if (commaount != 0 && commaount != 4 && commaount != 14)
            {
                const string outString = "Retention time format is wrong, the commaon count must be  in [0, 4, 14]";
                ErrorManager.AddWarning(ErrorType.WrongRetention, sheetName, item.RowNum, item.ColNum("retentiontime"), item.Use, outString);
                foreach (int col in item.ColNum("retentiontime"))
                {
                    ErrorReportManager.AddError(CharErrorType.W_WrongRetention_01, sheetName, item.RowNum, col, []);
                }
            }

            if (_UpdateWaitTimeWithPayloadWR(item) != item.Retention)
            {
                string outString =
                    $"Retention time \"{item.Retention}\" is not same to tool expect value \"{_UpdateWaitTimeWithPayloadWR(item)}\"";
                ErrorManager.AddWarning(ErrorType.WrongRetention, sheetName, item.RowNum, item.ColNum("retentiontime"), item.Use, outString);
                foreach (int col in item.ColNum("retentiontime"))
                {
                    ErrorReportManager.AddError(CharErrorType.W_WrongRetention_08, sheetName, item.RowNum, col, [item.Retention, _UpdateWaitTimeWithPayloadWR(item)]);
                }
            }

            if ((item.ExtraInits.Count > 0 || item.ExtraPLs.Count > 0) && item.Retention != "")
            {
                const string outString = "Retention time due to exist extra inits(10) or payloads(5)";
                ErrorManager.AddWarning(ErrorType.WrongRetention, sheetName, item.RowNum,
                    item.ColNum(new List<string> { "init10", "payload5", "retentiontime" }), item.Use, outString);
                foreach (int col in item.ColNum(new List<string> { "init10", "payload5", "retentiontime" }))
                {
                    ErrorReportManager.AddError(CharErrorType.W_WrongRetention_09, sheetName, item.RowNum, col, []);
                }
            }
        }
        /*
         * Purpose : complete read feature after writing data 
         * 1. when write data, trigger process until complete read to reset
         * 2. ETR pattern would reduce voltage => need more time to achieve(priority 1)
         * 3. EMA pattern need more time to achieve(priority lower than ETR if exist in same time)
         * 4. if common scenario, put wait time before read
         * 
         */
        private static string _UpdateWaitTimeWithPayloadWR(Characterization item)
        {
            string regWrite = @"^\dW\w*";
            string regRead = @"^\dR\w*";
            string regEtr = "ETR";
            string regEma = "EMA";
            string regInit = "^IN";
            var indexList = new List<int>();
            var payloads = item.PayloadPatterns.Values.ToList();
            int writeIndex = -1;

            int etrIndex = -1;
            int emaIndex = -1;
            int firstInitIndex = -1;

            for (int i = 0; i < payloads.Count; i++)
            {
                string payload = payloads[i];
                foreach (string subpattern in payload.Split(','))
                {
                    List<string> payloadSgmt = subpattern.Split('_').ToList();
                    if (payloadSgmt.Exists(p => Regex.IsMatch(p, regWrite, RegexOptions.IgnoreCase)))
                    {
                        writeIndex = i;
                    }
                    if (payloadSgmt.Exists(p => Regex.IsMatch(p, regEtr, RegexOptions.IgnoreCase)))
                    {
                        etrIndex = i;
                        indexList.Add(10 + i);
                    }
                    else if (emaIndex == -1 && payloadSgmt.Exists(p => Regex.IsMatch(p, regEma, RegexOptions.IgnoreCase)))
                    {
                        emaIndex = i;
                    }
                    else if (firstInitIndex == -1 &&
                             payloadSgmt.Exists(p => Regex.IsMatch(p, regInit, RegexOptions.IgnoreCase)))
                    {
                        firstInitIndex = i;
                    }

                    if (writeIndex != -1 && payloadSgmt.Exists(p => Regex.IsMatch(p, regRead, RegexOptions.IgnoreCase)))
                    {
                        if (etrIndex != -1)
                        {
                            //indexList.Add(10 + ETR_index);
                        }
                        else if (emaIndex != -1)
                        {
                            indexList.Add(10 + emaIndex);
                        }
                        else if (firstInitIndex != -1)
                        {
                            indexList.Add(10 + firstInitIndex);
                        }
                        else
                        {
                            indexList.Add(10 + i - 1);
                        }

                        writeIndex = -1;
                        etrIndex = -1;
                        emaIndex = -1;
                        firstInitIndex = -1;
                    }
                }
            }

            if (indexList.Count > 0)
            {
                var result = Enumerable.Repeat("", 15).ToList();
                foreach (int indexitem in indexList)
                {
                    result[indexitem] = "0.02";
                }

                return string.Join(",", result);
            }

            //if (index != -1)
            //{
            //    var result = Enumerable.Repeat("",15).ToList();
            //    result[10 + index] = "0.02";
            //    return string.Join(",", result);
            //}

            return item.Retention;
        }
        private void CheckRetentionFormatForNewTChar(Characterization item, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(item.Retention))
            {
                return;
            }

            List<string> segments = item.Retention.Split(',').ToList();
            foreach (string segment in segments)
            {
                if (!_regexRetentionForNewTChar.IsMatch(segment))
                {
                    ErrorMessages.Add(new ErrorMessage
                    {
                        ErrorLevel = ErrorLevel.Error,
                        ErrorType = ErrorType.IllegalForNewTChar,
                        SheetName = sheetName,
                        RowNum = item.RowNum,
                        ColList = new List<int> { _retentionColIdx },
                        Message = "Please follow {INIT/PL}{Index}:{RetentionTime}:{GuardBand}mV, e.g INIT2:0.02:+25mV .",
                        CommentsList = new List<string> { segment },
                    });
                    ErrorReportManager.AddError(CharErrorType.E_IllegalForNewTChar_02, sheetName, item.RowNum, _retentionColIdx, [],
                        new ErrorInfo() { Comments = new List<string> { segment } });
                    continue;
                }
                string patternIndex = segment.Split(':')[0];


                if (patternIndex.StartsWith("INIT"))
                {
                    patternIndex = patternIndex.Replace("INIT", "init");
                }
                else if (patternIndex.StartsWith("PL"))
                {
                    patternIndex = patternIndex.Replace("PL", "payload");
                }

                PatternCell patternCell = item.PatternCellList.FirstOrDefault(x => x.Header.Equals(patternIndex, System.StringComparison.OrdinalIgnoreCase));
                if (patternCell == null)
                {
                    ErrorMessages.Add(new ErrorMessage
                    {
                        ErrorLevel = ErrorLevel.Error,
                        ErrorType = ErrorType.IllegalForNewTChar,
                        SheetName = sheetName,
                        RowNum = item.RowNum,
                        ColList = new List<int> { _retentionColIdx },
                        Message = string.Format($"{patternIndex} pattern didn't be defined, please check."),
                        CommentsList = new List<string> { segment },
                    });
                    ErrorReportManager.AddError(CharErrorType.E_IllegalForNewTChar_03, sheetName, item.RowNum, _retentionColIdx, []);
                }
            }
        }
    }
}
