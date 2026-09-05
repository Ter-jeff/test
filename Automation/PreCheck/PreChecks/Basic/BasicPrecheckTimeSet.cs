using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Singleton;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Static;

using TestPlanLib.Basic;

namespace Automation.PreCheck.PreChecks.Basic
{
    public class BasicPrecheckTimeSet : BasicPreCheckBase
    {
        private static readonly Regex _regex1 = new Regex("^=");
        private static readonly Regex _regex2 = new Regex(@"Time Set[\t]Period");

        private readonly string _patternListFileName;
        private readonly string _timeSetPath;

        public BasicPrecheckTimeSet(string patternList, string timeSetPath) : base(null, null)
        {
            _patternListFileName = patternList;
            _timeSetPath = timeSetPath;
        }

        protected internal override bool CheckExist()
        {
            if (!File.Exists(_patternListFileName))
            {
                return true;
            }

            var patternList = AcTSetCategoryMapSingleton.Instance().PatternList.Select(x => x.Value).ToList();
            List<string> timeSetList = GetTSetFileList(patternList);
            bool result = CheckTSetFile(timeSetList, _timeSetPath);
            return result;
        }

        protected internal override bool CheckHeaders()
        {
            return true;
        }

        protected internal override bool CheckFormat()
        {
            return true;
        }

        protected internal override bool CheckBusiness()
        {
            return true;
        }

        private List<string> GetTSetFileList(List<PatternData> patList)
        {
            var tSetFileList = new List<string>();
            foreach (PatternData pattern in patList)
            {
                string timeSetVersion = pattern.TimeSetVersion;
                if (!tSetFileList.Contains(timeSetVersion) && timeSetVersion.ToUpper() != "NA")
                {
                    tSetFileList.Add(timeSetVersion);
                }
            }
            return tSetFileList;
        }

        private bool CheckTSetFile(List<string> tSetFileList, string tSetPath)
        {
            string currentFolder = Directory.GetCurrentDirectory();
            string tempFolder = Path.Combine(currentFolder, "Temp");
            if (Directory.Exists(tempFolder))
            {
                Directory.Delete(tempFolder, true);
            }

            DirectoryInfo dir = Directory.CreateDirectory(tempFolder);
            if (Directory.Exists(tSetPath))
            {
                foreach (string fileName in tSetFileList)
                {
                    string completeFileName;
                    if (fileName.IndexOf(".txt", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        completeFileName = fileName;
                    }
                    else
                    {
                        completeFileName = fileName + ".txt";
                    }
                    string fileFullPath = Path.Combine(tSetPath, completeFileName);
                    if (File.Exists(fileFullPath))
                    {
                        string tarFile = Path.Combine(tempFolder, completeFileName);
                        File.Copy(fileFullPath, tarFile, true);
                        _ = CheckOneSheet(tarFile);
                    }

                }
            }
            dir.Delete(true);
            return true;
        }

        private bool CheckOneSheet(string fullFileName)
        {
            string sheetName = Path.GetFileName(fullFileName);
            string[] lines = File.ReadAllLines(fullFileName);
            int startRowNum = 4;
            for (int i = startRowNum; i < lines.Length; i++)
            {
                if (_regex2.IsMatch(lines[i]))
                {
                    startRowNum = i + 1;
                    break;
                }
            }
            for (int i = startRowNum; i < lines.Length; i++)
            {
                string[] arr = lines[i].Split('\t');
                if (arr.Length <= 4)
                {
                    break;
                }

                string timeSet = arr[1];
                string clockPeriod = arr[2];
                string pin = arr[3];
                if (_regex1.IsMatch(clockPeriod))
                {
                    continue;
                }

                if (string.IsNullOrEmpty(clockPeriod))
                {
                    string message = $"Line:{i + 1} format is empty line";
                    ErrorReportManager.AddError(BasicErrorType.W_FormatError_01, sheetName, 0, 0, $"Line:{i + 1} format is empty line");
                    Response.Report("TimeSet Warning in " + sheetName + " " + message, EnumMessageLevel.Warning, percentage: -1);
                    continue;
                }

                if (!double.TryParse(clockPeriod, out _))
                {
                    string message = string.Format("Row:{3} The period [{0}] is illegal for tSet {1} on Pin {2}", clockPeriod, timeSet, pin, i + 1);
                    ErrorReportManager.AddError(BasicErrorType.E_FormatError_19, sheetName, 0, 0, string.Format("Row:{3} The period [{0}] is illegal for tSet {1} on Pin {2}", clockPeriod, timeSet, pin, i + 1), new string[] { clockPeriod, timeSet, pin, (i + 1).ToString() });
                    Response.Report("TimeSet Error in " + sheetName + " " + message, EnumMessageLevel.Error, percentage: -1);
                }
            }

            return true;
        }
    }
}
