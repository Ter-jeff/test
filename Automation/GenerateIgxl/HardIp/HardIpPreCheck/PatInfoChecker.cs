using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Utility.HardIP;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class PatInfoChecker
    {
        private Regex _isMlpsubPat = new Regex(@"\w+_BI_\w+_MLPSUB", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public void CheckPatInfo(Dictionary<string, HardIpSheet> planDic)
        {
            foreach (string sheetName in planDic.Keys)
            {
                foreach (HardIpPattern patternItem in planDic[sheetName].Rows)
                {
                    if (Regex.IsMatch(patternItem.MiscInfo, HardIpConstData.IgnorePatInfo, RegexOptions.IgnoreCase))
                    {
                        continue;
                    }
                    IEnumerable<string> checkPatterns = patternItem.Pattern.PatternSetList.SelectMany(p => p);

                    //HardIpReference patInfo = SearchInfo.GetHardIpInfo(patternItem);
                    foreach (string pattern in checkPatterns)
                    {

                        HardIpInfo patInfo = HardIpService.GetHardIpInfo(pattern);

                        if (pattern.Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase) ||
                            HardIpConstData.RegInsInPatt.IsMatch(pattern) ||
                            HardIpConstData.RegOpcodeInPatt.IsMatch(pattern))
                        {
                            continue;
                        }

                        if (!patInfo.PatInfoExist)
                        {
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_MissingPatternInPatInfo_01,
                                patternItem.SheetName,
                                patternItem.RowNum,
                                0,
                                [pattern]
                            );
                        }

                        if (patInfo.SendBitStr != "" && patInfo.SendBitName != "")
                        {
                            string[] strArray =
                                patInfo.SendBitStr.Split('+')
                                    .Where(a => a.Contains("_"))
                                    .Select(a => a.Split('_')[0])
                                    .ToArray();
                            string[] nameArray = patInfo.SendBitName.Split('+');
                            var nameList = new List<string>();
                            if (strArray.Length == nameArray.Length)
                            {
                                for (int i = 0; i < strArray.Length; i++)
                                {
                                    string pair = nameArray[i] + "&" + strArray[i];
                                    if (!nameList.Contains(pair))
                                    {
                                        nameList.Add(pair);
                                    }
                                    else
                                    {
                                        ErrorReportManager.AddError(
                                            HardIpErrorType.E_WrongSendInformation_01,
                                            patternItem.SheetName,
                                            patternItem.RowNum,
                                            0,
                                            [nameArray[i]]
                                        );
                                    }
                                }
                            }

                            if (
                                !patInfo.DigSrcSignalName.Equals(patInfo.VmVector,
                                    StringComparison.CurrentCultureIgnoreCase))
                            {
                                ErrorReportManager.AddError(
                                    HardIpErrorType.E_WrongDigSrcSignalName_01,
                                    patternItem.SheetName,
                                    patternItem.RowNum,
                                    0,
                                    [patInfo.DigSrcSignalName, patInfo.VmVector]
                                );
                            }
                        }


                        if (!string.IsNullOrEmpty(patInfo.MeasSeqStr) && string.IsNullOrEmpty(patInfo.CallSubrs))
                        {
                            //Indicate that the pattern has no call subr but has meas seq
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_MissingCallSubroutine_01,
                                patternItem.SheetName,
                                patternItem.RowNum,
                                0,
                                [patInfo.Payload]
                            );
                        }
                        else if (string.IsNullOrEmpty(patInfo.MeasSeqStr) && !string.IsNullOrEmpty(patInfo.CallSubrs))
                        {
                            //Indicate that the pattern has call subr but no meas seq
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_MissingMeasureSequence_01,
                                patternItem.SheetName,
                                patternItem.RowNum,
                                0,
                                [patInfo.Payload]
                            );
                        }
                        else if (!string.IsNullOrEmpty(patInfo.MeasSeqStr) && !string.IsNullOrEmpty(patInfo.CallSubrs) &&
                                 patInfo.MeasSeqStr.Split(',').Length != patInfo.CallSubrsCnt)
                        {
                            //Indicate that the pattern call subr count and meas seq count is not match
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_MismatchSequenceAndCallSubrsCnt_01,
                                patternItem.SheetName,
                                patternItem.RowNum,
                                0,
                                [patInfo.Payload]
                            );
                        }
                    }
                }
            }
        }

        public void CheckPatInfoAll(List<HardIpInfo> patInfos)
        {
            foreach (HardIpInfo patInfo in patInfos)
            {
                if (!string.IsNullOrEmpty(patInfo.DigSrcSignalName) && string.IsNullOrEmpty(patInfo.SendBit))
                {
                    ErrorReportManager.AddError(PatInfoType.E_WrongDigSrcSignalName_01, "", 0, 0, $"Pattern : \"{patInfo.Payload}\" , Sen bit is Empty while pattern info has DigSrc Signal Name", new string[] { patInfo.Payload });
                }

                if (!string.IsNullOrEmpty(patInfo.VmVector) && !string.IsNullOrEmpty(patInfo.Subroutine) && !_isMlpsubPat.IsMatch(patInfo.Payload))
                {
                    if (!patInfo.Subroutine.Equals(patInfo.VmVector + "_srm_meas", StringComparison.OrdinalIgnoreCase))
                    {
                        ErrorReportManager.AddError(PatInfoType.E_MismatchVmandSubr_01, "", 0, 0, $"Please Check the pattern \"{patInfo.Payload}\" , the vm vector and call subrs is not match.", new string[] { patInfo.Payload });
                    }
                }

                if (!string.IsNullOrEmpty(patInfo.Payload) && !string.IsNullOrEmpty(patInfo.VmVector) && !string.IsNullOrEmpty(patInfo.Version))
                {
                    if (!patInfo.VmVector.Equals(patInfo.Payload + "_" + patInfo.Version,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        ErrorReportManager.AddError(PatInfoType.E_MismatchGenericName_01, "", 0, 0, $"Please Check the pattern \"{patInfo.Payload}\" , the vm vector and generic name is not match.", new string[] { patInfo.Payload });
                    }
                }
                if (!string.IsNullOrEmpty(patInfo.Subroutine) && string.IsNullOrEmpty(patInfo.VmVector))
                {
                    ErrorReportManager.AddError(PatInfoType.E_MissingVmvector_01, "", 0, 0, $"Please Check the pattern \"{patInfo.Payload}\" , the vm vector is missing", new string[] { patInfo.Payload });
                }
            }
        }
    }
}
