using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenAc;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;
using CommonLib.Results;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;
using IgxlLib.Utility;

using LogLib.Utility;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.PatternListCsvFile;
using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;

public class TimeSetGenerator
{
    private const string ConMissingFile = "Missing File:";
    private const string Na = "NA";

    public Dictionary<string, List<int>> TimeSetWithWrongForamtRows { get; } = [];
    public List<string> TimeSetIncorrectFormat { get; } = [];
    private static readonly Regex _regex = new Regex(@"^.*[/](?<patternName>(.*))(\.atp\.gz)$", RegexOptions.IgnoreCase);
    private static readonly Regex _regex2 = new Regex(@"[_]\w+");
    private static readonly Regex _regex3 = new Regex(@"Timing Mode:[\t](?<str>\w*)", RegexOptions.IgnoreCase);
    private static readonly Regex _regex4 = new Regex(@"Master Timeset Name:[\t](?<str>\w*)", RegexOptions.IgnoreCase);
    private static readonly Regex _regex5 = new Regex(@"Time Domain:[\t\s]*(?<str>.*?)[\t\s]*(?=Strobe Ref Setup Name:|$)", RegexOptions.IgnoreCase);
    private static readonly Regex _regex6 = new Regex(@"Time Set[\t]Period");
    private static readonly Regex _regex7 = new Regex(".txt", RegexOptions.IgnoreCase);
    private static readonly Regex _regex8 = new Regex("^_", RegexOptions.IgnoreCase);
    private static readonly Regex _regex9 = new Regex(@"\s+");

    public TimeSetSheets GenerateFlow(List<PatternData> patList, string tsetPath, string tempPath)
    {
        List<string> tsetFileList = GetTsetFileList(patList);
        if (TestPlanStatic.BinCutInstanceSheets != null || TestPlanStatic.ScanInstanceSheets != null || TestPlanStatic.EvsInstanceSheets != null || TestPlanStatic.RtosSheets != null
            || TestPlanStatic.IdsSheets != null)
        {
            List<string> timeSet = GetTimeSetByInstanceSheet();
            List<string> timeSetForce = GetTimeSetByForceCondition();
            var timeSetWithVersion = timeSet.Select(x => AcTSetCategoryMapSingleton.Instance().GetTimeSetVersion(x)).Where(x => !string.IsNullOrEmpty(x)).ToList();
            tsetFileList = tsetFileList.Union(timeSetWithVersion, StringComparer.CurrentCultureIgnoreCase).Union(timeSetForce, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        TimeSetSheets timeSetSheets = GenTsetFile(tsetFileList, tsetPath, tempPath);
        SetDefaultValueToVaiable(timeSetSheets);

        if (LocalSpecs.Options.Device != EnumDevice.LCD)
        {
            GenScanTimeSetEqnBase(patList, timeSetSheets);
        }

        return timeSetSheets;
    }

    private List<string> GetTimeSetByInstanceSheet()
    {
        var timeset = new List<string>();
        if (TestPlanStatic.BinCutInstanceSheets != null)
        {
            timeset.AddRange(TestPlanStatic.BinCutInstanceSheets.SelectMany(x => x.Rows).Select(GetOriTimeSet).Distinct()
                .Where(x => !string.IsNullOrEmpty(x)));
        }

        if (TestPlanStatic.ScanInstanceSheets != null)
        {
            timeset.AddRange(TestPlanStatic.ScanInstanceSheets.SelectMany(x => x.Rows).Select(GetOriTimeSet).Distinct()
                .Where(x => !string.IsNullOrEmpty(x)));
        }

        if (TestPlanStatic.EvsInstanceSheets != null)
        {
            timeset.AddRange(TestPlanStatic.EvsInstanceSheets.SelectMany(x => x.Rows).Select(GetOriTimeSet).Distinct()
                .Where(x => !string.IsNullOrEmpty(x)));
        }

        return timeset.Distinct().ToList();
    }

    private List<string> GetTimeSetByForceCondition()
    {
        var timeset = new List<string>();

        if (TestPlanStatic.RtosSheets != null)
        {
            timeset.AddRange(TestPlanStatic.RtosSheets.PlanDic.Values.SelectMany(x => x.Rows).Select(r => r.RtosIdsTimeSetUsed).ToList());
        }

        if (TestPlanStatic.IdsSheets != null)
        {
            timeset.AddRange(TestPlanStatic.IdsSheets.PlanDic.Values.SelectMany(x => x.Rows).Select(r => r.RtosIdsTimeSetUsed).ToList());
        }

        return timeset.Distinct().ToList();

    }

    #region Change Scan Time Set as equation based
    private void GenScanTimeSetEqnBase(List<PatternData> pPatList, TimeSetSheets timeSetSheets)
    {
        if (LocalSpecs.CompileItem == null)
        {
            //No Compile Infomation
            return;
        }
        var timeSetNeedChangedDict = new Dictionary<string, List<string>>();

        for (int i = 0; i < pPatList.Count; i++)
        {
            PatternData tPatternSet = pPatList[i];
            string tPatternFileName = tPatternSet.FileVersion;
            string tTimeSet = tPatternSet.TimeSetVersion;

            // jump run when patternfilename is na or timeset is na.
            if (tPatternFileName.Equals("NA", StringComparison.OrdinalIgnoreCase) ||
                tTimeSet.Equals("NA", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string tUpperedPattenName = GetPattenName(tPatternFileName).ToUpper();

            if (!LocalSpecs.CompileItem.ContainsKey(tUpperedPattenName.ToUpper()))
            {
                continue;
            }
            // get target time set list
            CompileItem tCompPatternInfo = LocalSpecs.CompileItem[tUpperedPattenName];

            List<string> nameList = tCompPatternInfo.ScanSetupTSet.Split('|').ToList();
            foreach (string name in nameList)
            {
                if (name.Equals(""))
                {
                    continue;
                }
                if (!timeSetNeedChangedDict.ContainsKey(tTimeSet))
                {
                    timeSetNeedChangedDict[tTimeSet] = [];
                }

                if (!timeSetNeedChangedDict[tTimeSet].Contains(name))
                {
                    timeSetNeedChangedDict[tTimeSet].Add(name);
                }
            }
        }

        foreach (ComTimeSetBasicSheet timeSetBasicSheet in timeSetSheets)
        {
            if (!timeSetNeedChangedDict.ContainsKey(timeSetBasicSheet.Name))
            {
                continue;
            }
            foreach (TSet tset in timeSetBasicSheet.Rows)
            {
                var timeSet = (ComTimeSetBasic)tset;
                if (!timeSetNeedChangedDict[timeSetBasicSheet.Name].Exists(p => p.Equals(timeSet.Name, StringComparison.OrdinalIgnoreCase)) ||
                    JudgeTimeSetIsEqBased(timeSet))
                {
                    continue;
                }

                timeSetBasicSheet.AddShiftInTSetName(timeSet.Name);
                ChangeTimeSetValueForEqnBase(timeSet, timeSetBasicSheet.IsMultiShiftInTSet);

            }
        }
    }

    private string GetPattenName(string pPatternFullName)
    {
        if (FunctionSingleton.IsMatch(pPatternFullName, @"^.*[/](?<patternName>(.*))(\.atp\.gz)$"))
        {
            return _regex.Match(pPatternFullName).Groups["patternName"].ToString();
        }

        return "";
    }

    internal bool JudgeTimeSetIsEqBased(ComTimeSetBasic timeSet)
    {
        if (_regex2.IsMatch(timeSet.CyclePeriod))
        {
            return true;
        }
        foreach (TimingRow row in timeSet.TimingRows)
        {
            if (_regex.IsMatch(row.DriveOn) || _regex.IsMatch(row.DriveData)
                || _regex.IsMatch(row.DriveReturn) || _regex.IsMatch(row.DriveOff)
                || _regex.IsMatch(row.CompareOpen) || _regex.IsMatch(row.CompareClose))
            {
                return true;
            }
        }
        return false;
    }

    private void SetDefaultValueToVaiable(TimeSetSheets timeSetSheets)
    {
        var nWireVaiableVars = new HashSet<string>();

        List<ProtocolAwarePin> nwirePins = NwireSingleton.Instance().SettingInfo.NwirePins;
        if (nwirePins != null && nwirePins.Any())
        {
            foreach (ProtocolAwarePin nWirePin in nwirePins)
            {
                string nWireVaiable = nWirePin.CreateFreqVarName();
                nWireVaiableVars.Add(nWireVaiable);
            }
        }

        foreach (ComTimeSetBasicSheet timeSetBasicSheet in timeSetSheets)
        {
            foreach (TSet timeSetBasic in timeSetBasicSheet.Rows)
            {
                var timeSet = (ComTimeSetBasic)timeSetBasic;
                foreach (string subContexVaiiable in timeSet.SubContextVariable)
                {
                    if (!nWireVaiableVars.Contains(subContexVaiiable) && !timeSet.SubCommentVariable.ContainsKey(subContexVaiiable))
                    {
                        timeSet.SubCommentVariable.Add(subContexVaiiable, -1);
                    }
                }
            }
        }
    }

    internal void ChangeTimeSetValueForEqnBase(ComTimeSetBasic pTimeSetBasic, bool isMulitShift)
    {

        string warningString = !isMulitShift ? "" : "(Warning)";  // add warn to each parameter

        string shiftInFreq = SpecFormat.GenAcSpecSymbol(BasicInitial.ShiftInFreqVar) + warningString;

        string conBaseValue = "=1/_" + shiftInFreq;

        double freq20Mhz = 20e6;

        if (double.TryParse(pTimeSetBasic.CyclePeriod, out double period))
        {
            freq20Mhz = 1 / period;
        }


        string lBaseCyclePeriod = pTimeSetBasic.CyclePeriod;
        // set Cycle Period of time set
        pTimeSetBasic.CyclePeriod = conBaseValue;
        // set all info of sub rows


        string clockSVar = SpecFormat.GenAcSpecSymbol(BasicInitial.ClockS);
        string strobe = SpecFormat.GenAcSpecSymbol(BasicInitial.Strobe);

        // need remove waiting Cyprus  
        string clockEVar = SpecFormat.GenAcSpecSymbol(BasicInitial.ClockE);
        string cycleSVar = SpecFormat.GenAcSpecSymbol(BasicInitial.CycleS);


        // don't add into ac variable
        if (!isMulitShift)
        {
            pTimeSetBasic.SubCommentVariable.Add(shiftInFreq, freq20Mhz);
            pTimeSetBasic.SubCommentVariable.Add(clockSVar, 0);
            pTimeSetBasic.SubCommentVariable.Add(strobe, 0);

            // need remove waiting Cyprus  
            pTimeSetBasic.SubCommentVariable.Add(cycleSVar, 0);
            pTimeSetBasic.SubCommentVariable.Add(clockEVar, 0);

            foreach (KeyValuePair<string, double> oneRow in pTimeSetBasic.ShiftInReserve)
            {
                if (pTimeSetBasic.SubCommentVariable.ContainsKey(oneRow.Key))
                {
                    pTimeSetBasic.SubCommentVariable[oneRow.Key] = oneRow.Value;
                }
            }

        }

        foreach (TimingRow timingRow in pTimeSetBasic.TimingRows)
        {
            var allTimeSetEg = new List<string>
            {
                timingRow.DriveOn, timingRow.DriveData, timingRow.DriveReturn, timingRow.DriveOff,
                timingRow.CompareOpen
            };

            if (!timingRow.DataFmt.Equals("RL", StringComparison.OrdinalIgnoreCase) &&
                !timingRow.DataFmt.Equals("RH", StringComparison.OrdinalIgnoreCase))
            {
                timingRow.DriveOn = ChangeTimeSetValueDetail(lBaseCyclePeriod, allTimeSetEg, shiftInFreq, 0, false);
                timingRow.DriveData = ChangeTimeSetValueDetail(lBaseCyclePeriod, allTimeSetEg, shiftInFreq, 1, false);
                timingRow.CompareOpen = ChangeTimeSetValueDetail(lBaseCyclePeriod, allTimeSetEg, shiftInFreq, 4, false);
                continue;
            }
            timingRow.DriveOn = ChangeTimeSetValueDetail(lBaseCyclePeriod, allTimeSetEg, shiftInFreq, 0);
            timingRow.DriveData = ChangeTimeSetValueDetail(lBaseCyclePeriod, allTimeSetEg, shiftInFreq, 1);
            timingRow.DriveReturn = ChangeTimeSetValueDetail(lBaseCyclePeriod, allTimeSetEg, shiftInFreq, 2);
            timingRow.DriveOff = ChangeTimeSetValueDetail(lBaseCyclePeriod, allTimeSetEg, shiftInFreq, 3);
            timingRow.CompareOpen = ChangeTimeSetValueDetail(lBaseCyclePeriod, allTimeSetEg, shiftInFreq, 4);
        }
    }

    internal decimal GetEgValueInDecimal(string inStr)
    {
        DataTable dt = new DataTable();
        string? dValStr = dt.Compute(inStr.TrimStart('='), string.Empty).ToString();
        if (!decimal.TryParse(dValStr, NumberStyles.Float, CultureInfo.CurrentCulture, out decimal dVal))
        {
            throw new Exception(string.Format("Incorrect Timeset Format : {0}", inStr));
        }
        return dVal;
    }

    internal string ChangeTimeSetValueDetail(string pBaseCyclePeriod, List<string> allCyclePeriod,
        string pshiftInFreq, int index, bool isClk = true)
    {
        // if Current CyclePeriod is not a real value
        // do Nothing
        const int d1OnIdx = 1;
        const int d2ReturnIdx = 2;

        string lshiftInFreq = SpecFormat.SpecValuePrefix + pshiftInFreq;

        //STEP0. if in timeSet on/data/return/off/open field is empty, just return empty
        if (allCyclePeriod[index] == "")
        {
            return allCyclePeriod[index];
        }

        //STEP1. if not empty, calc ratio (Edge/Period) first
        //       expect only Data / RETURN / OPEN edge will contains value
        // transform Scientific Notation Number.
        decimal periodVal = GetEgValueInDecimal(pBaseCyclePeriod);
        decimal egVal = GetEgValueInDecimal(allCyclePeriod[index]);
        if (isClk && index == d2ReturnIdx)  //when calc D2 eg, need to know D1 value to calc the interval between them
        {
            GetEgValueInDecimal(allCyclePeriod[d1OnIdx]);
        }

        // calculate Ratio
        decimal tRatio = egVal / periodVal;
        // round off ratio to the 2nd decimal place
        string tRoundedRatio = Math.Round(tRatio, 3, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);

        //STEP2. Fill all fill, expect only Data / RETURN / OPEN edge will need to fill
        string tResult = "=";
        switch (index)
        {
            case 0:  //ON
                tResult = "0";
                break;
            case 3:  //OFF
                tResult = "0";
                break;
            case 1:  //DATA                    
                if (tRoundedRatio == "0")
                {
                    //if it's clock pin, even d1 edge is 0, still need to sweep in shiftIn CZ, so put _Clock_S_Var in the cell
                    //so clock in d1 = _Clock_S_Var/_ShiftIn_Freq_Var
                    tResult = isClk ? $"={tRoundedRatio}/{lshiftInFreq}" : "0";
                }
                else
                {
                    //clock pin d1 = _Clock_S_Var/_ShiftIn_Freq_Var, and _Clock_S_Var will give an initial value in percetage of the original timeSet
                    //Non clock d1 = Ratio/_ShiftIn_Freq_Var
                    tResult = $"={tRoundedRatio}/{lshiftInFreq}";
                }
                break;
            case 2:  //RETURN
                //Notice, In d2 portion, wanna keep the duty of the original timeSet, so the "ratio" here means interval between d1 and d2
                if (tRoundedRatio == "0")
                {
                    //clock pin d2 = (_Clock_S_Var+Diff)/_ShiftIn_Freq_Var
                    tResult = $"={tRoundedRatio}/{lshiftInFreq}";
                }
                else
                {
                    //clock pin d2 = (_Clock_S_Var+Diff)/_ShiftIn_Freq_Var
                    //Non clock d2 = Ratio/_ShiftIn_Freq_Var
                    tResult = $"={tRoundedRatio}/{lshiftInFreq}";
                }

                break;
            case 4:  //OPEN
                if (tRoundedRatio == "0")
                {
                    tResult = "0";
                }
                else
                {
                    //expect all C0 are equal
                    tResult = $"={tRoundedRatio}/{lshiftInFreq}";
                }

                break;
        }
        return tResult;
    }
    #endregion

    public List<string> GetTsetFileList(List<PatternData> patList)
    {
        var tsetFileList = new List<string>();
        foreach (PatternData pattern in patList)
        {
            string tset = pattern.TimeSetVersion;
            if (!tsetFileList.Contains(tset) && !IgnoredFileName(tset))
            {
                tsetFileList.Add(tset);
            }
        }
        return tsetFileList;
    }

    private static List<string> CopyTimeSetsToCommon(List<string> tsetFileList, string tsetPath, string tempPath)
    {
        List<string> srcTimeSetsPaths = [];
        foreach (string fileName in tsetFileList)
        {
            if (fileName == string.Empty)
            {
                continue;
            }

            string completeFileName = Path.ChangeExtension(fileName, ".txt");
            string fileFullPath = Path.Combine(tsetPath, completeFileName);
            if (File.Exists(fileFullPath))
            {
                string tarFile = Path.Combine(tempPath, completeFileName);
                File.Copy(fileFullPath, tarFile, true);
                srcTimeSetsPaths.Add(tarFile);
            }
            else
            {
                ErrorMessageBox.Show($"Can't find this file: {fileFullPath}. ", ConMissingFile);
            }
        }
        return srcTimeSetsPaths;
    }

    private TimeSetSheets GenTsetFile(
        List<string> tsetFileList,
        string tsetPath, string tempPath
    )
    {
        try
        {
            if (!Directory.Exists(tsetPath))
            {
                return [];
            }

            if (!Directory.Exists(tempPath))
            {
                Directory.CreateDirectory(tempPath);
            }

            List<string> srcTimeSetsPaths = CopyTimeSetsToCommon(tsetFileList, tsetPath, tempPath);
            return ReadTimeSetTxt1P4(srcTimeSetsPaths);
        }
        catch (Exception e)
        {
            throw new Exception("Error occurred in generating time set " + e.StackTrace);
        }
    }

    internal bool IgnoredFileName(string fileName)
    {
        if (fileName.ToUpper().Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
        {
            return true;
        }

        if (fileName.ToUpper().Equals(Na, StringComparison.CurrentCultureIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private static int GetRowNumber(string[] lines)
    {
        int lIStartRowNum = 4;
        for (int i = lIStartRowNum; i < lines.Length; i++)
        {
            if (_regex6.IsMatch(lines[i]))
            {
                lIStartRowNum = i + 1;
                break;
            }
        }
        return lIStartRowNum;
    }

    private static void AddSubCommentsToTimeSets(string line, Dictionary<string, ComTimeSetBasic> timeSetDatas)
    {
        if (line.Contains("=")) //	HTOL_Freq_VAR	=1000000, tset1Per	=1000.000*ns
        {
            string[] spt = line.Split('=');
            string varTok = spt[0].Trim(); // ^_ need replace 
            varTok = _regex8.Replace(varTok, "");
            string valueTok = spt[1].Trim();

            //HardIpUtilityMain.ResetUtilities();
            string valueSter = DataConvertor.ConvertUnits(valueTok);
            //Use HardIp Function
            bool isNumOk = double.TryParse(valueSter, out double dValue);
            if (isNumOk)
            {
                foreach (KeyValuePair<string, ComTimeSetBasic> subTsb in timeSetDatas)
                {
                    if (subTsb.Value.SubContextVariable.Contains(varTok))
                    //which means the variable appear on the comment is use by this time set
                    {
                        if (!subTsb.Value.SubCommentVariable.ContainsKey(varTok))
                        {
                            subTsb.Value.SubCommentVariable.Add(varTok, dValue);
                        }
                    }
                }
            }
            //Add error report...  under construct
        }
        else //just write variable name, and no equal char (ex: AAA, BBB )
        {
            string varTok = _regex9.Replace(line, "");
            double dValue = -1e9;
            foreach (KeyValuePair<string, ComTimeSetBasic> subTsb in timeSetDatas)
            {
                if (subTsb.Value.SubContextVariable.Contains(varTok))
                //which means the variable appear on the comment is use by this time set
                {
                    if (!subTsb.Value.SubCommentVariable.ContainsKey(varTok))
                    {
                        subTsb.Value.SubCommentVariable.Add(varTok, dValue);
                    }
                }
            }

            //Add error report...  under construct
            //-> for no initial value
        }
    }

    private Result<ComTimeSetBasicSheet> HandleTimeSet(string timeSetPath)
    {
        string[] lines = File.ReadAllLines(timeSetPath);

        string lStrStrobe = "";

        string lStrTimeMode = _regex3.Match(lines[2]).Groups["str"].ToString();
        string lStrMasterTs = _regex4.Match(lines[2]).Groups["str"].ToString();
        string lStrTimeDomain = _regex5.Match(lines[3]).Groups["str"].ToString();

        // support 1.4 ,2.3 timing row conveter 20180613 by JN
        TimeRow1P4Converter timeRowConverter = Converter(lines[0]);
        string sheetName = _regex7.Replace(Path.GetFileName(timeSetPath), "");
        var timeSetBasicSheet = new ComTimeSetBasicSheet(sheetName, lStrTimeMode, lStrMasterTs,
            lStrTimeDomain, lStrStrobe);
        var timeSetDatas = new Dictionary<string, ComTimeSetBasic>();
        int lIStartRowNum = GetRowNumber(lines);

        bool formaterror = false;
        bool startVarDefinitions = false;
        bool startSimulationClockSetup = false;

        for (int i = lIStartRowNum; i < lines.Length; i++)
        {
            string lStrTimeSet, lStrClockPeriod;
            string[] datas = lines[i].Split('\t');

            if (!datas.Any() ||
                !lines[i].Split(['\t'], StringSplitOptions.RemoveEmptyEntries).Any())
            {
                continue;
            }

            if (lines[i].ContainsIgnoreCase("Simulation clock setup".ToUpper()))
            {
                startSimulationClockSetup = true;
                continue;
            }
            if (lines[i].Contains("VAR Definitions") && !startVarDefinitions)
            {
                startVarDefinitions = true;
                continue;
            }
            if (startSimulationClockSetup)
            {
                continue;
            }
            if (startVarDefinitions)
            {
                AddSubCommentsToTimeSets(lines[i], timeSetDatas);

                continue;
            }

            lStrTimeSet = datas[1];
            lStrClockPeriod = datas[2];
            //ReadTimeRow1P4() add argument _contextVar for read equation base variable
            if (
                !ReadTimeRow(
                    datas,
                    timeRowConverter,
                    out TimingRow timingRow,
                    out List<string> contextVar,
                    out Dictionary<string, double> shiftFreqVar,
                    out formaterror)
            )
            {
                if (formaterror)
                {
                    break;
                }

                if (!TimeSetWithWrongForamtRows.ContainsKey(sheetName))
                {
                    TimeSetWithWrongForamtRows.Add(sheetName, new List<int>());
                }

                TimeSetWithWrongForamtRows[sheetName].Add(i + 1);
            }

            if (timeSetDatas.ContainsKey(lStrTimeSet))
            {
                timeSetDatas[lStrTimeSet].AddTimingRow(timingRow);
                foreach (string varTmp in contextVar)
                {
                    if (!timeSetDatas[lStrTimeSet].SubContextVariable.Contains(varTmp))
                    {
                        timeSetDatas[lStrTimeSet].SubContextVariable.Add(varTmp);
                    }
                }

                foreach (KeyValuePair<string, double> dicPair in shiftFreqVar)
                {
                    if (!timeSetDatas[lStrTimeSet].ShiftInReserve.ContainsKey(dicPair.Key))
                    {
                        timeSetDatas[lStrTimeSet].ShiftInReserve.Add(dicPair.Key, dicPair.Value);
                    }
                }
            }
            else
            {
                var timeSetBasic = new ComTimeSetBasic { Name = lStrTimeSet, CyclePeriod = lStrClockPeriod };
                timeSetBasic.AddTimingRow(timingRow);

                foreach (string varTmp in contextVar)
                {
                    timeSetBasic.SubContextVariable.Add(varTmp);
                }

                foreach (KeyValuePair<string, double> dicPair in shiftFreqVar)
                {
                    timeSetBasic.ShiftInReserve.Add(dicPair.Key, dicPair.Value);
                }

                timeSetDatas.Add(lStrTimeSet, timeSetBasic);
            }
        }

        if (formaterror)
        {
            if (!TimeSetIncorrectFormat.Contains(sheetName))
            {
                TimeSetIncorrectFormat.Add(sheetName);
            }

            return Result<ComTimeSetBasicSheet>.Fail(sheetName);
        }

        foreach (KeyValuePair<string, ComTimeSetBasic> keyValuePair in timeSetDatas)
        {
            timeSetBasicSheet.AddRow(keyValuePair.Value);
        }
        if (Path.GetFileNameWithoutExtension(timeSetPath).Contains("RTOS"))
        {
            AddRtosContain(timeSetBasicSheet);
        }

        return Result<ComTimeSetBasicSheet>.Ok(timeSetBasicSheet);
    }

    public TimeSetSheets ReadTimeSetTxt1P4(List<string> timeSetPathList)
    {
        TimeSetSheets multiTimeSetSheets = [];
        try
        {
            foreach (string timeSetPath in timeSetPathList)
            {
                if (!File.Exists(timeSetPath))
                {
                    continue;
                }

                try
                {
                    Result<ComTimeSetBasicSheet> result = HandleTimeSet(timeSetPath);
                    if (!result.Success)
                    {
                        continue;
                    }
                    multiTimeSetSheets.Add(result.Value);
                }
                catch (Exception)
                {
                    ErrorMessageBox.Show($"Parsing Timeset {timeSetPath} fail. Please check");
                }
            }
        }
        catch (Exception ex)
        {
            ErrorMessageBox.Show(string.Format(ex.ToString()));
        }
        return multiTimeSetSheets;
    }

    private void AddRtosContain(ComTimeSetBasicSheet timesetSheet)
    {
        Dictionary<string, string> uartPinDic = TestProgram.IgxlWorkBk.PinMapPair.Value.GetUartPinDic();
        foreach (KeyValuePair<string, string> pin in uartPinDic)
        {
            string pinName = pin.Value;
            string pinType = pin.Key;
            timesetSheet.Rows.ForEach(x => x.TimingRows.RemoveAll(row => row.PinGrpName.Equals(pinName, StringComparison.OrdinalIgnoreCase)));
            var timesetBasic = new ComTimeSetBasic { Name = pinType, CyclePeriod = "=(1/_UART_RATE_GLB_VAR)/3" };
            var timesetRow = new TimingRow
            {
                PinGrpName = pinName,
                PinGrpSetup = "PA",
                DataSrc = "PA",
                DataFmt = "NR",
                DriveOff = "0",
                DriveData = "0",
                CompareMode = "PA",
                CompareOpen = "=(_UART_STROBE_GLB_VAR/_UART_RATE_GLB_VAR)/3",
                EdgeMode = "Auto"
            };
            timesetBasic.SubContextVariable.Add("UART_RATE_GLB_VAR");
            timesetBasic.SubCommentVariable.Add("UART_RATE_GLB_VAR", 115200);
            timesetBasic.SubContextVariable.Add("UART_STROBE_GLB_VAR");
            timesetBasic.SubCommentVariable.Add("UART_STROBE_GLB_VAR", 0.7125);
            timesetBasic.TimingRows.Add(timesetRow);
            timesetSheet.AddRow(timesetBasic);

        }
    }

    private bool ReadTimeRow(string[] line, TimeRow1P4Converter converter, out TimingRow row, out List<string> subContextVar, out Dictionary<string, double> shiftInReserveVar, out bool wrongContent)
    {
        subContextVar = new List<string>();
        shiftInReserveVar = new Dictionary<string, double>();
        wrongContent = false;
        row = converter.ConvertTimeRow(line);

        GetContextVariable(row.PinGrpClockPeriod, ref subContextVar);
        GetContextVariable(row.DriveOn, ref subContextVar);
        GetContextVariable(row.DriveData, ref subContextVar);
        GetContextVariable(row.DriveReturn, ref subContextVar);
        GetContextVariable(row.DriveOff, ref subContextVar);
        GetContextVariable(row.CompareOpen, ref subContextVar);
        GetContextVariable(row.CompareClose, ref subContextVar);

        string cyclePeriod = line[2];
        if (cyclePeriod != "")
        {
            if (IsContextVariable(cyclePeriod))
            {
                GetContextVariable(cyclePeriod, ref subContextVar);
            }
            else
            {
                decimal periodVal = //get period value
                    GetEgValueInDecimal(cyclePeriod);

                if (periodVal != (decimal)0.0 && row.DriveData != "") //check D1
                {
                    if (!Regex.IsMatch(row.DriveData, @"_(?<var>[\d|\w]+)", RegexOptions.IgnoreCase))
                    {
                        DataTable dt = new DataTable();
                        string? driveData = dt.Compute(row.DriveData.TrimStart('='), string.Empty).ToString();
                        if (!decimal.TryParse(driveData, NumberStyles.Float, CultureInfo.CurrentCulture, out decimal _))
                        {
                            wrongContent = true;
                            return false;
                        }
                        decimal d1Val = GetEgValueInDecimal(driveData);
                        decimal tRatio = d1Val / periodVal;
                        tRatio = Math.Round(tRatio, 2);
                        if (Regex.IsMatch(row.DataFmt, "RL", RegexOptions.IgnoreCase)) //if (tRatio != (decimal)0)
                        {
                            if (!shiftInReserveVar.ContainsKey(SpecFormat.GenAcSpecSymbol(BasicInitial.ClockS)))
                            {
                                shiftInReserveVar.Add(SpecFormat.GenAcSpecSymbol(BasicInitial.ClockS), (double)tRatio);
                            }
                        }
                    }
                }

                if (periodVal != (decimal)0.0 && row.CompareOpen != "") //check C0
                {
                    if (!Regex.IsMatch(row.CompareOpen, @"_(?<var>[\d|\w]+)", RegexOptions.IgnoreCase))
                    {
                        DataTable dt = new DataTable();
                        string? compareOpen = dt.Compute(row.CompareOpen.TrimStart('='), string.Empty).ToString();
                        if (!decimal.TryParse(compareOpen, NumberStyles.Float, CultureInfo.CurrentCulture, out decimal _))
                        {
                            wrongContent = true;
                            return false;
                        }

                        decimal c0Val = GetEgValueInDecimal(compareOpen);
                        decimal tRatio = c0Val / periodVal;
                        tRatio = Math.Round(tRatio, 2);
                        if (tRatio != 0)
                        {
                            if (!shiftInReserveVar.ContainsKey(SpecFormat.GenAcSpecSymbol(BasicInitial.Strobe)))
                            {
                                shiftInReserveVar.Add(SpecFormat.GenAcSpecSymbol(BasicInitial.Strobe), (double)tRatio);
                            }
                        }
                    }
                }
            }
        }

        return !converter.NeedCompensate(line);
    }

    internal void GetContextVariable(string cell, ref List<string> subContextVar)
    {
        //cell context example:
        //=_RT_CLK32768_Freq_GLB 
        //=(1/_TCK_Freq_VAR)
        //=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR+_Strobe_VAR
        //=_Cycle_S_VAR+0.7/_ShiftIn_Freq_VAR+_Clock_E_VAR

        MatchCollection matches = Regex.Matches(cell, @"_(?<var>[\d|\w]+)");
        foreach (Match match in matches)
        {
            string contextVar = match.Groups["var"].ToString();
            if (contextVar != "" && !subContextVar.Contains(contextVar))
            {
                subContextVar.Add(contextVar);
            }
        }
    }

    internal bool IsContextVariable(string cell)
    {
        //cell context example:
        //=_RT_CLK32768_Freq_GLB 
        //=(1/_TCK_Freq_VAR)
        //=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR+_Strobe_VAR
        //=_Cycle_S_VAR+0.7/_ShiftIn_Freq_VAR+_Clock_E_VAR    
        return Regex.IsMatch(cell, @"_(?<var>[\d|\w]+)");
    }

    public static TimeRow1P4Converter Converter(string header)
    {
        if (Regex.IsMatch(header, "DTTimesetBasicSheet,version=2.3", RegexOptions.IgnoreCase))
        {
            return new TimeRow2P3Converter();
        }

        return new TimeRow1P4Converter();
    }

    public string GetOriTimeSet(BinCutInstanceRow binCutInstanceRow)
    {
        return binCutInstanceRow.TimeSet.Contains(':') ? binCutInstanceRow.TimeSet.Split(':').First() : binCutInstanceRow.TimeSet;
    }
}
