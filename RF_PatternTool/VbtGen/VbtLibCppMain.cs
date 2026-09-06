using System.Text.RegularExpressions;

using RfLib.Dvdc.Reader.CapturePostProcess;

namespace RF_PatternTool.VbtGen
{
    public class SelCppMainFunc
    {
        private List<string> _cppSetupName = new List<string>();

        public SelCppMainFunc(List<string> cppsetupname)
        {
            _cppSetupName = cppsetupname;
        }

        public void Get_Select_CPP_Main_Func(List<string> content)
        {
            Header(content);
            SetCppSetupName(content);
            Footer(content);
        }

        private static void Header(List<string> content)
        {
            content.Add("Public Function Select_CPP_Main_Func(CPP_CapWave As DspWave, CPP_SetupName As String)");
            content.Add("On Error GoTo errHandler");
            content.Add("");
            content.Add("\tSelect Case UCase(CPP_SetupName)");
            content.Add("");
        }

        private void SetCppSetupName(List<string> content)
        {
            foreach (string setupname in _cppSetupName)
            {
                content.Add($"\tCase \"{setupname.ToUpper()}\"");
                content.Add("");
                content.Add($"\t\tCall {setupname}(CPP_CapWave)");
                content.Add("");
            }
        }

        private static void Footer(List<string> content)
        {
            content.Add("\tCase Else");
            content.Add("");
            content.Add("\t\tCall SourceDriverDigCapProcess(CPP_CapWave, CPP_SetupName)");
            content.Add("");
            content.Add("\tEnd Select");
            content.Add("");
            content.Add("Exit Function");
            content.Add("errHandler:");
            content.Add("\tCall Print_Error_Message(Error_Info, \"VBT_LIB_CPP_Main\", \"Select_CPP_Main_Func\")");
            content.Add("\tIf AbortTest Then Exit Function Else Resume Next");
            content.Add("End Function");
            content.Add("");
        }

    }

    public class CppSetupFunc
    {
        public class ResultWave
        {
            public string WaveType = "";
            public string VarName = "";
            public int WaveNum = 0;
            public string WaveBW = "0";
            public string WaveTestName = "";
            public string WaveStoreName = "";
            public string WaveCalcStoreName = "";

            public ResultWave(string wavetype, string varname, int wavenum, string bw, string wavetestname, string wavestorename, string wavecalcstorename)
            {
                WaveType = wavetype;
                VarName = varname + (wavenum + 1);
                WaveNum = wavenum;
                WaveBW = bw;
                WaveTestName = wavetestname;
                WaveStoreName = wavestorename;
                WaveCalcStoreName = wavecalcstorename;
            }
        }

        private string _setupName;
        private List<ResultWave> _resultWaves = new List<ResultWave>();

        private HashSet<string> _hashStoreNames = new HashSet<string>();
        private HashSet<string> _hashDspArgs = new HashSet<string>();

        private int _calcResultNum = 0;
        private bool _isCalcBestCodeCPP = false;
        private bool _isNonrffwCalcFun = false;

        private List<string> _parInfoToWaveArgs = new List<string>() { "rangeWave1", "rangeWave2", "targetWave", "BestCodeCalcEnum" };
        private List<string> _runDspArgs = new List<string>() { "CPP_CapWave", "rawResultWave", "calcResultWave" };

        private static readonly Regex _regexNum = new Regex(@"^-?\d+(\.\d+)?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public CppSetupFunc(List<PostProcessSheetRow> cpprows)
        {
            int storeInd = 0;
            int calcStoreInd = 0;

            _setupName = cpprows.First().SetupName;

            foreach (PostProcessSheetRow row in cpprows)
            {
                if (row.StoreName.Any())
                {
                    _hashStoreNames.Add(row.StoreName);
                }

                if (row.CalcStoreName.Any())
                {
                    _hashStoreNames.Add(row.CalcStoreName);
                }

                if (row.CalcTestName.Any())
                {
                    _calcResultNum++;

                    if (row.CalcEquation.StartsWith("rffw_calc_RAW"))
                    {
                        _resultWaves.Add(new ResultWave("localRawResultWave", "Raw", storeInd, row.BitWidth, row.CalcTestName, row.StoreName, row.CalcStoreName));
                        storeInd++;
                    }
                    else
                    {
                        _resultWaves.Add(new ResultWave("localCalcResultWave", "Calc", calcStoreInd, row.BitWidth, row.CalcTestName, row.StoreName, row.CalcStoreName));
                        calcStoreInd++;
                    }
                }

                if (row.CalcEquation.Any())
                {
                    int firstInd = row.CalcEquation.IndexOf('(');
                    string calcFuncName = row.CalcEquation.Substring(0, firstInd);
                    string[] calcArgs = row.CalcEquation.Substring(firstInd + 1, row.CalcEquation.Length - 1 - firstInd - 1).Split(',');

                    if (calcFuncName == "rffw_calc_RAW")
                    {
                        _calcResultNum--;
                    }
                    else if (calcFuncName == "Calc_BestCode_CPP")
                    {
                        _isCalcBestCodeCPP = true;
                        for (int i = 0; i < 4; i++)
                        {
                            _parInfoToWaveArgs.Insert(i, "\"" + calcArgs[i] + "\"");
                        }

                        _resultWaves.RemoveAll(w => w.WaveType == "localRawResultWave");
                    }
                    else
                    {
                        foreach (string arg in calcArgs)
                        {
                            if (!_hashStoreNames.Contains(arg) &&
                                !_hashDspArgs.Contains(arg) &&
                                !_regexNum.IsMatch(arg))
                            {
                                _hashDspArgs.Add(arg);
                            }
                        }
                        _isNonrffwCalcFun = true;
                    }
                }
            }
        }


        public void Get_Cpp_Setup_Func(List<string> content, List<List<string>> efusebitDefRows)
        {
            bool expendCode = true;
            Header(content, "", "CPP_CapWave As DSPWave");
            Set_DefVar(content);
            Set_CalcOutput(content);
            Set_RunDSPFunc(content);
            if (_isCalcBestCodeCPP)
            {
                Set_CalcResultWaveFunc(content);
            }

            Set_MoveDSPOutput(content);
            if (expendCode)
            {
                Set_PrintTestLimits(content);
                if (!_isCalcBestCodeCPP)
                {
                    Set_DefVarToStoreDicAndSetFuse(content, efusebitDefRows);
                }

                if (_isCalcBestCodeCPP)
                {
                    Set_AddCalcBestCodeCPPFunc(content);
                }
                else if (_isNonrffwCalcFun)
                {
                    Set_StroeBCFunc(content, "calcResultWave", "Calc");
                }

                Footer(content, "");
            }
            else
            {
                Set_RunPrintAndStoreFunc(content);
                if (_isCalcBestCodeCPP)
                {
                    Set_AddCalcBestCodeCPPFunc(content);
                }
                else if (_isNonrffwCalcFun)
                {
                    Set_StroeBCFunc(content, "calcResultWave", "Calc");
                }

                Footer(content, "");

                Header(content, "_Print", "localRawResultWave As DSPWave, localCalcResultWave As DSPWave");
                Set_PrintTestLimits(content);
                Footer(content, "_Print");

                if (!_isCalcBestCodeCPP)
                {
                    Header(content, "_Store", "localRawResultWave As DSPWave, localCalcResultWave As DSPWave");
                    Set_DefVarToStoreDicAndSetFuse(content, efusebitDefRows);
                    Footer(content, "_Store");
                }
            }
        }

        private void Header(List<string> content, string type, string args)
        {
            content.Add($"Public Function {_setupName}{type}({args})");
            content.Add("On Error GoTo errHandler");
            content.Add("");
        }

        private void Set_DefVar(List<string> content)
        {
            //content.Add("\t'Step1: Define variable");

            content.Add("\tDim Site As Variant");
            content.Add("\tDim rawResultWave As New DSPWave");
            content.Add("\tDim calcResultWave As New DSPWave");
            content.Add("\tDim localRawResultWave As New DSPWave");
            content.Add("\tDim localCalcResultWave As New DSPWave");

            content.Add("");

            if (_isCalcBestCodeCPP)
            {
                content.Add("\tDim BestCodeCalcEnum As Long");
                content.Add("\tDim targetWave As New DSPWave");
                content.Add("\tDim rangeWave1 As New DSPWave, rangeWave2 As New DSPWave");
                content.Add("\tDim BestCodeWave As New DSPWave");
                content.Add("\tDim BestValue1 As New SiteDouble, BestValue2 As New SiteDouble");
                content.Add("\tDim sweepWave As New DSPWave");
                content.Add("");
            }

        }

        private void Set_CalcOutput(List<string> content)
        {
            //content.Add("\t'Step2: Define how many calc output");

            content.Add($"\tcalcResultWave.CreateConstant 999, {_calcResultNum}, DspDouble");
            content.Add("");
        }

        private void Set_RunDSPFunc(List<string> content)
        {
            //content.Add("\t'Step3: Run DSPfunction");

            if (_hashDspArgs.Any())
            {
                foreach (string dsparg in _hashDspArgs)
                {
                    content.Add($"\tDim {dsparg} As New SiteDouble");
                    content.Add($"\t{dsparg} = GetStoredMeasurement(\"{dsparg}\")");
                }
                _runDspArgs.AddRange(_hashDspArgs);
                content.Add("");
            }

            if (_isCalcBestCodeCPP)
            {
                content.Add($"\tCall ParserInfoToWave({string.Join(", ", _parInfoToWaveArgs)})");

                _runDspArgs.Add("rangeWave1");
                _runDspArgs.Add("rangeWave2");
                _runDspArgs.Add("targetWave");
                _runDspArgs.Add("BestCodeCalcEnum");
                _runDspArgs.Add("BestCodeWave");
                _runDspArgs.Add("BestValue1");
                _runDspArgs.Add("BestValue2");
                _runDspArgs.Add("sweepWave");
            }

            content.Add($"\tCall rundsp.{_setupName}_DSP({string.Join(", ", _runDspArgs)})");

            content.Add("");
        }
        private static void Set_CalcResultWaveFunc(List<string> content)
        {
            content.Add("\tIf Theexec.enableWord(\"Verify_CPP_DSP\") Then");
            content.Add("\t\tCall verifyCPPDSP_storeCalc(calcResultWave)");
            content.Add("\tEnd If");

            content.Add("");
        }
        private void Set_MoveDSPOutput(List<string> content)
        {
            //content.Add("\t'Step4: Move DSP output to local");

            content.Add("\tFor Each Site In TheExec.sites");
            content.Add("\t\tlocalRawResultWave = rawResultWave.Copy");
            if (_calcResultNum > 0)
            {
                content.Add("\t\tlocalCalcResultWave = calcResultWave.Copy");
            }

            content.Add("\tNext");
            content.Add("");
        }

        private void Set_RunPrintAndStoreFunc(List<string> content)
        {
            //content.Add("\t'Step5: Run Printfunction and Storefunction");

            content.Add($"\tCall {_setupName}_Print(localRawResultWave, localCalcResultWave)");
            if (!_isCalcBestCodeCPP)
            {
                content.Add($"\tCall {_setupName}_Store(localRawResultWave, localCalcResultWave)");
            }
            content.Add("");
        }

        private void Set_AddCalcBestCodeCPPFunc(List<string> content)
        {

            string fusename = _parInfoToWaveArgs[3];
            content.Add($"\tCall Print_calc_BestCode_Limit(BestCodeCalcEnum, targetWave, {fusename}, BestCodeWave, BestValue1, BestValue2, True)");
            content.Add("");

            int index = 0;
            foreach (string name in fusename.Split(';'))
            {
                string fn = name.Trim('\"');
                content.Add($"\tDim bestcode{index + 1} As New SiteDouble");
                content.Add($"\tbestcode{index + 1} = BestCodeWave.Element({index})");
                content.Add($"\tCall AddStoredMeasurement(\"{fn}\", bestcode{index + 1})");
                content.Add($"\tCall Wireless_HIP_eFuse_Write(\"{fn}\", \"{fn}\", vbNullString, vbNullString)");
                content.Add("");
                index++;
            }

            content.Add($"\tCall GetBestCodeFromEfuse({fusename}, BestCodeCalcEnum, rangeWave1, rangeWave2, sweepWave, BestCodeWave, BestValue1, BestValue2)");
            content.Add($"\tCall Print_calc_BestCode_Limit(BestCodeCalcEnum, targetWave, {fusename}, BestCodeWave, BestValue1, BestValue2, False)");

            content.Add("");

            Set_StroeBCFunc(content, "BestCodeWave", "BC");


        }
        private static void Set_StroeBCFunc(List<string> content, string wavetype, string functiontype)
        {
            content.Add("\tIf TheExec.enableWord(\"Verify_CPP_DSP\") Then ");
            content.Add($"\t\tCall verifyCPPDSP_store{functiontype}({wavetype})");
            content.Add("\tEnd If");
            content.Add("");
        }
        private void Set_PrintTestLimits(List<string> content)
        {
            int total = 0;
            if (_isCalcBestCodeCPP)
            {
                for (int i = 0; i < _resultWaves.Count(); i++)
                {
                    Set_ExecTestLimit(content, i, i, ref total, _resultWaves[i].WaveType,
                        $", ForceResults:=tlForceNone, TName:=\"{_resultWaves[i].WaveTestName}\", scaleType:=scaleNone");
                }
            }
            else
            {
                content.Add("\tDim i As Long");

                Dictionary<string, int> chkRawCalc = new Dictionary<string, int> { { "localRawResultWave", 0 }, { "localCalcResultWave", 0 } };

                string currType = _resultWaves.First().WaveType;
                int currCalcstop = _resultWaves.First().WaveNum;
                int currElement = 0;
                foreach (ResultWave wave in _resultWaves)
                {
                    if (currType != wave.WaveType)
                    {
                        if (wave.WaveType == "localRawResultWave")
                        {
                            Set_ExecTestLimit(content, chkRawCalc["localCalcResultWave"], currCalcstop, ref total, "localCalcResultWave",
                            ", ForceResults:=tlForceFlow, scaleType:=scaleNoScaling");
                            chkRawCalc[currType] = currCalcstop + 1;
                            chkRawCalc["localRawResultWave"] = currElement;
                            currType = wave.WaveType;
                        }
                        else if (wave.WaveType == "localCalcResultWave")
                        {
                            Set_ExecTestLimit(content, chkRawCalc["localRawResultWave"], currElement - 1, ref total, "localRawResultWave",
                            ", ForceResults:=tlForceFlow, scaleType:=scaleNoScaling");
                            chkRawCalc["localRawResultWave"] = currElement;
                            currType = wave.WaveType;
                        }
                        currCalcstop = wave.WaveNum;

                    }
                    currElement = currElement + 1;
                    currCalcstop = wave.WaveNum;
                }
                Set_ExecTestLimit(content, chkRawCalc[currType], currCalcstop, ref total, currType,
                    ", ForceResults:=tlForceFlow, scaleType:=scaleNoScaling");
            }

            content.Add("");

        }

        private static void Set_ExecTestLimit(List<string> content, int start, int stop, ref int total, string type, string args)
        {
            if (start == stop)
            {
                content.Add($"\tTheExec.Flow.TestLimit {type}.Element({start}){args}");
                total = total + 1;
            }
            else
            {
                content.Add($"\tFor i = {start} To {stop}");
                content.Add($"\t\tTheExec.Flow.TestLimit {type}.Element(i){args}");
                content.Add("\tNext");
                total = total + stop - start + 1;
            }
        }


        private void Set_DefVarToStoreDicAndSetFuse(List<string> content, List<List<string>> efusebitDefRows)
        {
            int count = 0;
            foreach (ResultWave wave in _resultWaves)
            {
                count++;
                content.Add($"\tDim {wave.VarName} As New SiteDouble");
                content.Add($"\tDim StoreRaw{count} As New SiteDouble"); // new rule 251229
                if (wave.VarName.IndexOf("Raw", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    content.Add($"\t{wave.VarName} = {wave.WaveType}.Element({count - 1})");
                }
                else
                {
                    content.Add($"\t{wave.VarName} = {wave.WaveType}.Element({wave.WaveNum})");
                }

                content.Add($"\tStoreRaw{count} = localRawResultWave.Element({count - 1})"); // new rule 251229

                if (string.IsNullOrEmpty(wave.WaveCalcStoreName))
                {
                    content.Add($"\tCall AddStoredMeasurement(\"{wave.WaveStoreName}\", {wave.VarName})");
                    content.Add($"\tCall AddStoredMeasurement(\"{wave.WaveStoreName}\", StoreRaw{count})"); // new rule 251229
                    content.Add($"\tCall CheckStoreNameAndAddStoreBW(\"{wave.WaveStoreName}\", {(string.IsNullOrEmpty(wave.WaveBW) ? "0" : wave.WaveBW)})");
                    content.Add($"\tCall CheckStoreNameAndAddStoreBW(\"{wave.WaveStoreName}\", {(string.IsNullOrEmpty(wave.WaveBW) ? "0" : wave.WaveBW)})");
                }
                else
                {
                    content.Add($"\tCall AddStoredMeasurement(\"{wave.WaveStoreName}\", StoreRaw{count})");
                    content.Add($"\tCall AddStoredMeasurement(\"{wave.WaveCalcStoreName}\", {wave.VarName})"); // new rule 251229
                    content.Add($"\tCall CheckStoreNameAndAddStoreBW(\"{wave.WaveStoreName}\", {(string.IsNullOrEmpty(wave.WaveBW) ? "0" : wave.WaveBW)})");
                    content.Add($"\tCall CheckStoreNameAndAddStoreBW(\"{wave.WaveCalcStoreName}\", {(string.IsNullOrEmpty(wave.WaveBW) ? "0" : wave.WaveBW)})"); // new rule 251229

                }

                if (efusebitDefRows.Find(p => p[0].Equals(wave.WaveStoreName, StringComparison.OrdinalIgnoreCase)) != null)
                {
                    content.Add($"\tCall Wireless_HIP_eFuse_Write(\"{wave.WaveStoreName}\", \"{wave.WaveStoreName}\", vbNullString, vbNullString)");
                }
            }

            content.Add("");
        }


        private void Footer(List<string> content, string type)
        {
            content.Add("Exit Function");
            content.Add("errHandler:");
            content.Add($"\tCall Print_Error_Message(Error_Info, \"VBT_LIB_CPP_Main\", \"{_setupName}{type}\")");
            content.Add("\tIf AbortTest Then Exit Function Else Resume Next");
            content.Add("End Function");
            content.Add("");
        }

    }

    public class VbtLibCppMain
    {

        public static void AddHeader(List<string> content, string vbtname)
        {
            content.Insert(0, $"Attribute VB_Name = \"{vbtname}\"");
            content.Insert(1, "Option Explicit");
            content.Insert(2, "");
        }

        public static void Write(List<List<string>> content, List<List<PostProcessSheetRow>> cpps, List<List<string>> efusebitDefRows)
        {
            List<string> funcont = new List<string>();

            SelCppMainFunc selCppMainFunc = new SelCppMainFunc(cpps.Select(cpp => cpp.First().SetupName).ToList());
            selCppMainFunc.Get_Select_CPP_Main_Func(funcont);

            foreach (List<PostProcessSheetRow> cpprows in cpps)
            {
                CppSetupFunc cppSetupFunc = new CppSetupFunc(cpprows);
                cppSetupFunc.Get_Cpp_Setup_Func(funcont, efusebitDefRows);

                if (funcont.Count() > 55000)
                {
                    content.Add(new List<string>(funcont));
                    funcont.Clear();
                }
            }

            if (funcont.Any())
            {
                content.Add(new List<string>(funcont));
            }
        }

    }
}
