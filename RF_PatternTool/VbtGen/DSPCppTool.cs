using System.Text.RegularExpressions;

using RfLib.Dvdc.Reader.CapturePostProcess;

namespace RF_PatternTool.VbtGen
{
    public class CppSetupDSPFunc
    {
        public class ResultDSPWave
        {
            public string RawDec = "";
            public string NewDSPWave = "";
            public string CalcFunc = "";
            public List<string> CalcArgs = new List<string>();

            public ResultDSPWave(string rawdecwave)
            {
                RawDec = rawdecwave;
            }
        }

        public class CalcEquation
        {
            public string CalcFunName = "";
            public List<string> CalcFunArgs = new List<string>();
            public List<string> CalcPrinInfos = new List<string>();

            public CalcEquation(string funname)
            {
                CalcFunName = funname;
            }
        }

        private string _dSPFuncName = "";

        private List<ResultDSPWave> _resultDspWaves = new List<ResultDSPWave>();
        //StoreName : "RawDec"
        private Dictionary<string, string> _mapCalcEquName = new Dictionary<string, string>();
        //CalcStoreName : "LocalCalcWave"
        private Dictionary<string, string> _mapCalcStorName = new Dictionary<string, string>();

        private HashSet<string> _hashValFromMain = new HashSet<string>();

        private List<CalcEquation> _allCalcEqus = new List<CalcEquation>();

        private int _storeNum = 1;

        //private int eachOneBit = 0;
        private List<int> _eachOneBit = new List<int>();

        private bool _isCalcBestCodeCPP = false;

        private static readonly Regex _regexNum = new Regex(@"^-?\d+(\.\d+)?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);


        public CppSetupDSPFunc(List<PostProcessSheetRow> cpprows)
        {
            int inputInd = 1;
            int tempInputWaveInd = 1;
            int locCalcWaveInd = 1;
            int mapCalcStorNameInd = 0;

            _dSPFuncName = cpprows.First().SetupName + "_DSP";

            foreach (PostProcessSheetRow row in cpprows)
            {
                if (row.StoreName.Any())
                {
                    if (int.TryParse(row.BitWidth, out int onebit))
                    {
                        _eachOneBit.Add(onebit);
                    }
                    _resultDspWaves.Add(new ResultDSPWave("RawDec" + _storeNum));
                    _mapCalcEquName.Add(row.StoreName, "RawDec" + _storeNum);
                    _storeNum++;
                }

                if (row.CalcEquation.Any())
                {
                    int firstInd = row.CalcEquation.IndexOf('(');
                    string calcFuncName = row.CalcEquation.Substring(0, firstInd);
                    if (calcFuncName == "Calc_BestCode_CPP")
                    {
                        _isCalcBestCodeCPP = true;
                        continue;
                    }
                    if (calcFuncName == "rffw_calc_RAW")
                    {
                        continue;
                    }

                    var calcFunc = new CalcEquation(calcFuncName + "_DSP");

                    string[] calcArgs = row.CalcEquation.Substring(firstInd + 1, row.CalcEquation.Length - 1 - firstInd - 1).Split(',');

                    if (calcFunc.CalcFunName == "hw_acc_avg_DSP" ||
                        calcFunc.CalcFunName == "rsb_best_delta_gain_value_DSP" ||
                        calcFunc.CalcFunName == "rsb_best_delta_gain_code_DSP")
                    {
                        calcFunc.CalcFunArgs.Add("tempInputWave" + tempInputWaveInd);

                        calcFunc.CalcPrinInfos.Add($"\tDim tempInputWave{tempInputWaveInd} As New DSPWave");
                        calcFunc.CalcPrinInfos.Add($"\ttempInputWave{tempInputWaveInd}.CreateConstant 999, {calcArgs.Length}, DspDouble");

                        for (int i = 0; i < calcArgs.Length; i++)
                        {
                            if (_mapCalcEquName.ContainsKey(calcArgs[i]))
                            {
                                calcFunc.CalcPrinInfos.Add($"\ttempInputWave{tempInputWaveInd}.Element({i}) = {_mapCalcEquName[calcArgs[i]]}");
                            }

                            if (_mapCalcStorName.ContainsKey(calcArgs[i]))
                            {
                                calcFunc.CalcPrinInfos.Add($"\ttempInputWave{tempInputWaveInd}.Element({i}) = {_mapCalcStorName[calcArgs[i]]}.Element(0)");
                            }
                        }

                        tempInputWaveInd++;
                    }
                    else
                    {
                        foreach (string arg in calcArgs)
                        {

                            if (_mapCalcStorName.ContainsKey(arg))
                            {
                                string argInput = "Input" + inputInd;
                                calcFunc.CalcPrinInfos.Add($"\tDim {argInput} As Double");
                                calcFunc.CalcPrinInfos.Add($"\t{argInput} = {_mapCalcStorName[arg]}.Element(0)");

                                calcFunc.CalcFunArgs.Add(argInput);
                                inputInd++;
                            }
                            else if (_mapCalcEquName.ContainsKey(arg))
                            {
                                calcFunc.CalcFunArgs.Add(_mapCalcEquName[arg]);
                            }
                            else
                            {
                                if (!_regexNum.IsMatch(arg))
                                {
                                    _hashValFromMain.Add($"ByVal {arg} As Double");
                                }

                                calcFunc.CalcFunArgs.Add(arg);
                            }
                        }
                    }

                    calcFunc.CalcFunArgs.Add("LocalCalcWave" + locCalcWaveInd);
                    calcFunc.CalcPrinInfos.Add($"\tDim LocalCalcWave{locCalcWaveInd} As New DSPWave");
                    calcFunc.CalcPrinInfos.Add($"\tCall {calcFunc.CalcFunName}({string.Join(", ", calcFunc.CalcFunArgs)})");

                    if (!string.IsNullOrEmpty(row.CalcStoreName))
                    {
                        locCalcWaveInd++;
                        mapCalcStorNameInd++;
                    }

                    _allCalcEqus.Add(calcFunc);
                }

                if (!string.IsNullOrEmpty(row.CalcStoreName))
                {
                    _mapCalcStorName.Add(row.CalcStoreName, "LocalCalcWave" + mapCalcStorNameInd);
                }
            }

        }

        public void GetCppSetupDSPFunc(List<string> content)
        {
            Header(content);
            SetDefVar(content);
            SetCheckInputWaveSplit(content);
            SetGettingDecWave(content);

            if (_allCalcEqus.Any())
            {
                SetAssignSiteDouble(content);
                SetRunCalcFunction(content);
                SetAssignToOutCalcWave(content);
            }

            Footer(content);
        }

        private void Header(List<string> content)
        {
            if (_isCalcBestCodeCPP)
            {
                content.Add($"Public Function {_dSPFuncName}(ByVal InputWave As DSPWave, ByRef rawDecWave As DSPWave, ByRef OutCalcWave As DSPWave, " +
                    string.Join(", ", _hashValFromMain) + (_hashValFromMain.Any() ? ", " : "") +
                    $"ByVal InRange1 As DSPWave, ByVal InRange2 As DSPWave, ByVal inTarget As DSPWave, ByVal inAlgo As Long, ByRef outBCwave As DSPWave, ByRef outBV1 As Double, ByRef outBV2 As Double, ByRef outSweepWave As DSPWave) As Long");
            }
            else
            {
                content.Add($"Public Function {_dSPFuncName}(InputWave As DSPWave, rawDecWave As DSPWave, OutCalcWave As DSPWave" +
                    (_hashValFromMain.Any() ? ", " : "") + string.Join(", ", _hashValFromMain) + ") As Long");
            }
            content.Add("");
        }

        private static void SetDefVar(List<string> content)
        {
            content.Add("\tDim bitWidthWave As New DSPWave");
            content.Add("");
        }

        private void SetCheckInputWaveSplit(List<string> content)
        {
            content.Add($"\tbitWidthWave.CreateConstant 999, {_storeNum - 1}{(_isCalcBestCodeCPP ? "" : ", DspDouble")}");
            content.Add("");
            if (_isCalcBestCodeCPP)
            {
                content.Add("\tDim i as long");
                content.Add($"\tFor i = 0 to {_storeNum - 1}");
                content.Add("\t\tbitWidthWave.Element(i) = 32");
                content.Add("\tNext");
            }
            else
            {
                for (int i = 0; i < _eachOneBit.Count(); i++)
                {
                    content.Add($"\tbitWidthWave.Element({i}) = {_eachOneBit[i]}");
                }
            }

            content.Add("");
        }

        private static void SetGettingDecWave(List<string> content)
        {
            content.Add("\tCall Split_Dspwave(InputWave, bitWidthWave, rawDecWave)");
            content.Add("");
        }

        private void SetAssignSiteDouble(List<string> content)
        {
            if (_isCalcBestCodeCPP)
            {
                content.Add("\tDim TempRawDec as double");
                content.Add("\tDim TempLocalCalcWave as new dspwave");
                content.Add("");
                content.Add($"\tFor i = 0 to {_storeNum - 1}");
                content.Add("\t\tTempRawDec = rawDecWave.Element(i)");
                content.Add("\t\tCall uclkfreq_DSP(TempRawDec, TempLocalCalcWave)");
                content.Add("\t\tOutCalcWave.Element(i) = TempLocalCalcWave.Element(0)");
                content.Add("\tNext");
            }
            else
            {
                for (int i = 1; i < _storeNum; i++)
                {
                    content.Add($"\tDim RawDec{i} As Double");
                }

                for (int i = 1; i < _storeNum; i++)
                {
                    content.Add($"\tRawDec{i} = rawDecWave.Element({i - 1})");
                }
            }

            content.Add("");
        }

        private void SetRunCalcFunction(List<string> content)
        {
            if (_isCalcBestCodeCPP)
            {
                ;
            }
            else
            {
                for (int i = 0; i < _allCalcEqus.Count(); i++)
                {
                    foreach (string definfo in _allCalcEqus[i].CalcPrinInfos)
                    {
                        content.Add(definfo);
                    }

                    content.Add("");
                }
            }
        }

        private void SetAssignToOutCalcWave(List<string> content)
        {
            if (_isCalcBestCodeCPP)
            {
                content.Add("\tDim BestCodeWave As New DSPWave");
                content.Add("");
                content.Add("\tBestCodeWave = OutCalcWave.copy");
                content.Add("");
                content.Add("\tCall GetBestCodeAndBestValue(inAlgo, BestCodeWave, inTarget, InRange1, InRange2, outSweepWave, outBCwave, outBV1, outBV2)");
            }
            else
            {
                int localCalInd = 0;
                string localCalStr = _mapCalcStorName.First().Value;
                for (int i = 0; i < _mapCalcStorName.Count(); i++)
                {
                    if (localCalStr != _mapCalcStorName.ElementAt(i).Value)
                    {
                        localCalStr = _mapCalcStorName.ElementAt(i).Value;
                        localCalInd = 0;
                    }
                    content.Add($"\tOutCalcWave.Element({i}) = {localCalStr}.Element({localCalInd})");
                    localCalInd++;
                }
            }

            content.Add("");
        }

        private static void Footer(List<string> content)
        {
            content.Add("End Function");
            content.Add("");
        }

    }

    public class DSPCppTool
    {

        public static void AddHeader(List<string> content, string vbtname)
        {
            content.Insert(0, $"Attribute VB_Name = \"{vbtname}\"");
            content.Insert(1, "Option Explicit");
            content.Insert(2, "");
        }

        public static void Write(List<List<string>> content, List<List<PostProcessSheetRow>> cpps)
        {
            List<string> toolcont = new List<string>();

            foreach (List<PostProcessSheetRow> cpprows in cpps)
            {
                CppSetupDSPFunc cppSetupDspFunc = new CppSetupDSPFunc(cpprows);
                cppSetupDspFunc.GetCppSetupDSPFunc(toolcont);

                if (toolcont.Count() > 55000)
                {
                    content.Add(new List<string>(toolcont));
                    toolcont.Clear();
                }
            }

            if (toolcont.Any())
            {
                content.Add(new List<string>(toolcont));
            }
        }

    }
}
