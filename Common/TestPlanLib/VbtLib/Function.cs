using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.IgxlBase;

using TestPlanLib.Static;

namespace TestPlanLib.VbtLib
{
    [DebuggerDisplay("{FunctionName}")]
    public class Function
    {
        public const string TestAutoSwitchJtagTdi = "Test_AutoSwitch:JTAG_TDI";

        public bool IsInterPose { get; set; }
        public bool IsFound { get; set; }
        public string Block { get; set; } = string.Empty;
        public string FunctionName { get; set; } = string.Empty;
        public string FullFunctionName
        {
            get
            {
                return Type == ".NET" ? string.Format($"{NameSpace}.{FunctionName}") : FunctionName;
            }
        }
        public string Type { get; set; } = "VBT";
        public string NameSpace { get; set; } = string.Empty;
        public string Parameters { get; set; } = string.Empty;
        public string ParameterDefaults { get; set; } = string.Empty;
        public List<string> ArgList { get; set; } = [];
        public string FileName { get; set; } = string.Empty;
        public bool CheckParam { get; set; }
        public Dictionary<string, int> PatternDic { get; set; } = [];
        public List<string> UsedPatternList { get; set; } = [];

        public Function()
        {
            ArgList = [.. Enumerable.Repeat("", 100)];
            Parameters = "";
            FileName = "";
            CheckParam = true;
            Block = "";
            IsFound = true;
        }

        public Function(string functionName)
        {
            FunctionName = functionName;
            ArgList = [.. Enumerable.Repeat("", 100)];
            Parameters = "";
            FileName = "";
            CheckParam = true;
        }

        public Function(Function function)
        {
            if (function == null)
            {
                return;
            }

            IsInterPose = function.IsInterPose;
            IsFound = function.IsFound;
            Block = function.Block;
            FunctionName = function.FunctionName;
            Type = function.Type;
            NameSpace = function.NameSpace;
            Parameters = function.Parameters;
            ParameterDefaults = function.ParameterDefaults;
            FileName = function.FileName;
            CheckParam = function.CheckParam;

            ArgList = function.ArgList != null ? [.. function.ArgList] : [];
            UsedPatternList = function.UsedPatternList != null ? [.. function.UsedPatternList] : [];

            PatternDic = function.PatternDic != null ? new Dictionary<string, int>(function.PatternDic) : [];
        }

        public Function Copy()
        {
            return new Function(this);
        }

        public virtual void SetParamValue(string paramName, string paramValue)
        {
            int index = Parameters.Split(',').ToList().FindIndex(s => s.EqualsIgnoreCase(paramName));
            List<string>? aliasList = VbtFunctionLibShared.ParamMappingList.FirstOrDefault(s => s.Find(a => a.EqualsIgnoreCase(paramName)) != null);
            if (index != -1)
            {
                if (PatternDic.ContainsValue(index))
                {
                    VbtFunctionLibShared.UsedPatternList.Add(paramValue);
                }

                ArgList[index] = paramValue;
            }
            else if (aliasList != null)
            {
                foreach (string aliasName in aliasList)
                {
                    index = Parameters.Split(',').ToList().FindIndex(s => s.EqualsIgnoreCase(aliasName));
                    if (index != -1)
                    {
                        if (PatternDic.ContainsValue(index))
                        {
                            VbtFunctionLibShared.UsedPatternList.Add(paramValue);
                        }

                        ArgList[index] = paramValue;
                        break;
                    }
                }
            }
            else if (CheckParam)
            {
                if (IsFound)
                {
                    VbtFunctionLibShared.CheckMissingParameter(FunctionName, paramName, Block, Type);
                }
            }
        }

        public string GetParamValue(string paramName)
        {
            string paramValue = "";
            int index = Parameters.Split(',').ToList().FindIndex(s => s.EqualsIgnoreCase(paramName));
            List<string>? aliasList = VbtFunctionLibShared.ParamMappingList.FirstOrDefault(s => s.Find(a => a.EqualsIgnoreCase(paramName)) != null);
            if (index != -1)
            {
                paramValue = ArgList[index];
            }
            else if (aliasList != null)
            {
                foreach (string aliasName in aliasList)
                {
                    index = Parameters.Split(',').ToList().FindIndex(s => s.EqualsIgnoreCase(aliasName));
                    if (index != -1)
                    {
                        paramValue = ArgList[index];
                        break;
                    }
                }
            }

            return paramValue;
        }

        public void ReplaceFunction(InstanceRow instanceRow, Dictionary<string, string> dictionary)
        {
            instanceRow.VbtName = FullFunctionName;
            List<string> args = [.. instanceRow.ArgList.Split(',')];
            var dic = new Dictionary<string, string>();
            if (instanceRow.VbtType == "VBT")
            {
                for (int i = 0; i < args.Count; i++)
                {
                    dic.Add(args[i], instanceRow.Args[i]);
                }
                foreach (KeyValuePair<string, string> item in dic)
                {
                    SetParamValue(item.Key, item.Value);
                    if (dictionary.ContainsKey(item.Key.ToUpper()))
                    {
                        SetParamValue(dictionary[item.Key.ToUpper()], item.Value);
                    }
                }
            }
            instanceRow.ArgList = Parameters;
            instanceRow.Args = ArgList;
        }
    }
}
