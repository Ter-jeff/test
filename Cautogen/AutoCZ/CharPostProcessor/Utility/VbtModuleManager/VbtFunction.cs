using System;
using System.Collections.Generic;
using System.Linq;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager
{
    public class VbtFunction
    {
        public string FileName { get; set; }
        public string FunctionName { get; set; }
        public string Parameters { get; set; }
        public List<string> ArgList { get; set; }
        public string Type { get; set; } = "VBT";
        public string NameSpace { get; set; }

        public string FullFunctionName
        {
            get
            {
                if (Type == ".NET")
                {
                    return string.Format($"{NameSpace}.{FunctionName}");
                }

                return FunctionName;
            }
        }

        public VbtFunction()
        {
            ArgList = Enumerable.Repeat("", 100).ToList();
            Parameters = "";
            FileName = "";
        }

        public VbtFunction(string functionName)
        {
            FunctionName = functionName;
            ArgList = Enumerable.Repeat("", 100).ToList();
            Parameters = "";
            FileName = "";
        }

        public void SetParamValue(string paramName, string paramValue, bool checkMissing = true)
        {
            int index = Parameters.Split(',').ToList().FindIndex(s => s.Equals(paramName, StringComparison.OrdinalIgnoreCase));//.IndexOf(paramName);
            if (index != -1)
            {
                ArgList[index] = paramValue;
            }
            else if (checkMissing)
            {
                VbtFunctionLib.CheckMissingParamter(FunctionName, paramName);
            }
        }
    }
}
