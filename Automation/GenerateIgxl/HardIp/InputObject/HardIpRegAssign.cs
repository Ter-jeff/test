using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class HardIpRegAssign
    {
        public RegisterAssignType Type { get; set; }
        public string SubBlockName { get; set; }

        public string GetSheetType(List<string> sheetTypeList)
        {
            if (IseFuseReadWriteType())
            {
                return "_eFuse";
            }

            if (sheetTypeList.Count == 0)
            {
                return "";
            }

            string regBlockType = $"(?<block>{string.Join("|", sheetTypeList)})";
            if (Regex.IsMatch(SubBlockName.Split('_')[0], regBlockType, RegexOptions.IgnoreCase))
            {
                return "_" + Regex.Match(SubBlockName.Split('_')[0], regBlockType, RegexOptions.IgnoreCase).Groups["block"].Value;
            }

            return "";
        }

        private List<string> _headers;
        public List<string> Headers
        {
            get
            {
                switch (Type)
                {
                    case RegisterAssignType.DigSrc_Equation:
                        return new List<string> { "Equation" };
                    case RegisterAssignType.DigSrc_Assignment:
                        return new List<string> { "Equation", "Register Name" };
                    case RegisterAssignType.CUS_Str_DigCapData:
                        return new List<string> { "Bits", "SgmName", "DicName" };
                    case RegisterAssignType.DigSrc_FlowForLoopIntegerName:
                        return new List<string> { "SrcName", "SgmName" };
                    //case RegisterAssignType.SweepTable:
                    //    return _headers;
                    case RegisterAssignType.Calc_Eqn:
                        return new List<string> { "Parametric0", "Parametric1", "Parametric2", "Parametric3" };
                    case RegisterAssignType.MeasPins:
                    case RegisterAssignType.MeasV_PinS:
                    case RegisterAssignType.MeasI_PinS:
                    case RegisterAssignType.MeasF_PinS_SingleEnd:
                    case RegisterAssignType.Interpose_PreMeas:
                    case RegisterAssignType.Meas_StoreName:
                    case RegisterAssignType.TestSequence:
                    case RegisterAssignType.CalcEquName:
                    case RegisterAssignType.CalcStoreName:
                    case RegisterAssignType.MeasI_Range:
                    case RegisterAssignType.ForceV_Val:
                    case RegisterAssignType.ForceI_Val:
                        int max = RegAssignList.Max(x => x.Count);
                        var headers = new List<string> { "Seq" };
                        for (int i = 0; i < max - 1; i++)
                        {
                            headers.Add("Parametric" + i);
                        }

                        return headers;
                    case RegisterAssignType.CUS_Str_MainProgram:
                    case RegisterAssignType.TrimDictionaryStoreName:
                        return new List<string> { "Parametric0" };
                    case RegisterAssignType.HIP_eFuse_Read:
                    case RegisterAssignType.HIP_eFuse_Write:
                        return new List<string> { "m_catename", "Dict_Store_Code_Name" };
                    default:
                        return _headers;
                }
            }
            set { _headers = value; }
        }

        public List<List<string>> RegAssignList = new List<List<string>>();

        public bool IseFuseReadWriteType()
        {
            return Type == RegisterAssignType.HIP_eFuse_Read || Type == RegisterAssignType.HIP_eFuse_Write;
        }
    }
}
