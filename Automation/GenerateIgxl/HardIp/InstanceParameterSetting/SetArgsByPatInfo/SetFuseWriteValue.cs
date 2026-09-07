using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    internal class SetFuseWriteValue : SetValueBase
    {
        private List<string> _miscInfoList = new List<string>();

        public SetFuseWriteValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
            ReservedMiscInfoKeys = new List<string>
            {
                "FuseType",
                "m_catename",
                "Dict_Store_Code_Name",
                "Flag_Name",
                "Efuse_Binary_Write_Flag",
            };
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            _miscInfoList = pattern.MiscInfo.Split(';').ToList();
            if (function.Type == ".NET")
            {
                CsProcess(function);
            }
        }
        private void CsProcess(Function function)
        {
            string fuseType = GetMiscArgContent(_miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("FuseType", StringComparison.OrdinalIgnoreCase)) ?? "");
            string mCatename = GetMiscArgContent(_miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("m_catename", StringComparison.OrdinalIgnoreCase)) ?? "");
            string dictStoreCodeName = GetMiscArgContent(_miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("Dict_Store_Code_Name", StringComparison.OrdinalIgnoreCase)) ?? "");
            string flagName = GetMiscArgContent(_miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("Flag_Name", StringComparison.OrdinalIgnoreCase)) ?? "");
            string efuseBinaryWriteFlag = GetMiscArgContent(_miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("Efuse_Binary_Write_Flag", StringComparison.OrdinalIgnoreCase)) ?? "");

            function.SetParamValue("bankName", fuseType);
            function.SetParamValue("fieldName", mCatename);
            function.SetParamValue("dictionaryName", dictStoreCodeName);
            function.SetParamValue("flagName", flagName);
            function.SetParamValue("isBinToDecEnable", efuseBinaryWriteFlag);
        }
        private string GetMiscArgContent(string miscStr)
        {
            if (miscStr.Split(':').Length > 1)
            {
                return miscStr.Split(':')[1].Trim();
            }
            return "";
        }
    }
}
