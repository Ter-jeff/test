using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetFuseReadValue : SetValueBase
    {
        private List<string> _miscInfoList = new List<string>();
        public SetFuseReadValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
            ReservedMiscInfoKeys = new List<string>
            {
                "FuseType",
                "m_catename",
                "Dict_Store_Code_Name",
                "dspwavesize",
                "Efuse_Read_Dec_Flag",
                "NonZero_Val_Chk"
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
            string dspWaveSize = GetMiscArgContent(_miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("dspwavesize", StringComparison.OrdinalIgnoreCase)) ?? "");
            string efuseReadDecFlag = GetMiscArgContent(_miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("Efuse_Read_Dec_Flag", StringComparison.OrdinalIgnoreCase)) ?? "");
            string nonZeroValChk = GetMiscArgContent(_miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("NonZero_Val_Chk", StringComparison.OrdinalIgnoreCase)) ?? "");

            function.SetParamValue("bankName", fuseType);
            function.SetParamValue("fieldName", mCatename);
            function.SetParamValue("dictionaryName", dictStoreCodeName);
            function.SetParamValue("sampleSize", dspWaveSize);
            function.SetParamValue("enableNonZeroCheck", nonZeroValChk);

            if (efuseReadDecFlag.ToLower() == "true")
            {
                function.SetParamValue("isDecToBinEnable", "false");
            }

            if (efuseReadDecFlag.ToLower() == "false")
            {
                function.SetParamValue("isDecToBinEnable", "true");
            }
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
