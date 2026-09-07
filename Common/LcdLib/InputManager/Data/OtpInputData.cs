using System.Collections.Generic;

using Automation.InputManager.Data;

using LcdLib.OTP.Reader;

namespace LcdLib.InputManager.Data
{
    public class OtpInputData : InputDataBase
    {
        public List<OtpPatternRow> OtpPatternRows { get; set; } = [];
    }
}
