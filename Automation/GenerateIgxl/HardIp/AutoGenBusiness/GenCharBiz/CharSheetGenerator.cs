using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;

using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenCharBiz
{
    public class CharSheetGenerator
    {
        private CharSheet _charSheet;
        private const string CharSheetName = "Char_HardIP";

        public CharSheet GenCharSheet(Dictionary<string, HardIpSheet> planDic)
        {
            _charSheet = new CharSheet(CharSheetName);
            foreach (string sheetName in planDic.Keys)
            {
                var charGenerator = new CharRowGenerator();
                _charSheet.Rows.AddRange(charGenerator.GenCharRow(planDic[sheetName].Rows));
            }
            return _charSheet;
        }
    }
}
