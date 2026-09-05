using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Utility.HardIP;

using IgxlLib.IgxlBase;

using TestPlanLib.BinNumber;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTableRow
{
    public abstract class BinTableRowGeneratorBase
    {
        public string SheetName;
        public string BlockName = string.Empty;
        public string SubBlockName = string.Empty;
        public List<string> ErrorBinNums;
        public string TimingAc = string.Empty;
        public string InstNameSubStr = string.Empty;
        public bool NoPattern;
        public BinNumResult BinLib;
        public HardIpPattern Pattern;
        public bool IsHipEfuseRead;

        protected BinTableRowGeneratorBase(string sheetName, List<string> errorBinNums)
        {
            SheetName = sheetName;
            ErrorBinNums = errorBinNums;
        }

        #region Main Methods
        public BinTableRow GenBinTableRow(HardIpPattern pattern, string voltage = "")
        {
            SetPattern(pattern);
            return GenBinTableRowForPattern(voltage);
        }
        #endregion

        #region Abstract Methods
        protected abstract BinTableRow GenBinTableRowForPattern(string voltage = "");

        protected abstract void SetPattern(HardIpPattern pattern);
        #endregion

        #region Create columns creater methods
        protected string CreateSortBin()
        {
            return BinLib.SoftBin.ToString("G15");
        }

        protected string CreateHardBin()
        {
            return BinLib.BinNumInfo.HardBin.ToString("G15");
        }

        protected string CreateResult()
        {
            return BinLib.BinNumInfo.Status;
        }
        #endregion

        #region Other Methods        

        protected void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            IsHipEfuseRead = false;
            SubBlockName = CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, BlockName);
            Pattern = pattern;
            BlockName = Regex.Replace(CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName), "wireless_|lcd_", "", RegexOptions.IgnoreCase);
            TimingAc = CommonGenerator.GetTimingAc(Pattern);
            InstNameSubStr = HardIpService.GetInstNameSubStr(Pattern.MiscInfo);
            NoPattern = HardIpService.IsNoPattern(pattern.Pattern.RealPatternName);
            BinLib = SearchInfo.IsHardipRtosSheet(pattern.SheetName) ? SearchInfo.GetRtosBin(pattern) : SearchInfo.GetHardIpBin(pattern);
            if (VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)))
            {
                return;
            }

            if (VbtFunctionLibShared.EfuseReadFunctionList.Exists(f => f.Equals(pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)))
            {
                IsHipEfuseRead = true;
            }
        }

        #endregion
    }
}
