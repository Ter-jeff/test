using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.InstCommon
{
    public class InstCommonOutput
    {
        private const string ConCommonInstSheet = "TestInst_Common";
        private const string SetPpmuClamp = "Set_PPMU_Clamp";

        public List<InstanceRow> InstanceRows { get; set; } = new List<InstanceRow>();

        public void AddDataToLocalSpec()
        {
            if (!TestProgram.IgxlWorkBk.TryGetTestInstCommon(out InstanceSheet instanceSheet))
            {
                instanceSheet = new InstanceSheet(ConCommonInstSheet, "", true);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, instanceSheet);
            }

            instanceSheet.AddRows(SetPpmuClampPara(InstanceRows));
        }

        internal List<InstanceRow> SetPpmuClampPara(List<InstanceRow> instanceRows)
        {
            var result = new List<InstanceRow>();
            foreach (InstanceRow instanceRow in instanceRows)
            {
                if (instanceRow == null)
                {
                    continue;
                }

                if (instanceRow.VbtName.Equals(SetPpmuClamp, StringComparison.OrdinalIgnoreCase))
                {
                    //Set PPMU Clamp parameters
                    if (!instanceRow.TestName.ContainsIgnoreCase("conti"))
                    {
                        SetPpmuClampPara(instanceRow);
                    }
                }
                result.Add(instanceRow);
            }
            return result;
        }

        /// <summary>
        /// Set PPMU Clamp according to Power table IO Pin groups
        /// Example: Pins_1p1v  1.1
        ///          Pins_1p8v  1.8
        /// </summary>
        /// <param name="row"></param>
        private void SetPpmuClampPara(InstanceRow row)
        {
            row.Args.Clear();
            var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value.GroupList.Where(p => p.PinType == PinMapConst.TypeIo && Regex.IsMatch(p.PinName, @"^Pins_\d{1}p\d+v$", RegexOptions.IgnoreCase)).ToList();
            foreach (PinGroup pin in pinMap)
            {
                string pinVal = pin.PinName.Replace("Pins_", "").Replace("p", ".").Replace("v", "");
                double.TryParse(pinVal, out double outPinVal);
                row.Args.Add(pin.PinName);
                row.Args.Add((outPinVal * 1.2).ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}
