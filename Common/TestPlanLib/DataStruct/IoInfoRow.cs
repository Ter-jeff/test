using System;
using System.Text.RegularExpressions;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;
using LogLib.Utility;

namespace TestPlanLib.DataStruct
{
    public partial class IoInfoRow
    {
        private const double ErrorValue = -54321;

        [GeneratedRegex("[a-z]", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"(?<value>\d+([.]\d+)?)\s*(?<unit>\w*)")]
        private static partial Regex MyRegex1();
        private static readonly Regex _regex = MyRegex();

        private string _driverMode = "Hiz";
        public bool IsMerged = false;
        public bool IsCustom = false;
        public bool NeedToGenLevel
        {
            get
            {
                return !string.IsNullOrEmpty(PinGrpName) &&
                        !string.IsNullOrEmpty(Nv) &&
                        !string.IsNullOrEmpty(Hv) &&
                        !string.IsNullOrEmpty(Lv) &&
                        !string.IsNullOrEmpty(Vil) &&
                        !string.IsNullOrEmpty(Vih) &&
                        !string.IsNullOrEmpty(Vol) &&
                        !string.IsNullOrEmpty(Voh);
            }
        }
        public string PinGrpName { get; set; } = string.Empty;
        public string Nv { get; set; } = string.Empty;
        public string Hv { get; set; } = string.Empty;
        public string Lv { get; set; } = string.Empty;
        public string Vil { get; set; } = string.Empty;
        public string Vih { get; set; } = string.Empty;
        public string Vol { get; set; } = string.Empty;
        public string Voh { get; set; } = string.Empty;
        public string Iol { get; set; } = string.Empty;
        public string Ioh { get; set; } = string.Empty;
        public string Vt { get; set; } = string.Empty;
        public string Vch { get; set; } = string.Empty;
        public string Vcl { get; set; } = string.Empty;
        public string DriverMode
        {
            get
            {
                return _driverMode;
            }
            set
            {
                string input = value.Trim();
                if (string.IsNullOrEmpty(value))
                {
                    return;
                }

                _driverMode = input;
            }
        }
        public string Type { get; set; } = string.Empty;
        public string TimeDomain { get; set; } = string.Empty;
        public string VilRatio
        {
            get
            {
                string ratio = IsUsePowerPinValue ? GetRatioFromExpress(Vil) : CalculateDivision(Vil, Nv);
                return ratio;
            }
        }
        public string VihRatio
        {
            get
            {
                string ratio = IsUsePowerPinValue ? GetRatioFromExpress(Vih) : CalculateDivision(Vih, Nv);
                return ratio;
            }
        }
        public string VolRatio
        {
            get
            {
                string ratio = IsUsePowerPinValue ? GetRatioFromExpress(Vol) : CalculateDivision(Vol, Nv);
                return ratio;
            }
        }
        public string VohRatio
        {
            get
            {
                string ratio = IsUsePowerPinValue ? GetRatioFromExpress(Voh) : CalculateDivision(Voh, Nv);
                return ratio;
            }
        }
        public string VtRatio
        {
            get
            {
                string ratio = IsUsePowerPinValue ? GetRatioFromExpress(Vt) : CalculateDivision(Vt, Nv);
                return ratio;
            }
        }
        public string HvRatio
        {
            get
            {
                string ratio = IsUsePowerPinValue ? "0" : CalculateDivision(Hv, Nv);
                return ratio;
            }
        }
        public string LvRatio
        {
            get
            {
                string ratio = IsUsePowerPinValue ? "0" : CalculateDivision(Lv, Nv);
                return ratio;
            }
        }

        public bool IsUsePowerPinValue
        {
            get
            {
                return _regex.IsMatch(Nv);
            }
        }

        /// <summary>
        /// Read pinGroup voltage from ioconti sheet
        /// </summary>
        /// <param name="pinGrp"></param>
        /// <returns></returns>
        public static string GetIoPinGroupVoltage(string pinGrp, PinMapSheet pinMapSheet, IoContiSheet ioContiSheet)
        {
            PinGroup? group = pinMapSheet.GetGroup(pinGrp);
            if (group == null)
            {
                Response.Report($"'{pinGrp}' used in IoInfo sheet, but missing in IO_PinMap sheet!", EnumMessageLevel.Error, 45);
                return "";
            }

            string firstPin = group.PinList[0].PinName;
            string volatge = ioContiSheet.GetVoltage(firstPin);
            if (string.IsNullOrEmpty(volatge))
            {
                Response.Report($"Can not find IO pin voltage from IO_Continuty sheet for pin group '{pinGrp}' used in IoInfo!", EnumMessageLevel.Error, 45);
                return "";
            }
            string value = MyRegex1().Match(volatge).Groups["value"].ToString();
            string unit = MyRegex1().Match(volatge).Groups["unit"].ToString();
            return string.IsNullOrEmpty(unit) ? value : volatge.ConvertNumber();
        }

        private static string GetRatioFromExpress(string str)
        {
            string ratio;
            if (string.IsNullOrEmpty(str) || str.Trim() == "0")
            {
                ratio = "0";
            }
            else if (!str.Contains('*'))
            {
                ratio = "1";
            }
            else
            {
                ratio = str.Split('*')[1].Trim();

                if (!double.TryParse(ratio, out _))
                {
                    throw new Exception("Can not recognize ratio from IoInfo specified value: " + str);
                }
            }
            return ratio;
        }

        private static string CalculateDivision(string dividend, string divisor)
        {
            double answer;
            try
            {
                if (divisor == "0")
                {
                    return "0";
                }
                double num1 = double.Parse(dividend);
                double num2 = double.Parse(divisor);
                answer = Math.Round(num1 / num2, 3);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                answer = ErrorValue;
            }
            return answer.ToString();
        }
    }
}
