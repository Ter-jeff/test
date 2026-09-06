using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using TestPlanLib.Static;

namespace TestPlanLib.DataStruct
{
    public class IoLevelsSheet(string sheetname)
    {
        #region Properity
        public string SheetName { get; set; } = sheetname;
        public List<IoLevelsRow> RowList { get; } = [];
        public Dictionary<string, int> HeaderIndex { get; } = [];

        #endregion
        #region Constructor
        #endregion

        #region IO_Conti
        public IoContiSheet ConvertIoContiSheet()
        {
            IoContiSheet ioContiSheet = new IoContiSheet();
            foreach (IoLevelsRow row in RowList)
            {
                IoContiRow ioContiRow = new IoContiRow
                {
                    BallName = row.PinName
                };
                if (row.IsTheSameRow)
                {
                    ioContiRow.IoVoltage = row.IoLevelDate[0].Vdd;
                }

                if (!string.IsNullOrEmpty(row.Fsdd))
                {
                    ioContiRow.FsDd = row.Fsdd;
                }
                else
                {
                    ioContiRow.FsDd = "DD";
                }

                ioContiSheet.AddRow(ioContiRow);
            }
            return ioContiSheet;
        }
        #endregion

        public static void CheckIoPins(IoContiSheet ioContiSheet, PinMapSheet pinMapSheet)
        {
            //List<Pin> pins = mapSheet.GetIoContinuityPins();
            List<Pin> pins = pinMapSheet.GetIoPins();
            foreach (Pin pin in pins)
            {
                if (!ioContiSheet.GetPinList().Exists(x => x.EqualsIgnoreCase(pin.PinName)))
                {
                    string errorMessage = $"The pin {pin.PinName} can not be found in sheet {NeededSheets.ContiIo}";
                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_12, NeededSheets.ContiIo, 0, 0, $"The pin {pin.PinName} can not be found in sheet {NeededSheets.ContiIo}", [pin.PinName, NeededSheets.ContiIo]);
                }
            }
        }

        public List<GlobalSpec> GenGlbSymbol()
        {
            List<GlobalSpec> globalSpecs = [];
            foreach (IoLevelsRow row in RowList)
            {
                string domain = row.Domain;
                if (string.IsNullOrEmpty(domain))
                {
                    continue;
                }

                List<GlobalSpec> others = [];
                List<Tuple<string, string>> othersText = [];
                foreach (IoLevelsItem data in row.IoLevelDate)
                {
                    string level = data.Level;
                    string glbSymbol1 = SpecFormat.GenGlbSpecSymbol(domain);
                    string value1 = SpecFormat.GenSpecValueSingleValue(data.Vdd);
                    if (!globalSpecs.Any(x => x.Symbol.EqualsIgnoreCase(glbSymbol1)))
                    {
                        globalSpecs.Add(new GlobalSpec(glbSymbol1, value1));
                    }
                    else
                    {
                        value1 = globalSpecs.Find(x => x.Symbol == glbSymbol1)!.Value;
                    }

                    string glbSymbol2 = SpecFormat.GenGlbSpecSymbol(data.Domain + "_" + level);
                    _ = double.TryParse(value1.Replace("=", ""), out double v1);
                    _ = double.TryParse(data.Vdd, out double v2);
                    double ratio = v1 == 0 ? 0 : v2 / v1;
                    othersText.Add(new Tuple<string, string>(data.Level, data.Vdd));
                    string value2 = SpecFormat.GenSpecValueSingleValue(SpecFormat.SpecValuePrefix + glbSymbol1 + "*" + ratio);
                    if (!globalSpecs.Any(x => x.Symbol == glbSymbol2 && x.Value == value2))
                    {
                        others.Add(new GlobalSpec(glbSymbol2, value2));
                    }
                }

                foreach (IGrouping<string, Tuple<string, string>> group in othersText.GroupBy(x => x.Item1))
                {
                    if (group.Select(x => x.Item2).Distinct().Count() != 1)
                    {
                        foreach (GlobalSpec other in others)
                        {
                            if (!globalSpecs.Any(x => x.Symbol == other.Symbol && x.Value == other.Value))
                            {
                                globalSpecs.Add(other);
                            }
                        }
                    }
                }
            }
            return globalSpecs;
        }

        public List<GlobalSpec> NewGenGlbSymbol()
        {
            List<GlobalSpec> globalSpecs = [];
            foreach (IoLevelsRow row in RowList)
            {
                //Default
                string domain = row.IsGroupPin ? row.Domain : row.PinName;
                {
                    IoLevelsItem data = row.IoLevelDate[0];
                    string level = data.Level;
                    string vdd = data.Vdd;
                    GetGlobalSpec(vdd, domain, level, ref globalSpecs, data, true);
                }

                if (!row.IsTheSameRow)
                {
                    foreach (IoLevelsItem data in row.IoLevelDate)
                    {
                        string level = data.Level;
                        string vdd = data.Vdd;
                        GetGlobalSpec(vdd, domain, level, ref globalSpecs, data, false);
                    }
                }
            }
            return globalSpecs;
        }

        private static void GetGlobalSpec(string vdd, string domain, string level, ref List<GlobalSpec> globalSpecs, IoLevelsItem ioLevelsItem, bool isTheSameDomain)
        {
            GlobalSpec globalSpec = IoLevelsRow.GetGlobalSpec(vdd, vdd, domain, level, isTheSameDomain);
            if (!globalSpecs.Any(x => x.Symbol.EqualsIgnoreCase(globalSpec.Symbol)))
            {
                globalSpecs.Add(globalSpec);
            }

            string vih = GetFactoer(ioLevelsItem.Vih, ioLevelsItem.Domain);
            GlobalSpec globalSpecVih = IoLevelsRow.GetGlobalSpec(vdd, vih, domain, level, isTheSameDomain, "VIH");
            if (!string.IsNullOrEmpty(vih) &&
                !globalSpecs.Any(x => x.Symbol.EqualsIgnoreCase(globalSpecVih.Symbol)))
            {
                globalSpecs.Add(globalSpecVih);
            }

            string vil = GetFactoer(ioLevelsItem.Vil, ioLevelsItem.Domain);
            GlobalSpec globalSpecVil = IoLevelsRow.GetGlobalSpec(vdd, vil, domain, level, isTheSameDomain, "VIL");
            if (!string.IsNullOrEmpty(vil) &&
                !globalSpecs.Any(x => x.Symbol.EqualsIgnoreCase(globalSpecVil.Symbol)))
            {
                globalSpecs.Add(globalSpecVil);
            }

            string voh = GetFactoer(ioLevelsItem.Voh, ioLevelsItem.Domain);
            GlobalSpec globalSpecVoh = IoLevelsRow.GetGlobalSpec(vdd, voh, domain, level, isTheSameDomain, "VOH");
            if (!string.IsNullOrEmpty(voh) &&
                !globalSpecs.Any(x => x.Symbol.EqualsIgnoreCase(globalSpecVoh.Symbol)))
            {
                globalSpecs.Add(globalSpecVoh);
            }

            string vol = GetFactoer(ioLevelsItem.Vol, ioLevelsItem.Domain);
            GlobalSpec globalSpecVol = IoLevelsRow.GetGlobalSpec(vdd, vol, domain, level, isTheSameDomain, "VOL");
            if (!string.IsNullOrEmpty(vol) &&
                !globalSpecs.Any(x => x.Symbol.EqualsIgnoreCase(globalSpecVol.Symbol)))
            {
                globalSpecs.Add(globalSpecVol);
            }
        }

        private static string GetFactoer(string value, string domain)
        {
            return !value.Contains('*') ? "" : value.Replace(domain, "").Replace("*", "").Trim();
        }
    }
}
