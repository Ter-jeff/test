using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace BinCutScriptLib.Reader
{
    public class TableGenertor
    {
        protected const string VopVar = "_VOP_VAR";
        protected const string Var = "_VAR";

        protected ExcelPackage ExcelPackage;
        protected ExcelWorksheet ExcelWorksheet;

        public TableGenertor()
        {
            ExcelPackage = new ExcelPackage(new FileInfo("Test"));
            ExcelWorksheet = ExcelPackage.Workbook.Worksheets.Add("Test");
        }

        public void CreateTable(string tempFolder, List<string> titles, InstanceSheet instanceSheet, List<Tuple<string, string>> dcCategorys, DcSpecSheet dcSpecSheet, GlobalSpecSheet globalSpecSheet, out VoltageTable vrsTable, out VoltageTable nvTable)
        {
            vrsTable = new VoltageTable(instanceSheet);
            nvTable = new VoltageTable(instanceSheet);

            if (instanceSheet == null || dcSpecSheet == null || globalSpecSheet == null)
            {
                return;
            }

            foreach (Tuple<string, string> dcCategory in dcCategorys)
            {
                var vrsPins = new List<string>();
                var nvPins = new List<string>();
                foreach (string pin in titles)
                {
                    GetPinNameFromDcSpec(pin, dcSpecSheet, out string vrsPinName, out string nvPinName);
                    GetDCspecValue(vrsPinName, nvPinName, dcSpecSheet, dcCategory.Item1, dcCategory.Item2, out string vrsValue, out string nvValue);

                    if (string.IsNullOrEmpty(vrsValue))
                    {
                        vrsPins.Add(pin + "= NA");
                    }
                    else
                    {
                        vrsPins.Add(pin + "=" + GetFomulaValue(vrsValue, globalSpecSheet));
                    }

                    if (string.IsNullOrEmpty(nvValue))
                    {
                        nvPins.Add(pin + "= NA");
                    }
                    else
                    {
                        nvPins.Add(pin + "=" + GetFomulaValue(nvValue, globalSpecSheet));
                    }
                }

                string key = dcCategory.Item1 + "_" + dcCategory.Item2;
                vrsTable.Table.TryAdd(key, vrsPins);
                nvTable.Table.TryAdd(key, nvPins);
            }
        }

        public void CreateTableCs(string tempFolder, List<string> titles, InstanceSheet instanceSheet, List<Tuple<string, string>> dcCategorys, DcSpecSheet dcSpecSheet, GlobalSpecSheet globalSpecSheet, out VoltageTable vrsAllTable, out VoltageTable nvAllTable)
        {
            vrsAllTable = new VoltageTable(instanceSheet);
            nvAllTable = new VoltageTable(instanceSheet);

            if (instanceSheet == null || dcSpecSheet == null || globalSpecSheet == null)
            {
                return;
            }

            foreach (Tuple<string, string> dcCategory in dcCategorys)
            {
                var vrsAllPins = new List<string>();
                var nvAllPins = new List<string>();
                var pins = new List<string>();

                foreach (string pin in titles)
                {
                    if (BinCutData.PinMapData!.IsGroupExist(pin))
                    {
                        foreach (Pin groupItem in BinCutData.PinMapData.GetGroup(pin)!.PinList)
                        {
                            pins.Add(groupItem.PinName);
                        }
                    }
                    else
                    {
                        pins.Add(pin);
                    }
                }

                foreach (string pin in pins)
                {
                    GetPinNameFromDcSpec(pin, dcSpecSheet, out string vrsPinName, out string nvPinName);
                    GetDCspecValue(vrsPinName, nvPinName, dcSpecSheet, dcCategory.Item1, dcCategory.Item2, out string vrsAllValue, out string nvAllValue);

                    if (string.IsNullOrEmpty(vrsAllValue))
                    {
                        vrsAllPins.Add(pin + "= NA");
                    }
                    else
                    {
                        vrsAllPins.Add(pin + "=" + GetFomulaValue(vrsAllValue, globalSpecSheet));
                    }

                    if (string.IsNullOrEmpty(nvAllValue))
                    {
                        nvAllPins.Add(pin + "= NA");
                    }
                    else
                    {
                        nvAllPins.Add(pin + "=" + GetFomulaValue(nvAllValue, globalSpecSheet));
                    }
                }

                string key = dcCategory.Item1 + "_" + dcCategory.Item2;
                vrsAllTable.Table.TryAdd(key, vrsAllPins);
                nvAllTable.Table.TryAdd(key, nvAllPins);
            }
        }

        protected static void GetPinNameFromDcSpec(string pinName, DcSpecSheet dcSpecSheet, out string vrsPinName, out string nvPinName)
        {
            nvPinName = GetValue(pinName, dcSpecSheet, Var);
            vrsPinName = GetValue(pinName, dcSpecSheet, VopVar);
            if (string.IsNullOrEmpty(vrsPinName))
            {
                vrsPinName = nvPinName;
            }
        }

        private static string GetValue(string pinName, DcSpecSheet dcSpecSheet, string type)
        {
            if (dcSpecSheet.Rows.Exists(x => x.Symbol.EqualsIgnoreCase(pinName + type)))
            {
                return pinName + type;
            }

            while (BinCutData.PinMapData!.GroupList.Any(x => x.PinName.EqualsIgnoreCase(pinName)))
            {
                string name = BinCutData.PinMapData.GroupList.First(x => x.PinName.EqualsIgnoreCase(pinName))
                        .PinList.First().PinName;
                if (dcSpecSheet.Rows.Exists(x => x.Symbol.EqualsIgnoreCase(name + type)))
                {
                    return name + type;
                }

                pinName = name;
            }
            return "";
        }

        protected string GetFomulaValue(string formula, GlobalSpecSheet globalSpecSheet)
        {
            try
            {
                formula = formula.TrimStart(' ').TrimStart('=');
                formula = ReplaceGlobalValue(formula, globalSpecSheet);

                // Use decimal for simple numeric literals to avoid double-precision
                // midpoint rounding differences (e.g. 0.8855 → 0.886, not 0.885).
                if (decimal.TryParse(formula, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal decVal))
                {
                    return Math.Round(decVal, 3, MidpointRounding.AwayFromZero).ToString("F3");
                }

                ExcelWorksheet.Cells["A1"].Formula = formula;
                ExcelWorksheet.Cells["A1"].Calculate();
                return $"{ExcelWorksheet.Cells["A1"].Value:F3}";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static string ReplaceGlobalValue(string formula, GlobalSpecSheet globalSpecSheet)
        {
            foreach (object item in Reg.RegexFormula.Matches(formula))
            {
                string itemStr = item.ToString() ?? "";
                if (globalSpecSheet.Rows.Exists(x => itemStr.EqualsIgnoreCase("_" + x.Symbol)))
                {
                    GlobalSpec row = globalSpecSheet.Rows.Find(x => itemStr.EqualsIgnoreCase("_" + x.Symbol))!;
                    string value = row.Value.TrimStart(' ').TrimStart('=');
                    string global = "_" + row.Symbol.ToUpper();
                    while (!double.TryParse(value, out double _))
                    {
                        value = ReplaceGlobalValue(value, globalSpecSheet);
                    }
                    return formula.Replace(global, value);
                }
            }
            return formula;
        }

        private static void GetDCspecValue(string vrsPinName, string nvPinName, DcSpecSheet dcSpecSheet, string dcSpec, string selector, out string vrsValue, out string nvValue)
        {
            vrsValue = "";
            nvValue = "";
            if (dcSpecSheet.CategoryList.Exists(x => x.EqualsIgnoreCase(dcSpec)))
            {
                int index = dcSpecSheet.CategoryList.FindIndex(x => x.EqualsIgnoreCase(dcSpec));
                if (dcSpecSheet.Rows.Exists(x => x.SelectorList.Exists(y => y.SelectorName!.EqualsIgnoreCase(selector))))
                {
                    if (!string.IsNullOrEmpty(vrsPinName))
                    {
                        DcSpec rowVrs = dcSpecSheet.Rows.Find(x => x.Symbol.StartsWithIgnoreCase(vrsPinName))!;
                        if (selector.EqualsIgnoreCase("Max"))
                        {
                            vrsValue = rowVrs.CategoryList[index].Max ?? "";
                        }
                        else if (selector.EqualsIgnoreCase("Typ"))
                        {
                            vrsValue = rowVrs.CategoryList[index].Typ ?? "";
                        }
                        else if (selector.EqualsIgnoreCase("Min"))
                        {
                            vrsValue = rowVrs.CategoryList[index].Min ?? "";
                        }
                    }

                    if (!string.IsNullOrEmpty(nvPinName))
                    {
                        DcSpec rowNs = dcSpecSheet.Rows.Find(x => x.Symbol.StartsWithIgnoreCase(nvPinName))!;
                        if (selector.EqualsIgnoreCase("Max"))
                        {
                            nvValue = rowNs.CategoryList[index].Max ?? "";
                        }
                        else if (selector.EqualsIgnoreCase("Typ"))
                        {
                            nvValue = rowNs.CategoryList[index].Typ ?? "";
                        }
                        else if (selector.EqualsIgnoreCase("Min"))
                        {
                            nvValue = rowNs.CategoryList[index].Min ?? "";
                        }
                    }
                }
            }
        }
    }
}
