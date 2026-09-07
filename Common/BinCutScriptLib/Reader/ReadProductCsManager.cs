using System.Collections.Generic;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;

using CommonLib.Extension;

namespace BinCutScriptLib.Reader
{
    internal class ReadProductCsManager
    {
        public static void CheckProductCs(ref SiteInfo[] siteInfoArray, ref OneTouchDown oneTouchDown)
        {
            //Flow Flow_efuse_BankRead Start 
            //50483        6     vdd_soc_ms004 Bincut                                                    -1       710.0000 mV    795.0000 mV          895.0000 mV    0.0000         0       
            //50484        6     vdd_soc_ms005 Bincut                                                    -1       790.0000 mV    880.0000 mV          950.0000 mV    0.0000         0       
            //50485        6     vdd_dcs_ddr_md001 Bincut                                                -1       580.0000 mV    610.0000 mV          645.0000 mV    0.0000         0       
            //50486        6     vdd_dcs_ddr_md002 Bincut                                                -1       590.0000 mV    640.0000 mV          705.0000 mV    0.0000         0       
            //Flow Flow_efuse_BankRead Stop

            var rows = new List<ReadProductCsRow>();
            int i;
            int startIndex = -1;
            for (i = 0; i < oneTouchDown.Lines.Count; i++)
            {
                if (oneTouchDown.Lines[i].Line.Contains("Flow Flow_efuse_BankRead Start"))
                {
                    startIndex = i;
                }
                else if (startIndex != -1)
                {
                    if (oneTouchDown.Lines[i].Line.Contains("Flow Flow_efuse_BankRead Stop"))
                    {
                        List<BinCutLineBase> lines = oneTouchDown.Lines.GetRange(startIndex + 1, i - startIndex - 1);
                        var reader = new ReadProductCsReader(lines);
                        rows = reader.ReadProductCsRows;
                        break;
                    }
                }
            }

            foreach (ReadProductCsRow row in rows)
            {
                int site = row.Site;
                _ = double.TryParse(row.Measured, out double productValue);
                if (row.MeasuredUnit.EqualsIgnoreCase("v"))
                {
                    productValue *= 1000.0;
                }

                string mode = BinCutAlgorithmService.GetModeByName(row.TestName);

                if (!siteInfoArray[site].EFuseValues.Exists(x => x.Name.EqualsIgnoreCase(mode)))
                {
                    var efuseRow = new EFuseRow { Name = mode, Value = productValue };
                    siteInfoArray[site].EFuseValues.Add(efuseRow);
                }
                //else
                //{
                //    IEnumerable<EFuseRow> eFuses = allDice[site].EFuseValues.Where(x => x.Name.Equals(mode, StringComparison.CurrentCultureIgnoreCase));
                //    foreach (EFuseRow eFuse in eFuses)
                //    {
                //        if (false)
                //        {
                //            eFuse.Value = productValue;
                //        }
                //        else
                //        {
                //            continue;
                //        }
                //    }
                //}
            }
        }
    }
}
