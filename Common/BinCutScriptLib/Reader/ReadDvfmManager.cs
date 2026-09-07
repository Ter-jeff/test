using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

namespace BinCutScriptLib.Reader
{
    internal class ReadDvfmManager
    {
        private const string Passbin = "PASSBIN";
        private const string VddProduct = "VDD_Product";
        private const string ReadDvfmToGradeVdd = "<Read_DVFM_To_GradeVDD>";

        public static bool ReadProduct(StreamWriter streamWriter, EnumJob enumJob, ref OneTouchDown oneTouchDown, ref SiteInfo[] siteInfoArray)
        {
            bool flag = ReadRows(oneTouchDown, out List<ReadDvfmRow> rows, out int _);

            //Set Bin
            IEnumerable<ReadDvfmRow> bins = rows.Where(x => x.Name.EqualsIgnoreCase(Passbin));
            foreach (ReadDvfmRow bin in bins)
            {
                int binNumber = (int)bin.Measured;
                if (siteInfoArray[bin.Site].Bin > binNumber)
                {
                    BinCutPrint.PrintCommomError(streamWriter, bin.Line, $"The Bin Number of Judge_Store_IDS {siteInfoArray[bin.Site].Bin} is bigger than DVFM {binNumber} !!!");
                }

                siteInfoArray[bin.Site].Bin = (int)bin.Measured;

            }

            IEnumerable<ReadDvfmRow> powers = rows.Where(x => x.Type.EqualsIgnoreCase(VddProduct));
            foreach (ReadDvfmRow power in powers)
            {
                //Set product value by efuse
                string mode = power.Name.Split('_').Last();
                int site = power.Site;
                double value = power.Measured;
                double gb = SiteInfoHelpers.GetEfuseGb(BinCutData.BinningTables[siteInfoArray[site].Bin - 1], mode, enumJob);
                if (!siteInfoArray[site].EFuseValues.Exists(x => x.Name.EqualsIgnoreCase(mode)))
                {
                    var row = new EFuseRow { Name = mode, Value = value, Gb = gb };
                    siteInfoArray[site].EFuseValues.Add(row);
                }
                else
                {
                    IEnumerable<EFuseRow> eFuses = siteInfoArray[site].EFuseValues.Where(x => x.Name.EqualsIgnoreCase(mode));
                    foreach (EFuseRow eFuse in eFuses)
                    {
                        eFuse.Value = value;
                        eFuse.Gb = gb;
                    }
                }
            }

            return flag;
        }

        private static bool ReadRows(OneTouchDown oneTouchDown, out List<ReadDvfmRow> readDvfmRows, out int oneTouchIndex)
        {
            readDvfmRows = [];
            bool? returnFlag = null;
            oneTouchIndex = GetStartIndex(oneTouchDown, ref returnFlag);
            if (returnFlag != null)
            {
                return (bool)returnFlag;
            }

            oneTouchIndex++;

            //<Read_DVFM_To_GradeVDD>
            //16351000 1     PASSBIN                                                                                                                                                 -1       1              1                  3              0              0       
            //16351000 2     PASSBIN                                                                                                                                                 -1       1              1                  3              0              0       
            //16351000 3     PASSBIN                                                                                                                                                 -1       1              1                  3              0              0       
            //16351001 1     VDD_PCPU_MC601 VDD Grade                                                                                                                                -1       N/A            646.8750 mV        N/A            0.0000 V       0       
            //16351002 1     VDD_PCPU_MC601 VDD Product                                                                                                                              -1       N/A            709.3750 mV        N/A            0.0000 V       0       
            for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (IsStopPoint(oneTouchDown, oneTouchIndex))
                {
                    break;
                }

                readDvfmRows.Add(oneTouchDown.Lines[oneTouchIndex].NewReadDvfmLine().GetReadDvfmLineRow());
            }
            return true;
        }

        private static int GetStartIndex(OneTouchDown oneTouchDown, ref bool? returnFlag)
        {
            bool isFoundCfg = false;
            int oneTouchIndex;
            for (oneTouchIndex = 0; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase(ReadDvfmToGradeVdd))
                {
                    isFoundCfg = true;
                    break;
                }
            }

            if (!isFoundCfg)
            {
                returnFlag = false;
            }
            return oneTouchIndex;
        }

        public static bool GetCp2PowerNames(OneTouchDown oneTouchDown, out List<string> powerNames)
        {
            bool flag = ReadRows(oneTouchDown, out List<ReadDvfmRow> rows, out int _);
            powerNames = [.. rows.Where(x => x.Type.StartsWithIgnoreCase("VDD_Product")).Select(x => x.Name).Distinct()];
            return flag;
        }

        private static bool IsStopPoint(OneTouchDown oneTouchDown, int oneTouchIndex)
        {
            return oneTouchDown.Lines[oneTouchIndex].IsDvfmStopPoint();
        }
    }
}
