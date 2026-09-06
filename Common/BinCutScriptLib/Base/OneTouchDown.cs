using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Extension;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Reader;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

namespace BinCutScriptLib.Base
{
    public partial class OneTouchDown
    {
        [GeneratedRegex(@"^\d")]
        private static partial Regex MyRegex();

        public List<BinCutLineBase> Lines { get; set; } = [];
        public List<InterpolationRow> EnRows { get; set; } = [];

        public Dictionary<int, int> Bins { get; internal set; } = [];

        public int FlowVddbinningHeaderStartIndex
        {
            get
            {
                foreach (BinCutLineBase line in Lines)
                {
                    if (line.Line.StartsWithIgnoreCase("*print: Flow_Vddbinning_Header"))
                    {
                        return line.LineNo;
                    }

                    if (line.Line.StartsWithIgnoreCase("*print: Flow_Vddbinning start*"))
                    {
                        return line.LineNo;
                    }
                }
                return -1;
            }
        }

        public void ClearAll()
        {
            Lines.Clear();
            EnRows.Clear();
        }

        public List<IdsSetWriteDecimalLine> GetIdsSetWriteDecimalLines(string text)
        {
            int no = 0;
            foreach (BinCutLineBase line in Lines)
            {
                if (line.Line.StartsWithIgnoreCase("<Judge_stored_IDS>"))
                {
                    no = line.LineNo;
                    break;
                }
            }

            if (no == 0)
            {
                return [.. Lines.Where(x => x.Line.Contains(text) ||
                                        (x.Line.Contains("Set eFuse") && x.Line.Contains("ids") && x.Line.Contains("mA")) ||
                                        (x.Line.Contains("eFuse_SetWriteVariable_SiteAware") && x.Line.Contains("ids") && x.Line.Contains("mA)"))).Select(line => line.NewIdsSetWriteDecimalLine())];
            }

            return [.. Lines.Where(x => x.Line.Contains(text) ||
                                    (x.Line.Contains("Set eFuse") && x.Line.Contains("ids") && x.Line.Contains("mA")) ||
                                    (x.Line.Contains("eFuse_SetWriteVariable_SiteAware") && x.Line.Contains("ids") && x.Line.Contains("mA)"))).Select(line => line.NewIdsSetWriteDecimalLine())
                .Where(x => x.LineNo < no)];
        }

        public List<BinCutLineBase> GetBinCutConfig()
        {
            var lines = new List<BinCutLineBase>();
            if (Lines.Count == 0)
            {
                return lines;
            }

            int start = Lines.FindIndex(x => x.Line.Contains("*print: BinCut Config start*"));
            if (start == -1)
            {
                return lines;
            }

            int end = Lines.FindIndex(start + 1, x => x.Line.Contains("*print: BinCut Config end*"));
            if (end == -1)
            {
                return lines;
            }

            lines = Lines.GetRange(start, end - start + 1);
            return lines;
        }

        public List<IdsOffLine> GetIdsOffLines()
        {
            int no = 0;
            foreach (BinCutLineBase line in Lines)
            {
                if (line.Line.StartsWithIgnoreCase("<IDS_OFF1IDS_"))
                {
                    no = line.LineNo;
                    break;
                }
            }
            if (no == 0)
            {
                return [.. Lines.Where(x => x.Line.Contains("HAC_MeasI_N_OFF1IDS_IDS_")).Select(line => line.NewIdsOffLine())];
            }

            return
                [.. Lines.Where(x => x.Line.Contains("HAC_MeasI_N_OFF1IDS_IDS_")).Select(line => line.NewIdsOffLine()).Where(x => x.LineNo > no)];
        }

        public List<IdsOnLine> GetIdsOnLines()
        {
            int no = 0;
            foreach (BinCutLineBase line in Lines)
            {
                if (line.Line.StartsWithIgnoreCase("<IDS_OnIDS"))
                {
                    no = line.LineNo;
                    break;
                }
            }
            if (no == 0)
            {
                return [.. Lines.Where(x => x.Line.Contains("HAC_MeasI_N_ONIDS")).Select(line => line.NewIdsOnLine())];
            }

            return
                [.. Lines.Where(x => x.Line.Contains("HAC_MeasI_N_ONIDS")).Select(line => line.NewIdsOnLine()).Where(x => x.LineNo > no)];
        }

        public List<ReadFromDsscLine1> GetReadfromDsscLines()
        {
            int no = 0;
            foreach (BinCutLineBase line in Lines)
            {
                if (line.Line.StartsWithIgnoreCase("<Judge_stored_IDS>"))
                {
                    no = line.LineNo;
                    break;
                }
            }
            if (no == 0)
            {
                return [.. Lines.Where(x => x.Line.Contains("Read from DSSC")).Select(line => line.NewReadfromDsscLine())];
            }

            return [.. Lines.Where(x => x.Line.Contains("Read from DSSC")).Select(line => line.NewReadfromDsscLine()).Where(x => x.LineNo < no)];
        }

        public List<ReadFromDsscLine1> GetSetFuseValueLinesCsharp()
        {
            int lineNo = 0;
            foreach (BinCutLineBase line in Lines)
            {
                if (line.Line.StartsWithIgnoreCase("<Judge_stored_IDS") || line.Line.StartsWithIgnoreCase("<Judge_stored_IDS_Csharp"))
                {
                    lineNo = line.LineNo;
                    break;
                }
            }
            if (lineNo == 0)
            {
                return [.. Lines.Where(x => x.Line.Contains("Set fuse value in cache")).Select(line => line.NewReadfromDsscLine())];
            }

            return [.. Lines.Where(x => x.Line.Contains("Set fuse value in cache")).Select(line => line.NewReadfromDsscLine()).Where(x => x.LineNo < lineNo)];
        }

        public bool GetIdsOnAndOff(ref SiteInfo[] siteInfoArray)
        {
            List<IdsOnLine> idsOnLines = GetIdsOnLines();
            for (int i = 0; i < idsOnLines.Count; i++)
            {
                string[] spt = idsOnLines[i].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (spt.Length >= 13)
                {
                    int site = int.Parse(spt[1]);
                    string pin = spt[3];
                    double value;
                    if (spt[5] == "N/A")
                    {
                        value = spt[7] == "uA" ? double.Parse(spt[6]) / 1000 : double.Parse(spt[6]);
                    }
                    else
                    {
                        value = spt[8] == "uA" ? double.Parse(spt[7]) / 1000 : double.Parse(spt[7]);
                    }
                    siteInfoArray[site].IdsOnList.Add(pin, value);
                }
            }
            List<IdsOffLine> idsOffLines = GetIdsOffLines();
            for (int i = 0; i < idsOffLines.Count; i++)
            {
                string[] spt = idsOffLines[i].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (spt.Length >= 13)
                {
                    int site = int.Parse(spt[1]);
                    string pin = spt[3];
                    double value;
                    if (spt[5] == "N/A")
                    {
                        value = spt[7] == "uA" ? double.Parse(spt[6]) / 1000 : double.Parse(spt[6]);
                    }
                    else
                    {
                        value = spt[8] == "uA" ? double.Parse(spt[7]) / 1000 : double.Parse(spt[7]);
                    }
                    siteInfoArray[site].IdsOffList.Add(pin, value);
                }
            }
            return idsOnLines.Count > 0 && idsOffLines.Count > 0;
        }

        public bool GetRealIds(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray)
        {
            List<IdsSetWriteDecimalLine> idsSetWriteDecimalLines = GetIdsSetWriteDecimalLines("IDS_SetWriteDecimal");
            int index = FlowVddbinningHeaderStartIndex;
            for (int i = 0; i < idsSetWriteDecimalLines.Count; i++)
            {

                idsSetWriteDecimalLines[i].GetIdsSetWriteDecimal(out string idsName, out double realVal, out int site, out double idsLsb);
                if (realVal == 0 && idsSetWriteDecimalLines[i].LineNo > index && index != -1)
                {
                    BinCutPrint.PrintIdsError(streamWriter, siteInfoArray, idsSetWriteDecimalLines[i].LineNo, site, idsSetWriteDecimalLines[i].Line);
                }

                var ids = new IdsData
                {
                    IdsName = idsName,
                    RealVal = realVal,
                    IdsLsb = idsLsb,
                    IdsType = EnumIdsType.SetWriteDecimal
                };

                if (BinCutConfig.EfuseFloor)
                {
                    ids.EfuseVal = BinCutConfig.IsSimulationMode ? Math.Round(ids.RealVal, 4) : BinCutAlgorithmService.FloorIds(ids.RealVal, ids.IdsLsb);
                }
                else
                {
                    ids.EfuseVal = BinCutConfig.IsSimulationMode ? Math.Round(ids.RealVal, 4) : BinCutAlgorithmService.CeilingIds(ids.RealVal, ids.IdsLsb);
                }

                for (int idsIdx = 0; idsIdx < siteInfoArray[site].RealIds.Count; idsIdx++)
                {
                    if (siteInfoArray[site].RealIds[idsIdx].IdsName.EqualsIgnoreCase(ids.IdsName))
                    {
                        siteInfoArray[site].RealIds.RemoveAt(idsIdx);
                    }
                }
                siteInfoArray[site].RealIds.Add(ids);
            }
            return idsSetWriteDecimalLines.Count > 0;
        }

        public bool GetRealIdsCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray)
        {
            List<IdsSetWriteDecimalLine> lines = GetIdsSetWriteDecimalLines("Stored ids value in cache");
            int index = FlowVddbinningHeaderStartIndex;
            for (int i = 0; i < lines.Count; i++)
            {
                lines[i].GetIdsSetWriteDecimalCs(out string idsName, out double realVal, out int site, out double idsLsb);
                if (realVal == 0 && lines[i].LineNo > index && index != -1)
                {
                    BinCutPrint.PrintIdsError(streamWriter, siteInfoArray, lines[i].LineNo, site, lines[i].Line);
                }

                var idsData = new IdsData
                {
                    IdsName = idsName,
                    RealVal = realVal,
                    IdsLsb = idsLsb,
                    IdsType = EnumIdsType.SetWriteDecimal
                };

                bool isSimulation = BinCutConfig.IsSimulationMode;
                bool applyFloor = BinCutConfig.EfuseFloor;

                idsData.EfuseVal = isSimulation ? Math.Round(idsData.RealVal, 4)
                    : applyFloor ? BinCutAlgorithmService.FloorIds(idsData.RealVal, idsData.IdsLsb) : BinCutAlgorithmService.CeilingIds(idsData.RealVal, idsData.IdsLsb);

                List<IdsData> realIds = siteInfoArray[site].RealIds;
                realIds.RemoveAll(ids => ids.IdsName.EqualsIgnoreCase(idsData.IdsName));
                realIds.Add(idsData);
            }
            return lines.Count > 0;
        }

        public bool GetEfuseIds(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<string> powerNames)
        {
            List<ReadFromDsscLine1> readFromDsscLines = GetReadfromDsscLines();
            int index = FlowVddbinningHeaderStartIndex;
            for (int i = 0; i < readFromDsscLines.Count; i++)
            {
                //Site(0) Read from DSSC :                        ids_vdd_pcpu [(MSB)0410:0400(LSB)] =  26.4000mA [00010000100]                      [0x084]        = 132  
                //Site(0) Read from DSSC :                        ids_vdd_ecpu [(MSB)0420:0411(LSB)] =   8.7000mA [0001010111]                         [0x057]        = 87     
                Match match = Reg.RegexDssc.Match(readFromDsscLines[i].Line);
                if (!match.Success)
                {
                    continue;
                }

                string siteStr = match.Groups["site"].ToString();
                int site = int.Parse(siteStr);
                string efuseItemName = match.Groups["Name"].ToString();

                #region Get IDS
                var ids = new IdsData
                {
                    EfuseVal = 0.0,
                    RealVal = 0.0,
                    IdsLsb = 0.1,
                    IdsName = efuseItemName
                };
                string valueStr1 = match.Groups["value1"].ToString();
                Match allMatchs1 = Reg.RegexValue.Match(valueStr1);
                if (allMatchs1.Length != 0)
                {
                    _ = double.TryParse(allMatchs1.Value, out double value);
                    ids.EfuseVal = value;
                    ids.IdsType = EnumIdsType.ReadFromDssc;
                    ids.RealVal = ids.EfuseVal;
                    if (value == 0 && readFromDsscLines[i].LineNo > index && index != -1)
                    {
                        BinCutPrint.PrintIdsError(streamWriter, siteInfoArray, readFromDsscLines[i].LineNo, site, readFromDsscLines[i].Line);
                    }
                }

                if (ids.EfuseVal != 0)
                {
                    for (int idsIdx = 0; idsIdx < siteInfoArray[site].RealIds.Count; idsIdx++) //後蓋前的寫法
                    {
                        if (siteInfoArray[site].RealIds[idsIdx].IdsName.EqualsIgnoreCase(ids.IdsName))
                        {
                            siteInfoArray[site].RealIds.RemoveAt(idsIdx);
                        }
                    }
                    siteInfoArray[site].RealIds.Add(ids);
                }
                #endregion

                #region Get Prouduct value
                //Site(1) Read from DSSC :                       vdd_gpu_mg002 [(MSB)0505:0498(LSB)] =  587.500mV [00111100]                         [0x3C]         = 60    , 509.375mV
                //Site(1) Read from DSSC :                       vdd_gpu_mg003 [(MSB)0513:0506(LSB)] =  656.250mV [01010010]                         [0x52]         = 82    , 578.125mV
                if (powerNames.Any(x => x.Contains(efuseItemName, StringComparison.OrdinalIgnoreCase)))
                {
                    string valueStr = Reg.RegexDssc.Match(readFromDsscLines[i].Line).Groups["value1"].ToString();
                    Match allMatchs = Reg.RegexValue.Match(valueStr);
                    if (allMatchs.Length != 0)
                    {
                        _ = double.TryParse(allMatchs.Value, out double value);
                        if (!siteInfoArray[site].EfuseProductValue.ContainsKey(efuseItemName.ToUpper()) && value != 0)
                        {
                            siteInfoArray[site].EfuseProductValue.Add(efuseItemName.ToUpper(), value);
                        }
                    }
                }
                #endregion
            }
            return readFromDsscLines.Count > 0;
        }

        public bool GetEfuseIdsCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<string> powerNames)
        {
            //Case1
            //[INFO]  [Site 0] Set fuse value in cache 'bkm_process' = 'SiteGeneric`1 { [0] = 3, , ,  }'.
            //[INFO]  [Site 1] Set fuse value in cache 'CFG_condition_01' = 'SiteGeneric`1 { , [1] = 3, ,  }'.
            //[INFO]  [Site 2] Set fuse value in cache 'CFG_condition_01' = 'SiteGeneric`1 { , , [2] = 3,  }'.
            //[INFO]  [Site 3] Set fuse value in cache 'CFG_condition_01' = 'SiteGeneric`1 { , , , [3] = 3 }'.
            //[INFO]  [Site 0] Set fuse value in cache 'CFG_condition_03' = 'SiteGeneric`1 { [0] = 12, , ,  }'.
            //[INFO]  [Site 1] Set fuse value in cache 'CFG_condition_03' = 'SiteGeneric`1 { , [1] = 12, ,  }'.
            //[INFO]  [Site 2] Set fuse value in cache 'CFG_condition_03' = 'SiteGeneric`1 { , , [2] = 12,  }'.
            //[INFO]  [Site 3] Set fuse value in cache 'CFG_condition_03' = 'SiteGeneric`1 { , , , [3] = 12 }'.
            //Case2
            //[INFO]  [Site 1] Set fuse value in cache 'CFG_condition_01' = 3 (Dec)
            List<ReadFromDsscLine1> lines = GetSetFuseValueLinesCsharp();
            int index = FlowVddbinningHeaderStartIndex;
            for (int i = 0; i < lines.Count; i++)
            {
                string siteStr = "";
                string efuseItemName = "";
                string valueStr1 = "";
                if (lines[i].Line.Contains('{'))
                {
                    Match match = Reg.RegexFuse1.Match(lines[i].Line);
                    if (!match.Success)
                    {
                        continue;
                    }

                    siteStr = match.Groups["site"].ToString();
                    efuseItemName = match.Groups["name"].ToString();
                    valueStr1 = match.Groups["value"].ToString();
                }
                else
                {
                    Match match = Reg.RegexFuse2.Match(lines[i].Line);
                    if (!match.Success)
                    {
                        continue;
                    }

                    siteStr = match.Groups["site"].ToString();
                    efuseItemName = match.Groups["name"].ToString();
                    valueStr1 = match.Groups["value"].ToString();
                }

                _ = int.TryParse(siteStr, out int site);
                var ids = new IdsData
                {
                    EfuseVal = 0.0,
                    RealVal = 0.0,
                    IdsLsb = 0.1,
                    IdsName = efuseItemName
                };

                Match matchValue = Reg.RegexValue.Match(valueStr1.Split('=').Last());
                if (matchValue.Success)
                {
                    if (double.TryParse(matchValue.Value, out double value))
                    {
                        ids.EfuseVal = value;
                        ids.IdsType = EnumIdsType.ReadFromDssc;
                        ids.RealVal = value;

                        if (value == 0 && lines[i].LineNo > index && index != -1)
                        {
                            BinCutPrint.PrintIdsError(streamWriter, siteInfoArray, lines[i].LineNo, site, lines[i].Line);
                        }

                        if (value != 0)
                        {
                            siteInfoArray[site].RealIds.RemoveAll(x => x.IdsName.EqualsIgnoreCase(ids.IdsName));
                            siteInfoArray[site].RealIds.Add(ids);
                        }

                        #region Set Prouduct value
                        if (powerNames.Any(x => x.Contains(efuseItemName, StringComparison.OrdinalIgnoreCase)))
                        {
                            if (!siteInfoArray[site].EfuseProductValue.ContainsKey(efuseItemName.ToUpper()) && value != 0)
                            {
                                siteInfoArray[site].EfuseProductValue.Add(efuseItemName.ToUpper(), value);
                            }
                        }
                        #endregion
                    }
                }
            }
            return lines.Count > 0;
        }

        public int GetBvCount()
        {
            return Lines.Count(x => x.Line.StartsWith("BV_") && !x.Line.Contains("EQN ="));
        }

        public int GetBvCountCs()
        {
            return Lines.Count(x => (x.Line.StartsWith("[INFO]") && x.Line.Contains("Bincut safe voltage:")) || x.Line.Contains("Bincut payload voltage:"));
        }

        public int GetDsscCnt()
        {
            //DSSC_SELSRAM_BV_Str,3,I=1,A=1,G=1,E=1,P=0,D=1,S=1
            //SELSRAM_Compare_Bit_Str,2,1111111(LSB->MSB),I=1,A=1,G=1,E=1,P=1,D=1,S=1
            //SRAM_Vth(DCVS),4,I=0.725V,A=0.728V,G=0.725V,D=0.725V,S=0.725V
            //SELSRAM_DSSC_Bit_Str,5,00000
            //SelSram_voltage,3,I=0.681V,A=0.721V,G=0.650V,D=0.709V,S=0.550V
            int count = Lines.Count(x => x.Line.StartsWithIgnoreCase("DSSC_SELSRAM_BV_Str") ||
                                         x.Line.StartsWithIgnoreCase("SELSRAM_Compare_Bit_Str") ||
                                         x.Line.StartsWithIgnoreCase("SRAM_Vth") ||
                                         x.Line.StartsWithIgnoreCase("SELSRAM_DSSC_Bit_Str") ||
                                         x.Line.StartsWithIgnoreCase("SelSram_voltage") ||
                                         x.Line.StartsWithIgnoreCase("Efuse_Product"));
            return count;
        }

        public List<BinCutLineBase> GetMissingLine()
        {
            var lines = new List<BinCutLineBase>();
            int bincutStartIdx = Lines.Exists(x => x.Line.Contains("*print: BinCut Config start*")) ? Lines.Find(x => x.Line.Contains("*print: BinCut Config start*"))!.LineNo : 0;
            lines.AddRange(Lines.Where(x => x.Line.StartsWith("BV_") && !x.Line.Contains("EQN =")));

            //DSSC_SELSRAM_BV_Str,3,I=1,A=1,G=1,E=1,P=0,D=1,S=1
            //SELSRAM_Compare_Bit_Str,2,1111111(LSB->MSB),I=1,A=1,G=1,E=1,P=1,D=1,S=1
            //SRAM_Vth(DCVS),4,I=0.725V,A=0.728V,G=0.725V,D=0.725V,S=0.725V
            //SELSRAM_DSSC_Bit_Str,5,00000
            //SelSram_voltage,3,I=0.681V,A=0.721V,G=0.650V,D=0.709V,S=0.550V
            lines.AddRange(Lines.Where(x => x.Line.StartsWithIgnoreCase("DSSC_SELSRAM_BV_Str") || x.Line.StartsWithIgnoreCase("SELSRAM_Compare_Bit_Str") || x.Line.StartsWithIgnoreCase("SRAM_Vth") || x.Line.StartsWithIgnoreCase("SELSRAM_DSSC_Bit_Str") || x.Line.StartsWithIgnoreCase("SelSram_voltage") || x.Line.StartsWithIgnoreCase("Efuse_Product")));

            //remove lines that are not in bincut section
            lines.RemoveAll(x => x.LineNo < bincutStartIdx);

            return lines;
        }

        public List<BinCutLineBase> GetMissingLineCsharp()
        {
            var lines = new List<BinCutLineBase>();
            int index = Lines.FindIndex(x => x.Line.StartsWithIgnoreCase("Flow PostBincut Stop"));
            List<BinCutLineBase> range = index != -1 ? Lines.GetRange(0, index) : Lines;
            lines.AddRange(range.Where(x => (x.Line.StartsWith("[INFO]") && x.Line.Contains("Bincut safe voltage:")) || x.Line.Contains("Bincut payload voltage:")));

            //Selsram
            IEnumerable<BinCutLineBase> selSram = range.Where(x => Reg.RegexSelSram1.IsMatch(x.Line.Split('(')[0]));
            lines.AddRange(selSram);
            return lines;
        }

        public Dictionary<string, Dictionary<string, bool>> GetMultiPinResult()
        {
            bool isFound = false;
            var pinResultDic = new Dictionary<string, Dictionary<string, bool>>();
            foreach (BinCutLineBase line in Lines)
            {
                if (isFound)
                {
                    if (line.Line.StartsWith('*') || string.IsNullOrEmpty(line.Line))
                    {
                        break;
                    }

                    List<string> arr = [.. line.Line.Split([' '], StringSplitOptions.RemoveEmptyEntries)];
                    if (arr.Count >= 7)
                    {
                        string site = arr[1];
                        string flag = arr[2];
                        if (!pinResultDic.ContainsKey(site))
                        {
                            pinResultDic.Add(site, []);
                        }

                        if (double.TryParse(arr[5], out double valueDouble))
                        {
                            bool value = Convert.ToBoolean(valueDouble);
                            if (!pinResultDic[site].TryAdd(flag, value))
                            {
                                pinResultDic[site][flag] = value;
                            }
                        }
                    }
                }
                if (line.Line.StartsWithIgnoreCase("<Datalog_Harvesting_Fuses_After_Harvesting_Descision>"))
                {
                    isFound = true;
                }
            }
            return pinResultDic;
        }

        public Dictionary<string, Dictionary<string, bool>> GetFlagState()
        {
            //*******************************************
            //*print: Check_flagstate_for_failflag start*
            //*******************************************
            //<Check_flagstate_for_failflag>

            bool isFound = false;
            var pinResultDic = new Dictionary<string, Dictionary<string, bool>>();
            int index = Lines.FindLastIndex(x => x.Line.StartsWithIgnoreCase("<Check_flagstate_for_failflag>"));
            if (index == -1)
            {
                return pinResultDic;
            }

            for (int i = index; i < Lines.Count; i++)
            {
                BinCutLineBase line = Lines[i];
                if (isFound)
                {
                    if (line.Line.StartsWithIgnoreCase("*print: Check_flagstate_for_failflag end*"))
                    {
                        break;
                    }

                    List<string> arr = [.. line.Line.Split([' '], StringSplitOptions.RemoveEmptyEntries)];
                    if (arr.Count >= 7)
                    {
                        string site = arr[1];
                        string flag = arr[2];
                        if (!pinResultDic.ContainsKey(site))
                        {
                            pinResultDic.Add(site, []);
                        }

                        if (double.TryParse(arr[5], out double valueDouble))
                        {
                            bool value = Convert.ToBoolean(valueDouble);
                            if (!pinResultDic[site].TryAdd(flag, value))
                            {
                                pinResultDic[site][flag] = value;
                            }
                        }
                    }
                }
                if (line.Line.StartsWithIgnoreCase("<Check_flagstate_for_failflag>"))
                {
                    isFound = true;
                }
                //if (line.Line.StartsWith("*print: Check_flagstate_for_failflag start*", StringComparison.CurrentCultureIgnoreCase))
                //    isFound = true;
            }
            return pinResultDic;
        }

        public bool GetPrintOutVddBinning(out List<BinCutLineBase> binCutLineBases)
        {
            binCutLineBases = [];
            bool isFoundPrintOutVdd = false;
            if (Lines.Count == 0)
            {
                return false;
            }

            int oneTouchIndex;
            for (oneTouchIndex = 0; oneTouchIndex < Lines.Count; oneTouchIndex++)
            {
                if (Lines[oneTouchIndex].Line.Contains("<PrintOutVddBinning>") || Lines[oneTouchIndex].Line.Contains("<PrintOutVddBinning_Csharp>"))
                {
                    isFoundPrintOutVdd = true;
                    break;
                }
            }

            if (!isFoundPrintOutVdd)
            {
                return false;
            }

            oneTouchIndex++;
            for (; oneTouchIndex < Lines.Count; oneTouchIndex++)
            {
                //STEP3a. read until a blank line
                //1580213  4     VDD_GPU_MG007 V                        -1       875.0000       928.1250           1.0344 K       0.0000         0       
                //1580214  4     VDD_GPU_MG101 C                        -1       528.1250       565.6250           612.5000       0.0000         0       
                if (Lines[oneTouchIndex].Line.Length == 0 || Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries).Length < 10
                    || !MyRegex().IsMatch(Lines[oneTouchIndex].Line.Trim()))
                {
                    break;
                }

                binCutLineBases.Add(Lines[oneTouchIndex]);
            }

            Lines.RemoveRange(0, oneTouchIndex);
            return true;
        }

        public bool GetPrintOutVddBinningCs(out List<BinCutLineBase> binCutLineBases)
        {
            return Lines.GetAndRemoveRange(out binCutLineBases, "print: Vddbinning_PrintOutVddBinning start", "print: Vddbinning_PrintOutVddBinning end");
        }

        public void GetHarvestConfigForEachTouchDown(SiteInfo[] siteInfoArray)
        {
            var harvtResults = new List<HarvestResult>();
            List<BinCutLineBase> lines = GetBinCutConfig();
            var harvLines = lines.Where(x => Reg.RegexHarvResult.IsMatch(x.Line)).ToList();
            foreach (BinCutLineBase harvLine in harvLines)
            {
                HarvestResult harvtResult = BinCutDatalogConfigReader.GetHarvestResult(harvLine.Line);
                harvtResult.Line = harvLine;
                harvtResults.Add(harvtResult);
            }

            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (lines.Count != 0)
                {
                    siteInfoArray[site].IsBinCutConfig = true;
                }

                if (harvtResults.Count != 0)
                {
                    HarvestResult? hrvResult = harvtResults.Find(x => x.Site == site);
                    if (hrvResult != null)
                    {
                        siteInfoArray[site].HarvesFlags = hrvResult.HarvesFlags;
                    }
                }
            }
        }

        public void GetHarvestConfigForEachTouchDownCs(SiteInfo[] siteInfoArray, EnumJob enumJob)
        {
            //New HarvFlag for C#
            //43235726         0     Harvest_Summary_After_Harvesting_Decision                    F_ECPU_TD_CORE_5_0_SUM_0                 -1       -1.0000        1.0000               1.0000         0.0000         0       

            int start = -1;
            start = Lines.FindIndex(x => x.Line.Contains("[INFO]  ----- BinCut Config start -----"));
            if (start == -1)
            {
                return;
            }

            int siteSortBinIndex = -1;

            if (enumJob.Equals(EnumJob.FT1) || enumJob.Equals(EnumJob.FT2))
            {
                siteSortBinIndex = Lines.FindIndex(start, x => x.Line.Contains(" Site    Sort     Bin"));
                if (siteSortBinIndex == -1)
                {
                    return;
                }

                int index = Lines.GetRange(start, siteSortBinIndex - start).FindLastIndex(x => x.Line.StartsWith("<Harvest_Summary_"));
                if (index != -1)
                {
                    List<BinCutLineBase> lines = Lines.GetRange(start + index, siteSortBinIndex - (start + index) + 1);
                    var reader = new HarvestReader(lines);
                    IEnumerable<HarvestRow> rows = reader.Rows.Where(x => x.Pin.StartsWithIgnoreCase("F_"));
                    for (int site = 0; site < siteInfoArray.Length; site++)
                    {
                        IEnumerable<HarvestRow> flags = rows.Where(x => x.Site == site);
                        foreach (HarvestRow flag in flags)
                        {
                            string value = double.TryParse(flag.Measured, out double measuredValue) && measuredValue == 1.0 ? "T" : "F";
                            siteInfoArray[site].HarvesFlags[flag.Pin] = value;
                        }
                    }
                }
            }
            else
            {
                int index = Lines.GetRange(0, start).FindLastIndex(x => x.Line.StartsWith("<Harvest_Summary_"));
                if (index != -1)
                {
                    List<BinCutLineBase> lines = Lines.GetRange(index, start - index + 1);
                    var reader = new HarvestReader(lines);
                    IEnumerable<HarvestRow> rows = reader.Rows.Where(x => x.Pin.StartsWithIgnoreCase("F_"));
                    for (int site = 0; site < siteInfoArray.Length; site++)
                    {
                        IEnumerable<HarvestRow> flags = rows.Where(x => x.Site == site);
                        if (flags.Any())
                        {
                            foreach (HarvestRow flag in flags)
                            {
                                string value = double.TryParse(flag.Measured, out double measuredValue) && measuredValue == 1.0 ? "T" : "F";
                                siteInfoArray[site].HarvesFlags[flag.Pin] = value;
                            }
                        }
                        else
                        {
                            siteSortBinIndex = Lines.FindIndex(start, x => x.Line.Contains(" Site    Sort     Bin"));
                            if (siteSortBinIndex == -1)
                            {
                                return;
                            }

                            int index2 = Lines.GetRange(start, siteSortBinIndex - start).FindLastIndex(x => x.Line.StartsWith("<Harvest_Summary_"));
                            if (index2 != -1)
                            {
                                List<BinCutLineBase> lines2 = Lines.GetRange(start + index2, siteSortBinIndex - (start + index2) + 1);
                                var reader2 = new HarvestReader(lines2);
                                IEnumerable<HarvestRow> rows2 = reader2.Rows.Where(x => x.Pin.StartsWithIgnoreCase("F_"));
                                for (site = 0; site < siteInfoArray.Length; site++)
                                {
                                    flags = rows2.Where(x => x.Site == site);
                                    foreach (HarvestRow flag in flags)
                                    {
                                        string value = double.TryParse(flag.Measured, out double measuredValue) && measuredValue == 1.0 ? "T" : "F";
                                        siteInfoArray[site].HarvesFlags[flag.Pin] = value;
                                    }
                                }
                            }
                        }

                    }

                }
            }
        }

        public bool GetAdjustVddBinning(out List<AdjustVddBinningRow> adjustVddBinningRows, out List<ProductIdentifierLineRow> productIdentifierLineRows)
        {
            //STEP1. Search until found BV
            adjustVddBinningRows = [];
            productIdentifierLineRows = [];
            bool isFoundAdjustVddBin = false;
            if (Lines.Count == 0)
            {
                return false;
            }

            int oneTouchIndex;
            for (oneTouchIndex = 0; oneTouchIndex < Lines.Count; oneTouchIndex++)
            {
                if (Lines[oneTouchIndex].Line.Contains("<Adjust_VddBinning>"))
                {
                    isFoundAdjustVddBin = true;
                    break;
                }
            }

            if (!isFoundAdjustVddBin)
            {
                return false;
            }

            oneTouchIndex++;
            //STEP3. Get all test limit and save to printOutLines
            for (; oneTouchIndex < Lines.Count; oneTouchIndex++)
            {
                if (Lines[oneTouchIndex].Line.Contains("Error"))
                {
                    return false;
                }

                //STEP3a. read until a blank line
                //1580300  0     VDD_SOC_MS001 E                        -1       1.0000         2.0000             4.0000         0.0000         0       
                //1580301  0     VDD_SOC_MS001 C                        -1       559.3750       600.0000           625.0000       0.0000         0       
                //1580302  0     VDD_SOC_MS001 V                        -1       631.2500       671.8750           696.8750       0.0000         0       
                //1580303  0     VDD_SOC_MS001 I                        -1       0.0000         22.6000            60.0000        0.0000         0       
                if (Lines[oneTouchIndex].Line.StartsWithIgnoreCase("SITE") && Lines[oneTouchIndex].Line.Contains("Adjust fusing bin num to binX"))
                {
                    continue;
                }

                if (Lines[oneTouchIndex].Line.StartsWithIgnoreCase("SITE") && Lines[oneTouchIndex].Line.Contains("[Set eFuse]  Product_Identifier ="))
                {
                    productIdentifierLineRows.Add(Lines[oneTouchIndex].NewProductIdentifierLine().GetIdentifierLineRow());
                }

                if (Lines[oneTouchIndex].Line.StartsWith("BV_"))
                {
                    break;
                }

                if (Lines[oneTouchIndex].Line.Contains('>'))
                {
                    break;
                }

                if (Lines[oneTouchIndex].Line.StartsWith("********************************"))
                {
                    break;
                }

                //2017/05/22 add for T-Sk add str for interpolation
                if (Lines[oneTouchIndex].Line.Contains("Start of Writing Vx voltages"))
                {
                    break;
                }

                string[] tok = Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (tok.Length < 10)
                {
                    continue;
                }

                //check site and type field must have only 1 char
                if (tok[1].Length != 1)
                {
                    continue;
                }

                adjustVddBinningRows.Add(Lines[oneTouchIndex].NewAdjustVddBinningLine().GetAdjustVddBinningRow());
            }
            //delete original data from oneTchLines to speed up parse
            Lines.RemoveRange(0, oneTouchIndex);

            return true;
        }

        public bool GetAdjustVddVddBinningCs(out List<AdjustVddBinningRow> adjustVddBinningRows, out List<ProductIdentifierLineRow> productIdentifierLineRows)
        {
            adjustVddBinningRows = [];
            productIdentifierLineRows = [];

            bool flag = Lines.GetAndRemoveRange(out List<BinCutLineBase> lines, "print: Vddbinning_Adjust_VddBinning start", "print: Vddbinning_Adjust_VddBinning end");
            foreach (BinCutLineBase line in lines)
            {
                if (line.Line.StartsWithIgnoreCase("[INFO]"))
                {
                    continue;
                }

                string[] arr = line.Line.Split([" K ", " "], StringSplitOptions.RemoveEmptyEntries);
                if (arr.Length < 10)
                {
                    continue;
                }

                AdjustVddBinningLine row = line.NewAdjustVddBinningLine();
                AdjustVddBinningRow adjustVddBinningRow = row.GetAdjustVddBinningRowCsharp();
                adjustVddBinningRows.Add(adjustVddBinningRow);
            }
            return flag;
        }

        private (bool flowControl, bool value) PowerBinningHarv(ref string sheetName, ref bool isFoundPrintOutVdd, out int oneTouchIndex, string key)
        {
            for (oneTouchIndex = 0; oneTouchIndex < Lines.Count; oneTouchIndex++)
            {
                if (Lines[oneTouchIndex].Line == key)
                {
                    return (flowControl: false, value: false);
                }

                if (string.IsNullOrEmpty(sheetName) && Lines[oneTouchIndex].Line.Contains("===   PwrBin Sheet :"))
                {
                    sheetName = Lines[oneTouchIndex].Line.Split(':')[1].Trim('=').Trim();
                    isFoundPrintOutVdd = true;
                    break;
                }

                if (Lines[oneTouchIndex].Line.Contains("<Power_Binning_Calculation>") || Lines[oneTouchIndex].Line.StartsWithIgnoreCase("<Power_Binning"))
                {
                    isFoundPrintOutVdd = true;
                    break;
                }
            }

            return (flowControl: true, value: default);
        }

        public bool GetPowerBinningHarvCs(out List<BinCutLineBase> binCutLineBases, out string sheetName, string key)
        {
            binCutLineBases = [];
            //STEP1. Search until found BV
            bool loopFlag = false;
            sheetName = "";
            //[INFO]  ==============================================
            //[INFO] ====== PwrBin Sheet: PwrScreen_CS100F ======
            //[INFO] ==============================================
            bool isFoundPrintOutVdd = false;
            if (Lines.Count == 0)
            {
                return false;
            }

            (bool flowControl, bool value) = PowerBinningHarv(ref sheetName, ref isFoundPrintOutVdd, out int oneTouchIndex, key);
            if (!flowControl)
            {
                return value;
            }

            sheetName = string.IsNullOrEmpty(sheetName) ? "PwrBin" : sheetName;
            if (!isFoundPrintOutVdd)
            {
                return false;
            }

            oneTouchIndex++;
            //41825000         1     PwrScreen_CS100F_VDD_PCPU_MP001_Ids - 1       N / A            230.20 mA N/ A            0.00           0
            //41825001         1     PwrScreen_CS100F_VDD_PCPU_MP001_Vbin - 1       N / A            640.00 mV N/ A            0.00           0
            //41825002         1     PwrScreen_CS100F_VDD_PCPU_MP001_C - 1       N / A            50.00 mV N/ A            0.00           0
            //41825003         1     PwrScreen_CS100F_VDD_PCPU_MP001_P_binned - 1       N / A            7.7288               N / A            0.0000         0
            //41825472         1     PwrScreen_CS100F_Power_Binning_P_total                                                                                                             -1       N/A            65.7153              N/A            0.0000         0       
            GetAllTestLimit(binCutLineBases, ref sheetName, ref loopFlag, ref oneTouchIndex, key);

            //delete original data from cpm to speed up parse
            Lines.RemoveRange(0, oneTouchIndex - 1);

            return loopFlag;
        }

        private void GetAllTestLimit(List<BinCutLineBase> binCutLineBases, ref string sheetName, ref bool loopFlag, ref int oneTouchIndex, string key)
        {
            //STEP3. Get all test limit and save to printOutLines
            for (; oneTouchIndex < Lines.Count; oneTouchIndex++)
            {
                if (Lines[oneTouchIndex].Line.Length == 0 || Lines[oneTouchIndex].Line.Contains(key))
                {
                    loopFlag = true;
                    break;
                }

                if (Lines[oneTouchIndex].Line.Contains("===   PwrBin Sheet :"))
                {
                    loopFlag = true;
                    break;
                }

                if (Lines[oneTouchIndex].Line.Contains("===   PwrBin Sheet :"))
                {
                    sheetName = Lines[oneTouchIndex].Line.Split(':')[1].Trim('=').Trim();
                }

                //48098610 0     PwrBinX_TNX_V3_VDD_SRAM_ANE_Ids                                    -1       N/A            1.0585             N/A            0.0000         0
                string[] patToken1 = Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (patToken1.Length >= 9 && patToken1[2].StartsWithIgnoreCase(sheetName))
                {
                    binCutLineBases.Add(Lines[oneTouchIndex]);
                }
            }
        }

        public int GetHarvestSourceCodeCntCount()
        {
            return Lines.Count(x => x.Line.Contains("HarvestSourceCode"));
        }
    }
}
