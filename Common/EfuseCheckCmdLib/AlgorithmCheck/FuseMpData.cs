using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    public class FuseMpData(string site)
    {
        private readonly Dictionary<string, string> _allDatas = [];
        private readonly Dictionary<string, List<string>> _allMlsbDataSets = [];
        private readonly Dictionary<string, bool> _allDataMode = [];
        public string Site = site;

        public void SetData(string fuseData)
        {
            if (fuseData.Split(',').Length <= 6)
            {
                return;
            }

            string[] fuseSgmts = fuseData.Split(',');
            string bank = fuseSgmts[1].Trim();
            string lsb = fuseSgmts[4].Trim();
            string msb = fuseSgmts[5].Trim();
            string data = fuseSgmts[6].Trim();
            if (_allDatas.TryAdd(bank, ""))
            {
                _allMlsbDataSets.Add(bank, []);
            }
            string mlsbData = $"{lsb},{msb}";
            if (_allMlsbDataSets[bank].Exists(x => x.EqualsIgnoreCase(mlsbData)))
            {
                //If the data is already add in dic, use the latest data
                _allMlsbDataSets[bank] = [];
                _allDatas[bank] = "";
            }
            _allMlsbDataSets[bank].Add(mlsbData);
            if (_allDataMode.TryGetValue(bank, out bool value) && value) //isDoubleBit
            {
                int loopCnt = data.Length / 2;
                for (int i = 0; i < loopCnt; i += 4)
                {
                    string preFixBinary = data.Substring(data.Length - (i * 2) - 8, 4);
                    string postFixBinary = data.Substring(data.Length - (i * 2) - 4, 4);
                    int resultOpeateByOr = Convert.ToInt32(preFixBinary, 16) | Convert.ToInt32(postFixBinary, 16);
                    if (resultOpeateByOr.ToString("X").PadLeft(4, '0') != preFixBinary)
                    {
                    }
                    if (resultOpeateByOr.ToString("X").PadLeft(4, '0') != postFixBinary)
                    {
                    }
                    _allDatas[bank] = resultOpeateByOr.ToString("X").PadLeft(4, '0') + _allDatas[bank];
                }
            }
            else
            {
                _allDatas[bank] = data + _allDatas[bank];
            }
        }

        public void SetDataMode(string modeData) //CONFIG,1,Double-Bits = True
        {
            if (modeData.Split(',').Length < 3)
            {
                return;
            }

            string[] modeSgmts = modeData.Split(',');
            string bank = modeSgmts[0].Trim();
            bool isDoubleBits = bool.Parse(modeSgmts[2].Split('=').Last());
            if (!_allDatas.ContainsKey(bank))
            {
                _allDataMode.Add(bank, isDoubleBits);
            }
            _allDataMode[bank] = isDoubleBits;
        }

        public EfuseDatalogItem? GetData(string bank, string msb, string lsb, string algo, string resolution, string baseVoltage)
        {
            if (!_allDatas.TryGetValue(bank, out string? value))
            {
                return null;
            }

            try
            {
                var result = new EfuseDatalogItem { Block = bank };
                double tmpdouble = 0.0;
                int indexMsb = int.Parse(msb) > int.Parse(lsb) ? Math.DivRem(int.Parse(msb), 4, out int tmpMsb) : Math.DivRem(int.Parse(lsb), 4, out tmpMsb);
                int indexLsb = int.Parse(msb) > int.Parse(lsb) ? Math.DivRem(int.Parse(lsb), 4, out int _) : Math.DivRem(int.Parse(msb), 4, out int _);
                string deriveData = value.Substring(value.Length - 1 - indexMsb, indexMsb - indexLsb + 1);
                string deriveDataBin = HexToBin(deriveData);
                deriveDataBin = deriveDataBin.Substring(4 - (tmpMsb + 1), Math.Abs(int.Parse(msb) - int.Parse(lsb)) + 1);
                result.RawData = deriveDataBin;
                if ((bank.EqualsIgnoreCase("ECID") && !algo.EqualsIgnoreCase("CRC")) ||
                    (bank.EqualsIgnoreCase("CONFIG") && (algo.EqualsIgnoreCase("LOTID") || algo.EqualsIgnoreCase("NUMERIC"))))
                {
                    deriveDataBin = Reverse(deriveDataBin);
                }

                if (deriveDataBin.Length % 4 != 0 && !algo.EqualsIgnoreCase("LOTID"))
                {
                    deriveDataBin = deriveDataBin.PadLeft(((deriveDataBin.Length / 4) + 1) * 4, '0');
                }

                switch (algo.ToUpper())
                {
                    case "LOTID":
                        result.Value = GetLotID(deriveDataBin);
                        break;
                    case "NUMERIC":
                        result.Value = Convert.ToInt32(deriveDataBin, 2).ToString();
                        break;
                    case "VDDBIN":
                        {
                            if (double.TryParse(resolution, out tmpdouble))
                            {
                                result.Value = ((Convert.ToInt32(deriveDataBin, 2) * tmpdouble) + double.Parse(baseVoltage)).ToString();
                            }
                            else
                            {
                                result.Value = (Convert.ToInt32(deriveDataBin, 2) + double.Parse(baseVoltage)).ToString();
                            }

                            break;
                        }
                    case "IDS":
                        if (double.TryParse(resolution, out tmpdouble))
                        {
                            result.Value = (Convert.ToInt32(deriveDataBin, 2) * tmpdouble).ToString();
                        }
                        else
                        {
                            result.Value = Convert.ToInt32(deriveDataBin, 2).ToString();
                        }

                        break;
                    case "BASE":
                        if (double.TryParse(resolution, out tmpdouble))
                        {
                            result.Value = ((Convert.ToInt32(deriveDataBin, 2) + 1) * tmpdouble).ToString();
                        }
                        else
                        {
                            result.Value = Convert.ToInt32(deriveDataBin, 2).ToString();
                        }

                        break;
                    case "CRC":
                        //1000110100101110
                        result.Value = "0x" + string.Join("", Enumerable.Range(0, deriveDataBin.Length / 8).Select(i => Convert.ToByte(deriveDataBin.Substring(i * 8, 8), 2).ToString("X2")));
                        break;
                    default:
                        if (deriveDataBin.Length >= 32)
                        {
                            if (deriveDataBin.Length % 8 != 0)
                            {
                                deriveDataBin = deriveDataBin.PadLeft(((deriveDataBin.Length / 8) + 1) * 8, '0');
                            }

                            result.Value = "0x" + string.Join("", Enumerable.Range(0, deriveDataBin.Length / 8).Select(i => Convert.ToByte(deriveDataBin.Substring(i * 8, 8), 2).ToString("X2")));
                        }
                        else
                        {
                            result.Value = Convert.ToInt32(deriveDataBin, 2).ToString();
                        }
                        break;
                }
                return result;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string GetLotID(string data)
        {
            string result = "";
            var datas = new List<string>();
            while (!string.IsNullOrEmpty(data))
            {
                string tmpData;
                if (data.Length >= 6)
                {
                    tmpData = data[..6];
                    data = data[6..];
                }
                else
                {
                    tmpData = data;
                    data = "";
                }
                datas.Add(tmpData);
                string ch;
                if (Convert.ToInt32(tmpData, 2) < 10)
                {
                    ch = Convert.ToInt32(tmpData, 2).ToString();
                }
                else
                {
                    ch = Convert.ToChar(Convert.ToInt32(tmpData, 2) + 55).ToString();
                }

                result += ch;
            }
            return result;
        }

        private static string HexToBin(string data)
        {
            return string.Join(string.Empty, data.Select(p => Convert.ToString(Convert.ToInt32(p.ToString(), 16), 2).PadLeft(4, '0')));
        }

        private static string Reverse(string data)
        {
            char[] arr = data.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }
    }
}
