using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace TestPlanLib.BinNumberLegacy
{
    public enum EnumBinNumKeyType
    {
        Category,
        Module,
        SubModule,
        Block,
        Level,
        NonType
    }

    public class NumberDef
    {
        public EnumBinNumKeyType KeywordType { set; get; }
        public string Keyword { set; get; } = "";
        public string Number { set; get; } = "";

        public bool Search(SoftBinNumPara softBinNumPara)
        {
            if (KeywordType == EnumBinNumKeyType.Category && Keyword.EqualsIgnoreCase(softBinNumPara.Category))
            {
                return true;
            }
            if (KeywordType == EnumBinNumKeyType.Block && Keyword.EqualsIgnoreCase(softBinNumPara.Block))
            {
                return true;
            }
            if (KeywordType == EnumBinNumKeyType.Module && Keyword.EqualsIgnoreCase(softBinNumPara.Module))
            {
                return true;
            }
            if (KeywordType == EnumBinNumKeyType.SubModule && Keyword.EqualsIgnoreCase(softBinNumPara.SubModule))
            {
                return true;
            }
            return KeywordType == EnumBinNumKeyType.Level && Keyword.EqualsIgnoreCase(softBinNumPara.Level);
        }
    }

    public class DigitalDef
    {
        public List<NumberDef> NumberDefList { set; get; } = [];
        public string ReferSheet { set; get; } = "";

        public void AddNumDef(string keyword, EnumBinNumKeyType enumBinNumKeyType, string numebr)
        {
            NumberDef def = new NumberDef
            {
                Keyword = keyword,
                KeywordType = enumBinNumKeyType,
                Number = numebr
            };
            NumberDefList.Add(def);
        }
    }

    public partial class SoftBinDigiDef
    {
        [GeneratedRegex("(?!Mbist)(?<pmode>M([a-zA-Z]){1}([a-zA-Z0-9]){1}(?<modenumber>[a-fA-F0-9|x|X]{3}))", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        public string Category { set; get; } = "";
        public EnumBinNumKeyType CategoryType { set; get; }
        public List<DigitalDef> DigitalList { set; get; } = [];
        public List<SoftBinDetail> ReferDetail { set; get; } = [];

        public static void GetBinCutDigFromMode(string performanceMode, out string dig23)
        {
            dig23 = "00";
            dig23 = $"{int.Parse(dig23):00}";
        }

        public void GetBinNumberBinCut(SoftBinNumPara softBinNumPara, out string number)
        {
            number = "";
            foreach (DigitalDef digi in DigitalList)
            {
                #region From ModuleList
                if (digi.NumberDefList.Count > 0)
                {
                    if (digi.NumberDefList.Exists(p => p.Search(softBinNumPara)))
                    {
                        number += digi.NumberDefList.Find(p => p.Search(softBinNumPara))!.Number;
                        continue;
                    }

                    GetBinCutDigFromMode(softBinNumPara.PerformanceMode, out string dig2);
                    if (digi.NumberDefList.Exists(p => p.Keyword.EqualsIgnoreCase("Performance mode")))
                    {
                        number += dig2;
                        continue;
                    }
                }
                #endregion

                #region See Low_speed_Mbist
                if (digi.ReferSheet?.Length != 0)
                {
                    SoftBinDetail detail = ReferDetail.Find(p => p.Category.EqualsIgnoreCase(softBinNumPara.Block))!;
                    detail.GetDetailNumber(softBinNumPara, out string detailNum);
                    number += detailNum;
                }
                #endregion
            }
            if (number.Length != 4)
            {
                number = "9999";
            }
        }

        public static void GetNewSubDigFromMode(Dictionary<string, string> performanceModeDic, string performanceMode, out string dig2, out string dig3)
        {
            dig2 = "0";
            dig3 = "0";
            if (_regex.IsMatch(performanceMode))
            {
                string number = MyRegex().Match(performanceMode).Groups["modenumber"].ToString();

                if (int.TryParse(number[..1], out int ldig2))
                {
                    dig2 = ldig2.ToString();
                }

                //dig2 = int.Parse(number.Substring(0, 1)).ToString();
                if (int.TryParse(number.AsSpan(2, 1), out int ldig3))
                {
                    dig3 = ldig3.ToString();
                }

                dig3 = performanceModeDic.ContainsKey(performanceMode.ToUpper()) ? performanceModeDic[performanceMode.ToUpper()] : dig3;
            }
        }

        public void GetBinNumber(Dictionary<string, string> performanceModeDic, SoftBinNumPara softBinNumPara, out string number)
        {
            number = "";
            foreach (DigitalDef digi in DigitalList)
            {
                #region From ModuleList
                if (digi.NumberDefList.Count > 0)
                {
                    if (digi.NumberDefList.Exists(p => p.Search(softBinNumPara)))
                    {
                        number += digi.NumberDefList.Find(p => p.Search(softBinNumPara))!.Number;
                        continue;
                    }

                    GetNewSubDigFromMode(performanceModeDic, softBinNumPara.PerformanceMode, out string dig2, out string dig3);

                    if (digi.NumberDefList.Exists(p => p.Keyword.EqualsIgnoreCase("Other Performance mode")))
                    {
                        number += dig2;
                        continue;
                    }

                    if (digi.NumberDefList.Exists(p => p.Keyword.EqualsIgnoreCase("Performance mode")))
                    {
                        number += dig3;
                        continue;
                    }

                    #region New by performanceMode
                    if (digi.NumberDefList.All(x => x.Keyword.StartsWithIgnoreCase("M")))
                    {
                        if (softBinNumPara.PerformanceMode.Length != 0)
                        {
                            string dig = "0";
                            int prefixLen = 0;
                            foreach (NumberDef numberDef in digi.NumberDefList)
                            {
                                if (softBinNumPara.PerformanceMode.StartsWithIgnoreCase(numberDef.Keyword) &&
                                    numberDef.Keyword.Length > prefixLen)
                                {
                                    dig = numberDef.Number;
                                    prefixLen = numberDef.Keyword.Length;
                                }
                                if (dig != "0")
                                {
                                    break;
                                }
                            }
                            number += dig;
                            continue;
                        }
                        number += "0";
                        continue;
                    }
                    #endregion
                }
                #endregion

                #region See Low_speed_Mbist
                if (digi.ReferSheet?.Length != 0)
                {
                    SoftBinDetail detail = ReferDetail.Find(p => p.Category.EqualsIgnoreCase(softBinNumPara.Block))!;
                    detail.GetDetailNumber(softBinNumPara, out string detailNum);
                    number += detailNum;
                    if (number == "1461")
                    {
                    }
                }
                #endregion
            }
            if (number.Length != 4)
            {
                number = "9999";
            }
        }
    }
}
