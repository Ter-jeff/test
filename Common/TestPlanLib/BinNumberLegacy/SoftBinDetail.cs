using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace TestPlanLib.BinNumberLegacy
{
    public class SoftBinDetail
    {
        public string SheetName { set; get; } = "";
        public string Category { set; get; } = "";
        public List<SoftDetailDigiDef> DigiDefList { set; get; } = [];

        public bool GetDetailNumber(SoftBinNumPara softBinNumPara, out string number)
        {
            bool lBResult = true;
            number = "";
            List<string> subDigList = [];
            foreach (SoftDetailDigiDef digiDef in DigiDefList)
            {
                List<SoftDetailNumDef> numberDigiList;
                if (subDigList.Count != 0)
                {
                    numberDigiList = [.. digiDef.NumDefList.Where(p => subDigList.Contains(p.Category))];
                }
                else
                {
                    numberDigiList = [.. digiDef.NumDefList];
                }

                if (numberDigiList.Exists(p => p.Match(softBinNumPara.ColumnContentDic)))
                {
                    number += numberDigiList.Find(p => p.Match(softBinNumPara.ColumnContentDic))!.Number;
                    subDigList.AddRange(numberDigiList.Find(p => p.Match(softBinNumPara.ColumnContentDic))!.Subdig);
                }
                else
                {
                    lBResult = false;
                }
            }
            return lBResult;
        }
    }

    public class SoftDetailCondition
    {
        public string Column { set; get; } = "";
        public string Keyword { set; get; } = "";
    }

    public class SoftDetailNumDef
    {
        public string Category { set; get; } = "";
        public string Number { set; get; } = "";
        public List<string> Subdig { set; get; } = [];
        public List<SoftDetailCondition> Conditions { set; get; } = [];

        public void AddCondition(string columnName, string keyword)
        {
            SoftDetailCondition condition = new SoftDetailCondition
            {
                Column = columnName,
                Keyword = keyword
            };
            Conditions.Add(condition);
        }

        public bool Match(Dictionary<string, string> columnDic)
        {
            foreach (SoftDetailCondition condition in Conditions)
            {
                bool oneCondition = GetByOneCondition(columnDic, condition);
                if (!oneCondition)
                {
                    return false;
                }
            }
            return true;
        }

        private static bool GetByOneCondition(Dictionary<string, string> columnDic, SoftDetailCondition softDetailCondition)
        {
            foreach (string conditionKey in softDetailCondition.Keyword.Split(','))
            {
                if (columnDic.Keys.ToList().Exists(p => p.EqualsIgnoreCase(softDetailCondition.Column)))
                {
                    string column = columnDic.Keys.ToList().Find(p => p.EqualsIgnoreCase(softDetailCondition.Column))!;
                    if (column.EqualsIgnoreCase("Level") ||
                        column.EqualsIgnoreCase("Voltage"))
                    {
                        if (conditionKey.EqualsIgnoreCase(columnDic[column]))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        if (Regex.IsMatch(columnDic[column], conditionKey, RegexOptions.IgnoreCase))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
    }

    public class SoftDetailDigiDef
    {
        public List<SoftDetailNumDef> NumDefList { set; get; } = [];
    }

}
