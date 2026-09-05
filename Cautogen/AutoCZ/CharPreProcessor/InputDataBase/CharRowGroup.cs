using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

namespace Cautogen.AutoCZ.CharPreProcessor.InputDataBase
{
    public class CharRowGroup : Characterization
    {
        /* properties */
        private List<Characterization> _charRows;
        private Characterization _targetCharRow;  // first non-TSMC row being used to hook the base row properities

        public override string TpName { get { return string.Join(",", _charRows.Select(p => p.TpName).ToList()); } }
        public override string IpUse1 { get { return string.Join(",", _charRows.Select(p => p.IpUse1).ToList()); } }
        public override string IpUse2 { get { return string.Join(",", _charRows.Select(p => p.IpUse2).ToList()); } }
        public override string IpUse3 { get { return string.Join(",", _charRows.Select(p => p.IpUse3).ToList()); } }
        public override string IpUse4 { get { return string.Join(",", _charRows.Select(p => p.IpUse4).ToList()); } }
        public override string IpUse5 { get { return string.Join(",", _charRows.Select(p => p.IpUse5).ToList()); } }

        /* constructor */
        public CharRowGroup(Characterization charRow)
        {
            _charRows = new List<Characterization> { charRow };
            _targetCharRow = new Characterization(charRow);
            Copy(charRow);
        }

        /* methods */
        public static List<CharRowGroup> GetCharRowGroups(IEnumerable<Characterization> charplansheet)
        {
            var charRowGroups = new List<CharRowGroup>();

            // No same charRowGroup, create new one and append into charRowGroups
            foreach (Characterization charRow in charplansheet.Where(charRow => !_UpdateCharRowGroupList(charRow, charRowGroups)))
            {
                charRowGroups.Add(new CharRowGroup(charRow));
            }

            foreach (CharRowGroup charRowGroup in charRowGroups)
            {
                charRowGroup._Sort();
            }

            return charRowGroups;
        }

        private static bool _UpdateCharRowGroupList(Characterization charRow, IEnumerable<CharRowGroup> charRowGroups)
        {
            // for same group, push the charRow into that charRowGroup
            foreach (CharRowGroup charRowGroup in charRowGroups.Where(charRowGroup => charRowGroup._IsSameGroup(charRow)))
            {
                charRowGroup._Append(charRow);
                return true;
            }
            return false;
        }

        private bool _IsSameGroup(Characterization charRow)
        {
            if (charRow.UserDef1.ToUpper() != "HAC")
            {
                return false;
            }

            if (charRow.DcCateName != _targetCharRow.DcCateName)
            {
                return false;
            }

            if (charRow.OtherSupplies == "multi")
            {
                return false;
            }

            if (charRow.OtherSupplies != _targetCharRow.OtherSupplies)
            {
                return false;
            }

            if (charRow.UserDef9 != _targetCharRow.UserDef9)
            {
                return false;
            }

            if (charRow.Group != _targetCharRow.Group)
            {
                return false;
            }

            return charRow.Payload1 == _targetCharRow.Payload1;
        }

        private void _Append(Characterization charRow)
        {
            _charRows.Add(charRow);

            // replace target charRow if TpName stars with TSMC
            if (!Regex.IsMatch(_targetCharRow.TpName, "^TSMC_", RegexOptions.IgnoreCase))
            {
                return;
            }

            _targetCharRow = new Characterization(charRow);
            Copy(charRow);
        }

        private void _Sort()
        {
            if (UserDef1.ToUpper() != "HAC")
            {
                return;
            }

            foreach (Characterization row in _charRows)
            {
                string pinAlias = PatInfoReader.PinUnderlineDict.Keys.FirstOrDefault(
                    p => Regex.IsMatch(row.UserDef4, p.Replace("_", ""), RegexOptions.IgnoreCase));
                row.PinName = pinAlias ?? row.UserDef4;
            }

            //group with measseq and order 
            IEnumerable<List<Characterization>> measGroup = _MeasSeqGroup(_charRows);

            //by sequence sorting
            var group = new List<Characterization>();
            foreach (KeyValuePair<int, List<Characterization>> newItem in measGroup.Select(_SplitDiffPin).Select(splitGroup => splitGroup.GroupBy(p => p.DiffType).OrderBy(p => p.Key).ToDictionary(p => p.Key, p => p.ToList())).SelectMany(newGroup => newGroup))
            {
                if (
                    newItem.Value.All(
                        p => p.MeasSeq != 999 && !Regex.IsMatch(p.TpName, "_IDS_", RegexOptions.IgnoreCase)))
                {
                    //single-end sorting
                    newItem.Value.Sort((x, y) =>
                    {
                        if (x.PinName.Equals(y.PinName, StringComparison.OrdinalIgnoreCase))
                        {
                            return string.Compare(x.TpName, y.TpName, StringComparison.CurrentCultureIgnoreCase);
                        }
                        else
                        {
                            return string.Compare(x.PinName, y.PinName, StringComparison.CurrentCultureIgnoreCase);
                        }


                    });
                }
                group.AddRange(newItem.Value);
            }
            _charRows = group;
        }

        //Generate CharRows group by MeasSeq and order
        private static IEnumerable<List<Characterization>> _MeasSeqGroup(IEnumerable<Characterization> item)
        {
            return item.GroupBy(p => p.MeasSeq).OrderBy(p => p.Key).ToDictionary(p => p.Key, p => p.ToList()).Values;
        }

        private static List<Characterization> _SplitDiffPin(List<Characterization> charGroup)
        {
            var newGroup = new List<Characterization>();
            foreach (Characterization item in charGroup)
            {
                //differential pin contains "MeasV/I/F"
                bool checkPin = false;//default: single-end
                if (Regex.IsMatch(item.UserDef4, @"\w+(DIFF|CM)\w+", RegexOptions.IgnoreCase))
                {
                    string[] pins = Regex.Split(item.UserDef4, "DIFF|CM", RegexOptions.IgnoreCase);
                    checkPin = pins.All(pin => UtilityMain.UtilityData.PinList.Keys.FirstOrDefault(p => p.Replace("_", "") == pin.ToLower()) != null);
                }
                if (checkPin)
                {
                    switch (item.UserDef2)
                    {
                        case "MeasV":
                            newGroup.Add(_GenerateDiffChar(item, "pPin"));
                            newGroup.Add(_GenerateDiffChar(item, "nPin"));
                            newGroup.Add(_GenerateDiffChar(item, "Vod"));
                            newGroup.Add(_GenerateDiffChar(item, "Vcm"));
                            break;
                        case "MeasI":
                            newGroup.Add(_GenerateDiffChar(item, "pPin"));
                            newGroup.Add(_GenerateDiffChar(item, "nPin"));
                            newGroup.Add(_GenerateDiffChar(item, "Vod"));
                            break;
                        case "MeasF":
                            newGroup.Add(_GenerateDiffChar(item, "Vod"));
                            break;
                    }
                }
                else  // single-end pin
                {
                    item.DiffType = 0;
                    newGroup.Add(item);
                }
            }
            return newGroup;
        }

        //Generate Different Char row according to type(pPin/nPin/Vod/Vcm)
        private static Characterization _GenerateDiffChar(Characterization item, string type)
        {
            var newItem = new Characterization();
            newItem.Copy(item);
            string replaceStr = "";
            const string expr = "(DIFF|CM)";
            switch (type)
            {
                case "pPin":
                    newItem.PinName = _SearchPNPin(newItem.UserDef4, "P");
                    replaceStr = Regex.Replace(newItem.UserDef4, expr, "DIFF1");
                    newItem.DiffType = 0;
                    break;
                case "nPin":
                    newItem.PinName = _SearchPNPin(newItem.UserDef4, "N");
                    replaceStr = Regex.Replace(newItem.UserDef4, expr, "DIFF0");
                    newItem.DiffType = 0;
                    break;
                case "Vod":
                    newItem.PinName = _SearchPNPin(newItem.UserDef4, "P");
                    replaceStr = Regex.Replace(newItem.UserDef4, expr, "DIFF");
                    newItem.DiffType = 1;
                    break;
                case "Vcm":
                    newItem.PinName = _SearchPNPin(newItem.UserDef4, "P");
                    replaceStr = Regex.Replace(newItem.UserDef4, expr, "CM");
                    newItem.DiffType = 2;
                    break;
            }
            newItem.TpName = newItem.TpName.Replace("_" + newItem.UserDef4 + "_", "_" + replaceStr + "_");


            return newItem;
        }

        //search PNPin from differential pin
        private static string _SearchPNPin(string diffPinName, string pn)
        {
            return pn == "P"
                ? Regex.Match(diffPinName, @"(?<Pin>\w+)(DIFF|CM)\w+", RegexOptions.IgnoreCase).Groups["Pin"].Value
                : Regex.Match(diffPinName, @"\w+(DIFF|CM)(?<Pin>\w+)", RegexOptions.IgnoreCase).Groups["Pin"].Value;
        }
    }
}
