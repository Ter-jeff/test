using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions
{
    public class CharInputItem
    {
        // data member
        public string InputStr { get { return _field + ":" + _value; } }

        private string _field = "";
        private string _value = "";

        // constructor
        public CharInputItem(string inputStr)  // e.g. VDD_CPU:V:0.3
        {
            string[] splitStr = inputStr.Split(':');
            switch (splitStr.Length)
            {
                case 2:
                    _field = splitStr[0];
                    _value = splitStr[1];
                    break;

                case 3:
                    _field = splitStr[0] + ":" + splitStr[1];
                    _value = splitStr[2];
                    break;
            }
        }

        // public method
        public bool UpdateVal(CharInputItem item)
        {
            if (!string.Equals(item._field.Split(':')[0], _field.Split(':')[0], StringComparison.CurrentCultureIgnoreCase))
            {
                return false;
            }

            _field = item._field;
            _value = item._value;
            return true;
        }
    }

    public class CharInputItemList
    {
        // member
        public string InputStr
        {
            get { return string.Join(";", (from item in _itemList select item.InputStr).ToList()); }
        }

        private List<CharInputItem> _itemList = new List<CharInputItem>();

        // contructor
        public CharInputItemList(string inputStr)
        {
            _itemList = (from o in inputStr.Split(';') where o != "" select new CharInputItem(o)).ToList();
            _PutTermToLast();
        }

        // public method
        public void Append(CharInputItem appendItem)
        {
            if (!_itemList.Any(item => item.UpdateVal(appendItem)))
            {
                _itemList.Add(appendItem);
            }

            _PutTermToLast();
        }

        public void Extend(CharInputItemList extendList)
        {
            foreach (CharInputItem item in extendList)
            {
                Append(item);
            }

            _PutTermToLast();
        }

        private void _PutTermToLast()
        {
            _itemList = _itemList.OrderBy(item => Regex.IsMatch(item.InputStr, "TERM", RegexOptions.IgnoreCase) ? 1 : 0).ToList();
        }

        public IEnumerator<CharInputItem> GetEnumerator()
        {
            return ((IEnumerable<CharInputItem>)_itemList).GetEnumerator();
        }
    }
}
