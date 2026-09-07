using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

namespace TestPlanLib.DataStruct
{

    public class DcCategoryInfo(string categoryName, bool isHardIpDcCategory = false)
    {
        #region field

        private const string CategoryNameFix = "_";
        private string? _test;
        private string _domain = "";
        private string _subtest = "";
        private string _pmodePatternVdip = "";
        private string _chiplet = "";

        #endregion
        #region constructor

        #endregion

        public string CategoryName { get; } = categoryName;
        public bool IsHardipDcCategory { get; } = isHardIpDcCategory;

        public string Test
        {
            get
            {
                return _test ??= CategoryName.Split('_')[0];
            }
        }

        public string Domain
        {
            get
            {
                _domain = CategoryName.Split('_').Length >= 2 ? CategoryName.Split('_')[1] : DcCategoryName.CategoryDefaultValue;
                return _domain;
            }
        }

        public string Subtest
        {
            get
            {
                _subtest = CategoryName.Split('_').Length >= 3 ? CategoryName.Split('_')[2] : DcCategoryName.CategoryDefaultValue;
                return _subtest;
            }
        }

        public string PmodePatternVdip
        {
            get
            {
                //update by laura at 2019/09/11 to support IDS category new format(Userdefine start from 4)
                if (Test == "IDS")
                {
                    return DcCategoryName.CategoryDefaultValue;
                }
                _pmodePatternVdip = CategoryName.Split('_').Length >= 4 ? CategoryName.Split('_')[3] : DcCategoryName.CategoryDefaultValue;
                return _pmodePatternVdip;
            }
        }

        public string UserDefined
        {
            get
            {
                //update by laura at 2019/09/11 to support IDS category new format(Userdefine start from 4)
                _ = Test == "IDS" ? 4 : 5;
                List<string> splitlst = [.. CategoryName.Split('_')];
                if (splitlst.Count >= 5)
                {
                    splitlst.RemoveAt(0);
                    splitlst.RemoveAt(0);
                    splitlst.RemoveAt(0);
                    splitlst.RemoveAt(0);
                    return string.Join("_", splitlst);
                }
                else
                {
                    return DcCategoryName.CategoryDefaultValue;
                }
            }
        }

        public string Chiplet(PowerInfoSheet powerInfoSheet)
        {
            _chiplet = CategoryName.Split('_').Last();
            if (powerInfoSheet != null)
            {
                if (powerInfoSheet.Rows.Exists(x => x.Chiplet.EqualsIgnoreCase(_chiplet)))
                {
                    return _chiplet;
                }
            }
            return "";
        }

        public string GetStandardCategoryName()
        {
            string result = string.Join(CategoryNameFix, [Test, Domain, Subtest, PmodePatternVdip, UserDefined]);
            return result;
        }
    }
}
