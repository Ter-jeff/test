#pragma warning disable

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace TerLintAnalyzer.Test
{
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "<Pending>")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0059:Unnecessary assignment of a value", Justification = "<Pending>")]
    public class SampleTest
    {
        public class IfCondition
        {
        }

        public const int ConstValue = 1;

        private static int _returnValue;
        public static int ReturnValue
        {
            get
            {
                return _returnValue + 1;
            }
            set
            {
                _returnValue = value;
            }
        }

        public int x1 = 1; // This is a trailing comment

        public int Comment()
        {
            int x = 1; // This is a trailing comment

            // This is a valid comment
            int y = 1;

            int z = 1; /* Multi-line trailing */

            return x + y + z;
        }

        public void MethodArgumentNaming(MyClass myClass, List<MyClass> myClasss, MyClass[] myClassArray)
        {
        }

        public void MethodArgumentNaming1(MyClass row, List<MyClass> users, MyClass[] userArray)
        {
        }

        public void MethodArgumentNaming2(MyClass row1, MyClass row2)
        {
        }
        public void MethodSpacing1()
        {
        }


        public void MethodSpacing2()
        {
        }
        // This is a comment
        public void MethodSpacing3()
        {
        }

        // This is a comment
        public void MethodSpacing4()
        {
        }

        #region MethodSpacing
        public void MethodSpacing5()
        {
        }
        #endregion

        public void MethodSpacing6()
        {
        }

        public void MethodWrapping(int a, int b)
        {
        }

        public void MethodWrapping1(
            int a,
            int b)
        {
        }

        public int MethodWrapping2(int a,
                int b) => a + b;

        [Obsolete]
        public int MethodWrapping3()
        {
            return 1;
        }

        private static readonly Regex _regex = new Regex("BI", RegexOptions.IgnoreCase);
        private static readonly Regex _regex1 = new Regex("RPI", RegexOptions.IgnoreCase);
        private static readonly Regex _regex2 = new Regex("_SI_", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public bool PreferStringMethods()
        {
            var r = new Regex("SimpleText");
            r.IsMatch("input");

            var r1 = new Regex("^SimpleText");
            r1.IsMatch("input");

            var r2 = new Regex("SimpleText$");
            r2.IsMatch("input");

            var r3 = new Regex("^SimpleText$");
            r3.IsMatch("input");

            string[] patItems = "input".Split('_');
            return _regex.IsMatch(patItems[6]) && _regex1.IsMatch(patItems[4]);
        }

        public void PreferStringMethods2()
        {
            var r = new Regex("SimpleText", RegexOptions.IgnoreCase);
            r.IsMatch("input");

            var r1 = new Regex("^SimpleText", RegexOptions.IgnoreCase);
            r1.IsMatch("input");

            var r2 = new Regex("SimpleText$", RegexOptions.IgnoreCase);
            r2.IsMatch("input");

            var r3 = new Regex("^SimpleText$", RegexOptions.IgnoreCase);
            bool _ = r3.IsMatch("input");
        }

        public void RegexInline()
        {
            bool result = Regex.IsMatch("input", "abc");
            bool result1 = Regex.IsMatch("input1", "abc1");
            bool result2 = Regex.IsMatch("input2", "[a-z]");
        }

        public void RegexInline1()
        {
            bool result = Regex.IsMatch("input", "abc");
            bool result2 = Regex.IsMatch("input2", "abc2");
            bool result3 = Regex.IsMatch("input3", "abc3");
            bool result4 = Regex.IsMatch("input4", "[a-z]");
            bool result5 = _regex2.IsMatch("input5");
        }

        public void RegexInline2()
        {
            string pattern = "abc";
            bool result = Regex.IsMatch("input", pattern);
        }

        public void RegexInline3()
        {
            var r = new Regex("^[a - z]+$");
            bool result = r.IsMatch("input");
        }

        private bool IsMatch(string a, string b) => true;

        public void RegexInline4()
        {
            bool result = IsMatch("input", "pattern");
        }

        public void StringAnalyzer()
        {
            string s1 = "Hello";
            string s2 = "hello";

            // Trigger: StringEqualsCaseInsensitive
            bool c1 = string.Equals(s1, s2, StringComparison.OrdinalIgnoreCase);
            bool c2 = !string.Equals(s1, s2, StringComparison.OrdinalIgnoreCase);

            bool c3 = string.Equals(s1, s2);
            bool c4 = !string.Equals(s1, s2);

            s1.EqualsIgnoreCase(s2);

            // Trigger: EqualsCaseInsensitive
            bool c5 = s1.Equals(s2, StringComparison.InvariantCultureIgnoreCase);
            bool c6 = !s1.Equals(s2, StringComparison.InvariantCultureIgnoreCase);

            bool c7 = s1.Equals(s2);
            bool c8 = !s1.Equals(s2);

            bool c51 = s1.StartsWith(s2, StringComparison.InvariantCultureIgnoreCase);
            bool c52 = s1.EndsWith(s2, StringComparison.InvariantCultureIgnoreCase);
            if (!("".StartsWith("DCVS", StringComparison.CurrentCultureIgnoreCase) || "".StartsWith("DCVI", StringComparison.CurrentCultureIgnoreCase)))
            {
            }

            // Trigger: ToUpperToLowerComparison
            s1.ToLower().StartsWith("h");

            // Trigger: ToUpperToLowerEqualityOperator
            bool c9 = s1.ToUpper() == "HELLO";
            bool c10 = s1.ToUpper() != "HELLO";

            bool c11 = s1 == "HELLO".ToLower();
            bool c120 = s1 != "HELLO".ToLower();

            // Trigger: StringComparerUsage
            StringComparer c = StringComparer.OrdinalIgnoreCase;

            Dictionary<string, string> pinNameList = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if ("srcFile".ToLower().IndexOf(".txt", StringComparison.OrdinalIgnoreCase) != -1)
            {
            }

            int rowind = new List<string>().FindIndex(row => row.ToLower() == "call" && row == "call2");

            bool r = "a" != null;
            var myClass = new MyClass();
            bool r1 = myClass.PatternSet == null;

            foreach (var patRow in myClass.PatSetRows)
            {
                if (patRow.PatternSet == null)
                {
                    continue;
                }
            }
        }

        public void ExcessiveSpacing()
        {
            int x = 1;
        }
    }

    public class MyClass
    {
        public string PatternSet { get; set; }
        public List<MyRow> PatSetRows { get; set; }
    }

    public class MyRow
    {
        public string PatternSet { get; set; }
    }

    enum MyEnum
    {
        Value1,
        Value2,
        Value3
    }
}
