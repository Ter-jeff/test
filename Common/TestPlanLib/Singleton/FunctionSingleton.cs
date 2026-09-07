using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace TestPlanLib.Singleton
{
    public sealed partial class FunctionSingleton
    {
        #region Singleton
        private const string PattenIsNumber = "^[0-9]+([.]{1}[0-9]+){0,1}";

        [GeneratedRegex(PattenIsNumber)]
        private static partial Regex MyRegex();

        // Fields
        private static FunctionSingleton? _instance;

        // Constructor
        private FunctionSingleton()
        {
        }

        // Methods
        public static FunctionSingleton Instance()
        {
            return _instance ??= new FunctionSingleton();
        }

        #endregion
        #region public method

        /// <summary>
        /// FormatStringForCompareNoUpper
        /// </summary>
        /// <param name="pString">pString</param>
        /// <returns></returns>
        public static string Normalization(string pString)
        {
            string lStrResult = pString.Trim();

            lStrResult = ReplaceEnter(lStrResult);

            lStrResult = ReplaceDoubleBlank(lStrResult);

            return lStrResult;
        }

        /// <summary>
        /// ReplaceEnter
        /// </summary>
        /// <param name="pString">pString</param>
        /// <returns></returns>
        public static string ReplaceSpecialCase(string pString)
        {
            return Regex.Escape(pString);
        }

        /// <summary>
        /// CompareStringList
        /// </summary>
        /// <param name="pSource">pSource</param>
        /// <param name="pCompare">pCompareList</param>
        /// <returns></returns>
        public static bool CompareString(string pSource, string pCompare, bool pIgnoreUnderLine = false)
        {
            if (string.IsNullOrEmpty(pSource))
            {
                return false;
            }

            string lstr1;
            string lstr2;

            foreach (string pCompareSplit in pCompare.Split('|'))
            {
                lstr1 = Normalization(pSource);
                lstr2 = Normalization(pCompareSplit);

                if (pIgnoreUnderLine)
                {
                    lstr1 = lstr1.Replace(' ', '_');
                    lstr2 = lstr2.Replace(' ', '_');
                }

                lstr2 = ReplaceSpecialCase(lstr2);

                if (lstr2.Contains("\\*") ||
                    lstr2.Contains("\\?"))
                {
                    lstr2 = lstr2.Replace("\\*", ".*");
                    lstr2 = lstr2.Replace("\\?", ".+");
                }
                lstr2 = "^" + lstr2 + "$";
                if (Regex.IsMatch(lstr1, lstr2, RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// IsMatch
        /// </summary>
        /// <param name="pSource"></param>
        /// <param name="pPatten"></param>
        /// <returns></returns>
        public static bool IsMatch(string pSource, string pPatten, bool pIgnoreUnderLine = false)
        {
            string lstr1 = Normalization(pSource);
            string lstr2 = pPatten;

            if (pIgnoreUnderLine)
            {
                lstr1 = lstr1.Replace(' ', '_');
                lstr2 = lstr2.Replace(' ', '_');
            }

            return Regex.IsMatch(lstr1, lstr2, RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// CompareStringList
        /// </summary>
        /// <param name="pSource">pSource</param>
        /// <param name="pCompareList">pCompareList</param>
        /// <returns></returns>
        public static bool CompareStringList(string pSource, List<string> pCompareList, bool pIgnoreUnderLine = false)
        {
            foreach (string compareString in pCompareList)
            {
                if (CompareString(pSource, compareString, pIgnoreUnderLine))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// IsInt
        /// </summary>
        /// <param name="pString">pString</param>
        /// <returns></returns>
        public static bool IsNumber(string pString)
        {
            return MyRegex().IsMatch(pString);
        }

        #endregion
        /// <summary>
        /// ReplaceDoubleBlank
        /// </summary>
        /// <param name="pString">pString</param>
        /// <returns></returns>
        public static string ReplaceDoubleBlank(string pString)
        {
            string lStrResult = pString;
            do
            {
                lStrResult = lStrResult.Replace("  ", " ");
            } while (lStrResult.Contains("  "));
            return lStrResult;
        }

        /// <summary>
        /// ReplaceEnter
        /// </summary>
        /// <param name="pString">pString</param>
        /// <returns></returns>
        public static string ReplaceEnter(string pString)
        {
            return pString.Replace("\n", " ");
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <param name="accuracyBit"></param>
        /// <returns></returns>
        public static decimal Divide(decimal left, decimal right, int accuracyBit)
        {
            return right == 0 ? 0 : decimal.Round(decimal.Divide(left, right), accuracyBit);
        }
    }
}
