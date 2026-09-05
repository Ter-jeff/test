using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Utility;

using LogLib.Utility;

using DataTable = System.Data.DataTable;

namespace Automation.Utility
{
    public class StepEvaluator
    {
        public static string ErrMSg;

        public static int Evaluate(string inStr)
        {
            ErrMSg = "";
            string[] tmp = inStr.Replace(" ", "").Split(',');
            string start = Replace(Replace(tmp[0], "-"), "+");
            string stop = Replace(Replace(tmp[1], "-"), "+");
            if (!Regex.IsMatch(start, "^-"))
            {
                start = "+" + start;
            }

            if (!Regex.IsMatch(stop, "^-"))
            {
                stop = "+" + stop;
            }

            var lstep = new List<string> { tmp[2] };
            List<string> lstart = Regex.Split(start, ",").ToList();
            List<string> lstop = Regex.Split(stop, ",").ToList();
            var lcheck = lstart.Select(a => a).ToList();
            foreach (string item in lcheck)
            {
                if (lstop.Contains(item))
                {
                    lstart.Remove(item);
                    lstop.Remove(item);
                }
            }
            double sweepRange = EvaluateExpression(GetCalStr(lstart, lstop));
            double stepSize = EvaluateExpression(CalStr(lstep));
            return Convert.ToInt32(sweepRange / stepSize);
        }

        private static double EvaluateExpression(string eqn)
        {
            var dt = new DataTable();
            try
            {
                object result = dt.Compute(eqn, string.Empty);
                return Convert.ToDouble(result);
            }
            catch (Exception ex)
            {
                ErrMSg = "Error! Can't do datatable EvaluateExpression! >> " + eqn;
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                return 0;
            }
        }

        internal static string GetCalStr(List<string> l1, List<string> l2)
        {
            string evalStrSt = CalStr(l1);
            string evalStrSp = CalStr(l2);
            return evalStrSp + "-(" + evalStrSt + ")";
        }

        internal static string CalStr(List<string> l1)
        {
            string evalStr = "";
            foreach (string item in l1)
            {
                string tmp = Regex.Replace(item, "V|A|Hz|OHM|S", "", RegexOptions.IgnoreCase);
                Match mc = Regex.Match(tmp, @"\d+(?<unit>[a-zA-Z])", RegexOptions.IgnoreCase);
                if (mc.Success)
                {
                    string scale = mc.Groups["unit"].ToString();
                    evalStr += tmp.Replace(scale, "*" + Math.Pow(10, UnitUtility.GetScale(mc.Groups["unit"].ToString())));
                }
                else
                {
                    evalStr += tmp;
                }
            }
            return evalStr;
        }

        internal static string Replace(string inStr, string rpls)
        {
            inStr = Regex.Replace(inStr, "\\" + rpls, "," + rpls);
            return inStr;
        }
    }
}
