using System.Text.RegularExpressions;

using ScghLib.Enums;
using ScghLib.Reader;

namespace ScghLib.Utility
{
    public partial class BistAction
    {
        private static readonly Regex _regex1 = BistActionHelpers.MyRegex();
        private static readonly Regex _regex = BistActionHelpers.MyRegex1();

        public static string SetActionParameter(BistProdFlowRow bistProdFlowRow)
        {
            string parameter = BistActionHelpers.MyRegex3().Match(bistProdFlowRow.Action).Groups["str"].ToString();
            return parameter.Trim().ToUpper();
        }

        public static string SetDefaultActionParameter(BistProdFlowRow bistProdFlowRow)
        {
            string parameter = BistActionHelpers.MyRegex2().Match(bistProdFlowRow.Action).Groups["str"].ToString();
            return parameter.Trim().ToUpper();
        }

        public static int SetActionValue(BistProdFlowRow bistProdFlowRow)
        {
            string value = BistActionHelpers.MyRegex4().Match(bistProdFlowRow.Action).Groups["str1"].ToString();
            return int.Parse(value);
        }

        public static string GetActionParameter(BistProdFlowRow bistProdFlowRow)
        {
            string parameter = BistActionHelpers.MyRegex5().Match(bistProdFlowRow.Action).Groups["str"].ToString();
            if (string.IsNullOrEmpty(parameter))
            {
                parameter = "(" + BistActionHelpers.MyRegex7().Match(bistProdFlowRow.Action).Groups["str"] + ")";
            }

            return parameter.Trim().ToUpper();
        }

        public static string GetRetentionTime(BistProdFlowRow bistProdFlowRow)
        {
            string waitTime = "20";
            if (_regex.IsMatch(bistProdFlowRow.Action))
            {
                waitTime = BistActionHelpers.MyRegex6().Match(bistProdFlowRow.Action).Groups["str"].ToString();
            }

            return waitTime;
        }

        public static string GetRetentionRampStep(BistProdFlowRow bistProdFlowRow)
        {
            string rampStep = "5";
            if (_regex1.IsMatch(bistProdFlowRow.Action))
            {
                rampStep = BistActionHelpers.MyRegex().Match(bistProdFlowRow.Action).Groups["str"].ToString();
            }
            return rampStep;
        }

        public static string GetLoopNum(BistProdFlowRow bistProdFlowRow)
        {
            string loopNum = "";
            if (BistActionHelpers.MyRegex8().IsMatch(bistProdFlowRow.Note))
            {
                loopNum = BistActionHelpers.MyRegex9().Match(bistProdFlowRow.Note).Groups["str"].ToString();
            }
            return loopNum;
        }

        public static string GetDomainByAction(BistProdFlowRow bistProdFlowRow)
        {
            string domain = "";
            if (BistActionHelpers.MyRegex11().IsMatch(bistProdFlowRow.Action))
            {
                domain = BistActionHelpers.MyRegex10().Match(bistProdFlowRow.Action).Groups["str"].ToString();
            }
            return domain;
        }

        public static BistActionType GetActionType(BistProdFlowRow bistProdFlowRow)
        {

            if (BistActionHelpers.MyRegex15().IsMatch(bistProdFlowRow.Note))
            {
                return BistActionType.LoopStart;
            }
            if (BistActionHelpers.MyRegex16().IsMatch(bistProdFlowRow.Note))
            {
                return BistActionType.LoopEnd;
            }
            if (!string.IsNullOrEmpty(bistProdFlowRow.Pattern) && !(BistActionHelpers.MyRegex13().IsMatch(bistProdFlowRow.Action) || BistActionHelpers.MyRegex12().IsMatch(bistProdFlowRow.Action)))
            {
                return BistActionType.RunPattern;
            }
            if (BistActionHelpers.MyRegex14().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.Check;
            }
            if (BistActionHelpers.MyRegex17().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.DomainStart;
            }
            if (BistActionHelpers.MyRegex18().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.DomainEnd;
            }
            if (BistActionHelpers.MyRegex19().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.FailCheck;
            }
            if (BistActionHelpers.MyRegex20().IsMatch(bistProdFlowRow.Action) || BistActionHelpers.MyRegex21().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.FailCheckScan;
            }
            if (BistActionHelpers.MyRegex22().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.Fail;
            }
            if (BistActionHelpers.MyRegex23().IsMatch(bistProdFlowRow.Action) || BistActionHelpers.MyRegex24().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.Get;
            }
            if (BistActionHelpers.MyRegex25().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.Pass;
            }
            if (BistActionHelpers.MyRegex26().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.Set;
            }
            if (BistActionHelpers.MyRegex27().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.SetDefault;
            }
            if (BistActionHelpers.MyRegex28().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.Retention;
            }
            if (BistActionHelpers.MyRegex29().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.RetentionVoltDrop;
            }
            if (BistActionHelpers.MyRegex30().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.CallSubFlow;
            }
            if (BistActionHelpers.MyRegex31().IsMatch(bistProdFlowRow.Action))
            {
                return BistActionType.SetSitVar;
            }
            else
            {
                return BistActionType.Null;
            }
        }
    }
}
