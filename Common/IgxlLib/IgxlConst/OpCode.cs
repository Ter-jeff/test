// ReSharper disable UnusedMember.Global

using System.Collections.Generic;

using CommonLib.Extension;

namespace IgxlLib.IgxlConst
{
    public class OpCode
    {
        public const string Test = "Test";
        public const string BinTable = "Bintable";
        public const string Call = "call";
        public const string If = "If";
        public const string Else = "Else";
        public const string ElseIf = "ElseIf";
        public const string EndIf = "EndIf";
        public const string FlagFalse = "flag-false";
        public const string FlagFalseAll = "flag-false-all";
        public const string FlagTrue = "flag-true";
        public const string FlagClear = "flag-clear";
        public const string FlagClearAll = "flag-clear-all";
        public const string Nop = "nop";
        public const string Print = "print";
        public const string CreateSiteVar = "create-site-var";
        public const string AssignSiteVar = "assign-site-var";
        public const string Return = "Return";
        public const string UseLimit = "Use-Limit";
        public const string Characterize = "characterize";
        public const string SetDevice = "set-device";
        public const string EnableFlowWd = "enable-flow-word";
        public const string DisableFlowWd = "disable-flow-word";
        public const string Loop = "loop";
        public const string EndLoop = "endloop";

        public const string Concurrent = "concurrent";
        public const string For = "For";
        public const string Next = "Next";
        public const string Break = "Break";
        public const string AssignInteger = "assign-integer";
        public const string CreateInteger = "create-integer";
        public const string SetAlarmBin = "set-alarm-bin";
        public const string SetErrorBin = "set-error-bin";

        public const string LimitsAll = "limits-all";
        public const string Stop = "stop";

        public static readonly HashSet<string> AllOpCodes = new(StringExtensions.IgnoreCase)
        {
            Test,
            BinTable,
            Call,
            If,
            Else,
            ElseIf,
            EndIf,
            FlagFalse,
            FlagFalseAll,
            FlagTrue,
            FlagClear,
            FlagClearAll,
            Nop,
            Print,
            CreateSiteVar,
            AssignSiteVar,
            Return,
            UseLimit,
            Characterize,
            SetDevice,
            EnableFlowWd,
            DisableFlowWd,
            Loop,
            EndLoop,
            Concurrent,
            For,
            Next,
            Break,
            AssignInteger,
            CreateInteger,
            SetAlarmBin,
            SetErrorBin,
            LimitsAll,
            Stop
        };
    }
}
