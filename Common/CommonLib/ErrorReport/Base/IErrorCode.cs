namespace CommonLib.ErrorReport.Base
{
    public interface IErrorCode
    {
        /// <summary>Structured code: CCSSSNNN (e.g. "PM001001").</summary>
        string FullCode { get; }

        /// <summary>
        /// Format string for the one-line error message.
        /// Use {0}, {1} … placeholders — never embed runtime values here.
        /// Call <see cref="ErrorCode.FormatMessage"/> at the call site to produce the final string.
        /// </summary>
        string MessageTemplate { get; }

        /// <summary>Fixed severity for this error code. Callers do not supply ErrorLevel separately.</summary>
        EnumErrorLevel ErrorLevel { get; }

        /// <summary>
        /// Human-readable debugging guidance: root cause, where to look, how to fix.
        /// See ErrorCode.Rules.md for writing rules.
        /// </summary>
        string Guidance { get; }
    }
}
