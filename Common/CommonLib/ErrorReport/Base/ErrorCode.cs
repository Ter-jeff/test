using System;
using System.ComponentModel;
using System.Reflection;

namespace CommonLib.ErrorReport.Base
{
    /// <summary>
    /// Immutable implementation of <see cref="IErrorCode"/>.
    /// </summary>
    public class ErrorCode : IErrorCode
    {
        private readonly int _subCode;
        public EnumErrorLevel ErrorLevel { get; }
        public EnumErrorCategory EnumErrorCategory { get; }
        public EnumErrorBehavior? EnumErrorBehavior { get; }
        public EnumErrorTarget? EnumErrorTarget { get; }

        public int Code { get; }

        public string FullCode
        {
            get
            {
                string eStr = ErrorLevel.ToString()[..1].ToUpper();

                FieldInfo? fi = EnumErrorCategory
                    .GetType()
                    .GetField(EnumErrorCategory.ToString());

                DescriptionAttribute[] attributes = fi == null
                    ? []
                    : (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

                string ccStr = attributes.Length > 0
                    ? attributes[0].Description
                    : EnumErrorCategory.ToString()[..2].ToUpper();

                string sssStr = _subCode.ToString("D3");
                string nnnStr = Code.ToString("D3");

                string fullCode = $"{eStr}{ccStr}{sssStr}{nnnStr}";

                if (fullCode.Length != 9)
                {
                    throw new InvalidOperationException(
                        $"Generated code '{fullCode}' violates the 9-character constraint.");
                }

                return fullCode;
            }
        }

        public string MessageTemplate { get; }

        public string Guidance { get; }

        /// <summary>
        /// New constructor.
        /// Use EnumErrorBehavior and EnumErrorTarget instead of EnumErrorSubGroup.
        /// </summary>
        public ErrorCode(
            EnumErrorCategory enumErrorCategory,
            EnumErrorBehavior enumErrorBehavior,
            EnumErrorTarget enumErrorTarget,
            int code,
            string template,
            EnumErrorLevel enumErrorLevel,
            string? guidance = null)
        {
            ErrorLevel = enumErrorLevel;
            EnumErrorCategory = enumErrorCategory;
            EnumErrorBehavior = enumErrorBehavior;
            EnumErrorTarget = enumErrorTarget;
            Code = code;
            MessageTemplate = template ?? string.Empty;
            Guidance = guidance ?? string.Empty;
            _subCode = ((int)EnumErrorBehavior * 100) + (int)EnumErrorTarget;
        }

        /// <summary>
        /// Format the message using the template and provided args.
        /// Returns the raw template if formatting fails.
        /// </summary>
        public string FormatMessage(params object[]? args)
        {
            if (args == null || args.Length == 0)
            {
                return MessageTemplate;
            }

            try
            {
                return string.Format(MessageTemplate, args);
            }
            catch
            {
                return MessageTemplate;
            }
        }
    }
}
