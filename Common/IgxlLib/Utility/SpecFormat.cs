namespace IgxlLib.Utility
{
    public static class SpecFormat
    {
        public const string GlbSuffix = "GLB";
        public const string VarSuffix = "VAR";
        public const string SpecValuePrefixEqual = "=";
        public const string SpecValuePrefix = "_";
        public const string DcSpecVopSuffix = "VOP";
        public const string Ifold = "Ifold";
        public const string IfoldEvs = "IfoldEVS";
        public const string IfoldConti = "IfoldConti";
        public const string IfoldHip = "IfoldHIP";

        public static string GenGlbSpecSymbol(string pinName)
        {
            return CreateSpecSymbol(pinName, GlbSuffix);
        }

        public static string GenAcSpecSymbol(string pinName)
        {
            return CreateSpecSymbol(pinName, VarSuffix);
        }

        public static string GenGlbRatio(string global, string ratio)
        {
            return "=_" + global + "*_" + ratio;
        }

        public static string CreateSpecSymbol(string src, string suffix)
        {
            return string.IsNullOrEmpty(suffix) ? src.Trim() : src.Trim() + "_" + suffix;
        }

        public static string GenGlbSpecSymbolAtLevelSheet(string pinName, string parameterName, string block = "")
        {
            string src = "_" + pinName + "_" + parameterName;
            if (!string.IsNullOrEmpty(block))
            {
                src += "_" + block;
            }

            return CreateSpecSymbol(src, GlbSuffix);
        }

        public static string GenDcSpecSymbolAtLevelSheet(string pinName, string parameterName, string parameterSyntax, string blockType = "", string chiplet = "")
        {
            string src = $"_{pinName}";
            if (!string.IsNullOrEmpty(parameterName))
            {
                src += $"_{parameterName}";
            }
            if (!string.IsNullOrEmpty(parameterSyntax))
            {
                src += $"_{parameterSyntax}";
            }
            src += $"_{VarSuffix}";
            if (!string.IsNullOrEmpty(blockType))
            {
                src += $"_{blockType}";
            }
            if (!string.IsNullOrEmpty(chiplet))
            {
                src += $"_{chiplet}";
            }
            return src;
        }

        public static string GenSpecValueSingleValue(string str)
        {
            // Add "=" before old value
            return SpecValuePrefixEqual + str;
        }
    }
}
