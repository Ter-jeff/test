using System;
using System.Collections.Generic;

using CommonLib.Enums;

using CommonReaderLib;

using ProjectConfigLib.ProjectConfig;

namespace Automation.Static
{
    public class Options
    {
        // ---------------- ProjectInfo ----------------
        public EnumDevice Device { get; set; } = EnumDevice.AP;
        public string FrcCalcType { get; } = "Common";

        // ---------------- Basic ----------------
        public bool IsIgnorePatternCheck { get; set; }
        public bool GenerateT0TXTestprogram { get; set; }
        public bool VreEnable { get; set; }
        public bool SkipIgLink { get; set; }
        public EnumInstrument InstrumentType { get; set; } = EnumInstrument.UFPlus;

        public bool IsUfp
        {
            get
            {
                switch (InstrumentType)
                {
                    case EnumInstrument.UFlex_UFPlus:
                        return true;

                    case EnumInstrument.UFPlus:
                        return true;

                    case EnumInstrument.UFlex:
                        return false;

                    default:
                        return false;
                }
            }
        }

        public bool IsSplitDcSpecs { get; }
        public bool ConvertJobToEnableWord { get; }
        public bool GenerateVSlewRate { get; }

        public int SafetyLineRowLimit { get; set; } = 50000;

        // ---------------- Scan ----------------

        // ---------------- Mbist ----------------
        public string AdpativeCoolingBlock { get; set; } = string.Empty;
        public string AdpativeCoolingSuBlock { get; set; } = string.Empty;
        public string CoolingLimitValue { get; } = "85C,100,105";

        // ---------------- eFuseConfig ----------------
        public bool MergeInitPatterns { get; set; }

        // ---------------- HardIP ----------------
        public bool SplitCzFlow { get; } = true;
        public string CharacterizationType { get; } = "CZ,DD";
        public bool SplitPatBurstPage { get; }
        public bool NoBinOutForAll { get; }
        public bool BinTableBeforeEachItem { get; }
        public bool SearchValidPatternRev { get; set; }
        public bool BypassPreCheck { get; set; }
        public bool StopByCritical { get; set; } = true;
        public string PseudoFuseFilePath { get; set; }

        public Options()
        {
        }

        public Options(ProjectConfigSingleton projectSingleton)
        {
            Device = ToDeviceEnum(projectSingleton.GetValue("Device", "Type"));

            FrcCalcType = projectSingleton.GetValue("Device", "FRCCalcType");
            InstrumentType = ToEnumInstrument(projectSingleton.GetValue("Basic", "IsUFP"));
            IsSplitDcSpecs = projectSingleton.GetBool("Basic", "SplitDcSpecs");
            ConvertJobToEnableWord = projectSingleton.GetBool("Basic", "ConvertJobToEnableWord");
            GenerateVSlewRate = projectSingleton.GetBool("Basic", "Generate VSlewRate");
            SafetyLineRowLimit = projectSingleton.GetInt("Basic", "SafetyLineRowLimit", 50000);

            AdpativeCoolingBlock = projectSingleton.GetValue("Mbist", "AdpativeCoolingBlock");
            AdpativeCoolingSuBlock = projectSingleton.GetValue("Mbist", "AdpativeCoolingSuBlock");
            CoolingLimitValue = projectSingleton.GetValue("Mbist", "CoolingLimitValue");

            MergeInitPatterns = projectSingleton.GetBool("eFuseConfig", "MergeInitPatterns");

            SplitCzFlow = projectSingleton.GetBool("HardIP", "SplitCzFlow");
            CharacterizationType = projectSingleton.GetValue("HardIP", "CharacterizationType");
            SplitPatBurstPage = projectSingleton.GetBool("HardIP", "SplitPatBurstPage");
            NoBinOutForAll = projectSingleton.GetBool("HardIP", "NoBinOutForAll");
            BinTableBeforeEachItem = projectSingleton.GetBool("HardIP", "BinTableBeforeEachItem");

            PseudoFuseFilePath = projectSingleton.GetValue("eFuseConfig", "PseudoFuseFilePath");
        }

        private EnumDevice ToDeviceEnum(string text)
        {
            if (Enum.TryParse(text, true, out EnumDevice parsed))
            {
                return parsed;
            }

            return EnumDevice.AP;
        }

        public Options(InputInfoSheet sheet)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (sheet?.Rows != null)
            {
                foreach (InputInfoRow row in sheet.Rows)
                {
                    if (!string.IsNullOrEmpty(row.Item))
                    {
                        dict[row.Item.Trim()] = row.Value?.Trim() ?? string.Empty;
                    }
                }
            }

            string Get(string key, string defaultValue = "")
            {
                return dict.TryGetValue(key, out string value) && !string.IsNullOrEmpty(value)
                    ? value
                    : defaultValue;
            }

            int GetInt(string key, int defaultValue = 0)
            {
                string value;
                if (dict.TryGetValue(key, out value) && !string.IsNullOrEmpty(value))
                {
                    int parsed;
                    return int.TryParse(value, out parsed) ? parsed : defaultValue;
                }

                return defaultValue;
            }

            bool GetBool(string key, bool defaultValue = false)
            {
                string v = Get(key);
                if (string.IsNullOrEmpty(v))
                {
                    return defaultValue;
                }

                return v.Equals("1", StringComparison.OrdinalIgnoreCase)
                       || v.Equals("true", StringComparison.OrdinalIgnoreCase)
                       || v.Equals("yes", StringComparison.OrdinalIgnoreCase)
                       || v.Equals("y", StringComparison.OrdinalIgnoreCase);
            }

            // --- Basic ---
            // excel text/number format issue
            Device = ToEnumDevice(Get("Type", "AP"));

            FrcCalcType = Get("FRCCalcType", "Common");
            IsIgnorePatternCheck = GetBool("skippatcheck");
            GenerateT0TXTestprogram = GetBool("GenerateT0TXTestprogram");
            InstrumentType = ToEnumInstrument(Get("InstrumentType", "UFlex"));
            IsSplitDcSpecs = GetBool("SplitDcSpecs");
            ConvertJobToEnableWord = GetBool("ConvertJobToEnableWord", true);
            GenerateVSlewRate = GetBool("Generate VSlewRate");
            SafetyLineRowLimit = GetInt("SafetyLineRowLimit", 50000);
            VreEnable = GetBool("vre_enable");

            // --- Scan ---

            // --- Mbist ---
            AdpativeCoolingBlock = Get("AdpativeCoolingBlock");
            AdpativeCoolingSuBlock = Get("AdpativeCoolingSuBlock");
            CoolingLimitValue = Get("CoolingLimitValue", "85C,100,105");

            // --- eFuseConfig ---
            MergeInitPatterns = GetBool("MergeInitPatterns");

            // --- HardIP ---
            SplitCzFlow = GetBool("SplitCzFlow");
            CharacterizationType = Get("CharacterizationType", "CZ,DD");
            SplitPatBurstPage = GetBool("SplitPatBurstPage");
            NoBinOutForAll = GetBool("NoBinOutForAll");
            BinTableBeforeEachItem = GetBool("BinTableBeforeEachItem");


            SearchValidPatternRev = GetBool("SearchValidPatternRev");
            BypassPreCheck = GetBool("BypassPreCheck");
            StopByCritical = GetBool("StopByCritical", true);


            PseudoFuseFilePath = Get("PseudoFuseFilePath");
        }

        private EnumDevice ToEnumDevice(string text)
        {
            if (Enum.TryParse(text, true, out EnumDevice parsed))
            {
                return parsed;
            }

            return EnumDevice.AP;
        }

        private EnumInstrument ToEnumInstrument(string text)
        {
            if (Enum.TryParse(text, true, out EnumInstrument parsed))
            {
                return parsed;
            }

            return text.Equals("True", StringComparison.CurrentCultureIgnoreCase) ? EnumInstrument.UFPlus : EnumInstrument.UFlex;
        }
    }
}
