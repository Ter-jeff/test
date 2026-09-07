using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IGDataXML.SheetClassMapping;
using IgxlLib.IgxlSheets;

namespace IgxlLib.Utility
{
    public static class IgxlLoaderHelpers
    {
        private static readonly ConcurrentDictionary<string, IGXL?> _igxlConfigCache = new();
        private static readonly Dictionary<string, EnumSheetType> _sheetTypeMappings = new(StringExtensions.IgnoreCase)
        {
            { "DTFlowtableSheet", EnumSheetType.DTFlowtableSheet },
            { "DTTestInstancesSheet", EnumSheetType.DTTestInstancesSheet },
            { "DTDCSpecSheet", EnumSheetType.DTDCSpecSheet },
            { "DTACSpecSheet", EnumSheetType.DTACSpecSheet },
            { "DTLevelSheet", EnumSheetType.DTLevelSheet },
            { "DTGlobalSpecSheet", EnumSheetType.DTGlobalSpecSheet },
            { "DTTimesetBasicSheet", EnumSheetType.DTTimesetBasicSheet },
            { "DTBintablesSheet", EnumSheetType.DTBintablesSheet },
            { "DTChanMap", EnumSheetType.DTChanMap },
            { "DTCharacterizationSheet", EnumSheetType.DTCharacterizationSheet },
            { "DTJobListSheet", EnumSheetType.DTJobListSheet },
            { "DTPatternSetSheet", EnumSheetType.DTPatternSetSheet },
            { "DTPatternSubroutineSheet", EnumSheetType.DTPatternSubroutineSheet },
            { "DTReferencesSheet", EnumSheetType.DTReferencesSheet },
            { "DTPinMap", EnumSheetType.DTPinMap },
            { "DTPortMapSheet", EnumSheetType.DTPortMapSheet }
        };

        public static string GetSheetVersion(IIgxlSheet igxlSheet, string igxlVersion)
        {
            IGXL? igxlConfig = _igxlConfigCache.GetOrAdd(igxlVersion, LoadConfigForVersion);
            if (igxlConfig != null)
            {
                return igxlConfig.SheetItemClass.First(p => p.sheetname == igxlSheet.IgxlSheetName).sheetversion;
            }

            return "";
        }

        private static IGXL? LoadConfigForVersion(string igxlVersion)
        {
            var assembly = Assembly.GetExecutingAssembly();
            string[] resourceNames = assembly.GetManifestResourceNames();
            IGXL? igxlConfig = null;
            foreach (string resourceName in resourceNames)
            {
                if (resourceName.Contains(".SheetClassMapping.v" + igxlVersion + "_ultraflex.xml"))
                {
                    igxlConfig = LoadConfig(assembly.GetManifestResourceStream(resourceName)!);
                }
            }
            return igxlConfig;
        }

        private static IGXL? LoadConfig(Stream stream)
        {
            IGXL? result;
            try
            {
                var xs = new XmlSerializer(typeof(IGXL));
                var sysData = (IGXL?)xs.Deserialize(stream);
                stream.Close();
                result = sysData;
            }
            catch (Exception e)
            {
                throw new Exception($"Failed to load IGXL sheet class mapping config: {e.Message}", e);
            }
            return result;
        }

        public static EnumSheetType GetIgxlSheetType(string? text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return EnumSheetType.DTUnknown;
            }

            string? matchedKey = _sheetTypeMappings.Keys.FirstOrDefault(text.ContainsIgnoreCase);
            if (matchedKey != null)
            {
                return _sheetTypeMappings[matchedKey];
            }

            if (text.StartsWithIgnoreCase("Rev:") && text.ContainsIgnoreCase("Bin Cut"))
            {
                return EnumSheetType.BinCut_eqn;
            }

            return EnumSheetType.DTUnknown;
        }

        public static EnumSheetType GetIgxlSheetType(Stream stream, string file)
        {
            string? firstLine;
            using (var sr = new StreamReader(stream))
            {
                firstLine = sr.ReadLine();
            }
            string ext = Path.GetExtension(file);
            if (ext.EqualsIgnoreCase(".bas"))
            {
                return EnumSheetType.Bas;
            }

            if (ext.EqualsIgnoreCase(".cls"))
            {
                return EnumSheetType.Cls;
            }

            return GetIgxlSheetType(firstLine);
        }
    }
}
