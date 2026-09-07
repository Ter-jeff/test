using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class IgxlSheetFactory
    {
        private readonly Dictionary<EnumSheetType, IIgxlSheetReader> _readers;

        public IgxlSheetFactory()
        {
            // This will now successfully find all your classes (like ReadLevelSheet)
            _readers = typeof(IIgxlSheetReader).Assembly.GetTypes()
                .Where(p => !p.IsInterface && !p.IsAbstract && p.IsClass)
                .Where(p => typeof(IIgxlSheetReader).IsAssignableFrom(p))
                .Select(p => (IIgxlSheetReader)Activator.CreateInstance(p)!)
                .ToDictionary(
                    reader =>
                    {
                        PropertyInfo? prop = reader.GetType().GetProperty("SupportedType");
                        if (prop != null)
                        {
                            return (EnumSheetType)prop.GetValue(reader)!;
                        }
                        return reader.SupportedType;
                    },
                    reader => reader
                );
        }

        public IIgxlSheet? CreateSheet(EnumSheetType enumSheetType, ExcelWorksheet excelWorksheet)
        {
            if (_readers.TryGetValue(enumSheetType, out IIgxlSheetReader? reader))
            {
                MethodInfo? method = reader.GetType().GetMethod("ReadSheet", [typeof(ExcelWorksheet)]);
                if (method != null)
                {
                    object? result = method.Invoke(reader, [excelWorksheet]);
                    return (IIgxlSheet?)result;
                }
            }
            return null;
        }
    }
}
