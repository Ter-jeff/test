using System;
using System.Collections.Generic;

using IgxlLib.Enums;
using IgxlLib.IgxlReader;

namespace IgxlLib
{
    internal sealed class IgxlSheetReaderRegistry
    {
        private readonly Dictionary<EnumSheetType, (IIgxlSheetReader Reader, Action<object, string> Assign)> _mappings = [];

        public IgxlSheetReaderRegistry(IgxlLoader igxlLoader)
        {
            IgxlSheetReaderRegistryCore.AddTo(_mappings, igxlLoader);
            IgxlSheetReaderRegistrySpec.AddTo(_mappings, igxlLoader);
            IgxlSheetReaderRegistryPattern.AddTo(_mappings, igxlLoader);
        }

        public (IIgxlSheetReader Reader, Action<object, string> Assign) GetMapping(EnumSheetType enumSheetType)
        {
            if (!_mappings.TryGetValue(enumSheetType, out (IIgxlSheetReader Reader, Action<object, string> Assign) mapping))
            {
                throw new ArgumentException($"Unsupported sheet type: {enumSheetType}");
            }

            return mapping;
        }
    }
}
