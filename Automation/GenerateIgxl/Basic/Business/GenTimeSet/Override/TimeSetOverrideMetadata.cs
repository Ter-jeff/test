using System.Collections.Generic;
using System.Linq;

using CommonLib.Tables;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Dto;

public sealed class TimeSetOverrideMetadata
{
    public string SheetName { get; init; }

    public int HeaderRow { get; init; }

    private readonly Dictionary<string, ColumnMetadata> _metaDict;

    public TimeSetOverrideMetadata(string sheetName, IEnumerable<ColumnMetadata> metadata, int headerRow)
    {
        SheetName = sheetName;
        _metaDict = metadata.ToDictionary(m => m.Name, m => m);
        HeaderRow = headerRow;
    }

    public int GetColumnIndex(string columnName)
    {
        if (_metaDict.TryGetValue(columnName, out ColumnMetadata? columnMetadata))
        {
            return columnMetadata.Index;
        }
        return -1;
    }
}
