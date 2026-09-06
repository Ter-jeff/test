namespace CommonLib.Tables;

public sealed record ColumnConfig(string Name, bool Required, bool ValueRequired);

public sealed record ColumnMetadata(string Name, int Index);
