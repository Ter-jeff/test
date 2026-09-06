namespace BinCutScriptLib.Base
{
    public class EFuseRow
    {
        public string Name = string.Empty;
        public double Value;
        public double Gb;

        public EFuseRow() { }

        public EFuseRow(EFuseRow eFuseRow)
        {
            if (eFuseRow == null)
            {
                return;
            }

            Name = eFuseRow.Name;
            Value = eFuseRow.Value;
            Gb = eFuseRow.Gb;
        }

        public EFuseRow Copy()
        {
            return new EFuseRow(this);
        }
    }
}
