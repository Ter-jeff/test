namespace CommonLib.Static
{
    public class MockAssemblyProvider : AssemblyProvider
    {
        public static readonly MockAssemblyProvider Instance = new();

        public override string GetFileVersion(string version)
        {
            return "2022.12.19.1";
        }
    }
}
