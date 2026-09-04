namespace CommonLib.Utility
{
    public class MockAssemblyProvider : AssemblyProvider
    {

        public override string GetFileVersion(string version)
        {
            return "2022.12.19.1";
        }

        public static MockAssemblyProvider Instance = new MockAssemblyProvider();
    }
}
