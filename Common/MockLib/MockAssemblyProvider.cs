using CommonLib.Static;

namespace MockLib
{
    public sealed class MockAssemblyProvider : AssemblyProvider
    {
        public static MockAssemblyProvider Instance { get; } = new MockAssemblyProvider();

        private MockAssemblyProvider()
        {
        }

        public override string GetFileVersion(string version)
        {
            return "2022.12.19.1";
        }

    }
}
