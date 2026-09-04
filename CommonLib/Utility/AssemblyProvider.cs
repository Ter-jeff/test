namespace CommonLib.Utility
{
    public class AssemblyProvider
    {
        public static AssemblyProvider Current = new AssemblyProvider();

        public virtual string GetFileVersion(string version)
        {
            return version;

        }
    }
}
