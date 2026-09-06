namespace CommonLib.Static
{
    public class AssemblyProvider
    {
        public static AssemblyProvider Current { get; set; } = new AssemblyProvider();

        public virtual string GetFileVersion(string version)
        {
            return version;
        }
    }
}
