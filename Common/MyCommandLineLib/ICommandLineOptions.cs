namespace MyCommandLineLib
{
    public interface ICommandLineOptions : IOptions
    {
        string GetOutputFolder();
    }
}
