namespace MyCommandLineLib
{
    public interface ICommandLineApplication
    {
        bool UseStateMachine { get; set; }

        ICommandLineOptions PreAction(ICommandLineOptions commandLineOptions);

        ICommandLineApplication PostAction(ICommandLineOptions commandLineOptions);

        ICommandLineOptions Start(ICommandLineOptions commandLineOptions);

        ICommandLineOptions End(ICommandLineOptions commandLineOptions);

        ICommandLineOptions ValidateInput(ICommandLineOptions commandLineOptions);

        ICommandLineApplication Execute(ICommandLineOptions commandLineOptions);
    }
}
