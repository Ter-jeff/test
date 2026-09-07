using System;

namespace MyCommandLineLib
{
    public class CommandLineException(string caller, string msg, Exception? exception) : Exception(msg, exception)
    {
        public string Caller { get; set; } = caller;

        public CommandLineException(string caller, string msg)
            : this(caller, msg, null)
        {
        }
    }
}
