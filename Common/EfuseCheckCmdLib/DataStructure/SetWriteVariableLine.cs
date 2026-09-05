using System;

namespace EfuseCheckCmdLib.DataStructure
{
    public class SetWriteVariableLine
    {
        public string Site;
        public string Key;
        public string Value;
        public string Format = "";
        public string TestName = "";
        public int Number = 0;

        public SetWriteVariableLine(string line)
        {
            System.Collections.Generic.List<string> array = [.. line.Split([' ', '=', '[', ']', '\t'], StringSplitOptions.RemoveEmptyEntries)];

            Site = array[2];
            Key = array[6];
            Value = array[7];
        }
    }
}
