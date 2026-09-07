using System;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    internal class XParseDatalogException(string msg) : Exception
    {
        public string ErrMsg = msg;

        public override string Message
        {
            get
            {
                string retStr = "[xParseDatalog] " + ErrMsg;
                return retStr;
            }
        }

        public override string ToString()
        {
            string retStr = "[xParseDatalog] Exception Occurs!";
            return retStr;
        }
    }
}
