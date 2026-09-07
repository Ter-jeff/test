using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Cautogen.common.ReaderWriter.Reader
{
    public class ReaderBase : IReader
    {
        /* properites */
        protected string FilePath = "";

        protected List<Action> CallbackFuncList;

        /* constructor */
        public ReaderBase(string filePath, Action callbackFunc = null)
        {
            FilePath = filePath;
            CallbackFuncList = new List<Action> { callbackFunc };
        }

        /* methods */
        protected bool IsFileExist()
        {
            return File.Exists(FilePath);
        }

        public virtual void Read()
        {
            throw new NotImplementedException();
        }

        public virtual void RunCallBack()
        {
            foreach (Action callbackFunc in CallbackFuncList
                .Where(callbackFunc => callbackFunc != null))
            {
                callbackFunc();
            }
        }
    }
}
