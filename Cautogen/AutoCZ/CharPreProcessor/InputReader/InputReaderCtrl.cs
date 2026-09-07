using System.Collections.Generic;

using Cautogen.common.ReaderWriter.Reader;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader
{
    public class InputReaderCtrl
    {
        /* properties */
        private readonly List<IReader> _readerList = new List<IReader>();

        /* constructor */
        public InputReaderCtrl(List<IReader> readerList)
        {
            _readerList = readerList;
        }

        /* methods */
        public void WorkFlow()
        {
            foreach (IReader reader in _readerList)
            {
                reader.Read();
            }
            foreach (IReader reader in _readerList)
            {
                reader.RunCallBack();
            }
        }
    }
}
