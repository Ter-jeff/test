using System.Collections.Generic;

namespace Cautogen.common.ReaderWriter.Reader
{
    public class ReaderCtrl
    {
        /* properties */
        private readonly List<IReader> _readerList;

        /* constructor */
        public ReaderCtrl(List<IReader> readerList)
        {
            _readerList = readerList;
        }

        /* methods */
        public void WorkFlow()
        {
            foreach (IReader reader in _readerList)
            {
                reader.Read();
                reader.RunCallBack();
            }
        }
    }
}
