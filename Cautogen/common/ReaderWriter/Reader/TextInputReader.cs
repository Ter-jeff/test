using System;
using System.IO;

namespace Cautogen.common.ReaderWriter.Reader
{
    public class TextInputReader : ReaderBase
    {
        /* constructor */
        public TextInputReader(string filePath)
            : base(filePath) { }

        public TextInputReader(string filePath, Action callbackFunc)
            : base(filePath, callbackFunc) { }

        /* methods */
        public override void Read()  // with open(filePath, "rb") as f_in: ...
        {
            if (!IsFileExist())
            {
                return;
            }

            FileStream fileStream = null;

            try
            {
                fileStream = new FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite);
                using (var textReader = new StreamReader(fileStream))
                {
                    fileStream = null;
                    _Read(textReader);
                }
            }
            finally
            {
                if (fileStream != null)
                {
                    fileStream.Dispose();
                }
            }
        }

        protected virtual void _Read(StreamReader textReader)  // override this method to read the actuall contnet
        {
            throw new NotImplementedException();
        }
    }
}
