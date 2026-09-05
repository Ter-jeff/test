using System.Drawing;

namespace Cautogen.Utility
{


    public class ProgressMsg
    {
        private Color _color;
        private string _msg;

        public ProgressMsg(string msg, Color color = default(Color))
        {
            _msg = msg;
            _color = color;
        }
    }
}
