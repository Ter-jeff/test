using System;
using System.Linq;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    public class SetWriteItem
    {
        public string Site = "";
        public string EfuseKey = "";
        public string ReferenceKey = "";
        public string ReferenceValue = "";
        public string ReverseBit = "";
        public int Number;

        public string GetValue()
        {
            try
            {
                if (string.IsNullOrEmpty(ReverseBit))
                {
                    return ReferenceValue;
                }

                if (!int.TryParse(ReverseBit, out int tmp))
                {
                    return ReferenceValue;
                }

                string value = Convert.ToString(int.Parse(ReferenceValue), 2).PadLeft(tmp, '0');
                value = new string([.. value.Reverse()]);
                return Convert.ToInt32(value, 2).ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }
    }
}
