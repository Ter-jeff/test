using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;

using OfficeOpenXml;

namespace Automation.Static
{
    public static class ScghStatic
    {
        private static ScghData _scghData;
        public static ScghData ScghData
        {
            get
            {
                if (_scghData == null)
                {
                    ExcelWorkbook wb = EpWorkbook.ScghWorkbook;
                    if (wb == null)
                    {
                        return null;
                    }
                    _scghData = ScghData.LoadScghData(wb);
                }
                return _scghData;
            }
        }

        public static void Clear()
        {
            _scghData = null;
        }
    }
}
