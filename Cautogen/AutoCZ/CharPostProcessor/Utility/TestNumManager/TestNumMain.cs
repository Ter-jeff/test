namespace Cautogen.AutoCZ.CharPostProcessor.Utility.TestNumManager
{
    public class TestNumMain
    {
        private static int _testNum;
        private const int HipStep = 1000;
        private const int NonHipStep = 2000;

        public static int GetTestNum(bool isHardIp = true)
        {
            if (isHardIp)
            {
                _testNum += HipStep;
            }
            else
            {
                _testNum += NonHipStep;
            }

            return _testNum;
        }

        public static void NextBlock()
        {
            int num = _testNum / 10000;
            _testNum = (num + 1) * 10000;
        }

        public static void Reset(string tnumStart)
        {
            int tnumShift = int.Parse(tnumStart) - 10000;
            _testNum = 0 + tnumShift;
        }
    }
}
