using IgxlLib;

using TestPlanLib.NonIgxlSheets;
using TestPlanLib.VbtLib;

namespace Automation.Static
{
    public static class TestProgram
    {
        private static IgxlWorkBook _igxlWorkBk;
        private static IgxlWorkBook _t0TxIgxlWorkBk;
        private static IgxlWorkBook _subProgIgxlWorkBk;
        private static NonIgxlSheets _nonIgxlSheetList;
        private static VbtFunctionLib _vbtFunctionLib;

        public static IgxlWorkBook IgxlWorkBk
        {
            get
            {
                return _igxlWorkBk ?? (_igxlWorkBk = new IgxlWorkBook());
            }
            set
            {
                _igxlWorkBk = value;
            }
        }
        public static IgxlWorkBook T0TxIgxlWorkBk
        {
            get
            {
                return _t0TxIgxlWorkBk ?? (_t0TxIgxlWorkBk = new IgxlWorkBook());
            }
            set
            {
                _t0TxIgxlWorkBk = value;
            }
        }
        public static IgxlWorkBook SubProgIgxlWorkBk
        {
            get
            {
                return _subProgIgxlWorkBk ?? (_subProgIgxlWorkBk = new IgxlWorkBook());
            }
            set
            {
                _subProgIgxlWorkBk = value;
            }
        }
        public static NonIgxlSheets NonIgxlSheetsList
        {
            get
            {
                return _nonIgxlSheetList ?? (_nonIgxlSheetList = new NonIgxlSheets());
            }
        }
        public static VbtFunctionLib VbtFunctionLib
        {
            get
            {
                return _vbtFunctionLib ?? (_vbtFunctionLib = new VbtFunctionLib());
            }
            set
            {
                _vbtFunctionLib = value;
            }
        }

        public static void Clear()
        {
            _igxlWorkBk = new IgxlWorkBook();
            _t0TxIgxlWorkBk = new IgxlWorkBook();
            _subProgIgxlWorkBk = new IgxlWorkBook();
            _nonIgxlSheetList = new NonIgxlSheets();
            _vbtFunctionLib = new VbtFunctionLib();
        }

        internal static void Print()
        {
            IgxlWorkBk.PrintAllSheets();
        }
    }
}
