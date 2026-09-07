using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.Utility.Atpg;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.Utility
{
    [TestClass]
    public class AtpgServiceTests : FunctionTestBase
    {
        private readonly List<BinCutFinalInstanceRow> _binCutFinalInstanceRows = [];
        private readonly List<Function> _functions = [];
        private readonly List<UserFunctionTableRow> _userFunctionTableRows = [];
        private readonly List<string> _insPats = [];
        private readonly List<string> _ufDigSrcPats = [];

        [TestInitialize]
        public void Setup()
        {
            _binCutFinalInstanceRows.Clear();
            _functions.Clear();

            _binCutFinalInstanceRows.Add(new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow
                {
                    UserFunction = "U1"
                }
            });
            _functions.Add(new Function
            {
                Parameters = "digSrcPin,digSrcEquation,digSrcAssignment",
                ArgList = ["JTAG_TDI", "C|C|D|||E", "C=Selsram();D=DSSC(D1);E=DSSC(E2)",]
            });
            _userFunctionTableRows.Add(new UserFunctionTableRow() { UserFunction = "U1", MultiFstpSetting = "D1;E2" });

            _binCutFinalInstanceRows.Add(new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow
                {
                    UserFunction = "U2"
                }
            });
            _functions.Add(new Function
            {
                Parameters = "digSrcPin,digSrcEquation,digSrcAssignment",
                ArgList = ["JTAG_TDI", "C|C|D|||E", "C=Selsram();D=DSSC(D1)",]
            });
            _userFunctionTableRows.Add(new UserFunctionTableRow() { UserFunction = "U2", MultiFstpSetting = "D1;" });

            _binCutFinalInstanceRows.Add(new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow
                {
                    UserFunction = "K2:U2"
                }
            });
            _functions.Add(new Function
            {
                Parameters = "digSrcPin,digSrcEquation,digSrcAssignment",
                ArgList = ["JTAG_TDI", "C|C|D|||E", "C=Selsram();D=DSSC(U2)",]
            });
            _userFunctionTableRows.Add(null!);

            _insPats.Clear();
            _insPats.Add("_DSSC1_");
            _insPats.Add("_2dssc_");
            _insPats.Add("pat1");
            _insPats.Add("pat2");
            _insPats.Add("patA");
            _insPats.Add("patB");

            _ufDigSrcPats.Clear();
            _ufDigSrcPats.Add("Pat1");
            _ufDigSrcPats.Add("PatB");
        }

        [TestMethod]
        public void SetDigSrcTest()
        {
            List<bool> resultArr = [];
            for (int i = 0; i < _binCutFinalInstanceRows.Count; i++)
            {
                Function function = new Function
                {
                    Parameters = "digSrcPin,digSrcEquation,digSrcAssignment",
                    ArgList = ["", "", ""]
                };
                AtpgService.SetDigSrc(_binCutFinalInstanceRows[i], _ufDigSrcPats, _userFunctionTableRows[i], null, "", _insPats, ref function);
                if (function.ArgList.SequenceEqual(_functions[i].ArgList))
                {
                    resultArr.Add(true);
                }
            }
            Assert.AreEqual(_binCutFinalInstanceRows.Count, resultArr.Count);
            Assert.IsTrue(resultArr.All(x => x));
        }
    }
}
