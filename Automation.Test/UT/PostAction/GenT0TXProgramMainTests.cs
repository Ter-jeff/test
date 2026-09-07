
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenT0TXProgram;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class GenT0TXProgramMainTests
    {
        private static readonly SubFlowSheet _subSheet = new("MySubFlow")
        {
            Rows = new FlowRows
                {
                    new FlowRow { Opcode = "test", Job = "CP1,CP2", Enable = "!CP1,!CP2" },
                    new FlowRow { Opcode = "test", Job = "HardIP_NV,Prod_CP1,Prod_CP2", Enable = "(Prod_CP1||Prod_FT2)&&Prod_CP2&&Prod_FT1" }
                }
        };
        private static readonly Dictionary<string, SubFlowSheet> _subflows = new() { { "s1", _subSheet } };
        private static readonly InstanceSheet _instanceSheet1 = new("MyInstacneSheetWithHarvest")
        {
            Rows = new InstanceRows
                {
                    new InstanceRow { TestName = "InstacneHarvest", ArgList = "IsHarvesting", Args = ["TRUE"] },
                    new InstanceRow { TestName = "InstacneNotHarvest", ArgList = "Atg1", Args = ["XXX"] }
                }
        };
        private static readonly InstanceSheet _instanceSheet2 = new("MyInstacneSheetWithoutHarvest")
        {
            Rows = new InstanceRows
                {
                    new InstanceRow { TestName = "Instacne1", ArgList = "Arg1", Args = ["XXX"] },
                    new InstanceRow { TestName = "Instacne2", ArgList = "Atg1", Args = ["XXX"] }
                }
        };
        private static readonly Dictionary<string, InstanceSheet> _instanceSheets = new() { { "i1", _instanceSheet1 }, { "i2", _instanceSheet2 } };
        private readonly GenT0TXProgramMain _main = new(_subflows, _instanceSheets);

        [TestMethod]
        public void ModifySubflowsTest()
        {
            // Act
            List<SubFlowSheet> subflow = _main.ModifySubflows();

            // Assert
            FlowRow row1 = subflow.First().Rows[0];
            FlowRow row2 = subflow.First().Rows[1];
            Assert.AreEqual("FT1,FT2", row1.Job);
            Assert.AreEqual("!FT1,!FT2", row1.Enable);
            Assert.AreEqual("HardIP_NV,Prod_FT1,Prod_FT2", row2.Job);
            Assert.AreEqual("(Prod_FT1||Prod_FT2)&&Prod_FT2&&Prod_FT1", row2.Enable);
        }

        [TestMethod]
        public void ModifyInstancesTest()
        {
            // Act
            List<InstanceSheet> instSheets = _main.ModifyInstances();
            InstanceSheet instanceHarvestSheet = instSheets[0];
            InstanceSheet instanceNotHarvestSheet = instSheets[1];

            // Assert
            Assert.AreEqual("MyInstacneSheetWithHarvest_T0TX", instanceHarvestSheet.Name);
            Assert.AreEqual("MyInstacneSheetWithoutHarvest", instanceNotHarvestSheet.Name);
            Assert.AreEqual("FALSE", instanceHarvestSheet.Rows[0].GetArgument("isHarvesting"));
        }
    }
}
