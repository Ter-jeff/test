using System.Collections.Generic;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class InstanceRowTests
    {
        [TestMethod]
        public void InstanceRow_DefaultConstructor_InitializesEmptyCollections()
        {
            // Arrange & Act
            var instanceRow = new InstanceRow();

            // Assert
            Assert.AreEqual("", instanceRow.SheetName);
            Assert.AreEqual(0, instanceRow.RowNum);
            Assert.AreEqual("", instanceRow.TestName);
            Assert.AreEqual(0, instanceRow.Args.Count);
        }

        [TestMethod]
        public void InstanceRow_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                // Act
                TestName = "Test1",
                VbtType = "vbt",
                VbtName = "VBT1",
                CalledAs = "Instance1",
                DcCategory = "DC_CAT1",
                DcSelector = "DC_SEL1",
                AcCategory = "AC_CAT1",
                AcSelector = "AC_SEL1",
                TimeSets = "TS1",
                EdgeSets = "ES1",
                PinLevels = "PL1",
                Comment = "Instance comment"
            };

            // Assert
            Assert.AreEqual("Test1", instanceRow.TestName);
            Assert.AreEqual("vbt", instanceRow.VbtType);
            Assert.AreEqual("VBT1", instanceRow.VbtName);
            Assert.AreEqual("Instance1", instanceRow.CalledAs);
            Assert.AreEqual("DC_CAT1", instanceRow.DcCategory);
        }

        [TestMethod]
        public void InstanceRow_SetArgsList_UpdatesCollection()
        {
            // Arrange
            var instanceRow = new InstanceRow();

            // Act
            instanceRow.Args.Add("Arg1");
            instanceRow.Args.Add("Arg2");
            instanceRow.Args.Add("Arg3");

            // Assert
            Assert.AreEqual(3, instanceRow.Args.Count);
            Assert.AreEqual("Arg1", instanceRow.Args[0]);
            Assert.AreEqual("Arg2", instanceRow.Args[1]);
        }

        [TestMethod]
        public void InstanceRow_InitList_CanBeInitialized()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                // Act
                InitList = ["Init1", "Init2"],
                PayloadList = ["Payload1", "Payload2"],
                FinalJobs = ["FinalJob1"]
            };

            // Assert
            Assert.AreEqual(2, instanceRow.InitList.Count);
            Assert.AreEqual(2, instanceRow.PayloadList.Count);
            Assert.AreEqual(1, instanceRow.FinalJobs.Count);
        }

        [TestMethod]
        public void InstanceRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var instanceRow = new InstanceRow();

            // Assert
            Assert.IsInstanceOfType(instanceRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void InstanceRow_AllStringProperties_CanBeSet()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                // Act
                TestName = "TestName",
                VbtType = "VbtType",
                VbtName = "VbtName",
                CalledAs = "CalledAs",
                DcCategory = "DcCategory",
                DcSelector = "DcSelector",
                AcCategory = "AcCategory",
                AcSelector = "AcSelector",
                TimeSets = "TimeSets",
                EdgeSets = "EdgeSets",
                PinLevels = "PinLevels",
                MixedSignalTiming = "MixedSignalTiming",
                Overlay = "Overlay",
                ArgList = "ArgList",
                Comment = "Comment"
            };

            // Assert
            Assert.AreEqual("TestName", instanceRow.TestName);
            Assert.AreEqual("VbtType", instanceRow.VbtType);
            Assert.AreEqual("VbtName", instanceRow.VbtName);
            Assert.AreEqual("CalledAs", instanceRow.CalledAs);
            Assert.AreEqual("DcCategory", instanceRow.DcCategory);
            Assert.AreEqual("DcSelector", instanceRow.DcSelector);
            Assert.AreEqual("AcCategory", instanceRow.AcCategory);
            Assert.AreEqual("AcSelector", instanceRow.AcSelector);
            Assert.AreEqual("TimeSets", instanceRow.TimeSets);
            Assert.AreEqual("EdgeSets", instanceRow.EdgeSets);
            Assert.AreEqual("PinLevels", instanceRow.PinLevels);
            Assert.AreEqual("MixedSignalTiming", instanceRow.MixedSignalTiming);
            Assert.AreEqual("Overlay", instanceRow.Overlay);
            Assert.AreEqual("ArgList", instanceRow.ArgList);
            Assert.AreEqual("Comment", instanceRow.Comment);
        }

        [TestMethod]
        public void InstanceRow_MultipleArgs_CanBeAdded()
        {
            // Arrange
            var instanceRow = new InstanceRow();

            // Act
            for (int i = 0; i < 10; i++)
            {
                instanceRow.Args.Add($"Arg{i}");
            }

            // Assert
            Assert.AreEqual(10, instanceRow.Args.Count);
            Assert.AreEqual("Arg0", instanceRow.Args[0]);
            Assert.AreEqual("Arg9", instanceRow.Args[9]);
        }

        [TestMethod]
        public void InstanceRow_InitList_MultipleItems()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                InitList = []
            };

            // Act
            instanceRow.InitList.Add("Init1");
            instanceRow.InitList.Add("Init2");
            instanceRow.InitList.Add("Init3");

            // Assert
            Assert.AreEqual(3, instanceRow.InitList.Count);
        }

        [TestMethod]
        public void InstanceRow_PayloadList_MultipleItems()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                PayloadList = []
            };

            // Act
            instanceRow.PayloadList.Add("Payload1");
            instanceRow.PayloadList.Add("Payload2");

            // Assert
            Assert.AreEqual(2, instanceRow.PayloadList.Count);
        }

        [TestMethod]
        public void InstanceRow_FinalJobs_MultipleItems()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                FinalJobs = []
            };

            // Act
            instanceRow.FinalJobs.Add("Job1");
            instanceRow.FinalJobs.Add("Job2");
            instanceRow.FinalJobs.Add("Job3");

            // Assert
            Assert.AreEqual(3, instanceRow.FinalJobs.Count);
        }

        [TestMethod]
        public void InstanceRow_SheetName_CanBeSet()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                // Act
                SheetName = "TestSheet"
            };

            // Assert
            Assert.AreEqual("TestSheet", instanceRow.SheetName);
        }

        [TestMethod]
        public void InstanceRow_RowNum_CanBeSet()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                // Act
                RowNum = 42
            };

            // Assert
            Assert.AreEqual(42, instanceRow.RowNum);
        }

        [TestMethod]
        public void InstanceRow_ComplexInitialization()
        {
            // Arrange & Act
            var instanceRow = new InstanceRow
            {
                SheetName = "Sheet1",
                RowNum = 5,
                TestName = "ComplexTest",
                VbtType = "Type1",
                VbtName = "VBT1",
                CalledAs = "Instance1",
                DcCategory = "DC1",
                DcSelector = "Sel1",
                AcCategory = "AC1",
                AcSelector = "AcSel1",
                TimeSets = "TS1",
                EdgeSets = "ES1",
                PinLevels = "PL1",
                MixedSignalTiming = "MST1",
                Overlay = "OVR1",
                ArgList = "Args1",
                Comment = "Test comment"
            };

            // Assert
            Assert.AreEqual("Sheet1", instanceRow.SheetName);
            Assert.AreEqual(5, instanceRow.RowNum);
            Assert.AreEqual("ComplexTest", instanceRow.TestName);
            Assert.AreEqual("Type1", instanceRow.VbtType);
        }

        [TestMethod]
        public void InstanceRow_Args_CanBeCleared()
        {
            // Arrange
            var instanceRow = new InstanceRow();
            instanceRow.Args.Add("Arg1");
            instanceRow.Args.Add("Arg2");
            Assert.AreEqual(2, instanceRow.Args.Count);

            // Act
            instanceRow.Args.Clear();

            // Assert
            Assert.AreEqual(0, instanceRow.Args.Count);
        }

        [TestMethod]
        public void InstanceRow_CopyConstructor_CreatesIsolatedDeepCopy()
        {
            // Arrange
            var source = new InstanceRow
            {
                TestName = "OriginalTest",
                VbtType = "OriginalType",
                VbtName = "OriginalVbt",
                CalledAs = "OriginalCall",
                DcCategory = "OriginalDcCat",
                DcSelector = "OriginalDcSel",
                AcCategory = "OriginalAcCat",
                AcSelector = "OriginalAcSel",
                TimeSets = "OriginalTs",
                EdgeSets = "OriginalEs",
                PinLevels = "OriginalPl",
                MixedSignalTiming = "OriginalMst",
                Overlay = "OriginalOvr",
                ArgList = "Param1",
                Comment = "OriginalComment",
                Args = ["Val1"],
                InitList = ["Init1"],
                PayloadList = ["Payload1"],
                FinalJobs = ["Job1"]
            };

            // Act
            var copy = new InstanceRow(source);

            // Assert
            Assert.AreNotSame(source, copy);
            Assert.AreEqual(source.TestName, copy.TestName);
            Assert.AreEqual(source.VbtType, copy.VbtType);
            Assert.AreNotSame(source.Args, copy.Args);
            Assert.AreEqual(source.Args[0], copy.Args[0]);
            Assert.AreNotSame(source.InitList, copy.InitList);
            Assert.AreEqual(source.InitList[0], copy.InitList[0]);
            Assert.AreNotSame(source.PayloadList, copy.PayloadList);
            Assert.AreEqual(source.PayloadList[0], copy.PayloadList[0]);
            Assert.AreNotSame(source.FinalJobs, copy.FinalJobs);
            Assert.AreEqual(source.FinalJobs[0], copy.FinalJobs[0]);
        }

        [TestMethod]
        public void InstanceRow_CopyMethod_ReturnsValidInstanceRowObject()
        {
            // Arrange
            var source = new InstanceRow { TestName = "MethodCopyTarget" };

            // Act
            InstanceRow result = source.Copy();

            // Assert
            Assert.IsNotNull(result);
            Assert.AreNotSame(source, result);
            Assert.AreEqual("MethodCopyTarget", result.TestName);
        }

        [TestMethod]
        public void InstanceRow_GetArgument_ReturnsCorrectValueCaseInsensitively()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                ArgList = "ParamA,ParamB,ParamC",
                Args = ["Alpha", "Beta", "Gamma"]
            };

            // Act & Assert
            // Verifies exact match
            Assert.AreEqual("Beta", instanceRow.GetArgument("ParamB"));
            // Verifies case insensitivity
            Assert.AreEqual("Alpha", instanceRow.GetArgument("parama"));
            // Verifies missing argument handles gracefully
            Assert.AreEqual("", instanceRow.GetArgument("ParamMissing"));
        }

        [TestMethod]
        public void InstanceRow_SetArgument_UpdatesExistingIndexOrPadsCollectionIfNeeded()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                ArgList = "Param1,Param2,Param3",
                Args = ["Val1", "Val2"]
            };

            // Act - Scenario A: Update index within tracking collection constraints
            instanceRow.SetArgument("Param2", "NewVal2");

            // Assert - Scenario A
            Assert.AreEqual("NewVal2", instanceRow.Args[1]);

            // Act - Scenario B: Set index outside tracking collection constraints requiring padding
            instanceRow.SetArgument("Param3", "PaddedVal3");

            // Assert - Scenario B
            Assert.AreEqual(3, instanceRow.Args.Count);
            Assert.AreEqual("PaddedVal3", instanceRow.Args[2]);
        }

        [TestMethod]
        public void InstanceRow_GetPinGrpFlags_ReturnsParsedTokensOrNullForBypassStrings()
        {
            // Arrange
            var instanceRow = new InstanceRow
            {
                ArgList = "Harv_FailFlag",             // Act & Assert - Case 1: Bypass string triggers null evaluation
                Args = ["HarvestPinGrpFlagTable"]
            };
            Assert.IsNull(instanceRow.GetPinGrpFlags());

            // Act & Assert - Case 2: Empty string triggers null evaluation
            instanceRow.Args = [""];
            Assert.IsNull(instanceRow.GetPinGrpFlags());

            // Act & Assert - Case 3: Valid delimiter string gets parsed cleanly
            instanceRow.Args = ["GrpA(Pin1);GrpB(Pin2)"];
            List<string> flags = instanceRow.GetPinGrpFlags();
            Assert.IsNotNull(flags);
            Assert.AreEqual(2, flags.Count);
            Assert.AreEqual("Pin1", flags[0]);
            Assert.AreEqual("Pin2", flags[1]);
        }

        [TestMethod]
        public void InstanceRow_MergeVbtMtdArgs_CombinesDelimitersCorrectly()
        {
            // Arrange
            var targetRow = new InstanceRow
            {
                ArgList = "Patset,Calc_Eqn,TestSequence,UnusedArg",
                Args = ["TgtPat", "TgtEqn", "TgtSeq", ""]
            };
            var incomingRow = new InstanceRow
            {
                ArgList = "Patset,Calc_Eqn,TestSequence,UnusedArg",
                Args = ["IncPat", "IncEqn", "IncSeq", ""]
            };

            // Act
            targetRow.MergeVbtMtdArgs(incomingRow);

            // Assert
            // Verifies the plus sign mapping loop execution
            Assert.AreEqual("IncPat+TgtPat", targetRow.GetArgument("Patset"));
            // Verifies the semicolon string delimiter mapping loop execution
            Assert.AreEqual("IncEqn;TgtEqn", targetRow.GetArgument("Calc_Eqn"));
            // Verifies the hash sign prefix string mapping loop execution
            Assert.AreEqual("IncSeq#TgtSeq", targetRow.GetArgument("TestSequence"));
            // Verifies skipped execution when both argument fields evaluate empty strings
            Assert.AreEqual("", targetRow.GetArgument("UnusedArg"));
        }
    }
}
