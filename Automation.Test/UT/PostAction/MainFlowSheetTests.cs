using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Business;
using Automation.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class MainFlowSheetTests
    {
        private static readonly string[] _enableModulesM1M2 = ["M1", "M2"];
        private static readonly string[] _enableModulesMxMy = ["MX", "MY"];

        [TestInitialize]
        public void Setup()
        {
            // Default option for every test
            LocalSpecs.Options.GenerateT0TXTestprogram = false;
        }

        // ------------------------------------------------------------
        // 01. CoreOverFailing: basics, flags, duplicates, trimming
        // ------------------------------------------------------------

        [TestMethod]
        public void ReadSheet_CoreOverFailing_ShouldParseBasicRows()
        {
            MainFlowSheet result;
            Dictionary<string, string> expected = new Dictionary<string, string>
            {
                { "Flow_CoreOverFailing", "" },
                { "Flow_CoreOverFailing_SC", "SC" },
                { "Flow_CoreOverFailing_BI", "BI" },
            };

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");

                // Fixed title row, 6 columns
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                // Minimal rows with TRUE/FALSE
                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "Flow_CoreOverFailing";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "FALSE";

                s.Cells[3, 1].Value = "TestPlan";
                s.Cells[3, 2].Value = "Harvest";
                s.Cells[3, 3].Value = "Flow_CoreOverFailing_SC";
                s.Cells[3, 4].Value = "TRUE";
                s.Cells[3, 5].Value = "TRUE";
                s.Cells[3, 6].Value = "FALSE";

                s.Cells[4, 1].Value = "TestPlan";
                s.Cells[4, 2].Value = "Harvest";
                s.Cells[4, 3].Value = "Flow_CoreOverFailing_BI";
                s.Cells[4, 4].Value = "TRUE";
                s.Cells[4, 5].Value = "TRUE";
                s.Cells[4, 6].Value = "FALSE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            // CoreOverFailingFlows is computed from Rows.First().SequencesNew sheet names using regex and suffix uppercasing
            Assert.IsTrue(expected.Count == result.CoreOverFailingFlows.Count && expected.SequenceEqual(result.CoreOverFailingFlows));
        }

        [TestMethod]
        public void ReadSheet_ShouldReturnEmpty_WhenOnlyHeader()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            // Should be empty when only header exists
            Assert.AreEqual(0, result.CoreOverFailingFlows.Count);
        }

        [TestMethod]
        public void ReadSheet_ShouldRespect_Enable_And_Include()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");

                // Fixed 6-column header
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                // Row 2: Enable=true (CP1 has TRUE), Include=TRUE  --> in EnableModules
                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "M1";
                s.Cells[2, 3].Value = "Flow_CoreOverFailing";
                // EnableWord, NOT used for sequence.Enable
                s.Cells[2, 4].Value = "TRUE";
                // CP1 non-empty -> sequence.Enable = true
                s.Cells[2, 5].Value = "TRUE";
                // Include -> EnableModules collects M1
                s.Cells[2, 6].Value = "TRUE";

                // Row 3: Enable=false (CP1 empty), Include=TRUE     --> also collected in EnableModules
                s.Cells[3, 1].Value = "TestPlan";
                s.Cells[3, 2].Value = "M2";
                s.Cells[3, 3].Value = "Flow_CoreOverFailing_SC";
                // EnableWord only
                s.Cells[3, 4].Value = "FALSE";
                // CP1 empty -> sequence.Enable = false  <-- key point
                // Include -> EnableModules collects M2
                s.Cells[3, 5].Value = "";
                s.Cells[3, 6].Value = "TRUE";

                // Row 4: Enable=true (CP1 TRUE), Include=FALSE      --> NOT collected in EnableModules
                s.Cells[4, 1].Value = "TestPlan";
                s.Cells[4, 2].Value = "M3";
                s.Cells[4, 3].Value = "Flow_CoreOverFailing_BI";
                // EnableWord only
                s.Cells[4, 4].Value = "TRUE";
                // CP1 -> sequence.Enable = true
                s.Cells[4, 5].Value = "TRUE";
                // not collected
                s.Cells[4, 6].Value = "FALSE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            // Enable behavior: determined by whether CP1 (job column) is empty, unrelated to the "Enable" column
            FlowSequenceNew seqR2 = result.Rows.SelectMany(r => r.SequencesNew).First(x => x.SheetName == "Flow_CoreOverFailing");
            // CP1 = TRUE -> true
            Assert.IsTrue(seqR2.Enable);

            FlowSequenceNew seqR3 = result.Rows.SelectMany(r => r.SequencesNew).First(x => x.SheetName == "Flow_CoreOverFailing_SC");
            // CP1 = ""   -> false
            Assert.IsFalse(seqR3.Enable);

            FlowSequenceNew seqR4 = result.Rows.SelectMany(r => r.SequencesNew).First(x => x.SheetName == "Flow_CoreOverFailing_BI");
            // CP1 = TRUE -> true
            Assert.IsTrue(seqR4.Enable);

            // Include behavior: only controls EnableModules collection and deduplication
            CollectionAssert.AreEquivalent(_enableModulesM1M2, result.EnableModules);
        }

        [TestMethod]
        public void ReadSheet_ShouldIgnore_NonCoreOverFailing_Family()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "Flow_Other";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "FALSE";

                s.Cells[3, 1].Value = "TestPlan";
                s.Cells[3, 2].Value = "Harvest";
                s.Cells[3, 3].Value = "Flow_CoreOverFailing_X1";
                s.Cells[3, 4].Value = "TRUE";
                s.Cells[3, 5].Value = "TRUE";
                s.Cells[3, 6].Value = "FALSE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            Assert.AreEqual(1, result.CoreOverFailingFlows.Count);
            Assert.IsTrue(result.CoreOverFailingFlows.ContainsKey("Flow_CoreOverFailing_X1"));
        }

        [TestMethod]
        public void ReadSheet_ShouldHandle_Duplicate_SheetName()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "Flow_CoreOverFailing_SC";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "FALSE";

                s.Cells[3, 1].Value = "TestPlan";
                s.Cells[3, 2].Value = "Harvest";
                s.Cells[3, 3].Value = "Flow_CoreOverFailing_SC";
                s.Cells[3, 4].Value = "TRUE";
                s.Cells[3, 5].Value = "TRUE";
                s.Cells[3, 6].Value = "FALSE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            Assert.AreEqual(1, result.CoreOverFailingFlows.Count);
            Assert.AreEqual("SC", result.CoreOverFailingFlows["Flow_CoreOverFailing_SC"]);
        }

        [TestMethod]
        public void ReadSheet_ShouldTrim_SheetName_And_KeepSuffix()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "  Flow_CoreOverFailing_BI  ";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "FALSE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            Assert.IsTrue(result.CoreOverFailingFlows.ContainsKey("Flow_CoreOverFailing_BI"));
            Assert.AreEqual("BI", result.CoreOverFailingFlows["Flow_CoreOverFailing_BI"]);
        }

        // ------------------------------------------------------------
        // 02. T0TX gate: off vs on, and flag filtering
        // ------------------------------------------------------------

        [TestMethod]
        public void ReadSheet_T0TX_ShouldBeIgnored_WhenFlagOff()
        {
            MainFlowSheet result;
            LocalSpecs.Options.GenerateT0TXTestprogram = false;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "Flow_T0TX_Core";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "TRUE";

                s.Cells[3, 1].Value = "TestPlan";
                s.Cells[3, 2].Value = "Harvest";
                s.Cells[3, 3].Value = "Flow_CoreOverFailing";
                s.Cells[3, 4].Value = "TRUE";
                s.Cells[3, 5].Value = "TRUE";
                s.Cells[3, 6].Value = "TRUE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            Assert.IsFalse(result.CoreOverFailingFlows.ContainsKey("Flow_T0TX_Core"));
            Assert.IsTrue(result.CoreOverFailingFlows.ContainsKey("Flow_CoreOverFailing"));
        }

        [TestMethod]
        public void ReadSheet_T0TX_ShouldBeCollected_WhenFlagOn()
        {
            // Flag on
            LocalSpecs.Options.GenerateT0TXTestprogram = true;

            MainFlowSheet result;
            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");

                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "Flow_T0TX_Core";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "TRUE";

                s.Cells[3, 1].Value = "TestPlan";
                s.Cells[3, 2].Value = "Harvest";
                s.Cells[3, 3].Value = "Flow_T0TX_Ext";
                s.Cells[3, 4].Value = "TRUE";
                s.Cells[3, 5].Value = "TRUE";
                s.Cells[3, 6].Value = "TRUE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            // Correct check: find T0TX sequences from Rows
            bool hasCore = result.Rows.SelectMany(m => m.SequencesNew).Any(x => x.SheetName == "Flow_T0TX_Core");
            bool hasExt = result.Rows.SelectMany(m => m.SequencesNew).Any(x => x.SheetName == "Flow_T0TX_Ext");
            Assert.IsTrue(hasCore);
            Assert.IsTrue(hasExt);
        }

        [TestMethod]
        public void ReadSheet_T0TX_ShouldRespect_Enable_And_Include()
        {
            MainFlowSheet result;
            // enable T0TX path
            LocalSpecs.Options.GenerateT0TXTestprogram = true;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");

                // Fixed 6-column header
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                // EnableWord only, not used for sequence.Enable
                s.Cells[1, 4].Value = "Enable";
                // Job flag column; non-empty -> sequence.Enable = true
                // Controls EnableModules only
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                // Row 2: kept, Enable=true (CP1="TRUE"), Include=TRUE -> collected
                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "MX";
                s.Cells[2, 3].Value = "Flow_T0TX_Core";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "TRUE";

                // Row 3: disabled, Enable=false (CP1=""), Include=TRUE -> still collected in EnableModules
                s.Cells[3, 1].Value = "TestPlan";
                s.Cells[3, 2].Value = "MY";
                s.Cells[3, 3].Value = "Flow_T0TX_Disabled";
                s.Cells[3, 4].Value = "FALSE";
                // empty -> sequence.Enable = false
                s.Cells[3, 5].Value = "";
                s.Cells[3, 6].Value = "TRUE";

                // Row 4: excluded, Enable=true (CP1="TRUE"), Include=FALSE -> not in EnableModules
                s.Cells[4, 1].Value = "TestPlan";
                s.Cells[4, 2].Value = "MZ";
                s.Cells[4, 3].Value = "Flow_T0TX_Excluded";
                s.Cells[4, 4].Value = "TRUE";
                s.Cells[4, 5].Value = "TRUE";
                s.Cells[4, 6].Value = "FALSE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }
            // present

            // T0TX sequences exist only when flag is true
            Assert.IsTrue(result.Rows.SelectMany(m => m.SequencesNew).Any(x => x.SheetName == "Flow_T0TX_Core"));
            // present
            Assert.IsTrue(result.Rows.SelectMany(m => m.SequencesNew).Any(x => x.SheetName == "Flow_T0TX_Disabled"));
            // present
            Assert.IsTrue(result.Rows.SelectMany(m => m.SequencesNew).Any(x => x.SheetName == "Flow_T0TX_Excluded"));

            // Per-row Enable is decided by job column (CP1) non-empty
            FlowSequenceNew seqCore = result.Rows.SelectMany(m => m.SequencesNew).First(x => x.SheetName == "Flow_T0TX_Core");
            // CP1="TRUE" -> true
            Assert.IsTrue(seqCore.Enable);

            FlowSequenceNew seqDisabled = result.Rows.SelectMany(m => m.SequencesNew).First(x => x.SheetName == "Flow_T0TX_Disabled");
            // CP1="" -> false
            Assert.IsFalse(seqDisabled.Enable);

            FlowSequenceNew seqExcluded = result.Rows.SelectMany(m => m.SequencesNew).First(x => x.SheetName == "Flow_T0TX_Excluded");
            // CP1="TRUE" -> true
            Assert.IsTrue(seqExcluded.Enable);

            // Include controls EnableModules only; dedupe implied
            CollectionAssert.AreEquivalent(_enableModulesMxMy, result.EnableModules);

            // Sanity: CoreOverFailingFlows is unrelated to T0TX
            Assert.IsFalse(result.CoreOverFailingFlows.ContainsKey("Flow_T0TX_Core"));
            Assert.IsFalse(result.CoreOverFailingFlows.ContainsKey("Flow_T0TX_Disabled"));
            // dict only for Flow_CoreOverFailing*
            Assert.IsFalse(result.CoreOverFailingFlows.ContainsKey("Flow_T0TX_Excluded"));
        }

        // ------------------------------------------------------------
        // 03. Header variants: Option/Options, SubProgram, WLFT1/WLFT2
        // ------------------------------------------------------------

        [TestMethod]
        public void ReadSheet_Should_Handle_Option_And_Options_Headers_Minimal()
        {
            MainFlowSheet opt1;
            MainFlowSheet opt2;

            using (ExcelPackage p = new ExcelPackage())
            {
                // Option at column 7
                ExcelWorksheet s1 = p.Workbook.Worksheets.Add("Flow_Main_Opt1");
                s1.Cells[1, 1].Value = "Source";
                s1.Cells[1, 2].Value = "Module";
                s1.Cells[1, 3].Value = "SheetName";
                s1.Cells[1, 4].Value = "Enable";
                s1.Cells[1, 5].Value = "CP1";
                s1.Cells[1, 6].Value = "Include";
                s1.Cells[1, 7].Value = "Option";

                s1.Cells[2, 1].Value = "TestPlan";
                s1.Cells[2, 2].Value = "M1";
                s1.Cells[2, 3].Value = "Flow_X1";
                s1.Cells[2, 4].Value = "TRUE";
                s1.Cells[2, 5].Value = "TRUE";
                s1.Cells[2, 6].Value = "FALSE";
                s1.Cells[2, 7].Value = "OPT_ONE";

                opt1 = new MainFlowSheetReaderNew().ReadSheet(s1);

                // Options at column 7
                ExcelWorksheet s2 = p.Workbook.Worksheets.Add("Flow_Main_Opt2");
                s2.Cells[1, 1].Value = "Source";
                s2.Cells[1, 2].Value = "Module";
                s2.Cells[1, 3].Value = "SheetName";
                s2.Cells[1, 4].Value = "Enable";
                s2.Cells[1, 5].Value = "CP1";
                s2.Cells[1, 6].Value = "Include";
                s2.Cells[1, 7].Value = "Options";

                s2.Cells[2, 1].Value = "TestPlan";
                s2.Cells[2, 2].Value = "M2";
                s2.Cells[2, 3].Value = "Flow_X2";
                s2.Cells[2, 4].Value = "TRUE";
                s2.Cells[2, 5].Value = "TRUE";
                s2.Cells[2, 6].Value = "FALSE";
                s2.Cells[2, 7].Value = "OPT_TWO";

                opt2 = new MainFlowSheetReaderNew().ReadSheet(s2);
            }

            // Jobs is built from header keys right of Enable, CP1 must be present
            CollectionAssert.Contains(opt1.Jobs, "CP1");
            CollectionAssert.Contains(opt2.Jobs, "CP1");

            // Option column should be detected at index 7 in both cases
            Assert.AreEqual(7, opt1.OptionCol);
            Assert.AreEqual(7, opt2.OptionCol);
        }

        [TestMethod]
        public void ReadSheet_Should_Expose_SubProgram_ColumnIndex_And_Add_All_Row()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main_Sub");
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";
                // will be mapped to MinimalSubFlowCol
                s.Cells[1, 7].Value = "SubProgram";

                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "M1";
                s.Cells[2, 3].Value = "Flow_Sub:SF1";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "FALSE";
                s.Cells[2, 7].Value = "SUBY";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            // 1-based column index equals 7 when SubProgram header is at column 7
            Assert.AreEqual(7, result.SubprogramCol);

            // "All" job should be added by ReadDataSub, now Rows is public so we can assert it
            Assert.IsTrue(result.Rows.Any(m => m.JobName == "All"));
        }

        [TestMethod]
        public void ReadSheet_Should_Parse_WLFT1_WLFT2_And_Map_To_FT1_FT2()
        {
            MainFlowSheet result;
            LocalSpecs.Options.GenerateT0TXTestprogram = true;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main_T0TX");
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";
                // maps to T0TX_Room group internally -> FT1 job
                s.Cells[1, 7].Value = "WLFT1";
                // maps to T0TX_Hot  group internally -> FT2 job
                s.Cells[1, 8].Value = "WLFT2";

                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "Flow_T0TX_B";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "FALSE";
                s.Cells[2, 6].Value = "TRUE";
                s.Cells[2, 7].Value = "TRUE";
                s.Cells[2, 8].Value = "TRUE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            // Header keys should be present
            CollectionAssert.Contains(result.Jobs, "WLFT1");
            CollectionAssert.Contains(result.Jobs, "WLFT2");

            // And Rows should contain JobName FT1 and FT2 produced by ReadT0TXDataMain
            IEnumerable<string> jobNames = result.Rows.Select(x => x.JobName);
            Assert.IsTrue(jobNames.Contains("FT1"));
            Assert.IsTrue(jobNames.Contains("FT2"));
        }

        // ------------------------------------------------------------
        // 04. EnableModules behavior (TRUE dedupe)
        // ------------------------------------------------------------

        [TestMethod]
        public void ReadSheet_Should_Collect_EnableModules_And_Dedupe()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                // M1 TRUE twice
                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "M1";
                s.Cells[2, 3].Value = "Flow_A";
                s.Cells[2, 6].Value = "TRUE";

                s.Cells[3, 1].Value = "TestPlan";
                s.Cells[3, 2].Value = "M1";
                s.Cells[3, 3].Value = "Flow_B";
                s.Cells[3, 6].Value = "TRUE";

                // M2 FALSE ignored
                s.Cells[4, 1].Value = "TestPlan";
                s.Cells[4, 2].Value = "M2";
                s.Cells[4, 3].Value = "Flow_C";
                s.Cells[4, 6].Value = "FALSE";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            List<string> enabled = result.EnableModules;
            Assert.AreNotEqual(null, enabled);
            Assert.AreEqual(1, enabled.Count);
            // ReadEnableModule collects only "TRUE" and dedupes modules
            Assert.AreEqual("M1", enabled[0]);
        }

        [TestMethod]
        public void ReadSheet_Should_Parse_Comment_And_Group_Columns()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");

                // Fixed title row, 6 columns
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                // Extra headers for this test
                s.Cells[1, 7].Value = "Comment";
                s.Cells[1, 8].Value = "Group";

                // One minimal data row
                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "Flow_CoreOverFailing";
                // Enable
                s.Cells[2, 4].Value = "TRUE";
                // CP1 flag TRUE
                s.Cells[2, 5].Value = "TRUE";
                // Include
                s.Cells[2, 6].Value = "FALSE";
                // Comment
                s.Cells[2, 7].Value = "CMT_1";
                // Group
                s.Cells[2, 8].Value = "G1";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            // Assert through public Rows -> SequencesNew
            FlowSequenceNew seq = result.Rows.SelectMany(m => m.SequencesNew).First(x => x.SheetName == "Flow_CoreOverFailing");
            // Comment column captured

            Assert.AreEqual("", seq.Comment);
            // Group column captured
            Assert.AreEqual("G1", seq.Group);
        }

        [TestMethod]
        public void ReadSheet_Should_Parse_BinTableEnable_AllSites_Column()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");

                // Fixed title row, 6 columns
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                // Extra header, matched by regex: ^BinTable Enable\s*\(all sites\)
                s.Cells[1, 7].Value = "BinTable Enable (all sites)";

                // One minimal data row
                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "Flow_CoreOverFailing_X1";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "FALSE";
                // Bintable enable word
                s.Cells[2, 7].Value = "BT_ON";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            FlowSequenceNew seq = result.Rows
                .SelectMany(m => m.SequencesNew)
                .First(x => x.SheetName == "Flow_CoreOverFailing_X1");
            // BinTable Enable (all sites) captured

            Assert.AreEqual("BT_ON", seq.BintableEnableWord);
        }

        [TestMethod]
        public void ReadSheet_Should_Parse_Option_Into_OptionDict()
        {
            MainFlowSheet result;

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet s = p.Workbook.Worksheets.Add("Flow_Main");

                // Fixed title row, 6 columns
                s.Cells[1, 1].Value = "Source";
                s.Cells[1, 2].Value = "Module";
                s.Cells[1, 3].Value = "SheetName";
                s.Cells[1, 4].Value = "Enable";
                s.Cells[1, 5].Value = "CP1";
                s.Cells[1, 6].Value = "Include";

                // Extra header for Option
                s.Cells[1, 7].Value = "Option";

                // One minimal data row with inline option pairs
                s.Cells[2, 1].Value = "TestPlan";
                s.Cells[2, 2].Value = "Harvest";
                s.Cells[2, 3].Value = "Flow_CoreOverFailing";
                s.Cells[2, 4].Value = "TRUE";
                s.Cells[2, 5].Value = "TRUE";
                s.Cells[2, 6].Value = "FALSE";
                // "Flag" becomes Flag=Flag by parser
                s.Cells[2, 7].Value = "A=1; B=2; Flag";

                result = new MainFlowSheetReaderNew().ReadSheet(s);
            }

            FlowSequenceNew seq = result.Rows
                .SelectMany(m => m.SequencesNew)
                .First(x => x.SheetName == "Flow_CoreOverFailing");

            // Option parsing behavior: key-value and key-only both supported into OptionDict
            Assert.AreEqual("1", seq.OptionDict["A"]);
            Assert.AreEqual("2", seq.OptionDict["B"]);
            Assert.AreEqual("Flag", seq.OptionDict["Flag"]);
        }
    }
}
