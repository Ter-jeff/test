using System;
using System.Diagnostics.CodeAnalysis;

using Automation.Const;
using Automation.GenerateIgxl.Basic;
using Automation.GenerateIgxl.BinCut;
using Automation.GenerateIgxl.BistBira;
using Automation.GenerateIgxl.EFuse;
using Automation.GenerateIgxl.EVS;
using Automation.GenerateIgxl.HardIp;
using Automation.GenerateIgxl.PostAction;
using Automation.GenerateIgxl.PreAction;
using Automation.GenerateIgxl.Scan;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.Enums;

using CommonReaderLib.PatternListCsv;

using LogLib.Static;

namespace Automation
{
    [ExcludeFromCodeCoverage]
    public class ValidateTestPlanMain
    {
        public void Run()
        {
            try
            {
                #region PreAction
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.PreAction).Down)
                {
                    Response.Report("Pre-Action Pre-Check ...", EnumMessageLevel.StartPoint);
                    using (var preActionMain = new PreActionMain())
                    {
                        preActionMain.PreCheckFlow(null);
                    }
                    Response.Report("Pre-Action Pre-Check Completed !", EnumMessageLevel.EndPoint);
                }
                #endregion

                #region Basic
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.Basic).Down)
                {
                    Response.Report("Basic Pre-Check ...", EnumMessageLevel.StartPoint);
                    using (var basicMain = new BasicMain())
                    {
                        basicMain.PreCheckFlow(null);
                    }
                    Response.Report("Basic Pre-Check Completed !", EnumMessageLevel.EndPoint);
                }
                #endregion

                #region eFuse
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.Efuse).Down)
                {
                    Response.Report("eFuse Pre-Check ...", EnumMessageLevel.StartPoint);
                    using (var eFuseMain = new EFuseMain())
                    {
                        eFuseMain.PreCheckFlow(null);
                    }
                    Response.Report("eFuse Pre-Check Completed !", EnumMessageLevel.EndPoint);
                }
                #endregion

                #region Scan
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.Scan).Down && LocalSpecs.IsModuleIncluded(BlockConst.Scan))
                {
                    Response.Report("Scan Pre-Check ...", EnumMessageLevel.StartPoint);
                    using (var scanMain = new ScanMain())
                    {
                        scanMain.PreCheckFlow(null);
                    }
                    Response.Report("SCAN Pre-Check Completed !", EnumMessageLevel.EndPoint);
                }
                #endregion

                #region Rtos
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.Rtos).Down && LocalSpecs.IsModuleIncluded(BlockConst.Rtos))
                {
                    Response.Report("Rtos Pre-Check ...", EnumMessageLevel.StartPoint);

                    Response.Report("Rtos Pre-Check Completed !", EnumMessageLevel.EndPoint);
                }
                #endregion

                #region Mbist
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.Mbist).Down)
                {
                    Response.Report("Mbist Pre-Check ...", EnumMessageLevel.StartPoint);
                    using (var bistBiraMain = new BistBiraMain())
                    {
                        bistBiraMain.PreCheckFlow(null);
                    }
                    Response.Report("Mbist Pre-Check Completed !", EnumMessageLevel.EndPoint);
                }
                #endregion

                #region EVS
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.Evs).Down && LocalSpecs.IsModuleIncluded(BlockConst.Evs))
                {
                    Response.Report("EVS Pre-Check ...", EnumMessageLevel.StartPoint);
                    using (var evsMain = new EvsMain())
                    {
                        evsMain.PreCheckFlow(null);
                    }
                    Response.Report("EVS Pre-Check Completed !", EnumMessageLevel.EndPoint);
                }
                #endregion

                #region HardIP
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.HardIp).Down)
                {
                    Response.Report("HardIP Pre-Check ...", EnumMessageLevel.StartPoint);
                    var hardIpParaData = new HardIpParaData(EnumBlock.HardIp);
                    using (var hardIpMain = new HardIpMain())
                    {
                        hardIpMain.PreCheckFlow(hardIpParaData);
                    }
                    Response.Report("HardIP Pre-Check Completed !", EnumMessageLevel.EndPoint);
                }
                #endregion

                #region BinCut
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.BinCut).Down && LocalSpecs.IsModuleIncluded(BlockConst.BinCut))
                {
                    Response.Report("BinCut Pre-Check ...", EnumMessageLevel.StartPoint);
                    using (var binCutMain = new BinCutMain())
                    {
                        binCutMain.PreCheckFlow(null);
                    }
                    Response.Report("BinCut Pre-Check Completed !", EnumMessageLevel.EndPoint);
                }
                #endregion

                #region PostAction
                Response.Report("Post-Action Pre-Check ...", EnumMessageLevel.StartPoint);
                using (var postActionMain = new PostActionMain())
                {
                    postActionMain.PreCheckFlow(null);
                }
                Response.Report("Post-Action Pre-Check Completed !", EnumMessageLevel.EndPoint);
                #endregion
            }
            catch (Exception e)
            {
                Response.Report("Meet an error validating. " + e.StackTrace, EnumMessageLevel.Error, 0);
            }
        }
    }
}
