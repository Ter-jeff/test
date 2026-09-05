using System;
using System.Collections.Generic;
using System.Linq;

namespace RfLib.Dvdc.Base
{
    public class RelayPathBusiness
    {
        public static void Write(List<string> content, List<RelayPathItem> relayPathItems, Dictionary<string, string> calPathToLut)
        {
            Header(content);
            var allDest = relayPathItems.SelectMany(p => p.GetPorts()).Distinct().ToList();
            RfPort(content, allDest);
            GetEnumName(content, allDest);
            GetrfPort(content, allDest);
            Init_Dictionary_PathName(content, relayPathItems, calPathToLut);
            Init_Dictionary_FromTo(content, relayPathItems, calPathToLut, allDest);
            ComponentSetup(content, relayPathItems);
        }

        private static void Header(List<string> content)
        {
            string time = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");

            content.Add("VERSION 1.0 CLASS");
            content.Add("BEGIN");
            content.Add("  MultiUse = -1  'True");
            content.Add("END");
            content.Add("Attribute VB_Name = \"Class_RFPathCtrlUtil\"");
            content.Add("Attribute VB_GlobalNameSpace = False");
            content.Add("Attribute VB_Creatable = False");
            content.Add("Attribute VB_PredeclaredId = False");
            content.Add("Attribute VB_Exposed = False");
            content.Add("option explicit");
            content.Add($"\'\'\'{time}\'\'\'");
            content.Add($"\'\'\'{"Cal-Path-Disabled"}\'\'");
            content.Add("Public dictPathName as New Dictionary");
            content.Add("Public dictFromTo as New Dictionary");
            content.Add("");

            #region StreamWrite sr
            //sr.WriteLine("option explicit");
            //sr.WriteLine("\'\'\'{0}\'\'\'",time);
            //sr.WriteLine("\'\'\'{0}\'\'", "Cal-Path-Disabled");
            //sr.WriteLine("Public dictPathName as New Dictionary");
            //sr.WriteLine("Public dictFromTo as New Dictionary");
            //sr.WriteLine();
            /*option explicit
                '''8/25/2022 5:43:02 PM'''
                '''Cal-Path-Disabled'''
                Public dictPathName as New Dictionary
                Public dictFromTo as New Dictionary
             */
            #endregion
        }

        private static void RfPort(List<string> content, List<string> ports)
        {
            content.Add("Public Enum rfPort");
            int i = 1;
            foreach (string port in ports)
            {
                content.Add($"\t{port}={i}");
                i++;
            }
            content.Add("End Enum");
            content.Add("");

            #region StreamWrite sr
            //sr.WriteLine("Public Enum rfPort");
            //int i = 1;
            //foreach (var port in ports)
            //{
            //    sr.WriteLine("{0}={1}",port,i);
            //    i++;
            //}
            //sr.WriteLine("End Enum");
            //sr.WriteLine();
            #endregion
        }

        private static void GetEnumName(List<string> content, List<string> ports)
        {
            content.Add("Public Function getEnumName(port as rfPort) as string");
            content.Add("Select Case port");
            int i = 1;
            foreach (string port in ports)
            {
                content.Add($"\tCase {i}:getEnumName=\"{port}\"");
                i++;
            }
            content.Add("End Select");
            content.Add("End Function");
            content.Add("");

            #region StreamWrite sr
            //sr.WriteLine("Public Function getEnumName(port as rfPort) as string");
            //sr.WriteLine("Select Case port");
            //int i = 1;
            //foreach (var port in ports)
            //{
            //    sr.WriteLine("Case {0}:getEnumName=\"{1}\"", i, port);
            //    i++;
            //}
            //sr.WriteLine("End Select");
            //sr.WriteLine("End Function");
            //sr.WriteLine();
            #endregion
        }

        private static void GetrfPort(List<string> content, List<string> ports)
        {
            content.Add("Public Function getrfPort(strPort As string) As rfPort");
            content.Add("Select Case strPort");
            int i = 1;
            foreach (string port in ports)
            {
                content.Add($"\tCase \"{port}\":getrfPort={i}");
                i++;
            }
            content.Add("End Select");
            content.Add("End Function");
            content.Add("");

            #region StreamWrite sr
            //sr.WriteLine("Public Function getrfPort(strPort As string) As rfPort");
            //sr.WriteLine("Select Case strPort");
            //int i = 1;
            //foreach (var port in ports)
            //{
            //    sr.WriteLine("Case \"{0}\":getrfPort={1}", port, i);
            //    i++;
            //}
            //sr.WriteLine("End Select");
            //sr.WriteLine("End Function");
            //sr.WriteLine();
            #endregion
        }

        private static void Init_Dictionary_PathName(List<string> content, List<RelayPathItem> relayPathItems, Dictionary<string, string> calDictionary)
        {
            content.Add("Public Function Init_Dictionary_PathName() as Long");
            foreach (RelayPathItem relayPathItem in relayPathItems)
            {
                if (calDictionary.TryGetValue(relayPathItem.PathName, out string? lut))
                {
                    content.Add($"\tdictPathName.Add \"{relayPathItem.PathName}\",\"{relayPathItem.GetPorts()[0]}:{relayPathItem.GetPorts()[1]}:{lut}\"");
                }
            }
            content.Add("End Function");
            content.Add("");

            #region StreamWrite sr
            //sr.WriteLine("Public Function Init_Dictionary_PathName() as Long");
            //foreach (var relayPathItem in relayPathItems)
            //{
            //    if (calDictionary.ContainsKey(relayPathItem.PathName))
            //    {
            //        sr.WriteLine("dictPathName.Add \"{0}\",\"{1}:{2}:{3}\"", relayPathItem.PathName, relayPathItem.GetPorts()[0], relayPathItem.GetPorts()[1], calDictionary[relayPathItem.PathName]);
            //    }
            //    else
            //    {
            //        ;
            //    }
            //}
            //sr.WriteLine("End Function");
            //sr.WriteLine();
            /*Public Function Init_Dictionary_PathName() as Long
	            dictPathName.Add "ARF_PADIO_TXOUT_5G_1_to_MW7G_MEAS","ARF_PADIO_TXOUT_5G_1:MW7G_MEAS:TXOUT_5G_1"
	            dictPathName.Add "ARF_PADIO_TXOUT_5G_1_to_ARF_PADIO_RXIN_5G_1","ARF_PADIO_TXOUT_5G_1:ARF_PADIO_RXIN_5G_1:TXOUT_5G_1_LB"
	            dictPathName.Add "ARF_PADIO_TXOUT_5G_0_to_MW7G_MEAS","ARF_PADIO_TXOUT_5G_0:MW7G_MEAS:TXOUT_5G_0"
	            dictPathName.Add "ARF_PADIO_TXOUT_5G_0_to_ARF_PADIO_RXIN_5G_0","ARF_PADIO_TXOUT_5G_0:ARF_PADIO_RXIN_5G_0:TXOUT_5G_0_LB"
	            dictPathName.Add "ARF_PADIO_RXIN_5G_0_to_MW7G_SRC","ARF_PADIO_RXIN_5G_0:MW7G_SRC:RXIN_5G_0"
                         * ....
	            dictPathName.Add "ARF_PADIO_CLKTEST_OUT_to_UW","ARF_PADIO_CLKTEST_OUT:UW:CLKTEST_OUT"
	            dictPathName.Add "ARF_PADIO_CLKTEST_OUT_to_UP1600","ARF_PADIO_CLKTEST_OUT:UP1600:N/A"
	            dictPathName.Add "ARF_PADIO_RXPLL_CLK_TEST_2G_to_UW","ARF_PADIO_RXPLL_CLK_TEST_2G:UW:RXPLL_CLK_TEST_2G"
	            dictPathName.Add "ARF_PADIO_RXPLL_CLK_TEST_2G_to_UP1600","ARF_PADIO_RXPLL_CLK_TEST_2G:UP1600:N/A"
	            dictPathName.Add "UPAC_TEST_to_BUFFER_to_CAP","UPAC_TEST:BUFFER_to_CAP:N/A"
            End Function
             */
            #endregion
        }

        private static void Init_Dictionary_FromTo(List<string> content, List<RelayPathItem> relayPathItems, Dictionary<string, string> calDictionary, List<string> allPorts)
        {
            content.Add("Public Function Init_Dictionary_FromTo() as Long");
            foreach (RelayPathItem relayPathItem in relayPathItems)
            {
                if (calDictionary.TryGetValue(relayPathItem.PathName, out string? lut))
                {
                    int fromIndex = allPorts.IndexOf(relayPathItem.GetPorts()[0]) + 1;
                    int toIndex = allPorts.IndexOf(relayPathItem.GetPorts()[1]) + 1;
                    content.Add(string.Format("\tdictFromTo.Add \"{0}:{1}\",\"{5}:{2}:{3}:{4}\"", fromIndex, toIndex, relayPathItem.GetPorts()[0], relayPathItem.GetPorts()[1], lut, relayPathItem.PathName));
                }
            }
            content.Add("End Function");
            content.Add("");

            #region StreamWrite sr
            //sr.WriteLine("Public Function Init_Dictionary_FromTo() as Long");
            //foreach (var relayPathItem in relayPathItems)
            //{
            //    if (calDictionary.ContainsKey(relayPathItem.PathName))
            //    {
            //        var fromIndex = allPorts.IndexOf(relayPathItem.GetPorts()[0]) + 1;
            //        var toIndex = allPorts.IndexOf(relayPathItem.GetPorts()[1]) + 1;
            //        sr.WriteLine("dictFromTo.Add \"{0}:{1}\",\"{5}:{2}:{3}:{4}\"",
            //            fromIndex, toIndex, relayPathItem.GetPorts()[0], relayPathItem.GetPorts()[1], calDictionary[relayPathItem.PathName], relayPathItem.PathName);
            //    }
            //    else
            //    {
            //        ;
            //    }
            //}
            //sr.WriteLine("End Function");
            //sr.WriteLine();
            /*Public Function Init_Dictionary_FromTo() as Long
	            dictFromTo.Add "1:2","ARF_PADIO_TXOUT_5G_1_to_MW7G_MEAS:ARF_PADIO_TXOUT_5G_1:MW7G_MEAS:TXOUT_5G_1"
	            dictFromTo.Add "1:3","ARF_PADIO_TXOUT_5G_1_to_ARF_PADIO_RXIN_5G_1:ARF_PADIO_TXOUT_5G_1:ARF_PADIO_RXIN_5G_1:TXOUT_5G_1_LB"
	            dictFromTo.Add "4:2","ARF_PADIO_TXOUT_5G_0_to_MW7G_MEAS:ARF_PADIO_TXOUT_5G_0:MW7G_MEAS:TXOUT_5G_0"
	            dictFromTo.Add "4:5","ARF_PADIO_TXOUT_5G_0_to_ARF_PADIO_RXIN_5G_0:ARF_PADIO_TXOUT_5G_0:ARF_PADIO_RXIN_5G_0:TXOUT_5G_0_LB"
	            dictFromTo.Add "5:6","ARF_PADIO_RXIN_5G_0_to_MW7G_SRC:ARF_PADIO_RXIN_5G_0:MW7G_SRC:RXIN_5G_0"
	            dictFromTo.Add "3:6","ARF_PADIO_RXIN_5G_1_to_MW7G_SRC:ARF_PADIO_RXIN_5G_1:MW7G_SRC:RXIN_5G_1"
	            dictFromTo.Add "7:8","ARF_PADIO_TX_LOGEN_TEST_5G_to_IFSRC_MEAS1:ARF_PADIO_TX_LOGEN_TEST_5G:IFSRC_MEAS1:TX_LOGEN_TEST_5G"
	            dictFromTo.Add "9:8","ARG_TXOUT_to_IFSRC_MEAS1:ARG_TXOUT:IFSRC_MEAS1:ARG_TXOUT"
                         * ....
	            dictFromTo.Add "60:16","ARF_PADIO_RXPLL_CLK_TEST_2G_to_UW:ARF_PADIO_RXPLL_CLK_TEST_2G:UW:RXPLL_CLK_TEST_2G"
	            dictFromTo.Add "60:23","ARF_PADIO_RXPLL_CLK_TEST_2G_to_UP1600:ARF_PADIO_RXPLL_CLK_TEST_2G:UP1600:N/A"
	            dictFromTo.Add "61:30","UPAC_TEST_to_BUFFER_to_CAP:UPAC_TEST:BUFFER_to_CAP:N/A"
            End Function
            */
            #endregion
        }

        private static void ComponentSetup(List<string> content, List<RelayPathItem> relayPathItems)
        {
            //ARF_PADIO_OUT_TEST_RXBB_2G_0-BUFFER-UW
            foreach (RelayPathItem relayPathItem in relayPathItems)
            {
                string setupName = relayPathItem.PathName.Replace("-", "_to_");
                content.Add($"Public Function Setup_{setupName}() as Long");
                var componentsStatusGroup =
                    relayPathItem.ComponentInfos.GroupBy(p => p.Type + "#" + p.Status)
                        .ToDictionary(p => p.Key, p => p.ToList());
                if (componentsStatusGroup.TryGetValue("Relay#1", out List<Component>? relayOn))
                {
                    content.Add($"\tthehdw.Utility.Pins(\"{string.Join(",", relayOn.Select(p => p.Name))}\").State = tlUtilBitOn");
                    content.Add($"\ttheexec.Datalog.WriteComment \"RelayON: {string.Join(",", relayOn.Select(p => p.Name))}\"");
                }
                if (componentsStatusGroup.TryGetValue("Relay#0", out List<Component>? relayOff))
                {
                    content.Add($"\tthehdw.Utility.Pins(\"{string.Join(",", relayOff.Select(p => p.Name))}\").State = tlUtilBitOff");
                    content.Add($"\ttheexec.Datalog.WriteComment \"RelayOFF: {string.Join(",", relayOff.Select(p => p.Name))}\"");
                }
                if (componentsStatusGroup.TryGetValue("Switch#1", out List<Component>? switchOn))
                {
                    content.Add($"\tthehdw.Digital.Pins(\"{string.Join(",", switchOn.Select(p => p.Name))}\").InitState = chInitHi");
                    content.Add($"\ttheexec.Datalog.WriteComment \"InitHI: {string.Join(",", switchOn.Select(p => p.Name))}\"");
                }
                if (componentsStatusGroup.TryGetValue("Switch#0", out List<Component>? switchOff))
                {
                    content.Add($"\tthehdw.Digital.Pins(\"{string.Join(",", switchOff.Select(p => p.Name))}\").InitState = chInitLo");
                    content.Add($"\ttheexec.Datalog.WriteComment \"InitLO: {string.Join(",", switchOff.Select(p => p.Name))}\"");
                }
                content.Add("\tTheExec.Datalog.WriteComment \"==========================================\"");
                content.Add("End Function");
                content.Add("");
            }

            #region StreamWrite sr
            ////ARF_PADIO_OUT_TEST_RXBB_2G_0-BUFFER-UW
            //foreach (var relayPathItem in relayPathItems)
            //{
            //    var setupName = relayPathItem.PathName.Replace("-", "_to_");
            //    sr.WriteLine("Public Function Setup_{0}() as Long", setupName);
            //    var componentsStatusGroup =
            //        relayPathItem.ComponentInfos.GroupBy(p => p.Type + "#" + p.Status)
            //            .ToDictionary(p => p.Key, p => p.ToList());
            //    if (componentsStatusGroup.ContainsKey("Relay#1"))
            //    {
            //        sr.WriteLine("thehdw.Utility.Pins(\"{0}\").State=tlUtilBitOn", string.Join(",", componentsStatusGroup["Relay#1"].Select(p => p.Name)));
            //        sr.WriteLine("theexec.Datalog.WriteComment \"RelayON: {0}\"", string.Join(",", componentsStatusGroup["Relay#1"].Select(p => p.Name)));
            //    }
            //    if (componentsStatusGroup.ContainsKey("Relay#0"))
            //    {
            //        sr.WriteLine("thehdw.Utility.Pins(\"{0}\").State=tlUtilBitOff", string.Join(",", componentsStatusGroup["Relay#0"].Select(p => p.Name)));
            //        sr.WriteLine("theexec.Datalog.WriteComment \"RelayOFF: {0}\"", string.Join(",", componentsStatusGroup["Relay#0"].Select(p => p.Name)));
            //    }
            //    if (componentsStatusGroup.ContainsKey("Switch#1"))
            //    {
            //        sr.WriteLine("thehdw.Digital.Pins(\"{0}\").InitState=chInitHi", string.Join(",", componentsStatusGroup["Switch#1"].Select(p => p.Name)));
            //        sr.WriteLine("theexec.Datalog.WriteComment \"InitHI: {0}\"", string.Join(",", componentsStatusGroup["Switch#1"].Select(p => p.Name)));
            //    }
            //    if (componentsStatusGroup.ContainsKey("Switch#0"))
            //    {
            //        sr.WriteLine("thehdw.Digital.Pins(\"{0}\").InitState=chInitLo", string.Join(",", componentsStatusGroup["Switch#0"].Select(p => p.Name)));
            //        sr.WriteLine("theexec.Datalog.WriteComment \"InitLO: {0}\"", string.Join(",", componentsStatusGroup["Switch#0"].Select(p => p.Name)));
            //    }
            //    sr.WriteLine("TheExec.Datalog.WriteComment \"==========================================\"");
            //    sr.WriteLine("End Function");
            //    sr.WriteLine();
            //}
            /*Public Function Setup_ARF_PADIO_OUT_TEST_RXBB_2G_0_to_BUFFER_to_UW() as Long
	            thehdw.Utility.Pins("K23,K24").State=tlUtilBitOn
	            theexec.Datalog.WriteComment "RelayON: K23,K24"
	            thehdw.Utility.Pins("K66").State = tlUtilBitOff
	            theexec.Datalog.WriteComment "RelayOFF: K66"
	            thehdw.Digital.Pins("U14_CTLA,U14_CTLB,U14_CTLC,U15_CTLA,U15_CTLB,U15_CTLC").InitState = chInitLo
	            theexec.Datalog.WriteComment "InitLO: U14_CTLA,U14_CTLB,U14_CTLC,U15_CTLA,U15_CTLB,U15_CTLC"
	            TheExec.Datalog.WriteComment "=========================================="
            End Function
             */
            #endregion
        }
    }
}
