using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

using CommonLib.DataStructure;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.Interop;
using CommonLib.Extension;
using CommonLib.ResponseManager;

using Microsoft.Office.Interop.Excel;

using Range = Microsoft.Office.Interop.Excel.Range;

using OfficeOpenXml;
using OfficeOpenXml.VBA;

using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace CommonLib.ErrorReport
{
    public class ErrorReportManager
    {
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int id);
        public static ErrorInstance ErrorInstance = ErrorInstance.Instance;
        public static bool ReportEnable = true;
        public static bool StopByCritical = true;

        #region Interop
        //For Autogen
        public static void GenAutogenErrorReport(string planCopyFile, string scghCopyFile,
            string binCutCopyFile, string postBinCutCopyFile, string modeSeqCopyFile)
        {
            if (!ReportEnable)
            {
                return;
            }

            Response.Report("Creating Error Report ...", MessageLevel.CheckPoint);
            Response.Report($"Total error count: {ErrorInstance.GetErrorCount()}", MessageLevel.Error);

            Workbook planWorkbook = OpenWorkbookIfExists(planCopyFile);
            Workbook scghWorkbook = OpenWorkbookIfExists(scghCopyFile);
            Workbook binCutWorkbook = OpenWorkbookIfExists(binCutCopyFile);
            Workbook binCutPostWorkbook = OpenWorkbookIfExists(postBinCutCopyFile);
            Workbook binCutModeSeqWorkbook = OpenWorkbookIfExists(modeSeqCopyFile);

            var workbookList = new List<Workbook>();
            if (binCutWorkbook != null)
            {
                workbookList.Add(binCutWorkbook);
            }

            if (scghWorkbook != null)
            {
                workbookList.Add(scghWorkbook);
            }

            if (planWorkbook != null)
            {
                workbookList.Add(planWorkbook);
            }

            if (binCutPostWorkbook != null)
            {
                workbookList.Add(binCutPostWorkbook);
            }

            if (binCutModeSeqWorkbook != null)
            {
                workbookList.Add(binCutModeSeqWorkbook);
            }

            try
            {
                Dictionary<Workbook, List<Error>> dic = GroupbyWorkbook(workbookList, ErrorInstance.GetErrorList(), planWorkbook);
                List<Type> typeList = ErrorInstance.GetErrorTypeList();
                int cnt = typeList.Count;
                foreach (KeyValuePair<Workbook, List<Error>> item in dic)
                {
                    InitialSummarySheet(item.Key);
                    for (int index = 0; index < typeList.Count; index++)
                    {
                        Type type = typeList[index];
                        List<Error> errorList = ErrorInstance.GetErrorsByType(item.Value, type);
                        string message = $"Type :{type.Name} , Count {errorList.Count}";
                        if (errorList.Count > 0)
                        {
                            Response.Report(message, MessageLevel.Error, (index * 100 / cnt));
                        }

                        ErrorReportInterop report = ErrorReportFactory.GetReport(type, errorList);
                        report.Write(item.Key);
                    }
                }

                foreach (Workbook workbook in workbookList)
                {
                    Application app = workbook.Parent;
                    app.DisplayAlerts = false;
                    app.AlertBeforeOverwriting = false;
                    app.Visible = false;
                    string fullName = workbook.FullName;
                    workbook.SaveAs(Path.ChangeExtension(fullName, ".xlsm"), XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                    if (File.Exists(fullName) && !Path.GetExtension(fullName).Equals(".xlsm"))
                    {
                        File.Delete(fullName);
                    }

                    if (dic.ContainsKey(workbook))
                    {
                        string outString =
                            "Errors occurred during test program generation. Please check the SummaryReport sheet in \n" +
                            fullName;
                        Response.Report(outString, MessageLevel.Error, 100);
                    }
                }

                foreach (Workbook workbook in workbookList)
                {
                    Application app = workbook.Parent;
                    workbook.Close(false);
                    var intPtr = new IntPtr(app.Hwnd);
                    int excelProcessId;
                    GetWindowThreadProcessId(intPtr, out excelProcessId);
                    var excelProcess = Process.GetProcessById(excelProcessId);
                    if (excelProcess != null)
                    {
                        excelProcess.Kill();
                        excelProcess.Dispose();
                    }
                }
            }

            catch (Exception e)
            {
                string outString = "Writing ErrorReport failed " + e.StackTrace;
                Response.Report(outString, MessageLevel.Error, 100);
            }

            finally
            {
                Response.Report("Error Report Completed!", MessageLevel.EndPoint);
            }
            //    foreach (var workbook in workbookList)
            //    {
            //        Application app = workbook.Parent;
            //        workbook.Close(false);
            //        IntPtr intPtr = new IntPtr(app.Hwnd);
            //        int excelProcessId;
            //        GetWindowThreadProcessId(intPtr, out excelProcessId);
            //        Process excelProcess = Process.GetProcessById(excelProcessId);
            //        if (excelProcess != null)
            //        {
            //            excelProcess.Kill();
            //            excelProcess.Dispose();
            //        }
            //    }
            //}
        }

        private static Workbook OpenWorkbookIfExists(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return null;
            }

            var app = new Application
            {
                Visible = false
            };
            return app.Workbooks.Open(Path.GetFullPath(filePath));
        }

        public static void InitialSummarySheet(Workbook wBook)
        {
            #region Delete sheet if alreadey exists
            foreach (Worksheet sheet in wBook.Worksheets)
            {
                if (sheet.Name == "SummaryReport" || sheet.Name.Contains("_ErrorReport"))
                {
                    wBook.Application.DisplayAlerts = false;
                    sheet.Delete();
                    break;
                }
            }
            #endregion

            Worksheet wSheet = wBook.Worksheets.Add(wBook.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
            wSheet.Name = "SummaryReport";
            wSheet.Cells[1, 1].Value = "ErrorPart";
            wSheet.Cells[1, 2].Value = "ErrorCount";
            wSheet.Cells[1, 3].Value = "WarnCount";
            wSheet.get_Range((Range)wSheet.Cells[1, 1], (Range)wSheet.Cells[1, 3]).Font.Bold = true;
        }

        private static Dictionary<Workbook, List<Error>> GroupbyWorkbook(List<Workbook> workBookList, List<Error> errors, Workbook planWorkbook)
        {
            var dic = new Dictionary<Workbook, List<Error>>();
            var sheetNames = new Dictionary<Workbook, List<string>>();
            foreach (Workbook workBook in workBookList)
            {
                sheetNames.Add(workBook, workBook.Worksheets.OfType<Worksheet>().Select(x => x.Name).ToList());
            }

            foreach (Error error in errors)
            {
                bool found = false;
                foreach (KeyValuePair<Workbook, List<string>> workBook in sheetNames)
                {
                    if (workBook.Value.Exists(x => x.Equals(error.SheetName)))
                    {
                        if (dic.ContainsKey(workBook.Key))
                        {
                            dic[workBook.Key].Add(error);
                        }
                        else
                        {
                            dic.Add(workBook.Key, new List<Error> { error });
                        }

                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    if (dic.ContainsKey(planWorkbook))
                    {
                        dic[planWorkbook].Add(error);
                    }
                    else
                    {
                        dic.Add(planWorkbook, new List<Error> { error });
                    }
                }
            }
            return dic;
        }
        #endregion

        #region Epplus
        public static void GenErrorReport(ExcelPackage excelPackage, List<string> copyFiles, string errorReportName, string summaryReport = "SummaryReport")
        {
            ErrorInstance.GenErrorReport(excelPackage, copyFiles, errorReportName, summaryReport);
        }

        public static void GenErrorReport(ExcelPackage excelPackage, string errorReportName)
        {
            ErrorInstance.GenErrorReport(excelPackage, errorReportName);
        }

        public static void GenErrorReportRaw(ExcelPackage excelPackage, string errorReportName)
        {
            ErrorInstance.GenErrorReportRaw(excelPackage, errorReportName);
        }

        public static bool AddMarcoFromBas(ExcelPackage excel)
        {
            bool flag = false;
            if (excel.Workbook.VbaProject == null)
            {
                excel.Workbook.CreateVBAProject();
            }

            const string libid = @"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\MSO.DLL#Microsoft Office 14.0 Object Library";
            if (!excel.Workbook.VbaProject.References.ToList().Exists(x => x.Libid.Equals(libid, StringComparison.CurrentCultureIgnoreCase)) && File.Exists(@"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\MSO.DLL"))
            {
                var refer = new ExcelVbaReference
                {
                    Libid = libid,
                    Name = @"Office"
                };
                excel.Workbook.VbaProject.References.Add(refer);
                flag = true;
            }

            if (excel.Workbook.VbaProject.References.ToList().Exists(x => x.Libid.Equals(libid, StringComparison.CurrentCultureIgnoreCase)))
            {
                const string moduleName = "LIB_General";
                ExcelVBAModule module = excel.IsExistModule(moduleName) ? excel.Workbook.VbaProject.Modules[moduleName] : excel.Workbook.VbaProject.Modules.AddModule(moduleName);
                var lines = new List<string>();
                lines.Add("Sub GoBack()");
                lines.Add("Attribute GoBack.VB_ProcData.VB_Invoke_Func = \"q\\n14\"");
                lines.Add("    sheetName = ThisWorkbook.BuiltinDocumentProperties(\"subject\").Value");
                lines.Add("    If sheetName <> \"\" Then");
                lines.Add("    ThisWorkbook.Activate");
                lines.Add("    Sheets(sheetName).Select");
                lines.Add("    End If");
                lines.Add("End Sub");
                module.Code = string.Join("\r\n", lines);

                const string moduleName1 = "ThisWorkbook";
                ExcelVBAModule module1 = excel.IsExistModule(moduleName1) ? excel.Workbook.VbaProject.Modules[moduleName1] : excel.Workbook.VbaProject.Modules.AddModule(moduleName1);
                lines = new List<string>();
                lines.Add("Private Sub Workbook_Open()");
                lines.Add("");
                lines.Add("    Dim ContextMenu As CommandBar");
                lines.Add("    Dim ctrl As CommandBarControl");
                lines.Add("    Set ContextMenu = Application.CommandBars(\"Cell\")");
                lines.Add("    ");
                lines.Add("    ' Delete the controls first to avoid duplicates.");
                lines.Add("    For Each ctrl In ContextMenu.Controls");
                lines.Add("        If ctrl.Tag = \"My_Cell_Control_Tag\" Then");
                lines.Add("            ctrl.Delete");
                lines.Add("        End If");
                lines.Add("    Next ctrl");
                lines.Add("");
                lines.Add("    ' Add one custom button to the Cell context menu.");
                lines.Add("    With ContextMenu.Controls.Add(Type:=msoControlButton)");
                lines.Add("        .OnAction = \"'\" & ThisWorkbook.Name & \"'!\" & \"GoBack\"");
                lines.Add("        .FaceId = 59");
                lines.Add("        .Caption = \"GoBack (Ctrl+q)\"");
                lines.Add("        .Tag = \"My_Cell_Control_Tag\"");
                lines.Add("    End With");
                lines.Add("    ");
                lines.Add("    ' Add a separator");
                lines.Add("    ContextMenu.Controls(2).BeginGroup = True");
                lines.Add("    Application.MacroOptions Macro:=\"GoBack\"");
                lines.Add("End Sub");
                lines.Add("");
                lines.Add("Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)");
                lines.Add("");
                lines.Add("    Me.BuiltinDocumentProperties(\"subject\") = Sh.Name");
                lines.Add("");
                lines.Add("End Sub");
                module1.Code = string.Join("\r\n", lines);
            }
            return flag;
        }
        #endregion

        public static void AddError(EnumErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(string errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(OtpErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(PmicErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(HardIpErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(EFuseErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(BasicErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(PreActionErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(MbistErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(ScanErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(EvsErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(HtolErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(RtosErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(MainFlowErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(RelayErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(BinCutErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(DuplicateInstance errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(DuplicateTestNumber errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(UnusedDcCategory errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(ZeroVoltageDcCategory errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(PatValtPinCheckerType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(PatInfoType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }
        public static void AddError(SsnErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(HarvestErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(PatternMissing errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }

        public static void AddError(FlowMainErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            Add(errorType, errorLevel, sheetName, rowNum, columnNum, message, comments);
        }
        //private static void Add(object errorType, ErrorLevel errorLevel, string sheetName, int rowNum, string message, params string[] comments)
        //{
        //    var errorNew = new Error();
        //    errorNew.ErrorType = errorType;
        //    errorNew.SheetName = sheetName;
        //    errorNew.RowNum = rowNum;
        //    errorNew.Comments = comments.ToList();
        //    errorNew.Message = message;
        //    errorNew.ErrorLevel = errorLevel;
        //    AddError(errorNew);
        //    return;
        //}

        private static void Add(object errorType, ErrorLevel errorLevel, string sheetName, int rowNum, int columnNum, string message, params string[] comments)
        {
            if (errorLevel == ErrorLevel.Critical && StopByCritical)
            {
                Response.Report(message, MessageLevel.Error, 0);
                Response.Report("Critical error occurred. Exiting...", MessageLevel.Error, 0);
                Environment.Exit(-1);
            }
            var errorNew = new Error
            {
                ErrorType = errorType,
                SheetName = sheetName,
                RowNum = rowNum
            };

            errorNew.ColNum = (columnNum < 1) ? 0 : columnNum;
            errorNew.Comments = comments.ToList();
            errorNew.Message = message;
            errorNew.ErrorLevel = errorLevel;
            AddError(errorNew);
        }

        public static bool Contains(string message)
        {
            return ErrorInstance.Exist(message);
        }

        public static void AddError(Error error)
        {
            ErrorInstance.AddError(error);
        }

        public static void AddErrors(List<Error> errors)
        {
            ErrorInstance.AddErrors(errors);
        }

        public static void ResetErrors()
        {
            ErrorInstance.Clear();
        }

        public static int GetErrorCountByType(object errorType)
        {
            Type type = errorType.GetType();
            int number = ErrorInstance.GetErrorCountByType(type);
            return number;

        }

        public static List<Error> GetErrorList()
        {
            return ErrorInstance.GetErrorList();
        }

        public static int GetErrorCount()
        {
            return ErrorInstance.GetErrorCount();
        }

        #region Print
        public static void PrintErrors(string file)
        {
            var lines = new List<string>();
            var headers = new List<string> { "SheetName", "ErrorType", "Link", "ErrorLevel", "RowNum", "ColNum", "Message", "Comments" };
            lines.Add(string.Join("\t", headers));
            foreach (Error error in ErrorInstance.GetErrorList())
            {
                var list = new List<string>
                {
                    error.SheetName,
                    error.ErrorType.ToString(),
                    error.Link,
                    error.ErrorLevel.ToString(),
                    error.RowNum.ToString(),
                    error.ColNum.ToString(),
                    error.Message,
                    string.Join("\t", error.Comments)
                };
                string line = string.Join("\t", list);
                lines.Add(line);
            }
            File.WriteAllLines(file, lines);
        }

        public static DataTable GetErrorInfo()
        {
            var table = new DataTable();
            table.Columns.Add("");
            table.Columns.Add("");
            List<Error> errorList = ErrorInstance.GetSortedErrorList();
            foreach (Error item in errorList)
            {
                string tLog = "Sheet:[" + item.SheetName + "], RowNum:[" + item.RowNum + "], Message:[" + item.Message + "]";
                string logType = "Report[" + item.ErrorLevel + "]";
                System.Data.DataRow row = table.NewRow();
                if (tLog.Length > 8000)
                {
                    row[0] = tLog.Substring(0, 7000);
                }
                else
                {
                    row[0] = tLog;
                }

                if (logType.Length > 8000)
                {
                    row[0] = logType.Substring(0, 7000);
                }

                row[1] = logType;
                table.Rows.Add(row);
            }
            return table;
        }
        #endregion
    }
}
