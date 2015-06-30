using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using APLPX.Common;
using APLPX.Client.localhost;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.ComponentModel;
using System.Runtime.InteropServices;
using APLPX.Common.Helper;



namespace APLPX.Client
{
    public class AplBusinessLayer : IDisposable
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        #region private instance

        /// <summary>
        ///Import file name. 
        /// </summary>
        private String s_Filename;

        /// <summary>
        ///The key reference for disposable. 
        /// </summary>
        private bool m_isDisposed;

        /// <summary>
        ///The service  method accessing bool out param result. 
        /// </summary>
        private bool result = false;

        /// <summary>
        ///the service  method accessing bool out param result specification.  
        /// </summary>
        private bool resultSpc = true;

        /// <summary>
        ///The service insance for web infaces access. 
        /// </summary>
        private localhost.AplService ServiceProxy;
        /// <summary>
        /// Delegate method for progress bar action 
        /// </summary>
        delegate void PostMethodDelegate(string fileExt, string ProcessedFileName, out bool result);

        /// <summary>
        ///Background worker for status bar progressing 
        /// </summary>
        private BackgroundWorker worker;

        /// <summary>
        ///The parser instance to be used for csv and txt file parsing. 
        /// </summary>
        private CsvTxtParser csvTxtParser;

        #endregion

        #region Constructor

        /// <summary>
        ///Constructor 
        /// </summary>
        public AplBusinessLayer()
        {
            //Service initialize
            ServiceProxy = new AplService();
        }
        #endregion

        #region Public Methods



        /// <summary>
        /// Testing connection staging db connection configuration 
        /// </summary>
        /// <param name="stagingDbModel"> Staging db mode</param>
        /// <returns>Connection status</returns>
        public string TestConnection(localhost.StagingDbConfig stagingDbModel)
        {
            try
            {
                return ServiceProxy.TestConnection(stagingDbModel);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Saving db configuration interface
        /// </summary>
        /// <param name="stagingDbModel">Staging db model</param>
        /// <returns> result</returns>
        public string SaveStagingDbConfig(localhost.StagingDbConfig stagingDbModel)
        {
            try
            {
                return ServiceProxy.SaveStagingDbConnString(stagingDbModel);
            }
            catch (Exception)
            {

                throw;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="userId"></param>
        /// <returns></returns>
        public DataTable GetStagingDBConfigByUser(int userId)
        {
            try
            {
                using (DataSet resultDs = ServiceProxy.GetStagingDbInfoByUser(userId, resultSpc))
                {
                    return resultDs.Tables.Count > 0 ? resultDs.Tables[0] : new DataTable();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        /// <summary>
        /// Get Staging Db Row count
        /// </summary>
        /// <param name="userId"></param>
        /// <returns></returns>
        public bool CheckStagingDBConfigByUser(int userId)
        {
            try
            {
                bool userIdSpec = true;
                ServiceProxy.CheckStagingDBConfigByUserId(userId, userIdSpec, out result, out resultSpc);
                return result;

            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// The import data which is uploadedfile's. 
        /// </summary>
        /// <returns> process result</returns>
        public bool ImportData(string ProcessedFileName,out BackgroundWorker progressValue)
        {
            worker = new BackgroundWorker();
            PostMethodDelegate dlgt;
            DateTime startTime = System.DateTime.Now;
            string fileExt = Path.GetExtension(ProcessedFileName).Replace(".", "");
            bool result = false;
            dlgt = new PostMethodDelegate(StartProgress);
            AsyncCallback asyncCallBack = new AsyncCallback(show);
            worker.WorkerReportsProgress = true;
            worker.ReportProgress(10);
           
            progressValue = worker;
            IAsyncResult ar = dlgt.BeginInvoke(fileExt, ProcessedFileName, out result, asyncCallBack, dlgt);            
            return result;
        }
        /// <summary>
        ///File processing status updated into processed file status. 
        /// </summary>
        /// <param name="rowNum"> current processed file row number</param>
        void StartProgress(string fileExt, string ProcessedFileName, out bool result)
        {
            DateTime startTime = System.DateTime.Now;
            result = false;
            switch (fileExt)
            {
                case Utility.TXT:
                    result = ExtractTxtFiles(ProcessedFileName, startTime);
                    break;
                case Utility.CSV:
                    result = ExtractCsvFiles(ProcessedFileName, startTime);
                    break;
                case Utility.XLS:
                case Utility.XLSX:
                    result = ExtractXLFiles(ProcessedFileName, startTime);
                    break;
            }


        }
        /// <summary>
        ///Async method to start the fileprocess 
        /// </summary>
        /// <param name="asyncResult"></param>
        void show(IAsyncResult asyncResult)
        {
            // PostMethodDelegate worker = (PostMethodDelegate)((AsyncResult)asyncResult).AsyncDelegate;
            // worker.EndInvoke(asyncResult);

        }

        /// <summary>
        /// Get result dataset interface 
        /// </summary>
        /// <returns></returns>
        public DataSet GetResultSet()
        {
            try
            {
                using (DataSet resultDs = ServiceProxy.GetResultSet())
                {
                    return resultDs;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        #endregion

        #region Private methods

        /// <summary>
        /// Extract the data from csv  file
        /// and business rule validation for extracted data,If valid the data to stagging to the database.
        /// Invalid data's to write error file against the uploaded file's. 
        /// </summary>
        /// <param name="fileName"></param>
        private bool ExtractCsvFiles(string fileName, DateTime startTime)
        {
            worker.ReportProgress(10, fileName);
            csvTxtParser = new CsvTxtParser();
            string[][] valuelist;
            string[] tempvalueCollection = new string[0];
            DataColumn extractDc;
            DataTable extractDataTable = new DataTable();
            int TotalRowCount = 0;
            int TotalErrCount = 0;
            string concatFileName = string.Empty;
            string errFilename = Path.GetFileNameWithoutExtension(fileName) + "_Err";
            string sErrFilename = Path.GetDirectoryName(fileName) + "\\" + errFilename + ".csv";


            worker.ReportProgress(20, fileName);
            using (var stream = File.OpenRead(fileName))
            using (var reader = new StreamReader(stream))
            {
                valuelist = csvTxtParser.Parse(reader);
            }

            for (var iterateParsingList = 0; iterateParsingList < valuelist.Length; iterateParsingList++)
            {

                string[] childList = valuelist[iterateParsingList];

                //Sanity Check: Split the values from delimeter.
                if (childList[0].ToString().Contains(","))
                {
                    tempvalueCollection = childList[0].ToString().Split(',');
                }
                else if (childList[0].ToString().Contains("|"))
                {
                    tempvalueCollection = childList[0].ToString().Split('|');
                }
                else if (childList[0].ToString().Contains("\t"))
                {
                    tempvalueCollection = childList[0].ToString().Split('\t');
                }
                else if (childList[0].ToString().Contains(" "))
                {
                    tempvalueCollection = childList[0].ToString().Split(' ');
                }
                else
                {
                    tempvalueCollection = childList;

                    if (tempvalueCollection.Length > 0)
                    {
                        if (!iterateParsingList.Equals(0))
                        {
                            string[] cloneCopy = tempvalueCollection;
                            string[] tempList = new string[0];
                            tempvalueCollection = new string[7];
                            for (int iterateAppendval = 0; iterateAppendval < cloneCopy.Length; iterateAppendval++)
                            {
                                if (cloneCopy[iterateAppendval].Contains(","))
                                {
                                    tempvalueCollection = cloneCopy[iterateAppendval].Split(',');
                                }
                                else
                                {
                                    tempvalueCollection[iterateAppendval] = cloneCopy[iterateAppendval].ToString();
                                }
                            }
                        }
                    }
                }

                if (iterateParsingList == 0)
                {
                    for (int iterateValueList = 0; iterateValueList <= (tempvalueCollection.Length - 1); iterateValueList++)
                    {
                        extractDc = new DataColumn(Regex.Replace(tempvalueCollection[iterateValueList].ToString().Trim(), "['\"]", ""));
                        extractDataTable.Columns.Add(extractDc);
                    }
                }
                else
                {
                    DataRow dr = extractDataTable.NewRow();
                    for (int iterateValueList = 0; iterateValueList <= (tempvalueCollection.Length - 1); iterateValueList++)
                    {
                        if (tempvalueCollection[iterateValueList] != null)
                        {
                            dr[iterateValueList] = Regex.Replace(tempvalueCollection[iterateValueList].ToString().Trim(), "['\"]", "");
                        }
                        else
                        {
                            dr[iterateValueList] = tempvalueCollection[iterateValueList];
                        }
                    }
                    extractDataTable.Rows.Add(dr);
                    worker.ReportProgress(30, fileName);
                }

            }
                worker.ReportProgress(35, fileName);

                using (DataSet extractedData = new DataSet())
                {
                    extractedData.Tables.Add(extractDataTable);
                    //Sanity Check...
                    if (extractedData.Tables[0].Rows.Count > 0)
                    {
                        TotalRowCount = extractedData.Tables[0].Rows.Count;
                        DataSet passedRows = new DataSet();

                        using (DataSet errorRowsDs = ServiceProxy.ApplyBRuleForImportedData(extractedData, out passedRows))
                        {
                            worker.ReportProgress(40, fileName);
                            string FileName = Path.GetFileName(fileName);
                            if (errorRowsDs.Tables[0].Rows.Count > 0)
                            {
                                TotalErrCount = errorRowsDs.Tables[0].Rows.Count;

                                //If BussinessRule Field rows to write error file 
                                WriteCSVErrorFile(errorRowsDs, FileName.Replace("csv", ""), FileName.Replace("csv", ""), sErrFilename);
                                worker.ReportProgress(60, fileName);
                                //Update error result table for imported faild data entry
                                ServiceProxy.UpdateResultTable(FileName.Replace("csv", "") + FileName.Replace("csv", ""), Path.GetExtension(fileName), TotalRowCount, true, passedRows.Tables[0].Rows.Count,
                                                              true, TotalErrCount, true, fileName, errFilename, sErrFilename, startTime, true, out result, out resultSpc);
                                worker.ReportProgress(80, fileName);
                            }
                            else
                            {
                                //Update error result table for imported data entry
                                ServiceProxy.UpdateResultTable(FileName.Replace("csv", "") + FileName.Replace("csv", ""), Path.GetExtension(fileName), TotalRowCount, true, passedRows.Tables[0].Rows.Count,
                                                              true, TotalErrCount, true, fileName, string.Empty, string.Empty, startTime, true, out result, out resultSpc);
                                worker.ReportProgress(90, fileName);
                            }
                        }
                        ServiceProxy.SaveImportData(passedRows, out result, out resultSpc);
                        worker.ReportProgress(100, fileName);

                    }
                }            
            return result;            
        }

        /// <summary>
        /// Extract the data from csv and xlxs file
        /// and business rule validation for extracted data,If valid the data to stagging to the database.
        /// Invalid data's to write error file against the uploaded file's. 
        /// </summary>
        /// <param name="fileName"></param>
        private bool ExtractXLFiles(string fileName, DateTime startTime)
        {
            Excel.Application xlApp=null;
            Excel.Workbook xlWorkBook=null;
            Excel.Worksheet xlWorkSheet=null;
            Excel.Range range;
            int rCnt = 0;
            int cCnt = 0;
            int TotalRowCount = 0;
            int TotalErrCount = 0;



            string concatFileName = string.Empty;
            string errFilename;
            string sErrFilename = "";
            uint processId = 0;
            try
            {

                s_Filename = fileName;
                errFilename = Path.GetFileNameWithoutExtension(fileName) + "_Err";

                if (Path.GetExtension(fileName) == ".xls")
                {
                    sErrFilename = Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xls";
                }
                else if (Path.GetExtension(fileName) == ".xlsx")
                {
                    sErrFilename = Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xlsx";
                }
                if (File.Exists(sErrFilename))
                {
                    File.Delete(sErrFilename);
                }
                xlApp = new Excel.Application();

                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                GetWindowThreadProcessId(new IntPtr(xlApp.Hwnd), out processId);    
                worker.ReportProgress(10,fileName);
                for (int iterateworkSheet = 1; iterateworkSheet <= xlWorkBook.Worksheets.Count; iterateworkSheet++)
                {

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(iterateworkSheet);
                    concatFileName = string.Concat(xlWorkBook.Name.Replace("xlsx", ""), xlWorkSheet.Name);
                    range = xlWorkSheet.UsedRange;
                    using (DataSet extractedDataCol = new DataSet())
                    {
                        DataTable dtCreate = new DataTable();
                        DataColumn dc;
                        List<string> valueList = new List<string>();
                        for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                        {
                            for (int iterateHeaderColumns = 1; iterateHeaderColumns <= range.Columns.Count; iterateHeaderColumns++)
                            {
                                if (rCnt.Equals(1))
                                {
                                    if (!string.IsNullOrWhiteSpace(Convert.ToString((range.Cells[rCnt, iterateHeaderColumns] as Excel.Range).Value2)))
                                    {
                                        dc = new DataColumn(Convert.ToString((range.Cells[rCnt, iterateHeaderColumns] as Excel.Range).Value2));
                                        dtCreate.Columns.Add(dc);
                                    }

                                }
                            }

                                valueList = new List<string>();
                                for (cCnt = 1; cCnt <= dtCreate.Columns.Count; cCnt++)
                                {
                                    if (!rCnt.Equals(1))
                                    {

                                        if ((range.Cells[rCnt, cCnt] as Excel.Range).Value2 != null)
                                        {
                                            valueList.Add(Regex.Replace(Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2), "['\"]", ""));
                                        }
                                        else
                                        {
                                            valueList.Add(Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2));
                                        }
                                    }
                                }
                            //worker.ReportProgress(20,fileName);
                            if (!rCnt.Equals(1))
                            {
                                DataRow dr = dtCreate.NewRow();
                                for (int iterateValuList = 0; iterateValuList <= (valueList.Count - 1); iterateValuList++)
                                {
                                    if (valueList[iterateValuList] != null)
                                    {
                                        dr[iterateValuList] = Regex.Replace(valueList[iterateValuList].ToString().Trim(), "['\"]", ""); 
                                    }
                                    else
                                    {
                                        dr[iterateValuList] = valueList[iterateValuList];
                                    }
                                }
                                dtCreate.Rows.Add(dr);
                            }

                        }
                        extractedDataCol.Tables.Add(dtCreate);
                        worker.ReportProgress(30,fileName);
                        //Sanity Check...
                        if (extractedDataCol.Tables[0].Rows.Count > 0)
                        {
                            TotalRowCount = extractedDataCol.Tables[0].Rows.Count;
                            DataSet passedRows = new DataSet();

                            using (DataSet errorRowsDs = ServiceProxy.ApplyBRuleForImportedData(extractedDataCol, out passedRows))
                            {
                                worker.ReportProgress(40,fileName);
                                if (errorRowsDs.Tables[0].Rows.Count > 0)
                                {
                                    TotalErrCount = errorRowsDs.Tables[0].Rows.Count;
                                    //If BussinessRule Field rows to write error file 
                                    WriteErrorFile(errorRowsDs, xlWorkBook.Name.Replace("xlsx", ""), xlWorkSheet.Name, fileName);
                                    worker.ReportProgress(60,fileName);
                                    //Update error result table for imported faild data entry
                                    ServiceProxy.UpdateResultTable(xlWorkBook.Name.Replace("xlsx", "") + xlWorkSheet.Name, Path.GetExtension(fileName), TotalRowCount, true, passedRows.Tables[0].Rows.Count,
                                                                  true, TotalErrCount, true,fileName, errFilename, sErrFilename, startTime, true, out result, out resultSpc);
                                    worker.ReportProgress(80,fileName);
                                }
                                else
                                {
                                    //Update error result table for imported data entry
                                    ServiceProxy.UpdateResultTable(xlWorkBook.Name.Replace("xlsx", "") + xlWorkSheet.Name, Path.GetExtension(fileName), TotalRowCount, true, passedRows.Tables[0].Rows.Count,
                                                                  true, TotalErrCount, true,fileName,string.Empty,string.Empty, startTime, true, out result, out resultSpc);
                                    worker.ReportProgress(80,fileName);
                                }
                            }
                            ServiceProxy.SaveImportData(passedRows, out result, out resultSpc);
                            worker.ReportProgress(100,fileName);

                        }

                    }

                }
                
                return result;

            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                xlWorkBook.Close(0);
                if (processId != 0)
                {
                    Process excelProcess = Process.GetProcessById((int)processId);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }               
            }
        }

        /// <summary>
        ///Error filecreation for xlsx and xls file types. 
        /// </summary>
        /// <param name="dsErrorvalues"> error data's</param>
        /// <param name="workbookname"> error file name</param>
        /// <param name="worksheetname">sheet name</param>
        private void WriteErrorFile(DataSet dsErrorvalues, string workbookname, string worksheetname, string fileName)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                return;
            }

            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Worksheet newWorksheet = null;
            object misValue = System.Reflection.Missing.Value;
            string data = null;
            string errFilename;
            Boolean bWorksheet = false;
            int i = 0;
            int j = 0;
            Boolean bWorkbookopen = false;
            uint processId = 0;
            try
            {
                errFilename = Path.GetFileNameWithoutExtension(fileName) + "_Err";
                if (Path.GetExtension(fileName) == ".xls")
                {
                    string sErrFilename = Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xls";
                }
                else if (Path.GetExtension(fileName) == ".xlsx")
                {
                    string sErrFilename = Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xlsx";
                }



                if (File.Exists(Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xls") || File.Exists(Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xlsx"))
                {
                    if (Path.GetExtension(fileName) == ".xls" && File.Exists(Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xls"))
                    {
                        xlWorkBook = xlApp.Workbooks.Open(Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xls", 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, true, 0, true, 1, 0);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        bWorkbookopen = true;
                    }
                    else if (Path.GetExtension(fileName) == ".xlsx" && File.Exists(Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xlsx"))
                    {
                        xlWorkBook = xlApp.Workbooks.Open(Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xlsx", 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, true, 0, true, 1, 0);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        bWorkbookopen = true;
                    }
                    if (bWorkbookopen)
                    {
                        GetWindowThreadProcessId(new IntPtr(xlApp.Hwnd), out processId);
                        foreach (Excel.Worksheet sheet in xlWorkBook.Sheets)
                        {
                            // Check the name of the current sheet
                            if (sheet.Name == worksheetname)
                            {
                                bWorksheet = true;
                                break;
                            }
                        }

                        if (!bWorksheet)
                        {
                            newWorksheet = (Excel.Worksheet)xlApp.Worksheets.Add();
                            newWorksheet.Name = worksheetname;
                            xlWorkSheet = newWorksheet;
                        }
                        else
                        {
                            xlWorkSheet.Delete();
                            newWorksheet = (Excel.Worksheet)xlApp.Worksheets.Add();
                            xlWorkSheet = newWorksheet;
                            xlWorkSheet.Name = worksheetname;
                        }
                    }
                }
                else
                {

                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Name = worksheetname;
                }

                for (int Columnvalue = 0; Columnvalue <= dsErrorvalues.Tables[0].Columns.Count - 1; Columnvalue++)
                {
                    if (!dsErrorvalues.Tables[0].Columns[Columnvalue].ToString().ToLower().Contains("column"))
                    {
                        data = dsErrorvalues.Tables[0].Columns[Columnvalue].ToString();
                        xlWorkSheet.Cells[1, Columnvalue + 1] = data;
                    }
                }

                for (int tableCount = 0; tableCount <= dsErrorvalues.Tables.Count - 1; tableCount++)
                {
                    for (i = 0; i <= dsErrorvalues.Tables[tableCount].Rows.Count - 1; i++)
                    {
                        for (j = 0; j <= dsErrorvalues.Tables[tableCount].Columns.Count - 1; j++)
                        {
                            if (!dsErrorvalues.Tables[tableCount].Rows[i].ItemArray[j].ToString().ToLower().Contains("column"))
                            {
                                data = dsErrorvalues.Tables[tableCount].Rows[i].ItemArray[j].ToString();
                                xlWorkSheet.Cells[i + 2, j + 1] = data;
                            }
                        }
                    }
                }
                xlApp.DisplayAlerts = false;
                if (Path.GetExtension(fileName) == ".xls")
                {
                    xlWorkBook.SaveAs(Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xls", misValue);
                }
                else if (Path.GetExtension(fileName) == ".xlsx")
                {
                    xlWorkBook.SaveAs(Path.GetDirectoryName(fileName) + "\\" + errFilename + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
                    xlWorkBook.Close(0);
                }
                xlApp.Quit();

            }
            catch (Exception ex)
            {

            }
            finally
            {
                if (processId != 0)
                {
                    Process excelProcess = Process.GetProcessById((int)processId);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }
            }

        }
        /// <summary>
        ///Error filecreation for xlsx and xls file types. 
        /// </summary>
        /// <param name="dsErrorvalues"> error data's</param>
        /// <param name="workbookname"> error file name</param>
        /// <param name="worksheetname">sheet name</param>
        private void WriteCSVErrorFile(DataSet dsErrorvalues, string workbookname, string worksheetname, string errorFilename)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                return;
            }


            Excel.Workbook xlWorkBook=null;
            Excel.Worksheet xlWorkSheet=null;
            Excel.Worksheet newWorksheet = null;
            object misValue = System.Reflection.Missing.Value;
            string data = null;
            Boolean bWorksheet = false;
            int i = 0;
            int j = 0;
            uint processId = 0;
            try
            {
                //Sanity check: The Excel add ins allow max size of file name is 31.. The exist file name size has triming.
                if (worksheetname.Length > 31)
                {
                    worksheetname = worksheetname.Remove(31, worksheetname.Length - 31);
                }

                if (File.Exists(errorFilename))
                {
                    xlWorkBook = xlApp.Workbooks.Open(errorFilename, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, true, 0, true, 1, 0);
                    GetWindowThreadProcessId(new IntPtr(xlApp.Hwnd), out processId); 
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    foreach (Excel.Worksheet sheet in xlWorkBook.Sheets)
                    {
                        // Check the name of the current sheet
                        if (sheet.Name == worksheetname)
                        {
                            bWorksheet = true;
                            break;
                        }
                    }

                    if (!bWorksheet)
                    {
                        newWorksheet = (Excel.Worksheet)xlApp.Worksheets.Add();
                        newWorksheet.Name = worksheetname;
                        xlWorkSheet = newWorksheet;
                    }
                    else
                    {
                        xlWorkSheet.Delete();
                        newWorksheet = (Excel.Worksheet)xlApp.Worksheets.Add();
                        xlWorkSheet = newWorksheet;
                        xlWorkSheet.Name = worksheetname;
                    }
                }
                else
                {

                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Name = worksheetname;
                }

                for (int Columnvalue = 0; Columnvalue <= dsErrorvalues.Tables[0].Columns.Count - 1; Columnvalue++)
                {
                    data = dsErrorvalues.Tables[0].Columns[Columnvalue].ToString();
                    xlWorkSheet.Cells[1, Columnvalue + 1] = data;
                }

                for (int tableCount = 0; tableCount <= dsErrorvalues.Tables.Count - 1; tableCount++)
                {
                    for (i = 0; i <= dsErrorvalues.Tables[tableCount].Rows.Count - 1; i++)
                    {
                        for (j = 0; j <= dsErrorvalues.Tables[tableCount].Columns.Count - 1; j++)
                        {
                            data = dsErrorvalues.Tables[tableCount].Rows[i].ItemArray[j].ToString();
                            xlWorkSheet.Cells[i + 2, j + 1] = data;
                        }
                    }
                }
                xlApp.DisplayAlerts = false;
                xlWorkBook.SaveAs(errorFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, false, Type.Missing, Type.Missing, Type.Missing);
                xlWorkBook.Close(0);
                xlApp.Quit();
            }
            catch (Exception ex)
            {

            }
            finally
            {
                if (processId != 0)
                {
                    Process excelProcess = Process.GetProcessById((int)processId);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }
            }

        }

        /// <summary>
        ///File scraping in txt file's 
        /// </summary>
        /// <param name="fileName"> file name</param>
        /// <returns> result </returns>
        private bool ExtractTxtFiles(string fileName, DateTime startTime)
        {
            int TotalRowCount = 0;
            int TotalErrCount = 0;
            string errFilename;
            string sErrFilename;
            string format = "|";
            DataColumn extractDc;
            DataTable extractDataTable = new DataTable();
            try
            {
                errFilename = Path.GetFileNameWithoutExtension(fileName) + "_Err";
                sErrFilename = Path.GetDirectoryName(fileName) + "\\" + errFilename + ".txt";
                if (File.Exists(sErrFilename))
                {
                    File.Delete(sErrFilename);
                }


                CsvTxtParser csvTxtParser = new CsvTxtParser();
                string[][] results;
                string[] tempvalueCollection = new string[0];

                using (var stream = File.OpenRead(fileName))
                using (var reader = new StreamReader(stream))
                {
                    results = csvTxtParser.Parse(reader);
                }

                for (var iterateParsingList = 0; iterateParsingList < results.Length; iterateParsingList++)
                {
                    string[] childList = results[iterateParsingList];

                    if (childList[0].ToString().Contains("|"))
                    {
                        tempvalueCollection = childList[0].ToString().ToLower().Split('|');
                        format = "|";
                    }
                    else if (childList[0].ToString().Contains(";"))
                    {
                        tempvalueCollection = childList[0].ToString().ToLower().Split(';');
                        format = ";";
                    }
                    else if (childList[0].ToString().Contains("\t"))
                    {
                        tempvalueCollection = childList[0].ToString().ToLower().Split('\t');
                        format = "\t";
                    }
                    else if (childList[0].ToString().Contains(" "))
                    {
                        tempvalueCollection = childList[0].ToString().ToLower().Split(' ');
                        format = " ";
                    }
                    else
                    {
                        tempvalueCollection = childList;

                        if (tempvalueCollection.Length > 0)
                        {
                            if (!iterateParsingList.Equals(0))
                            {
                                string[] cloneCopy = tempvalueCollection;
                                string[] tempList = new string[0];
                                tempvalueCollection = new string[7];
                                for (int iterateAppendval = 0; iterateAppendval < cloneCopy.Length; iterateAppendval++)
                                {
                                    if (cloneCopy[iterateAppendval].Contains(","))
                                    {
                                        tempvalueCollection = cloneCopy[iterateAppendval].Split(',');
                                    }
                                    else
                                    {
                                        tempvalueCollection[iterateAppendval] = cloneCopy[iterateAppendval].ToString();
                                    }
                                }
                            }
                        }
                    }

                    if (iterateParsingList == 0)
                    {
                        for (int iterateValueList = 0; iterateValueList <= (tempvalueCollection.Length - 1); iterateValueList++)
                        {
                            extractDc = new DataColumn(Regex.Replace(tempvalueCollection[iterateValueList].ToString().Trim(), "['\"]", ""));
                            extractDataTable.Columns.Add(extractDc);
                        }
                    }
                    else
                    {
                        DataRow dr = extractDataTable.NewRow();
                        for (int iterateValueList = 0; iterateValueList <= (tempvalueCollection.Length - 1); iterateValueList++)
                        {
                            if (tempvalueCollection[iterateValueList] != null)
                            {
                                dr[iterateValueList] = Regex.Replace(tempvalueCollection[iterateValueList].ToString().Trim(), "['\"]", "");
                            }
                            else
                            {
                                dr[iterateValueList] = tempvalueCollection[iterateValueList];
                            }
                        }
                        extractDataTable.Rows.Add(dr);
                        worker.ReportProgress(30, fileName);
                    }
                }
                worker.ReportProgress(10, fileName);
                using (DataSet dataCollection = new DataSet())
                {                   
                    dataCollection.Tables.Add(extractDataTable);
                    worker.ReportProgress(30, fileName);

                    //Sanity Check...
                    if (dataCollection.Tables[0].Rows.Count > 0)
                    {
                        TotalRowCount = dataCollection.Tables[0].Rows.Count;
                        DataSet passedRows = new DataSet();

                        using (DataSet errorRowsDs = ServiceProxy.ApplyBRuleForImportedData(dataCollection, out passedRows))
                        {
                            worker.ReportProgress(40, fileName);
                            if (errorRowsDs.Tables[0].Rows.Count > 0)
                            {
                                TotalErrCount = errorRowsDs.Tables[0].Rows.Count;
                                //If BussinessRule Field rows to write error file                                 
                                WriteErrortxtFile(errorRowsDs.Tables[0], sErrFilename, format);
                                worker.ReportProgress(60, fileName);
                                //Update error result table for imported faild data entry
                                ServiceProxy.UpdateResultTable(Path.GetFileName(fileName), Path.GetExtension(fileName), TotalRowCount, true, passedRows.Tables[0].Rows.Count,
                                                              true, TotalErrCount, true,fileName, errFilename, sErrFilename, startTime, true, out result, out resultSpc);
                                worker.ReportProgress(80, fileName);
                            }
                            else
                            {
                                //Update error result table for imported data entry
                                ServiceProxy.UpdateResultTable(Path.GetFileName(fileName), Path.GetExtension(fileName), TotalRowCount, true, passedRows.Tables[0].Rows.Count,
                                                              true, TotalErrCount, true, fileName, string.Empty, string.Empty, startTime, true, out result, out resultSpc);
                                worker.ReportProgress(80, fileName);
                            }
                        }
                        ServiceProxy.SaveImportData(passedRows, out result, out resultSpc);
                        worker.ReportProgress(100, fileName);

                    }
                }
                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        ///Write error txt file while process failed in import data. 
        /// </summary>
        /// <param name="datatable">failed rows</param>
        /// <param name="file">error file name</param>
        private void WriteErrortxtFile(DataTable datatable, string file, string format)
        {
            StreamWriter str = new StreamWriter(file, false, System.Text.Encoding.Default);

            string Columns = string.Empty;
            foreach (DataColumn column in datatable.Columns)
            {
                if (format != "space")
                {
                    Columns += column.ColumnName + format;
                }
                else if (format == "space")
                {
                    Columns += '"' + column.ColumnName + '"' + ' ';
                }
            }
            if (format != "space")
            {
                str.WriteLine(Columns.Remove(Columns.Length - 1, 1));
            }
            else if (format == "space")
            {
                str.WriteLine(Columns);
            }


            foreach (DataRow datarow in datatable.Rows)
            {
                string row = string.Empty;

                foreach (object items in datarow.ItemArray)
                {
                    if (format != "space")
                    {
                        row += items.ToString() + format;
                    }
                    else if (format == "space")
                    {
                        row += '"' + items.ToString() + '"' + ' ';
                    }

                }

                if (format != "space")
                {
                    str.WriteLine(row.Remove(row.Length - 1, 1));
                }
                else if (format == "space")
                {
                    str.WriteLine(row);
                }



            }
            str.Flush();
            str.Close();

        }

        /// <summary>
        ///Relaese the memory of object 
        /// </summary>
        /// <param name="obj"></param>
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool isDisposing)
        {
            if (!m_isDisposed)
            {
                if (isDisposing)
                {
                }
                m_isDisposed = true;
            }
        }

        ~AplBusinessLayer()
        {
            Dispose(false);
        }

        #endregion

        #endregion
    }
}
