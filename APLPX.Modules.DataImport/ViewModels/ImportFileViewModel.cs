using GongSolutions.Wpf.DragDrop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using ReactiveUI;
using System.Collections.ObjectModel;
using APLPX.Modules.DataImport.Models;
using APLPX.Common;
using APLPX.Client;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System.Windows.Input;
using System.Threading;
using System.Collections;
using System.Runtime.Remoting.Messaging;
using System.Data;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text.RegularExpressions;
using APLPX.Common.Helper;



namespace APLPX.Modules.DataImport.ViewModels
{
    public class ImportFileViewModel : ReactiveObject, IDropTarget
    {

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        /// <summary>
        /// Persist Imported files observable collection instance 
        /// </summary>
        private ObservableCollection<ImportFileModel> m_importedFileLists;

        /// <summary>
        /// Delegate method for progress bar action 
        /// </summary>
        delegate void PostMethodDelegate(int currentRow);

        /// <summary>
        /// Object declaration for APLBusiness layer 
        /// </summary>
        private AplBusinessLayer aplBusinessL;

        /// <summary>
        /// Property for Result visibility changes 
        /// </summary>
        private bool m_isResultFeatureModuleVisible;

        /// <summary>
        ///Property using Drag over menu. 
        /// </summary>
        private bool m_isDragOverFlag;

        /// <summary>
        ///TODO 
        /// </summary>
        private ObservableCollection<ResultModel> m_resultset;

        /// <summary>
        ///TODO 
        /// </summary>
        private int m_maxScreenwidth;


        #region Constructor
        /// <summary>
        /// Default constructor initialization 
        /// </summary>
        public ImportFileViewModel()
        {
            InitializeCommand();
            m_importedFileLists = new ObservableCollection<ImportFileModel>();
            ImportedFileCollection = new ObservableCollection<ImportFileModel>();

            m_resultset = new ObservableCollection<ResultModel>();
            aplBusinessL = new AplBusinessLayer();
            //Check all default value set to true
            IsSelectAll = true;
                       

        }
        #endregion

        #region Properties
        /// <summary>
        ///Persist the  (drag n drop) and import files.
        /// </summary>
        public ObservableCollection<ImportFileModel> ImportedFileCollection
        {
            get { return m_importedFileLists; }
            set
            {
                m_importedFileLists = value;
                this.RaiseAndSetIfChanged(ref m_importedFileLists, value);
            }
        }


        private bool _isSelectAll;
        public bool IsSelectAll
        {
            get { return _isSelectAll; }
            set
            {
                _isSelectAll = value;
                if (_isSelectAll == true)
                {
                    SelectAll();
                }
                else
                {
                    DeselectAll();
                }
                this.RaisePropertyChanged("IsSelectAll");
            }
        }

        /// <summary>
        /// Property  to visible planning feature list 
        /// </summary>
        public bool IsDragOverFlag
        {
            get
            {
                return m_isDragOverFlag;
            }
            private set
            {
                this.RaiseAndSetIfChanged(ref m_isDragOverFlag, value);
            }
        }

        /// <summary>
        /// Property for showing result view
        /// </summary>
        public bool IsResultFeatureModuleVisible
        {
            get
            {
                return m_isResultFeatureModuleVisible;
            }
            set
            {
                m_isResultFeatureModuleVisible = value;
                if (IsResultFeatureModuleVisible)
                {
                    ResultDataCommandExecuted();
                }


            }
        }
        /// <summary>
        ///property for getting import file result set 
        /// </summary>
        public ObservableCollection<ResultModel> ImportFileResult
        {
            get { return m_resultset; }
            set
            {
                m_resultset = value;
            }
        }

        private ResultModel _selectedRow { get; set; }
        public ResultModel SelectedRow
        {
            get { return _selectedRow; }

            set
            {
                _selectedRow = value;                               
            }
        }

        private Thickness m_importDatamargin;
        public Thickness ImportDataMargin 
        {
            get
            {
                return m_importDatamargin;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_importDatamargin,value);
            }
        }

        private Thickness m_resultSetmargin;
        public Thickness ResultSetMargin
        {
            get
            {
                return m_resultSetmargin;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_resultSetmargin, value);
            }
        }

        private bool m_enableFullScreenBtn;
        public bool EnableFullScreenBtn
        {
            get 
            {
                return m_enableFullScreenBtn;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_enableFullScreenBtn, value);
            }
        }

        private bool m_enableExistFullScreenBtn;
        public bool EnableExistFullScreenBtn
        {
            get
            {
                return m_enableExistFullScreenBtn;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_enableExistFullScreenBtn, value);
            }
        }

        /// <summary>
        ///TODO : Max screen size 
        /// </summary>
        public int MaxScreenWidth
        {
            get
            {
                return m_maxScreenwidth;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_maxScreenwidth, value);
            }
        }

        #endregion

        #region Command

        /// <summary>
        ///Command property for browsefile 
        /// </summary>
        public ReactiveCommand<object> BrowseFileCmd { get; private set; }
        ///// <summary>
        /////COmmand property for Validatefile 
        ///// </summary>
        //public ReactiveCommand<object> ValidateFileCmd { get; private set; }
        /// <summary>
        ///COmmand property for ImportData 
        /// </summary>
        public ReactiveCommand<object> ImportDataCmd { get; private set; }    
        /// <summary>
        /// Command to display the Admin module features.
        /// </summary>
        public ReactiveCommand<object> RemoveUploadFileCmd { get; private set; }

        /// <summary>
        ///Refresh result set command 
        /// </summary>
        public ReactiveCommand<object> RefreshResultsetCmd { get; private set; }
        /// <summary>
        /// Command to display the Rowselection value.
        /// </summary>
        public ReactiveCommand<object> RowClickCmd { get; private set; }

        /// <summary>
        ///TODO 
        /// </summary>
        public ReactiveCommand<object> FullScreenCmd { get; private set; }

        /// <summary>
        ///TODO 
        /// </summary>
        public ReactiveCommand<object> ExistFullScreenCmd { get; private set; }

        #endregion

        #region Private Methods

        /// <summary>
        ///Check all uploaded files 
        /// </summary>
        private void SelectAll()
        {
            foreach(var iterateModel in ImportedFileCollection)
            {
                iterateModel.IsSelected = true;
            }
        }

        /// <summary>
        ///Un check all Uploaded files
        /// </summary>
        private void DeselectAll()
        {
            foreach (var iterateModel in ImportedFileCollection)
            {
                iterateModel.IsSelected = false;
            }
        }
        /// <summary>
        ///Command Initialize 
        /// </summary>
        private void InitializeCommand()
        {
            //Initialize Admin module fetaures visibility command.
            BrowseFileCmd = ReactiveCommand.Create();
            BrowseFileCmd.Subscribe(x => BrowseFileCommandExecuted(x));
           
            ImportDataCmd = ReactiveCommand.Create();
            ImportDataCmd.Subscribe(x => ImportDataCommandExecuted(x));

            RemoveUploadFileCmd = ReactiveCommand.Create();
            RemoveUploadFileCmd.Subscribe(x => RemoveUploadFileCommandExecuted(x));

            RefreshResultsetCmd = ReactiveCommand.Create();
            RefreshResultsetCmd.Subscribe(x => ResultDataCommandExecuted());

            RowClickCmd = ReactiveCommand.Create();
            RowClickCmd.Subscribe(x => RowFileOpenCommandExecuted(x));

            FullScreenCmd = ReactiveCommand.Create();
            FullScreenCmd.Subscribe(x => FullScreenCommandExecuted());

            ExistFullScreenCmd = ReactiveCommand.Create();
            ExistFullScreenCmd.Subscribe(x => ExistFullScreenCommandExecuted());
        }

        /// <summary>
        ///TODO: Max screen size 
        /// </summary>
        private void FullScreenCommandExecuted()
        {
            //double test = System.Windows.SystemParameters.;
            double test =System.Windows.SystemParameters.PrimaryScreenWidth;
            // if (obj != System.Windows.WindowState.Minimized) 
           // {
                MaxScreenWidth = Convert.ToInt16(test -150);
                ImportDataMargin = new Thickness(1140, 10, 10, 637);
                ResultSetMargin = new Thickness(1100, 0, 0, -4);
                EnableFullScreenBtn = false;
                EnableExistFullScreenBtn = true;
           // }
        }

        /// <summary>
        ///TODO 
        /// </summary>
        private void ExistFullScreenCommandExecuted()
        {
            MaxScreenWidth = 900;
            ImportDataMargin = new Thickness(841, 10, 10, 637);
            ResultSetMargin = new Thickness(800, 0, 0, -5);
            EnableFullScreenBtn = true;
            EnableExistFullScreenBtn = false;
        }

        /// <summary>
        ///Remove uploaded file from  import file collection
        /// </summary>
        /// <param name="sender">Currently selected file</param>
        private void RemoveUploadFileCommandExecuted(object sender)
        {
            List<ImportFileModel> tempCollection = ImportedFileCollection.ToList();

            foreach(var importModel in tempCollection)
            {
                if(importModel.IsSelected)
                {
                    ImportedFileCollection.Remove(importModel);
                }
            }
        }
        /// <summary>
        /// Binding data to the result grid
        /// </summary>
        private void ResultDataCommandExecuted()
        {

            try
            {
                using (DataSet resultSetDs = aplBusinessL.GetResultSet())
                {
                    //Sanity Check
                    if (resultSetDs.Tables[0].Rows.Count > 0)
                    {
                        ImportFileResult.Clear();
                        foreach (DataRow dr in resultSetDs.Tables[0].Rows)
                        {                           
                                ImportFileResult.Add(new ResultModel
                                {
                                    FileName = dr["FileName"].ToString(),
                                    Format = dr["Format"].ToString(),
                                    RowsInFile = dr["RowsInFile"].ToString(),
                                    RowsImported = dr["RowsImported"].ToString(),
                                    RowsWithError = dr["RowsWithError"].ToString(),
                                    ErrorFilePath = dr["ErrorFilePath"].ToString(),
                                    ErrorFileName = dr["ErrorFileName"].ToString(),
                                    StartTime = dr["StartTime"].ToString(),
                                    EndTime = dr["EndTime"].ToString(),
                                    Duration = dr["Duration"].ToString(),
                                    SoucreFilePath = dr["SourceFilePath"].ToString(),
                                });
                        }
                    }
                    else
                    {
                        ImportFileResult.Clear();
                    }
                }
            }
            catch (Exception ex)
            {                
                EventAgg._eventAggregator.GetEvent<ErrorMessageEvent>().Publish(ex.Message);
            }
        }
        /// <summary>
        ///Remove uploaded file from  import file collection
        /// </summary>
        /// <param name="sender">Currently selected file</param>
        private void RowFileOpenCommandExecuted(object sender)
        {
            int index = ImportFileResult.IndexOf(SelectedRow as ResultModel);
            if (index > -1)
            {
                if(SelectedRow.ErrorFilePath!=null && SelectedRow.ErrorFilePath!="")
                {
                    if(File.Exists(SelectedRow.ErrorFilePath))
                    {
                        System.Diagnostics.Process.Start(SelectedRow.ErrorFilePath);
                        EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish("");
                    }
                    else
                    {
                        EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish("The error file does not exists.");
                    }
                }
                else
                {
                    if(SelectedRow.RowsWithError.Equals("0") && SelectedRow.ErrorFilePath.Equals(""))
                    {
                        EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish("There is no error file.");
                    }
                }
             
            }
        }
        /// <summary>
        ///Browse file button binding commande execution 
        /// </summary>
        /// <param name="sender"></param>
        public void BrowseFileCommandExecuted(object sender)
        {            
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Images (*.xls;*.xlsx;*.txt;*.csv)|*.xls;*.xlsx;*.txt;*.csv";
            Nullable<bool> result = dlg.ShowDialog();
            var extension = Path.GetExtension(dlg.FileName);
            try
            {
                if (result == true)
                {  
                    //Validate the selected file format and valid content.If invalid file cannot upload for import data.
                    string validateMsg = ValidateUploadedFiles(dlg.FileName);
                    if (validateMsg.Equals("Processed"))
                    {
                        EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish("");
                        string filename = dlg.SafeFileName;

                        if (extension != null && extension.Equals(".xlsx") || extension.Equals(".xls") || extension.Equals(".txt") || extension.Equals(".csv"))
                        {
                            FileInfo fileInfo = new FileInfo(dlg.FileName);
                            ImportedFileCollection.Add(new ImportFileModel { Name = Path.GetFileName(dlg.FileName), Path = dlg.FileName, Size = fileInfo.Length, Type = extension, DateModified = fileInfo.CreationTime.ToString(), IsSelected = true });
                        }
                        else
                        {
                            EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish(string.Format("{0} is an unsupported file type. Supported file types are (.xlsx, .xls, .csv, .txt).", extension));
                        }
                    }
                    else
                    {
                        EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish(string.Format(validateMsg));
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }           
        }

        private string ValidateUploadedFiles(string fileName)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            int rCnt = 0;
            int cCnt = 0;
            var extension = Path.GetExtension(fileName).Replace(".","");
            string result = "Processed";
            bool notValid = false;
            int notValidcount = 0;
            string message = string.Empty;
            uint processId = 0;
            try
            {
                if (extension != null && extension.Equals(Utility.XLSX) || extension.Equals(Utility.XLS))
                {
                    xlApp = new Excel.Application();
                    try
                    {
                        xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                        GetWindowThreadProcessId(new IntPtr(xlApp.Hwnd), out processId);
                        for (int iterateworkSheet = 1; iterateworkSheet <= xlWorkBook.Worksheets.Count; iterateworkSheet++)
                        {

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(iterateworkSheet);
                            range = xlWorkSheet.UsedRange;
                            using (DataSet extractedDataCol = new DataSet())
                            {                                
                                List<string> valueList = new List<string>();
                                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                                {
                                    if (rCnt.Equals(1))
                                    {
                                        for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                                        {
                                            if (cCnt <= 7)
                                            {
                                                if (rCnt.Equals(1))
                                                {
                                                    if ((range.Cells[rCnt, cCnt] as Excel.Range).Value2 != null)
                                                    {
                                                        if (!notValid)
                                                        {
                                                            valueList.Add(Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2).ToLower());
                                                        }
                                                        else if (notValidcount < cCnt)
                                                        {
                                                            return "Invalid File";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        notValid = true;
                                                        notValidcount = cCnt;
                                                    }
                                                }
                                            }

                                        }

                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                              
                                if(!valueList.Exists(element => element == "sku"))
                                {
                                    message = "Sku";
                                }
                                if (!valueList.Exists(element => element == "name"))
                                {
                                    message += " Name";
                                }
                                if (!valueList.Exists(element => element == "company"))
                                {
                                    message += " Company";
                                }
                                if (!valueList.Exists(element => element == "price"))
                                {
                                    message += " Price";
                                }
                                if (!valueList.Exists(element => element == "crawldate") && !valueList.Exists(element => element == "crawl_date"))
                                {
                                    message += " CrawlDate";
                                }                              

                                if(!string.IsNullOrEmpty(message))
                                {
                                    message += " columns are missing in source file.";

                                    return message;
                                }

                            }

                        }
                    }
                    catch (Exception)
                    {
                        return "File is corrupted";
                    }
                }
                else if (extension.Equals(Utility.CSV))
                {
                    CsvTxtParser csvTxtParser = new CsvTxtParser();
                    string[][] results;
                    string[] tempvalueCollection = new string[0];
                    int fileHeaderCount = 0;
                    int dataRowsCount = 0;
                    using (var stream = File.OpenRead(fileName))
                    using (var reader = new StreamReader(stream))
                    {
                        results = csvTxtParser.Parse(reader);
                    }

                    for (var iterateParsingList = 0; iterateParsingList < results.Length; iterateParsingList++)
                    {

                        string[] childList = results[iterateParsingList];

                        if (childList[0].ToString().Contains(","))
                        {
                            tempvalueCollection = childList[0].ToString().ToLower().Split(',');
                        }
                        else if (childList[0].ToString().Contains("|"))
                        {
                            tempvalueCollection = childList[0].ToString().ToLower().Split('|');
                        }
                        else if (childList[0].ToString().Contains(";"))
                        {
                            tempvalueCollection = childList[0].ToString().ToLower().Split(';');
                        }
                        else if (childList[0].ToString().Contains("\t"))
                        {
                            tempvalueCollection = childList[0].ToString().ToLower().Split('\t');
                        }
                        else if (childList[0].ToString().Contains(" "))
                        {
                            tempvalueCollection = childList[0].ToString().ToLower().Split(' ');
                        }
                        else
                        { 
                             tempvalueCollection = childList;

                             if (tempvalueCollection.Length > 0)
                             {
                                 if (iterateParsingList.Equals(0))
                                 {
                                     fileHeaderCount = tempvalueCollection.Length;
                                     if (!tempvalueCollection.Any(element => element.ToLower().Contains("sku")))
                                     {
                                         message = "Sku";
                                     }
                                     if (!tempvalueCollection.Any(element => element.ToLower().Contains("name")))
                                     {
                                         message += " Name";
                                     }
                                     if (!tempvalueCollection.Any(element => element.ToLower().Contains("company")))
                                     {
                                         message += " Company";
                                     }
                                     if (!tempvalueCollection.Any(element => element.ToLower().Contains("price")))
                                     {
                                         message += " Price";
                                     }
                                     if (!tempvalueCollection.Any(element => element.ToLower().Contains("crawldate")) && !tempvalueCollection.Any(element => element.ToLower().Contains("crawl_date")))
                                     {
                                         message += " CrawlDate";
                                     }
                                     if (!string.IsNullOrEmpty(message))
                                     {
                                         message += " columns are missing in source file.";

                                         return message;
                                     }
                                 }
                                 else
                                 {
                                     int count = 0;
                                     string[] tempList = new string[0];
                                     foreach (var iteration in tempvalueCollection)
                                     {
                                         if (iteration.Contains(","))
                                         {
                                             tempList = iteration.Split(',');
                                         }
                                         else
                                         {
                                             count = count + 1;
                                         }
                                     }
                                     dataRowsCount = count + tempList.Length;
                                     break;
                                 }
                             }
                        }

                        if (iterateParsingList == 0)
                        {
                            fileHeaderCount = tempvalueCollection.Length;
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("sku")))
                            {
                                message = "Sku";
                            }
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("name")))
                            {
                                message += " Name";
                            }
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("company")))
                            {
                                message += " Company";
                            }
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("price")))
                            {
                                message += " Price";
                            }
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("crawldate")) && !tempvalueCollection.Any(element => element.ToLower().Contains("crawl_date")))
                            {
                                message += " CrawlDate";
                            }
                            if (!string.IsNullOrEmpty(message))
                            {
                                message += " columns are missing in source file.";

                                return message;
                            }
                        }
                        else
                        {
                            dataRowsCount = tempvalueCollection.Length;
                            break;
                        }                                                

                    }
                    if (fileHeaderCount < dataRowsCount)
                    {
                        string fName = Path.GetFileName(fileName);
                        message += string.Format("'{0}'  file have some mismatch values.", fName);
                    }
                    if (!string.IsNullOrEmpty(message))
                    {
                        return message;
                    }
                }
                else if (extension.Equals(Utility.TXT))
                {
                    try
                    {
                        string[] lines = File.ReadAllLines(fileName);
                    }
                    catch (Exception)
                    {
                        return "File is corrupted";
                    }
                    CsvTxtParser csvTxtParser = new CsvTxtParser();
                    string[][] results;
                    string[] tempvalueCollection = new string[0];
                    int fileHeaderCount = 0;
                    int dataRowsCount = 0;
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
                        }
                        else if (childList[0].ToString().Contains(";"))
                        {
                            tempvalueCollection = childList[0].ToString().ToLower().Split(';');
                        }
                        else if (childList[0].ToString().Contains("\t"))
                        {
                            tempvalueCollection = childList[0].ToString().ToLower().Split('\t');
                        }
                        else if (childList[0].ToString().Contains(" "))
                        {
                            tempvalueCollection = childList[0].ToString().ToLower().Split(' ');
                        }
                        else
                        {
                            tempvalueCollection = childList;

                            if(tempvalueCollection.Length >0)
                            {
                                if (iterateParsingList.Equals(0))
                                {
                                    fileHeaderCount = tempvalueCollection.Length;
                                    if (!tempvalueCollection.Any(element => element.ToLower().Contains("sku")))
                                    {
                                        message = "Sku";
                                    }
                                    if (!tempvalueCollection.Any(element => element.ToLower().Contains("name")))
                                    {
                                        message += " Name";
                                    }
                                    if (!tempvalueCollection.Any(element => element.ToLower().Contains("company")))
                                    {
                                        message += " Company";
                                    }
                                    if (!tempvalueCollection.Any(element => element.ToLower().Contains("price")))
                                    {
                                        message += " Price";
                                    }
                                    if (!tempvalueCollection.Any(element => element.ToLower().Contains("crawldate")) && !tempvalueCollection.Any(element => element.ToLower().Contains("crawl_date")))
                                    {
                                        message += " CrawlDate";
                                    }
                                    if (!string.IsNullOrEmpty(message))
                                    {
                                        message += " columns are missing in source file.";

                                        return message;
                                    }
                                }
                                else
                                {
                                    int count = 0;
                                    string[] tempList =new string[0];
                                    foreach (var iteration in tempvalueCollection)
                                    {
                                        if (iteration.Contains(","))
                                        {
                                           tempList = iteration.Split(',');
                                        }
                                        else
                                        {
                                            count = count + 1;
                                        }
                                    }
                                    dataRowsCount = count + tempList.Length;
                                    break;
                                }
                            }
                        }
                        if (iterateParsingList == 0)
                        {
                            fileHeaderCount = tempvalueCollection.Length;
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("sku")))
                            {
                                message = "Sku";
                            }
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("name")))
                            {
                                message += " Name";
                            }
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("company")))
                            {
                                message += " Company";
                            }
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("price")))
                            {
                                message += " Price";
                            }
                            if (!tempvalueCollection.Any(element => element.ToLower().Contains("crawldate")) && !tempvalueCollection.Any(element => element.ToLower().Contains("crawl_date")))
                            {
                                message += " CrawlDate";
                            }
                            if (!string.IsNullOrEmpty(message))
                            {
                                message += " columns are missing in source file.";

                                return message;
                            }
                        }
                        else
                        {
                            dataRowsCount = tempvalueCollection.Length;
                            break;
                        } 
                    }
                    if (fileHeaderCount < dataRowsCount)
                    {
                        string fName = Path.GetFileName(fileName);
                        message += string.Format("'{0}'  file have some mismatch values.", fName);
                    }
                    if (!string.IsNullOrEmpty(message))
                    {
                        return message;
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (extension.Equals(Utility.XLSX) || extension.Equals(Utility.XLS))
                {
                    xlWorkBook.Close(0);
                    if (processId != 0)
                    {
                        Process excelProcess = Process.GetProcessById((int)processId);
                        excelProcess.CloseMainWindow();
                        excelProcess.Refresh();
                        excelProcess.Kill();
                    }
                    //releaseObject(xlWorkBook);
                    //releaseObject(xlApp);   
                }
                            
            }
            return result;
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
            }
        }      

        /// <summary>
        ///Import data file peocessing in import data button command binding execution 
        /// </summary>
        /// <param name="sender"></param>
        private void ImportDataCommandExecuted(object sender)
        {
            PostMethodDelegate dlgt;
            int progressValue=0;
            //Sanity check
            try
            {
                if (ImportedFileCollection.Count > 0)
                {
                    if (ImportedFileCollection.Any(x => x.IsSelected == true))
                    {
                        for (int iteratePreProcessFile = 0; iteratePreProcessFile < ImportedFileCollection.Count; iteratePreProcessFile++)
                        {
                            if (ImportedFileCollection[iteratePreProcessFile].IsSelected)
                            {
                                ImportedFileCollection[iteratePreProcessFile].FileStatusVisible = false;
                                ImportedFileCollection[iteratePreProcessFile].FileProgressVisible = true;

                                dlgt = new PostMethodDelegate(StartProgress);
                                AsyncCallback asyncCallBack = new AsyncCallback(show);
                                IAsyncResult ar = dlgt.BeginInvoke(iteratePreProcessFile, asyncCallBack, dlgt);

                            }
                        }
                    }
                    else
                    {
                        EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish("Select the files to import");
                    }
                }
                else
                {
                    EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish("Please upload the file");
                }
            }
            catch (Exception ex)
            {
                EventAgg._eventAggregator.GetEvent<ErrorMessageEvent>().Publish(ex.Message);
            }
        }

        #region BackgroundWorker Events

        /// <summary>
        ///Async method to start the fileprocess 
        /// </summary>
        /// <param name="asyncResult"></param>
        void show(IAsyncResult asyncResult)
        {
      
            PostMethodDelegate dlgt = (PostMethodDelegate)asyncResult.AsyncState;
            dlgt.EndInvoke(asyncResult);
        }

        /// <summary> 
        ///File processing status updated into processed file status. 
        /// </summary>
        /// <param name="rowNum"> current processed file row number</param>
       void StartProgress(int rowNum)
        {
            BackgroundWorker worker = new BackgroundWorker();
           
            worker.WorkerReportsProgress = true;
            aplBusinessL.ImportData(ImportedFileCollection[rowNum].Path,out worker);                    

            worker.ProgressChanged += new ProgressChangedEventHandler(this.ReportProgress);
           
        }
       public void ReportProgress(object sender, ProgressChangedEventArgs e)
       {
           for (int i = 0; i < ImportedFileCollection.Count;i++ )
           {
               if(e.UserState.ToString()== ImportedFileCollection[i].Path.ToString())
               {
                   if (e.ProgressPercentage != 100)
                   {
                       ImportedFileCollection[i].Progress = e.ProgressPercentage;
                       ImportedFileCollection[i].FileProgressVisible = true;
                       ImportedFileCollection[i].FileStatusVisible = false;
                   }
                   else if (e.ProgressPercentage == 100)
                   {
                       ImportedFileCollection[i].FileProgressVisible = false;
                       ImportedFileCollection[i].FileStatusVisible = true;
                       ImportedFileCollection[i].FileStatus = "Processed";
                   }
               }
           }
              
       }

        #endregion        
        /// <summary>
        ///The DragOver method will determine the supported file or not. 
        /// </summary>
        /// <param name="dropInfo"></param>
        void IDropTarget.DragOver(IDropInfo dropInfo)
        {
            var dragFileList = ((DataObject)dropInfo.Data).GetFileDropList().Cast<string>();
            dropInfo.Effects = dragFileList.Any(item =>
            {
                var extension = Path.GetExtension(item);
                if (extension != null && extension.Equals(".xlsx") || extension.Equals(".xls") || extension.Equals(".txt") || extension.Equals(".csv"))
                {
                    IsDragOverFlag = true;
                }
                return extension != null && extension.Equals(".xlsx") || extension.Equals(".xls") || extension.Equals(".txt") || extension.Equals(".csv");
            }) ? DragDropEffects.Copy : DragDropEffects.None;            

        }

        /// <summary>
        ///Check the Droped file is valid
        ///If the valid files only observer the file collection to process
        /// </summary>
        /// <param name="dropInfo"></param>
        void IDropTarget.Drop(IDropInfo dropInfo)
        {
            IsDragOverFlag = false;
            var dragFileList = ((DataObject)dropInfo.Data).GetFileDropList().Cast<string>();
            dropInfo.Effects = dragFileList.Any(item =>
            {
                var extension = Path.GetExtension(item);
                //Validate the selected file format and valid content.If invalid file cannot upload for import data.
                string validateMsg = ValidateUploadedFiles(item);
                if (validateMsg.Equals("Processed"))
                {
                    
                    FileInfo fileInfo = new FileInfo(item);
                    EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish("");
                    string filename = item;

                    if (extension != null && extension.Equals(".xlsx") || extension.Equals(".xls") || extension.Equals(".txt") || extension.Equals(".csv"))
                    {
                        ImportedFileCollection.Add(new ImportFileModel { Name = Path.GetFileName(item), Path = item, Size = fileInfo.Length, Type = extension, DateModified = fileInfo.CreationTime.ToString(), IsSelected = true });
                    }
                    else
                    {
                        EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish(string.Format("{0} is an unsupported file type. Supported file types are (.xlsx, .xls, .csv, .txt).", extension));
                    }
                }
                else
                {
                    EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish(string.Format(validateMsg));
                }                
                return extension != null && extension.Equals(".xlsx") || extension.Equals(".xls") || extension.Equals(".txt") || extension.Equals(".csv");
            }) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        #endregion
    }   
}
