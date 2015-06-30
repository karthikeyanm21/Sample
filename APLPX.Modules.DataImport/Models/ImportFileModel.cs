using ReactiveUI;
using System;
using System.Collections.Generic;


namespace APLPX.Modules.DataImport.Models
{
    public class ImportFileModel : ReactiveObject
    {
        #region private memers

        private string m_name;
        private string m_path;
        private string m_datemodified;
        private string m_type;
        private long m_size;
        private bool m_isselected;
        private int m_progress; 
        private string m_filestatus;
        private bool m_filestatusvisible;
        private bool m_fileprogressvisible;

        #endregion

        /// <summary>
        ///Get and Set the Name of the file
        /// </summary>
        public string Name 
        {
            get
            {
                return m_name;
            }
             set
            {
                this.RaiseAndSetIfChanged(ref m_name, value);
            }
        }

        /// <summary>
        ///Get and set the path of the file. 
        /// </summary>
        public string Path 
        {
            get
            {
                return m_path;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_path, value);
            }
        }

        /// <summary>
        ///Get and set the Type of the file. 
        /// </summary>
        public string DateModified 
        {
            get
            {
                return m_datemodified;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_datemodified, value);
            }
        }

        /// <summary>
        ///Get and set the Type of the file. 
        /// </summary>
        public string Type 
        {
            get
            {
                return m_type;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_type, value);
            }
        }

        /// <summary>
        ///Get and set the Size of the file. 
        /// </summary>
        public long Size 
        {
            get
            {
                return m_size;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_size, value);
            }
        }

        /// <summary>
        ///Get and set the Select file to process. 
        /// </summary>
        public bool IsSelected 
        {
            get
            {
                return m_isselected;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_isselected, value);
            }
        }

        /// <summary>
        ///Get and set the Error message of files. 
        /// </summary>
        public string FileStatus 
        {
            get
            {
                return m_filestatus;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_filestatus, value);
            }
        }

        /// <summary>
        ///Get and set the progress value property
        /// </summary>
        public int Progress
        {
            get
            {
                return m_progress;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_progress, value);
            }
        }

        /// <summary>
        ///Get and set the File status message 
        /// </summary>
        public bool FileStatusVisible
        {
            get
            {
                return m_filestatusvisible;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_filestatusvisible, value);
            }
        }

        /// <summary>
        /// Get and set the File progress visible status
        /// </summary>
        public bool FileProgressVisible
        {
            get
            {
                return m_fileprogressvisible;
            }
            set
            {
                this.RaiseAndSetIfChanged(ref m_fileprogressvisible, value);
            }
        }
    }
}
