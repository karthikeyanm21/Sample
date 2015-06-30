using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ReactiveUI;

namespace APLPX.Modules.DataImport.Models
{
    public class ResultModel : ReactiveObject
    {
        /// <summary>
        /// Gets or Sets Filename. Ready to be bind to UI.
        /// </summary>
        private string m_filename;
        public string FileName
        {
            get { return m_filename; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_filename, value);
            }
        }

        /// <summary>
        /// Gets or Sets Format. Ready to be bind to UI.
        /// </summary>
        private string m_format;
        public string Format
        {
            get { return m_format; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_format, value);
            }
        }

        /// <summary>
        /// Gets or Sets RowsinFile. Ready to be bind to UI.
        /// </summary>
        private string m_rowsinfile;
        public string RowsInFile
        {
            get { return m_rowsinfile; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_rowsinfile, value);
            }
        }

        /// <summary>
        /// Gets or Sets Rowsimported. Ready to be bind to UI.
        /// </summary>
        private string m_rowsimported;
        public string RowsImported
        {
            get { return m_rowsimported; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_rowsimported, value);
            }
        }
        /// <summary>
        /// Gets or Sets Rows with Error. Ready to be bind to UI.
        /// </summary>
        private string m_rowswitherror;
        public string RowsWithError
        {
            get { return m_rowswitherror; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_rowswitherror, value);
            }
        }

        /// <summary>
        /// Gets or Sets original resource file path. Ready to be bind to UI.
        /// </summary>
        private string m_sourcefilepath;
        public string SoucreFilePath
        {
            get { return m_sourcefilepath; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_sourcefilepath, value);
            }
        }

        /// <summary>
        /// Gets or Sets Error file path. Ready to be bind to UI.
        /// </summary>
        private string m_errorfilepath;
        public string ErrorFilePath
        {
            get { return m_errorfilepath; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_errorfilepath, value);
            }
        }

        /// <summary>
        /// Gets or Sets Authentication. Ready to be bind to UI.
        /// </summary>
        private string m_errorfilename;
        public string ErrorFileName
        {
            get { return m_errorfilename; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_errorfilename, value);
            }
        }
        /// <summary>
        /// Gets or Sets Authentication. Ready to be bind to UI.
        /// </summary>
        private string m_StartTime;
        public string StartTime
        {
            get { return m_StartTime; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_StartTime, value);
            }
        }
        /// <summary>
        /// Gets or Sets Authentication. Ready to be bind to UI.
        /// </summary>
        private string m_EndTime;
        public string EndTime
        {
            get { return m_EndTime; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_EndTime, value);
            }
        }
        /// <summary>
        /// Gets or Sets Authentication. Ready to be bind to UI.
        /// </summary>
        private string m_Duration;
        public string Duration
        {
            get { return m_Duration; }
            set
            {

                this.RaiseAndSetIfChanged(ref m_Duration, value);
            }
        }
    }
}
