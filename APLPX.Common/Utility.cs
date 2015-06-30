using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APLPX.Common
{
    public static class Utility
    {
        #region Const variable

        /// <summary>
        ///File extention 
        /// </summary>
        public const string TXT = "txt";
        public const string XLS = "xls";
        public const string XLSX = "xlsx";
        public const string CSV = "csv";

        /// <summary>
        ///Error Code define while staging db connections failed. 
        /// </summary>
        public const string CredantialsFailureCode = "18456";
        public const string DatabaseNameFailureCode = "4060";
        public const string ServerNameFailureCode = "53";

        #endregion

        #region

        public static DateTime DateTimeParser(string dateTimeString)
        {
            DateTime parseDateTime = new DateTime();
            if ((dateTimeString.Contains("-")) || (dateTimeString.ToString().Contains("/")))
            {
                parseDateTime = DateTime.FromOADate(Convert.ToDouble(Convert.ToDateTime(dateTimeString).ToOADate()));
            }
            else
            {
                parseDateTime = DateTime.FromOADate(Convert.ToDouble(dateTimeString));
            }

            return parseDateTime;
        }
        #endregion
    }
}
