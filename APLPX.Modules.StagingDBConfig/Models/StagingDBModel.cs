using System;
using System.Collections.Generic;

namespace APLPX.Modules.StagingDBConfig.Models
{
    public class StagingDBModel
    {
        /// <summary>
        /// Property for Servertype
        /// </summary>
        public string ServerType { get; set; }

        /// <summary>
        /// Property for Servername
        /// </summary>
        public string ServerName { get; set; }

        /// <summary>
        /// Property for Authentication
        /// </summary>
        public string Authentication { get; set; }

        /// <summary>
        /// Property for Username 
        /// </summary>
        public string LogIn { get; set; }

        /// <summary>
        /// Property for Password
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Property for Databasename
        /// </summary>
        public string DatabaseName { get; set; }
    }
}
