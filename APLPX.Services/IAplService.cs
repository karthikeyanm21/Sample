using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace APLPX.Services
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract]
    public interface IAplService
    {

        #region Apl web interface
        [OperationContract]
        bool CheckStagingDBConfigByUserId(int userid);
        
        [OperationContract]
        DataSet GetStagingDbInfoByUser(int userId);

        [OperationContract]
        string TestConnection(StagingDbConfig stagingDbConfig);

        [OperationContract]
        string SaveStagingDbConnString(StagingDbConfig stagingDbConfigInfo);

        [OperationContract]
        bool SaveImportData(DataSet stagingData);

        [OperationContract]
        DataSet ApplyBRuleForImportedData(DataSet importedData, out DataSet passedRows);

        [OperationContract]
        DataSet GetResultSet();

        [OperationContract]
        bool UpdateResultTable(string fileName, string fileExt, int totalRows, int importedRows, int errorRows, string sourceFilePath, string errorFileName, string errorFilePath, DateTime startTime);

        #endregion

    }


    // Use a data contract as illustrated in the sample below to add composite types to service operations.
    [DataContract]
    public class StagingDbConfig
    {
        /// <summary>
        /// Property for Servertype
        /// </summary>
        [DataMember]
        public string ServerType { get; set; }

        /// <summary>
        /// Property for Servername
        /// </summary>
        [DataMember]
        public string ServerName { get; set; }

        /// <summary>
        /// Property for Authentication
        /// </summary>
        [DataMember]
        public string Authentication { get; set; }

        /// <summary>
        /// Property for Username 
        /// </summary>
        [DataMember]
        public string LogIn { get; set; }

        /// <summary>
        /// Property for Password
        /// </summary>
        [DataMember]
        public string Password { get; set; }

        /// <summary>
        /// Property for Databasename
        /// </summary>
        [DataMember]
        public string DatabaseName { get; set; }
    }
}
