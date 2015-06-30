using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Resources;
using System.Security.Cryptography;
using System.Text;
using System.Configuration;
using System.Reflection;
using APLPX.DataAccess.App_Code;

namespace APLPX.DataAccess
{
    public class APLPostProxy 
    {
        #region Private instance

        /// <summary>
        /// A connection string with the global database
        /// specified as the initial catalog.
        /// </summary>
        private static string m_GlobalDBConnectionStr = "";

        /// <summary>
        /// A property for getting and setting the connection
        /// string in which the global database is specified
        /// as the initial catalog.
        /// </summary>
        public static string GlobalDBConnectionStr
        {
            get
            {
                return m_GlobalDBConnectionStr;
            }

            set
            {
                m_GlobalDBConnectionStr = value;
            }
        }     

        #endregion

        #region Constructor

        /// <summary>
        ///Constructor 
        /// </summary>
        public APLPostProxy()
        {
           GlobalDBConnectionStr= GetGlobalConnectionString();     
        }
        #endregion

        #region Admin Feature

        /// <summary>
        /// Staging DB connection configuration validation.
        /// </summary>
        /// <param name="stagingDBConnectionString">Staging DB connection string</param>
        /// <returns>bool result</returns>
        public string TestConnectionString(string stagingDBConnectionString)
        {
            string message = "Test Connection Successful.";

            if (string.IsNullOrWhiteSpace(stagingDBConnectionString))
            {
                message = "Failed: missing sql connection string";
            }
            else
            {
                try
                {                                                          
                    using (SqlConnection conn = new SqlConnection(stagingDBConnectionString))
                    {
                        conn.Open();
                    }
                }
                catch (SqlException sqle)
                {
                    message = sqle.Number.ToString();                  
                }
            }
            return message;
        }

        /// <summary>
        /// Save staging DB configuration 
        /// </summary>
        /// <param name="stagingDbConfigInfo">User defined staging Db config info</param>
        /// <returns>Db inserted result</returns>
        public string SaveStagingDbConfig(string[] stagingDbConfigInfo)
        {
            string resultMessage = "Database Configuration Save Successful";

            string queryBuilder = string.Format("IF EXISTS(SELECT 1 FROM APLPRICEEXPERTCONFIG WHERE USERID={6}) BEGIN UPDATE APLPRICEEXPERTCONFIG SET SERVERNAME='{0}',SERVERTYPE='{1}',AUTHENTICATION='{2}',USERNAME='{3}',PASSWORD='{4}',DATABASENAME='{5}',USERID={6} WHERE USERID={6} END ELSE BEGIN INSERT INTO [APLPRICEEXPERTCONFIG]([SERVERNAME],[SERVERTYPE],[AUTHENTICATION],[USERNAME],[PASSWORD],[DATABASENAME],[USERID]) VALUES('{0}','{1}','{2}','{3}','{4}','{5}',{6}) END", 
                                                stagingDbConfigInfo[0], stagingDbConfigInfo[1],stagingDbConfigInfo[2],stagingDbConfigInfo[3],Encrypt(stagingDbConfigInfo[4]),stagingDbConfigInfo[5],1);
            SqlTransaction transaction = null;

            using (SqlConnection conn = new SqlConnection(GlobalDBConnectionStr))
            {
                try
                {
                    conn.Open();
                    transaction = conn.BeginTransaction(IsolationLevel.ReadUncommitted);
                    bool result= ExecuteNonQuery(transaction, queryBuilder) != 0 ? true : false;
                    transaction.Commit();
                   // return result;
                }
                catch (SqlException sqle)
                {                    
                    resultMessage = sqle.Number.ToString();
                    transaction.Rollback();
                }
            }
            return resultMessage;
        }

        /// <summary>
        /// Populate  user configured staging DB Info
        /// </summary>
        /// <returns>Dataset</returns>
        public DataSet GetStagingDBConfig(int userId)
        {
            //TODO:Based on logged in user id to get  staging DB configuration.
            string getStagingDbConfigQuery = string.Format("SELECT [SERVERNAME],[SERVERTYPE],[AUTHENTICATION],[USERNAME],[PASSWORD],[DATABASENAME] FROM [APLPRICEEXPERTCONFIG] WHERE USERID={0}",userId);
            SqlTransaction transaction = null;
            DataSet stagingConfigDs =new DataSet();            
            using (SqlConnection conn = new SqlConnection(GlobalDBConnectionStr))
            {
                try
                {
                    conn.Open();
                    transaction = conn.BeginTransaction(IsolationLevel.ReadUncommitted);
                    using (SqlCommand cmdGetStagingConfig = APLDBConnectionManager.GetCommand(transaction))
                    {
                        cmdGetStagingConfig.CommandText = getStagingDbConfigQuery;
                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmdGetStagingConfig))
                        {
                            transaction.Commit();
                            adapter.Fill(stagingConfigDs);
                        }
                    }
                }
                catch (SqlException ex)
                {
                    if (ex.Number.Equals(53) || ex.Number.Equals(5))
                    {
                        return stagingConfigDs;
                    }
                    else
                    {
                        transaction.Rollback();
                        throw ex;
                    }
                }
            }

            if (stagingConfigDs.Tables.Count > 0 && stagingConfigDs.Tables[0].Rows.Count > 0)
            {
                stagingConfigDs.Tables[0].Rows[0]["Password"] = Decrypt(stagingConfigDs.Tables[0].Rows[0]["Password"].ToString());
            }

            return stagingConfigDs;
        }

        #endregion

        #region Import Feature

        /// <summary>
        /// Save import data  DB configuration
        /// </summary>
        /// <param name="importedData"></param>
        /// <returns></returns>
        public bool SaveImportData(DataSet importedData)
        {
            SqlTransaction transaction = null;
            DateTime CrawlDate = System.DateTime.Now;
            try
            {               
                using (SqlConnection conn = new SqlConnection())
                {
                    //TODO: current logged in user id must pass the parameter to get statging db configuration.
                    int userId = 1;
                    conn.ConnectionString = GetStagingDBConnectionStr(userId);     
                    //END:

                    conn.Open();
                    transaction = conn.BeginTransaction(IsolationLevel.ReadUncommitted);
                   
                        for (int iterateDsdata = 0; iterateDsdata <= (importedData.Tables[0].Rows.Count - 1); iterateDsdata++)
                        {
                            int inStockval = 0;
                            if(importedData.Tables[0].Columns.Contains("In_Stock"))
                            {
                                inStockval = Convert.ToInt16(importedData.Tables[0].Rows[iterateDsdata]["In_Stock"]);
                            }
                            else if(importedData.Tables[0].Columns.Contains("InStock"))
                            {
                                inStockval = Convert.ToInt16(importedData.Tables[0].Rows[iterateDsdata]["InStock"]);
                            }
                            if (importedData.Tables[0].Columns.Contains("CrawlDate"))
                            {
                                CrawlDate = DateTimeParser(importedData.Tables[0].Rows[iterateDsdata]["CrawlDate"].ToString());
                            }
                            else if (importedData.Tables[0].Columns.Contains("Crawl_Date"))
                            {
                                CrawlDate = DateTimeParser(importedData.Tables[0].Rows[iterateDsdata]["Crawl_Date"].ToString());
                            }
                            else if (importedData.Tables[0].Columns.Contains("crawl_date"))
                            {
                                CrawlDate = DateTimeParser(importedData.Tables[0].Rows[iterateDsdata]["crawl_date"].ToString());
                            }
                         
                        string ifExistRows = string.Format("IF EXISTS(SELECT 1 FROM APLSTOCKKEEPING WHERE SKU='{0}' AND companyName ='{1}')" + "BEGIN UPDATE APLSTOCKKEEPING SET PRODUCTNAME='{2}', PRODUCTPRICE={3} , SHIPPINGCOST={4},INSTOCK={5},Crawl_Date='{6}',LASTMODIFIEDDATE='{7}' WHERE SKU ='{0}' AND COMPANYNAME ='{1}' END " +"ELSE BEGIN INSERT INTO APLSTOCKKEEPING VALUES('{0}','{2}','{1}',{3},{4},'{5}','{6}','{7}','{8}') END",importedData.Tables[0].Rows[iterateDsdata]["Sku"], importedData.Tables[0].Rows[iterateDsdata]["Company"], importedData.Tables[0].Rows[iterateDsdata]["Name"],importedData.Tables[0].Rows[iterateDsdata]["Price"], importedData.Tables[0].Rows[iterateDsdata]["Shipping"], inStockval,CrawlDate, System.DateTime.Now, System.DateTime.Now);
                    
                            ExecuteNonQuery(transaction, ifExistRows);
                        }
                        transaction.Commit();
                }               
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                throw ex;
            }
            return true;
        }        

        /// <summary>
        /// Import data process status will update to result table.
        /// </summary>       
        public bool UpdateResult(string fileName,string fileExt,int totalRows,int importedRows,int errorRows,string sourceFilePath,string errorFileName,string errorFilePath,DateTime startTime)
        {
            bool result = false;
            SqlTransaction transaction = null;
            try
            {
                TimeSpan Duration=System.DateTime.Now.Subtract(startTime);
                string resultVal = string.Format("('{0}','{1}',{2},{3},{4},'{5}','{6}','{7}','{8}','{9}','{10}')",
                                         fileName, fileExt, totalRows, importedRows, errorRows, errorFilePath, errorFileName,startTime,System.DateTime.Now,String.Format("{0:0.00}", Duration.TotalSeconds)   +" Sec",sourceFilePath);
                using (SqlConnection conn = new SqlConnection())
                {

                    //TODO: current logged in user id must pass the parameter to get statging db configuration.
                    int userId = 1;
                    conn.ConnectionString = GetStagingDBConnectionStr(userId);
                    //END:
                    conn.Open();
                    transaction = conn.BeginTransaction(IsolationLevel.ReadUncommitted);

                   int queryResult = ExecuteNonQuery(transaction, string.Format("IF EXISTS(SELECT 1 FROM RESULT WHERE FILENAME ='{0}' AND ROWSINFILE={1})" +
                                                                                 "BEGIN UPDATE RESULT SET ROWSINFILE={1}, ROWSIMPORTED={3} , ROWSWITHERROR={4},ERRORFILEPATH='{5}',ERRORFILENAME='{6}',STARTTIME='{7}',ENDTIME='{8}',DURATION='{9}',SOURCEFILEPATH='{10}' WHERE FILENAME ='{0}' AND ROWSINFILE={1} END " +
                                                                                 " ELSE BEGIN INSERT INTO RESULT(FILENAME,FORMAT,ROWSINFILE,ROWSIMPORTED,ROWSWITHERROR,ERRORFILEPATH,ERRORFILENAME,STARTTIME,ENDTIME,DURATION,SOURCEFILEPATH) VALUES {2}; END", fileName, totalRows, resultVal, importedRows, errorRows, errorFilePath, errorFileName, startTime, System.DateTime.Now, String.Format("{0:0.00}", Duration.TotalSeconds) + " Sec", sourceFilePath));
                    transaction.Commit();
                   result = queryResult !=0? true:false;
                }
            }
            catch (Exception)
            {
                transaction.Rollback();
                return false;
            }
            return result;
        }

        /// <summary>
        /// Get import data process result set
        /// </summary>
        /// <returns>Dataset</returns>
        public DataSet GetResultDBConfig()
        {
            ////The query generated by least processed file result sets.
            string getImportDataResultQuery = string.Format(" SELECT [FILENAME],[FORMAT],[ROWSINFILE],[ROWSIMPORTED],[ROWSWITHERROR],[ERRORFILEPATH],[ERRORFILENAME],[STARTTIME],[ENDTIME],[DURATION],[SOURCEFILEPATH] FROM [RESULT] WHERE [ENDTIME] > DATEADD(MINUTE, -5, GETDATE()) ORDER BY [ENDTIME] DESC");

            DataSet ResultDBConfigDs = new DataSet();
            SqlTransaction transaction = null;

            using (SqlConnection conn = new SqlConnection())
            {
                try
                {
                    //TODO: current logged in user id must pass the parameter to get statging db configuration.
                    int userId = 1;
                    conn.ConnectionString = GetStagingDBConnectionStr(userId);
                    //END:
                    conn.Open();
                    transaction = conn.BeginTransaction(IsolationLevel.ReadUncommitted);
                    using (SqlCommand cmdGetImportDataResult = APLDBConnectionManager.GetCommand(transaction))
                    {
                        cmdGetImportDataResult.CommandText = getImportDataResultQuery;

                        using (SqlDataAdapter sqlDa = new SqlDataAdapter(cmdGetImportDataResult))
                        {
                            transaction.Commit();
                            sqlDa.Fill(ResultDBConfigDs);
                        }
                    }
                }
                catch (SqlException)
                {
                    transaction.Rollback();
                    throw;
                }
            }
            return ResultDBConfigDs;
        }

        /// <summary>
        /// Check the current user have Staging DB config 
        /// </summary>
        /// <returns>Dataset</returns>
        public bool CheckUserStagingDBConfig(int userId)
        {
            //TODO:Based on logged in user id to get staging DB configuration.
            string getStagingDbConfigQuery =string.Format("SELECT COUNT(*) FROM APLPRICEEXPERTCONFIG WHERE USERID={0}",userId);
            SqlTransaction transaction = null;
            DataSet stagingConfigDs = new DataSet();
            bool result =false;
            using (SqlConnection conn = new SqlConnection(GlobalDBConnectionStr))
            {
                try
                {
                    conn.Open();
                    transaction = conn.BeginTransaction(IsolationLevel.ReadUncommitted);
                    using (SqlCommand cmdGetStagingConfig = APLDBConnectionManager.GetCommand(transaction))
                    {
                        cmdGetStagingConfig.CommandText = getStagingDbConfigQuery;
                         result = Convert.ToBoolean(cmdGetStagingConfig.ExecuteScalar());
                        transaction.Commit();
                        //using (SqlDataReader adapter = new SqlDataAdapter(cmdGetStagingConfig))
                        //{
                        //    transaction.Commit();
                        //    adapter.Fill(stagingConfigDs);
                        //}
                    }
                    return result; 
                }
                catch (SqlException ex)
                {
                       result = false;             
                        transaction.Rollback();
                        throw ex;                        
                }
            }

        }

        #endregion

        #region Private methods
        /// <summary>
        /// Utility method to excute the database query.
        /// </summary>
        /// <param name="conn">Database connection</param>
        /// <param name="query">Executed query with values</param>
        /// <returns></returns>
        private int ExecuteNonQuery(SqlTransaction transaction, string executeQuery)
        {
            try
            {
                using (SqlCommand sqlCommand = APLDBConnectionManager.GetCommand(transaction))
                {
                    sqlCommand.CommandText = executeQuery;
                    return sqlCommand.ExecuteNonQuery();                                      
                }
            }
            catch (Exception)
            {               
                throw;
            }
        }

        /// <summary>
        /// Encrypting the User details 
        /// </summary>
        /// <param name="clearText"></param>
        /// <returns></returns>
        private string Encrypt(string clearText)
        {
            try
            {
                ResourceManager RM = new ResourceManager("APLPX.DataAccess.Resources", Assembly.GetExecutingAssembly());
                string EncryptionKey = RM.GetString("CipherKey").ToString();

                byte[] clearBytes = Encoding.Unicode.GetBytes(clearText);
                using (Aes encryptor = Aes.Create())
                {
                    Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                    encryptor.Key = pdb.GetBytes(32);
                    encryptor.IV = pdb.GetBytes(16);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write);
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();                       
                        clearText = Convert.ToBase64String(ms.ToArray());
                    }
                }
            }
            catch (Exception)
            {
                
                throw;
            }
            return clearText;
        }

        /// <summary>
        /// Decrypting the User details
        /// </summary>
        /// <param name="cipherText"></param>
        /// <returns></returns>
        private string Decrypt(string cipherText)
        {
            ResourceManager RM = new ResourceManager("APLPX.DataAccess.Resources", Assembly.GetExecutingAssembly());
            string EncryptionKey = RM.GetString("CipherKey").ToString();
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write);
                    cs.Write(cipherBytes, 0, cipherBytes.Length);
                    cs.Close();
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }     

        /// <summary>
        /// Declaring Global connection string 
        /// </summary>
        private string GetGlobalConnectionString()
        {
            return ConfigurationManager.ConnectionStrings["AplDefaultConnectionString"].ConnectionString;
        }

        /// <summary>
        ///Get staging db configuration. 
        /// </summary>
        /// <param name="userID">Logged in user id</param>
        /// <returns>staging connection string</returns>
        private string GetStagingDBConnectionStr(int userID)
        {
            string stagingDbConnectionStr = string.Empty;
            using (SqlConnection sqlCon = new SqlConnection())
            {
                sqlCon.ConnectionString = GlobalDBConnectionStr;
                string getStagingConnection = string.Format("SELECT * FROM APLPRICEEXPERTCONFIG WHERE USERID={0}", userID); 
                sqlCon.Open();
                using (SqlDataAdapter sqlDA = new SqlDataAdapter(getStagingConnection, sqlCon))
                {
                    using (DataSet stagingConDs = new DataSet())
                    {
                        sqlDA.Fill(stagingConDs);

                        stagingDbConnectionStr = string.Format("Data Source={0};Initial Catalog={1};User={2};Password={3};",
                                                stagingConDs.Tables[0].Rows[0]["serverName"].ToString(), stagingConDs.Tables[0].Rows[0]["databaseName"].ToString(),
                                                stagingConDs.Tables[0].Rows[0]["userName"].ToString(), Decrypt(stagingConDs.Tables[0].Rows[0]["password"].ToString()));
                    }
                }
            }
            return stagingDbConnectionStr;
        }

        /// <summary>
        ///Utility method to date time parsing 
        /// </summary>
        /// <param name="dateTimeString">string datetime value</param>
        /// <returns>datetime</returns>
        private DateTime DateTimeParser(string dateTimeString)
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
