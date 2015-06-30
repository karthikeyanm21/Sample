using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using APLPX.DataAccess;
using System.Data;
using System.Diagnostics;
using APLPX.Services.BusinessRuleEngine;
using System.Text.RegularExpressions;
using System.Web.Configuration;
using System.Xml;

namespace APLPX.Services
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.PerCall)] 
    public class AplService : IAplService
    {
        private APLPostProxy dataAccessProxy;

        private Dictionary<string, string> m_errormessages;

        #region Contractor
        /// <summary>
        ///TODO 
        /// </summary>
        public AplService()
        {
           dataAccessProxy = new APLPostProxy();
           m_errormessages = new Dictionary<string, string>();
        }
        #endregion

        #region Properties

        /// <summary>
        ///Perist APL Custom error messages 
        /// </summary>
        public Dictionary<string,string> APLErrorMessages
        {
            get
            {
                return m_errormessages;
            }
            set
            {
                m_errormessages = value;
            }
        }
        #endregion

        #region Public methods

        /// <summary>
        /// The public interface for Testing staging DB connection
        /// </summary>
        /// <param name="stagingDbConfig">Staging Db config</param>
        /// <returns></returns>
       public string TestConnection(StagingDbConfig stagingDbConfig)
        {            
            string connectionString = string.Format("Data Source={0};Initial Catalog={1};User={2};Password={3};", stagingDbConfig.ServerName, stagingDbConfig.DatabaseName, stagingDbConfig.LogIn, stagingDbConfig.Password);
            return dataAccessProxy.TestConnectionString(connectionString);
        }

        /// <summary>
       /// The public interface for saving import data 
        /// </summary>
        /// <param name="stagingData"></param>
        /// <returns></returns>
       public bool SaveImportData(DataSet stagingData)
        {
            try
            {
                return dataAccessProxy.SaveImportData(stagingData);
            }
            catch (Exception)
            {
                
                throw;
            }
        }

       /// <summary>
       /// The public interface for saving staging DB connection 
       /// </summary>
       /// <param name="stagingDbConfigInfo"></param>
       /// <returns></returns>
       public string SaveStagingDbConnString(StagingDbConfig stagingDbConfigInfo)
       {
           string[] stagingDbConfigStr = new string[] { stagingDbConfigInfo.ServerName, stagingDbConfigInfo.ServerType, stagingDbConfigInfo.Authentication, stagingDbConfigInfo.LogIn, stagingDbConfigInfo.Password, stagingDbConfigInfo.DatabaseName };

           try
           {
               return dataAccessProxy.SaveStagingDbConfig(stagingDbConfigStr);
           }
           catch (Exception)
           {
               
               throw;
           }
       }

       /// <summary>
       /// The public interface for Applying business rules
       /// </summary>
       /// <param name="importedData"></param>
       /// <returns></returns>
       public DataSet ApplyBRuleForImportedData(DataSet importedData, out DataSet passedRows)
       {
           int RowCount;
           string sErrMsg = string.Empty;
           List<DataRow> errRownumbers = new List<DataRow>();

           try
           {
               List<APLConfigProperties> mapList = new List<APLConfigProperties>();

               using (DataSet dsErrorValues = new DataSet())
               {
                   DataTable dtErrorvalue = new DataTable();
                   Boolean bItemAdded = false;

                   using (DataSet ds = new DataSet())
                   {
                       //Read the Business rule from resource file
                       ds.ReadXml(WebConfigurationManager.AppSettings["RuleFilePath"]);

                       foreach (DataRow iterateRow in ds.Tables[0].Rows)
                       {
                           //Add business rule collection from resource file.
                           mapList.Add(new APLConfigProperties() { Scope = iterateRow["Scope"].ToString(), ValidationRegEx = iterateRow["ValidationRegEx"].ToString(), ErrorMsg = iterateRow["ErrorMsg"].ToString() });
                       }

                   }
                   dtErrorvalue = importedData.Tables[0].Clone();
                   dtErrorvalue.Columns.Add("ErrorMsg");
                   RowCount = importedData.Tables[0].Rows.Count - 1;

                   for (int iterateImportValues = 0; iterateImportValues <= importedData.Tables[0].Rows.Count - 1; iterateImportValues++)
                   {
                       AplBusinessRuleEngine buildRule = new AplBusinessRuleEngine();

                       bItemAdded = false;
                       sErrMsg = string.Empty;

                       if (!string.IsNullOrWhiteSpace(importedData.Tables[0].Rows[iterateImportValues]["Sku"].ToString()))
                       {
                           buildRule.Sku = importedData.Tables[0].Rows[iterateImportValues]["Sku"].ToString();
                       }
                       else
                       {
                           bItemAdded = true;
                           sErrMsg += String.Format("Sku must not contain null");
                       }

                       if (!string.IsNullOrWhiteSpace(importedData.Tables[0].Rows[iterateImportValues]["Name"].ToString()))
                       {
                           buildRule.Name = importedData.Tables[0].Rows[iterateImportValues]["Name"].ToString();
                       }
                       else
                       {
                           bItemAdded = true;
                           sErrMsg += String.Format("Name must not contain null");
                       }

                       if (!string.IsNullOrWhiteSpace(importedData.Tables[0].Rows[iterateImportValues]["Company"].ToString()))
                       {
                           buildRule.Company = importedData.Tables[0].Rows[iterateImportValues]["Company"].ToString();  
                       }
                       else
                       {
                           bItemAdded = true;
                           sErrMsg += String.Format("Company must not contain null");
                       }                                                                         

                       if (StringValRegEx(importedData.Tables[0].Rows[iterateImportValues]["Price"].ToString()))
                       {
                           buildRule.Price = !string.IsNullOrWhiteSpace(importedData.Tables[0].Rows[iterateImportValues]["Price"].ToString()) ? Convert.ToDouble(importedData.Tables[0].Rows[iterateImportValues]["Price"].ToString()) : 0;
                       }  
                       else
                       {
                           bItemAdded = true;
                           sErrMsg += String.Format("Price must not be string values.");
                       }

                       if (importedData.Tables[0].Columns.Contains("Shipping"))
                       {
                           if (StringValRegEx(importedData.Tables[0].Rows[iterateImportValues]["Shipping"].ToString()))
                           {
                               buildRule.Shipping = !string.IsNullOrWhiteSpace(importedData.Tables[0].Rows[iterateImportValues]["Shipping"].ToString()) ? Convert.ToDouble(importedData.Tables[0].Rows[iterateImportValues]["Shipping"]) : 0;
                           }
                           else
                           {
                               bItemAdded = true;
                               sErrMsg += String.Format("Shipping must not be string values.");
                           }
                       }                      
                       if (importedData.Tables[0].Columns.Contains("In_Stock"))
                       {
                           if (StringValRegEx(importedData.Tables[0].Rows[iterateImportValues]["In_Stock"].ToString()))
                           {
                               buildRule.InStock = !string.IsNullOrWhiteSpace(importedData.Tables[0].Rows[iterateImportValues]["In_Stock"].ToString()) ? Convert.ToDouble(importedData.Tables[0].Rows[iterateImportValues]["In_Stock"]) : 0;
                           }
                           else
                           {
                               bItemAdded = true;
                               sErrMsg += String.Format("In_Stock must not be string values.");
                           }
                       }
                       else if (importedData.Tables[0].Columns.Contains("InStock"))
                       {
                           if (StringValRegEx(importedData.Tables[0].Rows[iterateImportValues]["InStock"].ToString()))
                           {
                               buildRule.InStock = !string.IsNullOrWhiteSpace(importedData.Tables[0].Rows[iterateImportValues]["InStock"].ToString()) ? Convert.ToDouble(importedData.Tables[0].Rows[iterateImportValues]["InStock"]) : 0;
                           }
                           else
                           {
                               bItemAdded = true;
                               sErrMsg += String.Format("InStock must not be string values.");
                           }
                       }
                       if (importedData.Tables[0].Columns.Contains("CrawlDate"))
                       {
                           if (StringValRegEx(importedData.Tables[0].Rows[iterateImportValues]["CrawlDate"].ToString()))
                           {
                               if (!string.IsNullOrWhiteSpace(importedData.Tables[0].Rows[iterateImportValues]["CrawlDate"].ToString()))
                               {
                                   buildRule.CrawlDate = DateTimeParser(importedData.Tables[0].Rows[iterateImportValues]["CrawlDate"].ToString());
                               }
                               else
                               {
                                   bItemAdded = true;
                                   sErrMsg += String.Format("CrawlDate must not be empty.");
                               }
                           }
                           else
                           {
                               try
                               {
                                   buildRule.CrawlDate = Convert.ToDateTime(importedData.Tables[0].Rows[iterateImportValues]["CrawlDate"].ToString());
                               }
                               catch
                               {
                                   bItemAdded = true;
                                   sErrMsg += String.Format("CrawlDate must not be string values.");
                               }
                           }
                       }

                       if (string.IsNullOrWhiteSpace(sErrMsg))
                       {
                           for (int iterateRuleList = 0; iterateRuleList < mapList.Count; iterateRuleList++)
                           {
                               string validationRegex = mapList[iterateRuleList].ValidationRegEx;
                               string[] validationValues = validationRegex.Split(null);
                               string prodName = validationValues[0];
                               bool IsNumber = false;
                               Regex regex = new Regex(@"^\d$");
                               if (!regex.IsMatch(validationValues[2]))
                               {
                                   IsNumber = false;
                               }
                               else
                               {
                                   IsNumber = true;
                               }

                               if (IsNumber)
                               {
                                   if (validationRegex.Contains("=="))
                                   {
                                       buildRule.ErrorMsg = mapList[iterateRuleList].ErrorMsg;
                                       if (AplBRuleEngine.Apply(buildRule, "equal", prodName, Convert.ToDouble(validationValues[2])) != "Processed")
                                       {
                                           if (sErrMsg == string.Empty)
                                           {
                                               sErrMsg = sErrMsg + buildRule.ErrorMsg;
                                           }
                                           else
                                           {
                                               sErrMsg = sErrMsg + "," + buildRule.ErrorMsg;
                                           }

                                           bItemAdded = true;
                                       }


                                   }
                                   else if (validationRegex.Contains(">="))
                                   {

                                       buildRule.ErrorMsg = mapList[iterateRuleList].ErrorMsg;
                                       if (AplBRuleEngine.Apply(buildRule, "greater_than_equal", prodName, Convert.ToDouble(validationValues[2])) != "Processed")
                                       {
                                           if (sErrMsg == string.Empty)
                                           {
                                               sErrMsg = sErrMsg + buildRule.ErrorMsg;
                                           }
                                           else
                                           {
                                               sErrMsg = sErrMsg + Environment.NewLine + buildRule.ErrorMsg;
                                           }
                                           bItemAdded = true;
                                       }
                                   }
                                   else if (validationRegex.Contains("<="))
                                   {

                                       buildRule.ErrorMsg = mapList[iterateRuleList].ErrorMsg;
                                       if (AplBRuleEngine.Apply(buildRule, "less_than_equal", prodName, Convert.ToDouble(validationValues[2])) != "Processed")
                                       {
                                           if (sErrMsg == string.Empty)
                                           {
                                               sErrMsg = sErrMsg + buildRule.ErrorMsg;
                                           }
                                           else
                                           {
                                               sErrMsg = sErrMsg + Environment.NewLine + buildRule.ErrorMsg;
                                           }
                                           bItemAdded = true;

                                       }
                                   }
                                   else if (validationRegex.Contains("<"))
                                   {

                                       buildRule.ErrorMsg = mapList[iterateRuleList].ErrorMsg;
                                       if (AplBRuleEngine.Apply(buildRule, "less_than", prodName, Convert.ToDouble(validationValues[2])) != "Processed")
                                       {
                                           if (sErrMsg == string.Empty)
                                           {
                                               sErrMsg = sErrMsg + buildRule.ErrorMsg;
                                           }
                                           else
                                           {
                                               sErrMsg = sErrMsg + Environment.NewLine + buildRule.ErrorMsg;
                                           }
                                           bItemAdded = true;

                                       }
                                   }
                                   else if (validationRegex.Contains(">"))
                                   {

                                       buildRule.ErrorMsg = mapList[iterateRuleList].ErrorMsg;
                                       if (AplBRuleEngine.Apply(buildRule, "greater_than", prodName, Convert.ToDouble(validationValues[2])) != "Processed")
                                       {
                                           if (sErrMsg == string.Empty)
                                           {
                                               sErrMsg = sErrMsg + buildRule.ErrorMsg;
                                           }
                                           else
                                           {
                                               sErrMsg = sErrMsg + Environment.NewLine + buildRule.ErrorMsg;
                                           }
                                           bItemAdded = true;

                                       }
                                   }
                               }
                               else
                               {
                                   if (validationRegex.Contains("notnull"))
                                   {
                                       buildRule.ErrorMsg = mapList[iterateRuleList].ErrorMsg;
                                       if (AplBRuleEngine.Apply(buildRule, "notnull", prodName, validationValues[2]) != "Processed")
                                       {
                                           if (sErrMsg == string.Empty)
                                           {
                                               sErrMsg = sErrMsg + buildRule.ErrorMsg;
                                           }
                                           else
                                           {
                                               sErrMsg = sErrMsg + Environment.NewLine + buildRule.ErrorMsg;
                                           }

                                           bItemAdded = true;
                                       }
                                   }
                               }
                           }
                       }
                       if (bItemAdded)
                       {
                           dtErrorvalue.ImportRow(importedData.Tables[0].Rows[iterateImportValues]);
                           dtErrorvalue.Rows[dtErrorvalue.Rows.Count - 1]["ErrorMsg"] = sErrMsg;
                           errRownumbers.Add(importedData.Tables[0].Rows[iterateImportValues]);

                       }

                   }

                   foreach (DataRow errRow in errRownumbers)
                   {
                       importedData.Tables[0].Rows.Remove(errRow);

                   }
                   dsErrorValues.Tables.Add(dtErrorvalue);
                   passedRows = importedData;
                   return dsErrorValues;
               }
           }
           catch (Exception)
           {

               throw;
           }
       }
       
       /// <summary>
       /// The public interface for updating result table
       /// </summary>
       /// <param name="fileName">Processed file name</param>
       /// <param name="fileExt"> file extention</param>
       /// <param name="totalRows"> total rows </param>
       /// <param name="importedRows">valid process rows count</param>
       /// <param name="errorRows">error rows count</param>
       /// <param name="errorFileName"> error file name</param>
       /// <param name="errorFilePath"> error file location</param>
       /// <returns> update result</returns>
       public bool UpdateResultTable(string fileName, string fileExt, int totalRows, int importedRows, int errorRows,string sourceFilePath, string errorFileName, string errorFilePath,DateTime startTime)
       {
           try
           {
               return dataAccessProxy.UpdateResult(fileName, fileExt, totalRows, importedRows, errorRows,sourceFilePath, errorFileName, errorFilePath, startTime);
           }
           catch (Exception)
           {
               
               throw;
           }
       }

       /// <summary>
       ///TODO: 
       /// </summary>
       /// <returns></returns>
       public DataSet GetResultSet()
       {
           try
           {
               return dataAccessProxy.GetResultDBConfig();
           }
           catch (Exception)
           {
               
               throw;
           }
       }       

       /// <summary>
       ///Get Staging DB config by user id 
       /// </summary>
       /// <param name="userId"></param>
       /// <returns></returns>
       public DataSet GetStagingDbInfoByUser(int userId)
       {
           try
           {
               return dataAccessProxy.GetStagingDBConfig(userId);
           }
           catch (Exception)
           {
               
               throw;
           }
       }      
        
       /// <summary>
       ///Check the logged in user have staging db config or not  
       /// </summary>
       /// <param name="userid"> logged in user id</param>
       /// <returns>bool</returns>
       bool IAplService.CheckStagingDBConfigByUserId(int userid)
       {
           try
           {
               return dataAccessProxy.CheckUserStagingDBConfig(userid);
           }
           catch (Exception)
           {

               throw;
           }
       }
       #endregion

        #region private methods

        /// <summary>
       ///Utility method for check string value has numeric 
       /// </summary>
       /// <param name="inputVal">input string</param>
       /// <returns>result</returns>
        private bool StringValRegEx(string inputVal)
       {
           bool result = false;
           int intVal;
           if (int.TryParse(inputVal, out intVal))
           {
               return true;
           }
           double dVal;
           if (double.TryParse(inputVal, out dVal))
           {
               return true;
           }

           return result;
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


        /// <summary>
        ///TODO -Implementation in progress 
        /// </summary>
        private void GetAPLErrorMsgs()
        {
            XmlDocument myXmlDocument = new XmlDocument();
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            myXmlDocument.Load(@"E:\APLImportData\APLPX.Services\Helper\APLErrorMsgs.xml");
            XmlNodeList list = myXmlDocument.SelectNodes("/AplErrorMessage/ErrorMessages");
            foreach (XmlNode stats in list)
            {
                APLErrorMessages.Add(stats["ErrorCode"].InnerText, stats["ErrorMsg"].InnerText);
            }
        }
        #endregion

    }
}
