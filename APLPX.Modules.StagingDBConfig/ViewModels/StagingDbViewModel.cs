using System;
using System.Collections.Generic;
using System.Linq;
using ReactiveUI;
using System.Windows.Input;
using Microsoft.Practices.Prism.Commands;
using APLPX.Modules.StagingDBConfig.Models;
using System.Collections.ObjectModel;
using System.Windows;
using Microsoft.Practices.Prism.PubSubEvents;
using APLPX.Common;
using APLPX.Client;
using System.Data;


namespace APLPX.Modules.StagingDBConfig.ViewModels
{
    public class StagingDbViewModel : ReactiveObject
    {
        #region Private instance

        /// <summary>
        /// Creating Staging DB model instance 
        /// </summary>
        private StagingDBModel stagingdbDom;

        /// <summary>
        /// Creating Server type observable collection instance 
        /// </summary>
        private ObservableCollection<string> _stagingServerType;

        /// <summary>
        /// Creating Server Authentication observable collection instance 
        /// </summary>
        private readonly ObservableCollection<string> _stagingAutentication;       

        /// <summary>
        /// Creating AdminResources instance
        /// </summary>
        private AplBusinessLayer aplBusinessL;

        #endregion
     
        #region Constructor
        /// <summary>
        /// Default constructor initialization 
        /// </summary>
        public StagingDbViewModel() 
        {
            InitializeCommand();
             stagingdbDom = new StagingDBModel();
             aplBusinessL = new AplBusinessLayer();
            _stagingServerType = new ObservableCollection<string>();
            _stagingAutentication = new ObservableCollection<string>();

            this.SimulateStagingData();
            this.PopulateStagingDbConfig();
            ServerNameHighlight = CredentialsHighlight= DatabaseNameHighlight= "white";
           
        }
        /// <summary>
        /// Get Staging DB values for Validation
        /// </summary>
        /// <returns></returns>
        public static bool StagingDBCount()
        {
            try
            {
                //TODO: Logged in User id must pass the arguement.
                int userId = 1;
                using (AplBusinessLayer aplBusiness = new AplBusinessLayer())
                {
                    return aplBusiness.CheckStagingDBConfigByUser(userId);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }


        #endregion     

        #region commands

        /// <summary>
        /// Command to check staging config to test in database server.
        /// </summary>
        public ReactiveCommand<object> TestConnectionCmd { get; private set; }

        /// <summary>
        /// Command to save staging config to database server.
        /// </summary>
        public ReactiveCommand<object> SaveDBStagingCmd { get; private set; }
        #endregion       

        #region Properties
        /// <summary>
        /// Gets or Sets ServerType. Ready to be bind to UI.
        /// Impelments INotifyPropertyChanged which enables the binded element to refresh itself whenever the value changes.
        /// </summary>
        private string m_serverType;
        public string ServerType
        {
            get { return m_serverType; }
            set
            {
                stagingdbDom.ServerType = value;
                this.RaiseAndSetIfChanged(ref m_serverType, value);
            }
        }

        /// <summary>
        /// Gets or Sets ServerName. Ready to be bind to UI.
        /// Impelments INotifyPropertyChanged which enables the binded element to refresh itself whenever the value changes.
        /// </summary>
        private string m_serverName;
        public string ServerName
        {
            get { return m_serverName; }
            set
            {
                stagingdbDom.ServerName = value;
                this.RaiseAndSetIfChanged(ref m_serverName, value);
            }
        }

        /// <summary>
        /// Gets or Sets Authentication. Ready to be bind to UI.
        /// Impelments INotifyPropertyChanged which enables the binded element to refresh itself whenever the value changes.
        /// </summary>
        private string m_authentication;
        public string Authentication
        {
            get { return m_authentication; }
            set
            {
                stagingdbDom.Authentication = value;
                this.RaiseAndSetIfChanged(ref m_authentication, value);
                this.RaisePropertyChanged("AllowLogInCredential");
            }
        }

        /// <summary>
        /// Gets or Sets Authentication. Ready to be bind to UI.
        /// Impelments INotifyPropertyChanged which enables the binded element to refresh itself whenever the value changes.
        /// </summary>
        private string m_login;
        public string LogIn
        {
            get { return m_login; }
            set
            {
                stagingdbDom.LogIn = value;
                this.RaiseAndSetIfChanged(ref m_login, value);
            }
        }
        /// <summary>
        /// Gets or Sets Authentication. Ready to be bind to UI.
        /// Impelments INotifyPropertyChanged which enables the binded element to refresh itself whenever the value changes.
        /// </summary>
        private string m_password;
        public string Password
        {
            get { return m_password; }
            set
            {
                stagingdbDom.Password = value;
                this.RaiseAndSetIfChanged(ref m_password, value);
            }
        }

        /// <summary>
        /// Gets or Sets Authentication. Ready to be bind to UI.
        /// Impelments INotifyPropertyChanged which enables the binded element to refresh itself whenever the value changes.
        /// </summary>
        private string m_databasename;
        public string DatabaseName
        {
            get { return m_databasename; }
            set
            {
                stagingdbDom.DatabaseName = value;
                this.RaiseAndSetIfChanged(ref m_databasename, value);
            }
        }

        private string m_serverNameboderbrushcolor;
        public string ServerNameHighlight
        {
            get { return m_serverNameboderbrushcolor; }
            set
            {
                this.RaiseAndSetIfChanged(ref m_serverNameboderbrushcolor, value);
            }
        }

        private string m_credentialsboderbrushcolor;
        public string CredentialsHighlight
        {
            get { return m_credentialsboderbrushcolor; }
            set
            {
                this.RaiseAndSetIfChanged(ref m_credentialsboderbrushcolor, value);
            }
        }

        private string m_databseNameboderbrushcolor;
        public string DatabaseNameHighlight
        {
            get { return m_databseNameboderbrushcolor; }
            set
            {
                this.RaiseAndSetIfChanged(ref m_databseNameboderbrushcolor, value);
            }
        }


        /// <summary>
        /// Gets the Staging server type. Used to maintain the server type.
        /// Since this observable collection it makes sure all changes will automatically reflect in UI 
        /// as it implements both CollectionChanged, PropertyChanged;
        /// </summary>
        public ObservableCollection<string> StagingServerType { get { return _stagingServerType; } }

        /// <summary>
        ///TODO: 
        /// </summary>
        public ObservableCollection<string> StagingAutentication { get { return _stagingAutentication; } }

        /// <summary>
        /// Gets whether or not the user should be able to select a Authentication Mode.
        /// </summary>
        public bool AllowLogInCredential
        {
            get { return (Authentication == "Windows Authentication" ? false : true); }
        }

        #endregion

        #region Private methods

        /// <summary>
        ///Simulate Staging config info 
        /// </summary>
        private void SimulateStagingData()
        {
            StagingServerType.Add("Database Engine");
            StagingServerType.Add("Analysis Services");
            StagingServerType.Add("Reporting Services");
            StagingServerType.Add("Integration Services");

            StagingAutentication.Add("Windows Authentication");
            StagingAutentication.Add("SQL Server Authentication");

            ServerType = "Database Engine";
            Authentication = "SQL Server Authentication";
        }


        /// <summary>
        /// Initialize actions command for staging DB.
        /// </summary>
        private void InitializeCommand()
        {
            TestConnectionCmd = ReactiveCommand.Create();
            TestConnectionCmd.Subscribe(x => TestConnectionCommandExecuted(x));

            SaveDBStagingCmd = ReactiveCommand.Create();
            SaveDBStagingCmd.Subscribe(x => SaveDBStagingConfig(x));
           
        }


        /// <summary>
        /// Testing the database connection
        /// </summary>
        /// <param name="sender"></param>
        private void TestConnectionCommandExecuted(object sender)
        {
            //Always create a new instance of Destination before adding. 
            //Otherwise we will endup sending the same instance that is binded, to the BL which will cause complications
            Client.localhost.StagingDbConfig destination = new Client.localhost.StagingDbConfig { ServerType = ServerType, ServerName = ServerName, Authentication = Authentication, LogIn = LogIn, Password = Password, DatabaseName = DatabaseName };
            string testConResult = string.Empty;
            try
            {
                if (!string.IsNullOrWhiteSpace(ServerName) && !string.IsNullOrWhiteSpace(DatabaseName) && (!string.IsNullOrWhiteSpace(LogIn) || !string.IsNullOrWhiteSpace(Password)))
                {
                     testConResult = aplBusinessL.TestConnection(destination);

                    if (testConResult.Equals(Utility.CredantialsFailureCode))
                    {
                        testConResult = "User credentials failed";
                        //TODO :we need to improve the error field hight binding command
                        CredentialsHighlight = "red";
                        DatabaseNameHighlight = "white";
                        ServerNameHighlight = "white";
                        //END
                    }
                    else if (testConResult.Equals(Utility.DatabaseNameFailureCode))
                    {
                        testConResult = "Database name is not valid";
                        //TODO :we need to improve the error field hight binding command
                        DatabaseNameHighlight = "red";
                        CredentialsHighlight = "white";
                        ServerNameHighlight = "white";
                        //END
                    }
                    else if (testConResult.Equals(Utility.ServerNameFailureCode))
                    {
                        testConResult = "Server name is not valid";
                        //TODO :we need to improve the error field hight binding command
                        ServerNameHighlight = "red";
                        CredentialsHighlight = "white";
                        DatabaseNameHighlight = "white";
                        //END
                    }
                    else
                    {
                        //TODO :we need to improve the error field hight binding command
                        ServerNameHighlight = "white";
                        CredentialsHighlight = "white";
                        DatabaseNameHighlight = "white";
                        //END
                    }
                }
                else
                {
                    if (Authentication.Equals("SQL Server Authentication"))
                    {
                        CredentialsHighlight = string.IsNullOrWhiteSpace(LogIn) && string.IsNullOrWhiteSpace(Password) ? "red" : "white";
                    }                  
                    ServerNameHighlight = string.IsNullOrWhiteSpace(ServerName) ? "red": "white";
                    DatabaseNameHighlight = string.IsNullOrWhiteSpace(DatabaseName) ? "red" : "white";
                    testConResult = "Server Name and Database Name is mandatory";
                }
                EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish(testConResult);                     
            }
            catch (Exception ex)
            {
                EventAgg._eventAggregator.GetEvent<ErrorMessageEvent>().Publish(ex.Message);
            }      
        }
        private string Testconnection()
        {
            //Always create a new instance of Destination before adding. 
            //Otherwise we will endup sending the same instance that is binded, to the BL which will cause complications
            Client.localhost.StagingDbConfig destination = new Client.localhost.StagingDbConfig { ServerType = ServerType, ServerName = ServerName, Authentication = Authentication, LogIn = LogIn, Password = Password, DatabaseName = DatabaseName };
            string testConResult = string.Empty;
            try
            {
                if (!string.IsNullOrWhiteSpace(ServerName) && !string.IsNullOrWhiteSpace(DatabaseName) && (!string.IsNullOrWhiteSpace(LogIn) || !string.IsNullOrWhiteSpace(Password)))
                {
                    testConResult = aplBusinessL.TestConnection(destination);

                    if (testConResult.Equals(Utility.CredantialsFailureCode))
                    {
                        testConResult = "User credentials failed";
                        //TODO :we need to improve the error field hight binding command
                        CredentialsHighlight = "red";
                        DatabaseNameHighlight = "white";
                        ServerNameHighlight = "white";
                        //END
                    }
                    else if (testConResult.Equals(Utility.DatabaseNameFailureCode))
                    {
                        testConResult = "Database name is not valid";
                        //TODO :we need to improve the error field hight binding command
                        DatabaseNameHighlight = "red";
                        CredentialsHighlight = "white";
                        ServerNameHighlight = "white";
                        //END
                    }
                    else if (testConResult.Equals(Utility.ServerNameFailureCode))
                    {
                        testConResult = "Server name is not valid";
                        //TODO :we need to improve the error field hight binding command
                        ServerNameHighlight = "red";
                        CredentialsHighlight = "white";
                        DatabaseNameHighlight = "white";
                        //END
                    }
                    else
                    {
                        //TODO :we need to improve the error field hight binding command
                        ServerNameHighlight = "white";
                        CredentialsHighlight = "white";
                        DatabaseNameHighlight = "white";
                        //END
                    }
                }
                else
                {
                    if (Authentication.Equals("SQL Server Authentication"))
                    {
                        CredentialsHighlight = string.IsNullOrWhiteSpace(LogIn) && string.IsNullOrWhiteSpace(Password) ? "red" : "white";
                    }
                    ServerNameHighlight = string.IsNullOrWhiteSpace(ServerName) ? "red" : "white";
                    DatabaseNameHighlight = string.IsNullOrWhiteSpace(DatabaseName) ? "red" : "white";
                    testConResult = "Server Name and Database Name is mandatory";
                }
                EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish(testConResult);
            }
            catch (Exception ex)
            {
                EventAgg._eventAggregator.GetEvent<ErrorMessageEvent>().Publish(ex.Message);
            }
            return testConResult;
        }
        /// <summary>
        /// Add operation of the AddDestinationCmd.
        /// Operation that will be performormed on the control click.
        /// </summary>
        /// <param name="obj"></param>
        private void SaveDBStagingConfig(object sender)
        {
            if (Testconnection() == "Test Connection Successful.")
            {
                //Always create a new instance of Destination before adding. 
                //Otherwise we will endup sending the same instance that is binded, to the BL which will cause complications
                Client.localhost.StagingDbConfig saveStagingConConfig = new Client.localhost.StagingDbConfig { ServerType = ServerType, ServerName = ServerName, Authentication = Authentication, LogIn = LogIn, Password = Password, DatabaseName = DatabaseName };
                string result = string.Empty;

                try
                {
                    //Add Destination will be successful only if the Destination information valid.  
                    if (!string.IsNullOrWhiteSpace(ServerName) && !string.IsNullOrWhiteSpace(DatabaseName) && (!string.IsNullOrWhiteSpace(LogIn) && !string.IsNullOrWhiteSpace(Password)))
                    {
                        result = aplBusinessL.SaveStagingDbConfig(saveStagingConConfig);

                        if (result.Equals(Utility.CredantialsFailureCode))
                        {
                            result = "User credentials failed";
                            //TODO :we need to improve the error field hight binding command
                            CredentialsHighlight = "red";
                            DatabaseNameHighlight = "white";
                            ServerNameHighlight = "white";
                            //END
                        }
                        else if (result.Equals(Utility.DatabaseNameFailureCode))
                        {
                            result = "Database name is not valid";
                            //TODO :we need to improve the error field hight binding command
                            DatabaseNameHighlight = "red";
                            CredentialsHighlight = "white";
                            ServerNameHighlight = "white";
                            //END
                        }
                        else if (result.Equals(Utility.ServerNameFailureCode))
                        {
                            result = "Server name is not valid";
                            //TODO :we need to improve the error field hight binding command
                            ServerNameHighlight = "red";
                            CredentialsHighlight = "white";
                            DatabaseNameHighlight = "white";
                            //END
                        }
                        else
                        {
                            //TODO :we need to improve the error field hight binding command
                            CredentialsHighlight = "white";
                            DatabaseNameHighlight = "white";
                            ServerNameHighlight = "white";
                            //END
                        }
                    }
                    else
                    {
                        if (Authentication.Equals("SQL Server Authentication"))
                        {

                            if (string.IsNullOrWhiteSpace(LogIn) || string.IsNullOrWhiteSpace(Password))
                            {
                                CredentialsHighlight = "red";
                                result = "SQL Server Credentials is mandatory.";
                            }
                            else
                            {
                                CredentialsHighlight = "red";
                            }

                        }
                        if (string.IsNullOrWhiteSpace(ServerName))
                        {
                            ServerNameHighlight = "red";
                            result += " Server Name is mandatory.";
                        }
                        else
                        {
                            ServerNameHighlight = "white";
                        }

                        if (string.IsNullOrWhiteSpace(DatabaseName))
                        {
                            DatabaseNameHighlight = "red";
                            result += " Database Name is mandatory.";
                        }
                        else
                        {
                            DatabaseNameHighlight = "white";
                        }
                    }

                    EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Publish(result);
                }
                catch (Exception ex)
                {
                    EventAgg._eventAggregator.GetEvent<ErrorMessageEvent>().Publish(ex.Message);
                }
            }
        }

        /// <summary>
        ///Populate staging db connection string by logged in user 
        /// </summary>
        private void PopulateStagingDbConfig()
        {
            //TODO: Logged in User id must pass the arguement.
            int userId = 1;
            try
            {
                using (DataTable stagingDt = aplBusinessL.GetStagingDBConfigByUser(userId))
                {
                    if (stagingDt.Rows.Count > 0)
                    {
                        ServerName = stagingDt.Rows[0]["ServerName"].ToString();
                        ServerType = stagingDt.Rows[0]["ServerType"].ToString();
                        Authentication = stagingDt.Rows[0]["Authentication"].ToString();
                        LogIn = stagingDt.Rows[0]["UserName"].ToString();
                        Password = stagingDt.Rows[0]["Password"].ToString();
                        DatabaseName = stagingDt.Rows[0]["DatabaseName"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                EventAgg._eventAggregator.GetEvent<ErrorMessageEvent>().Publish(ex.Message);
            }
        }

        #endregion
    }
}
