using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

using MahApps.Metro;
using ReactiveUI;
using APLPX.Modules.StagingDBConfig.ViewModels;
using APLPX.Common;
using Microsoft.Practices.Prism.PubSubEvents;
using Microsoft.Practices.Prism.Mvvm;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System.Threading.Tasks;


namespace APLPX.UI.Main.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        public List<AccentColorMenuData> AccentColors { get; private set; }
        public List<AppThemeMenuData> AppThemes { get; private set; }


        /// <summary>
        ///Property variable for visibility changes 
        /// </summary>
        private bool m_isFeatureListVisible;

        /// <summary>
        /// Property for Planning visibility changes 
        /// </summary>
        private bool m_isPlanningFeatureVisible;

        /// <summary>
        /// Property for Staging form visibility changes 
        /// </summary>
        private bool m_isStagingDBFetureModuleVisible;

        /// <summary>
        ///  Property for Import Menu visbility changes 
        /// </summary>
        private bool m_isImportMenufeatureVisible;

        /// <summary>
        ///Property for Status bar messages 
        /// </summary>
        private string _currentStatusBarText;



        #region Constructor

        public MainViewModel()
        {
            InitializeCommand();

            // create accent color menu items for the demo
            this.AccentColors = ThemeManager.Accents
                                            .Select(a => new AccentColorMenuData() { Name = a.Name, ColorBrush = a.Resources["AccentColorBrush"] as Brush })
                                            .ToList();

            // create metro theme color menu items for the demo
            this.AppThemes = ThemeManager.AppThemes
                                           .Select(a => new AppThemeMenuData() { Name = a.Name, BorderColorBrush = a.Resources["BlackColorBrush"] as Brush, ColorBrush = a.Resources["WhiteColorBrush"] as Brush })
                                           .ToList();
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Initialize actions command for main model.
        /// </summary>
        private void InitializeCommand()
        {
            //Initialize Admin module fetaures visibility command.
            AdminModuleCmd = ReactiveCommand.Create();
            AdminModuleCmd.Subscribe(x => AdminModuleCommandExecuted(x));

            //
            PlaningModuleCmd = ReactiveCommand.Create();
            PlaningModuleCmd.Subscribe(x => PlanningModuleCommandExecuted(x));

            //Initialize selected feature's  command.
            StagingDBFormCmd = ReactiveCommand.Create();
            StagingDBFormCmd.Subscribe(x => StagingDBFormCommandExecuted(x));


            ImportMenuCmd = ReactiveCommand.Create();
            ImportMenuCmd.Subscribe(x => ImportMenuVisibleCommandExecuted(x));


            //Set up event listeners for pending/completed operations.
           EventAgg._eventAggregator.GetEvent<StatusBarEvent>().Subscribe(args => StatusBarUpdate(args));

           //Set up event listeners for pending/completed operations.
           EventAgg._eventAggregator.GetEvent<ErrorMessageEvent>().Subscribe(args => ErrorMsgEvent(args));

        }


        /// <summary>
        ///Set visibility to admin feature list.
        /// </summary>
        /// <param name="sender"></param>
        private void AdminModuleCommandExecuted(object sender)
        {
            IsPlanningFeatureListVisibile = false;
            IsAdminFeatureListVisibile = true;
        }

        /// <summary>
        ///Set visibility to Planning feature list.
        /// </summary>
        /// <param name="sender"></param>
        private void PlanningModuleCommandExecuted(object sender)
        {
            IsAdminFeatureListVisibile = false;
            IsPlanningFeatureListVisibile = true;
        }

        /// <summary>
        ///Status bar message updates 
        /// </summary>
        /// <param name="args"></param>
        private void StatusBarUpdate(string args)
        {
            CurrentStatusBarMessage = args;
        }

        /// <summary>
        ///Error Message event  
        /// </summary>
        /// <param name="errorMsg"></param>
        private void ErrorMsgEvent(string errorMsg)
        {
            var mainview = System.Windows.Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
            mainview.ShowMessageAsync("Error", errorMsg);
        }

        /// <summary>
        ///Set visibility to StagingDB feature form.
        /// </summary>
        /// <param name="sender"></param>
        private void StagingDBFormCommandExecuted(object sender)
        {
            IsStagingDBFeatureModuleVisible = true;
            IsImportMenuFeatureVisible = false;
        }

        /// <summary>
        /// Set visibility to Import menu feature form.
        /// </summary>
        /// <param name="sender"></param>
        private void ImportMenuVisibleCommandExecuted(object sender)
        {
            try
            {
                bool result = APLPX.Modules.StagingDBConfig.ViewModels.StagingDbViewModel.StagingDBCount();
                if (result)
                {
                    IsStagingDBFeatureModuleVisible = false;
                    IsImportMenuFeatureVisible = true;
                }
                else
                {
                    var mainview0 = System.Windows.Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                    System.Threading.Tasks.Task<MessageDialogResult> result1 = mainview0.ShowMessageAsync("Staging DB Setting", "Please configure staging DB before process the file.", MessageDialogStyle.Affirmative);
                    IsPlanningFeatureListVisibile = false;
                    IsAdminFeatureListVisibile = true;
                    IsStagingDBFeatureModuleVisible = true;
                    IsImportMenuFeatureVisible = false;
                }
            }
            catch (Exception ex)
            {
                EventAgg._eventAggregator.GetEvent<ErrorMessageEvent>().Publish(ex.Message);
            }
        }
     
        #endregion

        #region commands

        /// <summary>
        /// Command to display the Admin module features.
        /// </summary>
        public ReactiveCommand<object> AdminModuleCmd { get; private set; }

        /// <summary>
        /// Command to display the planning module features.
        /// </summary>
        public ReactiveCommand<object> PlaningModuleCmd { get; private set; }

        /// <summary>
        /// Command to display the StagingDB features.
        /// </summary>
        public ReactiveCommand<object> StagingDBFormCmd { get; private set; }


        /// <summary>
        /// Command to Import menu from  button.
        /// </summary>
        public ReactiveCommand<object> ImportMenuCmd { get; private set; }

        #endregion

        #region Properties

        /// <summary>
        /// Gets/sets the current status text.
        /// </summary>
        public string CurrentStatusBarMessage
        {
            get { return _currentStatusBarText; }
            set { this.RaiseAndSetIfChanged(ref _currentStatusBarText, value); }
        }

        /// <summary>
        /// Property  to visible Admin feature list 
        /// </summary>
        public bool IsAdminFeatureListVisibile
        {
            get
            {
                return m_isFeatureListVisible;
            }
            private set
            {
                this.RaiseAndSetIfChanged(ref m_isFeatureListVisible, value);
            }
        }

        /// <summary>
        /// Property  to visible planning feature list 
        /// </summary>
        public bool IsPlanningFeatureListVisibile
        {
            get
            {
                return m_isPlanningFeatureVisible;
            }
            private set
            {
                this.RaiseAndSetIfChanged(ref m_isPlanningFeatureVisible, value);
            }
        }

        /// <summary>
        /// Property for StagingDBFeatureModule
        /// </summary>
        public bool IsStagingDBFeatureModuleVisible
        {
            get
            {
                return m_isStagingDBFetureModuleVisible;
            }
            private set
            {
                this.RaiseAndSetIfChanged(ref m_isStagingDBFetureModuleVisible, value);
            }
        }


        /// <summary>
        ///Propery for get and set planning module feature visibilty 
        /// </summary>
        public bool IsImportMenuFeatureVisible
        {
            get
            {
                return m_isImportMenufeatureVisible;
            }
            private set
            {
                this.RaiseAndSetIfChanged(ref m_isImportMenufeatureVisible, value);
            }
        }

        #endregion
    }


    /// <summary>
    ///MahApps Metro theme and accent color menu style appling class 
    /// </summary>
    public class AccentColorMenuData
    {
        public string Name { get; set; }
        public Brush BorderColorBrush { get; set; }
        public Brush ColorBrush { get; set; }

        private ICommand changeAccentCommand;

        public ICommand ChangeAccentCommand
        {
            get { return this.changeAccentCommand ?? (changeAccentCommand = new RelayCommand { CanExecuteDelegate = x => true, ExecuteDelegate = x => this.DoChangeTheme(x) }); }
        }

        protected virtual void DoChangeTheme(object sender)
        {
            var theme = ThemeManager.DetectAppStyle(Application.Current);
            var accent = ThemeManager.GetAccent(this.Name);
            ThemeManager.ChangeAppStyle(Application.Current, accent, theme.Item1);
        }
    }

    public class AppThemeMenuData : AccentColorMenuData
    {
        protected override void DoChangeTheme(object sender)
        {
            var theme = ThemeManager.DetectAppStyle(Application.Current);
            var appTheme = ThemeManager.GetAppTheme(this.Name);
            ThemeManager.ChangeAppStyle(Application.Current, theme.Item2, appTheme);
        }
    }
}
