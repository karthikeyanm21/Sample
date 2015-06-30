using System;
using System.Collections.Generic;
using System.Linq;
using ReactiveUI;
using System.Windows.Input;

namespace APLPX.UI.Main.ViewModels
{
    public class ViewModelBase : ReactiveObject,IDisposable
    {
        private bool m_isDisposed;
        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool isDisposing)
        {
            if (!m_isDisposed)
            {
                if (isDisposing)
                {
                }
                m_isDisposed = true;
            }
        }

        ~ViewModelBase()
        {
            Dispose(false);
        }

        #endregion
    }

    public class RelayCommand : ICommand
    {
        public Predicate<object> CanExecuteDelegate { get; set; }
        public Action<object> ExecuteDelegate { get; set; }

        public bool CanExecute(object parameter)
        {
            if (CanExecuteDelegate != null)
                return CanExecuteDelegate(parameter);
            return true; // if there is no can execute default to true
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void Execute(object parameter)
        {
            if (ExecuteDelegate != null)
                ExecuteDelegate(parameter);
        }
    }
}
