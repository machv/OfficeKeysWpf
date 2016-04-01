using System;
using System.Windows.Input;

namespace Mach.Wpf.Mvvm
{
    /// <summary>
    /// An <see cref="ICommand"/> whose delegates can be attached for <see cref="Execute"/> and <see cref="CanExecute"/>.
    /// </summary>
    public class DelegateCommand : ICommand
    {
        private readonly Action _action;
        private readonly ICommandOnCanExecute _canExecute;

        /// <summary>
        /// Delegate when CanExecute is called on the command. This can be null.
        /// </summary>
        /// <param name="parameter">Data used by the command.</param>
        /// <returns></returns>
        public delegate bool ICommandOnCanExecute(object parameter);

        /// <summary>
        /// Initializes a new instance of <see cref="DelegateCommand"/> class.
        /// </summary>
        /// <param name="action">Delegate to execute when Execute is called on the command.</param>
        public DelegateCommand(Action action)
        {
            _action = action;
        }

        /// <summary>
        /// Initializes a new instance of <see cref="DelegateCommand"/> class.
        /// </summary>
        /// <param name="action">Delegate to execute when Execute is called on the command.</param>
        /// <param name="canExecute">Delegate to execute when CanExecute is called on the command. This can be null.</param>
        public DelegateCommand(Action action, ICommandOnCanExecute canExecute)
        {
            _action = action;
            _canExecute = canExecute;
        }

        /// <summary>
        /// Defines the method that determines whether the command can execute in its current state.
        /// </summary>
        /// <param name="parameter">Data used by the command.</param>
        /// <returns>true if this command can be executed; otherwise, false.</returns>
        public bool CanExecute(object parameter)
        {
            if (_canExecute == null)
            {
                return true;
            }

            return _canExecute(parameter);
        }

        /// <summary>
        /// Defines the method to be called when the command is invoked.
        /// </summary>
        /// <param name="parameter">Data used by the command.</param>
        public void Execute(object parameter)
        {
            _action();
        }

#pragma warning disable 67
        /// <summary>
        /// Occurs when changes occur that affect whether or not the command should execute.
        /// </summary>
        public event EventHandler CanExecuteChanged;

        /// <summary>
        /// Fires <see cref="CanExecuteChanged"/> event.
        /// </summary>
        public void RaiseCanExecuteChanged()
        {
            if (CanExecuteChanged != null)
            {
                CanExecuteChanged(this, EventArgs.Empty);
            }
        }
#pragma warning restore 67
    }
}