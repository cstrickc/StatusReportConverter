using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace StatusReportConverter.Commands
{
    public class AsyncRelayCommand : ICommand
    {
        private readonly Func<Task> execute;
        private readonly Func<bool>? canExecute;
        private bool isExecuting;

        public AsyncRelayCommand(Action execute, Func<bool>? canExecute = null)
            : this(() => Task.Run(execute), canExecute)
        {
        }

        public AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null)
        {
            this.execute = execute ?? throw new ArgumentNullException(nameof(execute));
            this.canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object? parameter)
        {
            return !isExecuting && (canExecute?.Invoke() ?? true);
        }

        public async void Execute(object? parameter)
        {
            if (CanExecute(parameter))
            {
                try
                {
                    isExecuting = true;
                    CommandManager.InvalidateRequerySuggested();
                    await execute();
                }
                finally
                {
                    isExecuting = false;
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }
    }
}