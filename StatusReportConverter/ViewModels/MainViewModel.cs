using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;
using StatusReportConverter.Commands;
using StatusReportConverter.Models;
using StatusReportConverter.Services;

namespace StatusReportConverter.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly IDocumentConverterService converterService;
        private readonly ILogger<MainViewModel> logger;
        private StatusReport statusReport;
        private bool isConverting;
        private string statusMessage;

        public StatusReport StatusReport
        {
            get => statusReport;
            set
            {
                statusReport = value;
                OnPropertyChanged();
            }
        }

        public bool IsConverting
        {
            get => isConverting;
            set
            {
                isConverting = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanConvert));
            }
        }

        public string StatusMessage
        {
            get => statusMessage;
            set
            {
                statusMessage = value;
                OnPropertyChanged();
            }
        }

        public bool CanConvert => !IsConverting && 
            !string.IsNullOrWhiteSpace(StatusReport.InputHtmlPath) && 
            !string.IsNullOrWhiteSpace(StatusReport.OutputWordPath);

        public ICommand BrowseInputCommand { get; }
        public ICommand BrowseOutputCommand { get; }
        public ICommand ConvertCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand AddRiskCommand { get; }
        public ICommand DeleteRiskCommand { get; }
        public ICommand SubmitCurrentWeekCommand { get; }
        public ICommand SubmitNextWeekCommand { get; }

        public MainViewModel(IDocumentConverterService converterService, ILogger<MainViewModel> logger)
        {
            this.converterService = converterService;
            this.logger = logger;
            
            statusReport = new StatusReport();
            statusMessage = "Ready";
            isConverting = false;

            BrowseInputCommand = new RelayCommand(BrowseInput);
            BrowseOutputCommand = new RelayCommand(BrowseOutput);
            ConvertCommand = new AsyncRelayCommand(ConvertAsync, () => CanConvert);
            CancelCommand = new RelayCommand(Cancel);
            AddRiskCommand = new RelayCommand(AddRisk);
            DeleteRiskCommand = new RelayCommand<Risk>(DeleteRisk);
            SubmitCurrentWeekCommand = new RelayCommand(SubmitCurrentWeek);
            SubmitNextWeekCommand = new RelayCommand(SubmitNextWeek);

            logger.LogInformation("MainViewModel initialized");
            
            if (!converterService.ValidateLicense())
            {
                StatusMessage = "Warning: Running in evaluation mode";
            }
        }

        private void BrowseInput()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select HTML Status Report",
                Filter = "HTML Files (*.html;*.htm)|*.html;*.htm|All Files (*.*)|*.*",
                InitialDirectory = AppDomain.CurrentDomain.BaseDirectory
            };

            if (dialog.ShowDialog() == true)
            {
                StatusReport.InputHtmlPath = dialog.FileName;
                OnPropertyChanged(nameof(StatusReport));
                OnPropertyChanged(nameof(CanConvert));
                logger.LogInformation("Input file selected: {Path}", dialog.FileName);
            }
        }

        private void BrowseOutput()
        {
            var dialog = new SaveFileDialog
            {
                Title = "Save Word Document As",
                Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*",
                DefaultExt = "docx",
                InitialDirectory = AppDomain.CurrentDomain.BaseDirectory,
                FileName = $"StatusReport_{DateTime.Now:yyyyMMdd}.docx"
            };

            if (dialog.ShowDialog() == true)
            {
                StatusReport.OutputWordPath = dialog.FileName;
                OnPropertyChanged(nameof(StatusReport));
                OnPropertyChanged(nameof(CanConvert));
                logger.LogInformation("Output file selected: {Path}", dialog.FileName);
            }
        }

        private async void ConvertAsync()
        {
            try
            {
                IsConverting = true;
                StatusMessage = "Converting...";
                logger.LogInformation("Starting conversion");

                var success = await converterService.ConvertHtmlToWordAsync(StatusReport);

                if (success)
                {
                    StatusMessage = "Conversion completed successfully!";
                    logger.LogInformation("Conversion completed successfully");
                }
                else
                {
                    StatusMessage = "Conversion failed. Check logs for details.";
                    logger.LogError("Conversion failed");
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
                logger.LogError(ex, "Conversion error");
            }
            finally
            {
                IsConverting = false;
            }
        }

        private void Cancel()
        {
            logger.LogInformation("User cancelled operation");
            StatusMessage = "Operation cancelled";
        }

        private void AddRisk()
        {
            var newRisk = new Risk();
            StatusReport.Risks.Add(newRisk);
            logger.LogInformation("Added new risk: {Id}", newRisk.Id);
        }

        private void DeleteRisk(Risk? risk)
        {
            if (risk != null && StatusReport.Risks.Contains(risk))
            {
                StatusReport.Risks.Remove(risk);
                logger.LogInformation("Deleted risk: {Id}", risk.Id);
            }
        }

        private void SubmitCurrentWeek()
        {
            if (!string.IsNullOrWhiteSpace(StatusReport.CurrentWeekStatus))
            {
                logger.LogInformation("Current week status updated");
                StatusMessage = "Current week status updated";
            }
        }

        private void SubmitNextWeek()
        {
            if (!string.IsNullOrWhiteSpace(StatusReport.NextWeekGoals))
            {
                logger.LogInformation("Next week goals updated");
                StatusMessage = "Next week goals updated";
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}