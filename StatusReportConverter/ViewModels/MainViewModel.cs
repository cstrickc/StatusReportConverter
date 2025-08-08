using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;
using StatusReportConverter.Commands;
using StatusReportConverter.Constants;
using StatusReportConverter.Models;
using StatusReportConverter.Services;

namespace StatusReportConverter.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly IDocumentConverterService converterService;
        private readonly IHtmlParserService htmlParserService;
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

        public ICommand BrowseInputCommand { get; private set; }
        public ICommand BrowseOutputCommand { get; private set; }
        public ICommand ConvertCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }
        public ICommand AddRiskCommand { get; private set; }
        public ICommand DeleteRiskCommand { get; private set; }
        public ICommand SubmitCurrentWeekCommand { get; private set; }
        public ICommand SubmitNextWeekCommand { get; private set; }

        public MainViewModel(IDocumentConverterService converterService, IHtmlParserService htmlParserService, ILogger<MainViewModel> logger)
        {
            this.converterService = converterService;
            this.htmlParserService = htmlParserService;
            this.logger = logger;
            
            statusReport = new StatusReport();
            statusMessage = AppConstants.StatusMessages.READY;
            isConverting = false;

            InitializeCommands();
            
            logger.LogInformation("MainViewModel initialized");
            
            if (!converterService.ValidateLicense())
            {
                StatusMessage = AppConstants.StatusMessages.EVALUATION_MODE;
            }
        }

        private void InitializeCommands()
        {
            BrowseInputCommand = new RelayCommand(BrowseInput);
            BrowseOutputCommand = new RelayCommand(BrowseOutput);
            ConvertCommand = new AsyncRelayCommand(ConvertAsync, () => CanConvert);
            CancelCommand = new RelayCommand(Cancel);
            AddRiskCommand = new RelayCommand(AddRisk);
            DeleteRiskCommand = new RelayCommand<Risk>(DeleteRisk);
            SubmitCurrentWeekCommand = new RelayCommand(SubmitCurrentWeek);
            SubmitNextWeekCommand = new RelayCommand(SubmitNextWeek);
        }

        private void BrowseInput()
        {
            var dialog = new OpenFileDialog
            {
                Title = AppConstants.FileDialogs.INPUT_DIALOG_TITLE,
                Filter = AppConstants.FileDialogs.HTML_FILTER,
                InitialDirectory = AppDomain.CurrentDomain.BaseDirectory
            };

            if (dialog.ShowDialog() == true)
            {
                var extractedReport = htmlParserService.ExtractContentFromHtml(dialog.FileName);
                
                StatusReport.InputHtmlPath = dialog.FileName;
                StatusReport.CurrentWeekStatus = extractedReport.CurrentWeekStatus;
                StatusReport.NextWeekGoals = extractedReport.NextWeekGoals;
                
                StatusReport.Risks.Clear();
                foreach (var risk in extractedReport.Risks)
                {
                    StatusReport.Risks.Add(risk);
                }
                
                OnPropertyChanged(nameof(StatusReport));
                OnPropertyChanged(nameof(CanConvert));
                logger.LogInformation("Input file selected and content extracted: {Path}", dialog.FileName);
                StatusMessage = "HTML content loaded and parsed successfully";
            }
        }

        private void BrowseOutput()
        {
            var dialog = new SaveFileDialog
            {
                Title = AppConstants.FileDialogs.OUTPUT_DIALOG_TITLE,
                Filter = AppConstants.FileDialogs.WORD_FILTER,
                DefaultExt = AppConstants.FileDialogs.DEFAULT_EXTENSION,
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
                StatusMessage = AppConstants.StatusMessages.CONVERTING;
                logger.LogInformation("Starting conversion");

                var success = await converterService.ConvertHtmlToWordAsync(StatusReport);

                StatusMessage = success ? 
                    AppConstants.StatusMessages.SUCCESS : 
                    AppConstants.StatusMessages.FAILED;
                    
                logger.LogInformation("Conversion {Status}", success ? "succeeded" : "failed");
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
            StatusMessage = AppConstants.StatusMessages.CANCELLED;
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
                StatusMessage = AppConstants.StatusMessages.CURRENT_WEEK_UPDATED;
            }
        }

        private void SubmitNextWeek()
        {
            if (!string.IsNullOrWhiteSpace(StatusReport.NextWeekGoals))
            {
                logger.LogInformation("Next week goals updated");
                StatusMessage = AppConstants.StatusMessages.NEXT_WEEK_UPDATED;
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}