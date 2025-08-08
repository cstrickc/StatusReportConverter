using System;
using System.IO;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;
using StatusReportConverter.Services;
using StatusReportConverter.ViewModels;
using StatusReportConverter.Views;
using DotNetEnv;

namespace StatusReportConverter
{
    public partial class App : Application
    {
        private ServiceProvider? serviceProvider;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            try
            {
                var envPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".env");
                if (File.Exists(envPath))
                {
                    Env.Load(envPath);
                }
                else
                {
                    System.Windows.MessageBox.Show($".env file not found at: {envPath}", "Configuration Warning");
                }

                ConfigureLogging();
                ConfigureServices();

                var mainWindow = serviceProvider!.GetRequiredService<MainWindow>();
                mainWindow.Show();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Application startup error: {ex.Message}\n\n{ex.StackTrace}", "Error");
                Shutdown();
            }
        }

        private void ConfigureLogging()
        {
            var logPath = Environment.GetEnvironmentVariable(Constants.AppConstants.Environment.LOG_PATH) 
                ?? Constants.AppConstants.Environment.DEFAULT_LOG_PATH;
            var logLevel = Environment.GetEnvironmentVariable(Constants.AppConstants.Environment.LOG_LEVEL) 
                ?? Constants.AppConstants.Environment.DEFAULT_LOG_LEVEL;
            
            Directory.CreateDirectory(logPath);

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Is(ParseLogLevel(logLevel))
                .WriteTo.File(Path.Combine(logPath, Constants.AppConstants.Logging.LOG_FILE_PATTERN), 
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: Constants.AppConstants.Logging.RETAINED_FILE_COUNT)
                .CreateLogger();
        }

        private Serilog.Events.LogEventLevel ParseLogLevel(string level)
        {
            return level.ToLower() switch
            {
                "verbose" => Serilog.Events.LogEventLevel.Verbose,
                "debug" => Serilog.Events.LogEventLevel.Debug,
                "information" => Serilog.Events.LogEventLevel.Information,
                "warning" => Serilog.Events.LogEventLevel.Warning,
                "error" => Serilog.Events.LogEventLevel.Error,
                "fatal" => Serilog.Events.LogEventLevel.Fatal,
                _ => Serilog.Events.LogEventLevel.Information
            };
        }

        private void ConfigureServices()
        {
            var services = new ServiceCollection();

            services.AddSingleton<ILogger<IDocumentConverterService>>(provider =>
            {
                var factory = new LoggerFactory().AddSerilog();
                return factory.CreateLogger<IDocumentConverterService>();
            });
            
            services.AddSingleton<ILogger<IHtmlParserService>>(provider =>
            {
                var factory = new LoggerFactory().AddSerilog();
                return factory.CreateLogger<IHtmlParserService>();
            });
            
            services.AddSingleton<ILogger<IChartService>>(provider =>
            {
                var factory = new LoggerFactory().AddSerilog();
                return factory.CreateLogger<IChartService>();
            });

            services.AddSingleton<ILogger<MainViewModel>>(provider =>
            {
                var factory = new LoggerFactory().AddSerilog();
                return factory.CreateLogger<MainViewModel>();
            });

            services.AddSingleton<IDocumentConverterService, DocumentConverterService>();
            services.AddSingleton<IHtmlParserService, HtmlParserService>();
            services.AddSingleton<IChartService, ChartService>();
            services.AddSingleton<MainViewModel>();
            services.AddSingleton<MainWindow>();

            serviceProvider = services.BuildServiceProvider();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Log.CloseAndFlush();
            serviceProvider?.Dispose();
            base.OnExit(e);
        }
    }
}