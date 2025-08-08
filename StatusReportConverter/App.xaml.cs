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

            Env.Load();

            ConfigureLogging();
            ConfigureServices();

            var mainWindow = serviceProvider!.GetRequiredService<MainWindow>();
            mainWindow.Show();
        }

        private void ConfigureLogging()
        {
            var logPath = Environment.GetEnvironmentVariable("LOG_PATH") ?? "./logs/";
            var logLevel = Environment.GetEnvironmentVariable("LOG_LEVEL") ?? "Information";
            
            Directory.CreateDirectory(logPath);

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Is(ParseLogLevel(logLevel))
                .WriteTo.File(Path.Combine(logPath, "statusreport-.log"), 
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7)
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

            services.AddSingleton<ILogger<MainViewModel>>(provider =>
            {
                var factory = new LoggerFactory().AddSerilog();
                return factory.CreateLogger<MainViewModel>();
            });

            services.AddSingleton<IDocumentConverterService, DocumentConverterService>();
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