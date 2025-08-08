# Status Report Converter

WPF application for converting HTML status reports to Word documents with customizable content.

## Features

- Convert HTML status reports to Word format (.docx)
- Edit current week status and next week goals
- Manage project risks with add/edit/delete functionality
- Automatic table header repetition for multi-page tables
- Page numbering (Page X of Y format)
- Professional headers with "Enterprise AI Status Report"
- Graph/chart support from HTML

## Prerequisites

- .NET 8.0 SDK or later
- Windows 10/11
- Visual Studio 2022 or VS Code with C# extension

## Build and Run

### Using Visual Studio
1. Open `StatusReportConverter.sln`
2. Press F5 to build and run

### Using Command Line
```bash
# Navigate to project directory
cd "C:\Development Projects\Cook Street Client Pages\Weekly Status Reports"

# Restore NuGet packages
dotnet restore

# Build the project
dotnet build

# Run the application
dotnet run --project StatusReportConverter
```

## Configuration

The application uses environment variables configured in `StatusReportConverter/.env`:
- `LOG_LEVEL`: Logging level (Default: Information)
- `LOG_PATH`: Path for log files (Default: ./logs/)
- `LICENSE_PATH`: Path to Aspose license file

## Usage

1. Launch the application
2. Click "Browse..." to select your HTML status report
3. Click "Browse..." to specify the output Word document location
4. (Optional) Edit current week status in the text box
5. (Optional) Edit next week goals in the text box
6. (Optional) Add/edit/delete risks in the risk management table
7. Click "Convert" to generate the Word document

## License Requirements

This application uses Aspose.Words for .NET. Place your `Aspose.TotalProductFamily.lic` file in the root directory for full functionality. Without a license, the application runs in evaluation mode with watermarks.

## Project Structure

```
StatusReportConverter/
├── Constants/          # Application constants
├── Commands/          # WPF command implementations
├── Models/            # Data models
├── Services/          # Business logic services
├── Utils/             # Helper utilities
├── ViewModels/        # MVVM ViewModels
├── Views/             # WPF Views (XAML)
└── Resources/         # Styles and resources
```

## Security Features

- Input validation and sanitization
- Path traversal protection
- XSS prevention
- Secure logging with Serilog
- No hardcoded credentials

## Repository

https://github.com/cstrickc/StatusReportConverter