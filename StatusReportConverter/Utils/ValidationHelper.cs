using System;
using System.IO;
using System.Text.RegularExpressions;

namespace StatusReportConverter.Utils
{
    public static class ValidationHelper
    {
        private static readonly Regex HtmlTagRegex = new Regex(@"<script[^>]*>.*?</script>|<iframe[^>]*>.*?</iframe>", 
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        
        private static readonly Regex PathTraversalRegex = new Regex(@"\.\.[\\/]");
        
        public static bool ValidateHtmlFile(string filePath, out string errorMessage)
        {
            errorMessage = string.Empty;
            
            if (string.IsNullOrWhiteSpace(filePath))
            {
                errorMessage = "File path cannot be empty";
                return false;
            }
            
            if (PathTraversalRegex.IsMatch(filePath))
            {
                errorMessage = "Invalid file path detected";
                return false;
            }
            
            if (!File.Exists(filePath))
            {
                errorMessage = "File does not exist";
                return false;
            }
            
            var extension = Path.GetExtension(filePath).ToLower();
            if (extension != ".html" && extension != ".htm")
            {
                errorMessage = "File must be an HTML file";
                return false;
            }
            
            try
            {
                var content = File.ReadAllText(filePath);
                if (HtmlTagRegex.IsMatch(content))
                {
                    errorMessage = "HTML file contains potentially unsafe content";
                    return false;
                }
            }
            catch (Exception ex)
            {
                errorMessage = $"Error reading file: {ex.Message}";
                return false;
            }
            
            return true;
        }
        
        public static bool ValidateOutputPath(string filePath, out string errorMessage)
        {
            errorMessage = string.Empty;
            
            if (string.IsNullOrWhiteSpace(filePath))
            {
                errorMessage = "Output path cannot be empty";
                return false;
            }
            
            if (PathTraversalRegex.IsMatch(filePath))
            {
                errorMessage = "Invalid output path detected";
                return false;
            }
            
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                try
                {
                    Directory.CreateDirectory(directory);
                }
                catch (Exception ex)
                {
                    errorMessage = $"Cannot create output directory: {ex.Message}";
                    return false;
                }
            }
            
            return true;
        }
        
        public static string SanitizeText(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return string.Empty;
            
            input = Regex.Replace(input, @"<[^>]*>", string.Empty);
            input = input.Replace("&", "&amp;")
                        .Replace("<", "&lt;")
                        .Replace(">", "&gt;")
                        .Replace("\"", "&quot;")
                        .Replace("'", "&#39;");
            
            return input.Trim();
        }
    }
}