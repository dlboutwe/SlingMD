using System;
using System.IO;
using System.Text;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class FileService
    {
        private readonly ObsidianSettings _settings;

        public FileService(ObsidianSettings settings)
        {
            _settings = settings;
        }

        public ObsidianSettings GetSettings()
        {
            return _settings;
        }

        public void WriteUtf8File(string filePath, string content)
        {
            // Ensure directory exists
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Write file with UTF-8 encoding without BOM
            using (var writer = new StreamWriter(filePath, false, new UTF8Encoding(false)))
            {
                writer.Write(content);
            }
        }

        public string CleanFileName(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            // Replace invalid characters with underscore
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string cleanName = string.Join("_", input.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));

            // Replace additional problematic characters
            cleanName = cleanName.Replace("\"", "")
                               .Replace("'", "")
                               .Replace("`", "")
                               .Replace(":", "")
                               .Replace(";", "")
                               .Trim();

            return cleanName;
        }

        public void LaunchObsidian(string vaultName, string filePath)
        {
            try
            {
                // Replace backslashes with forward slashes for Obsidian URLs
                string normalizedPath = filePath.Replace('\\', '/');
                
                // Remove file extension if present
                if (normalizedPath.EndsWith(".md"))
                {
                    normalizedPath = normalizedPath.Substring(0, normalizedPath.Length - 3);
                }
                
                // Create and launch the Obsidian URL
                string obsidianUrl = $"obsidian://open?vault={Uri.EscapeDataString(vaultName)}&file={Uri.EscapeDataString(normalizedPath)}";
                System.Diagnostics.Process.Start(obsidianUrl);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Failed to launch Obsidian: {ex.Message}", "SlingMD", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        public bool EnsureDirectoryExists(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public string GetInboxPath() => _settings.GetInboxPath();
    }
} 