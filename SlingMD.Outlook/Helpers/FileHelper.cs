using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Helpers
{
    public static class FileHelper
    {
        public static string CleanFileName(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            // Replace invalid characters with a dash
            string invalidChars = Regex.Escape(new string(Path.GetInvalidFileNameChars()));
            string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

            return Regex.Replace(input, invalidRegStr, "-").Trim();
        }

        public static void WriteUtf8File(string filePath, string content)
        {
            // Ensure the directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));

            // Write the file with UTF-8 encoding
            using (var stream = new StreamWriter(filePath, false, new UTF8Encoding(false)))
            {
                stream.Write(content);
            }
        }

        public static void LaunchObsidian(string vaultName, string filePath)
        {
            string encodedPath = Uri.EscapeDataString(filePath);
            string command = $"obsidian://open?vault={vaultName}&file={encodedPath}";
            
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = command,
                UseShellExecute = true
            });
        }
    }
} 