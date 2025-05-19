using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Handles loading, rendering and construction of markdown templates. The service tries multiple
    /// probe paths to locate a named template allowing users and developers to override built-in
    /// placeholders with custom files.
    /// </summary>
    public class TemplateService
    {
        private readonly FileService _fileService;

        public TemplateService(FileService fileService)
        {
            _fileService = fileService;
        }

        /// <summary>
        /// Attempts to locate <paramref name="templateName"/> in a range of common folders (deployment
        /// directory, executing assembly location, current working directory). The first hit is returned
        /// as raw text.
        /// </summary>
        /// <returns>Template contents or <c>null</c> when the file could not be found.</returns>
        public string LoadTemplate(string templateName)
        {
            // Try multiple locations for the template file
            string[] possibleTemplatePaths = new[]
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", templateName),
                Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Templates", templateName),
                Path.Combine(Directory.GetCurrentDirectory(), "Templates", templateName),
                Path.Combine(Environment.CurrentDirectory, "Templates", templateName)
            };

            foreach (var path in possibleTemplatePaths)
            {
                if (File.Exists(path))
                {
                    return File.ReadAllText(path);
                }
            }

            return null;
        }

        /// <summary>
        /// Naïve string replacement renderer that swaps out <c>{{key}}</c> placeholders with the values
        /// supplied in <paramref name="replacements"/>.
        /// </summary>
        public string ProcessTemplate(string templateContent, Dictionary<string, string> replacements)
        {
            if (string.IsNullOrEmpty(templateContent)) return string.Empty;

            string result = templateContent;
            foreach (var replacement in replacements)
            {
                result = result.Replace($"{{{{{replacement.Key}}}}}", replacement.Value);
            }
            return result;
        }

        /// <summary>
        /// Produces a YAML front-matter block from the supplied dictionary. Lists are automatically
        /// serialised as YAML arrays and <see cref="DateTime"/> values use the <c>yyyy-MM-dd HH:mm</c>
        /// format.
        /// </summary>
        public virtual string BuildFrontMatter(Dictionary<string, object> metadata)
        {
            var frontMatter = new StringBuilder();
            frontMatter.AppendLine("---");

            foreach (var item in metadata)
            {
                if (item.Value == null) continue;

                // Always render 'tags' as a YAML list if it's a list/enumerable
                if (item.Key == "tags" && item.Value is IEnumerable<string> tagList)
                {
                    if (tagList.Any())
                    {
                        frontMatter.AppendLine($"{item.Key}: ");
                        foreach (var value in tagList)
                        {
                            frontMatter.AppendLine($"  - \"{value}\"");
                        }
                    }
                    else
                    {
                        frontMatter.AppendLine($"{item.Key}: []");
                    }
                }
                else if (item.Value is string strValue)
                {
                    frontMatter.AppendLine($"{item.Key}: \"{strValue}\"");
                }
                else if (item.Value is DateTime dtValue)
                {
                    frontMatter.AppendLine($"{item.Key}: {dtValue:yyyy-MM-dd HH:mm}");
                }
                else if (item.Value is IEnumerable<string> listValue)
                {
                    if (listValue.Count() > 0)
                    {
                        frontMatter.AppendLine($"{item.Key}: ");
                        foreach (var value in listValue)
                        {
                            frontMatter.AppendLine($"  - \"{value}\"");
                        }
                    }
                    else
                    {
                        frontMatter.AppendLine($"{item.Key}: []");
                    }
                }
                else
                {
                    frontMatter.AppendLine($"{item.Key}: {item.Value}");
                }
            }

            frontMatter.AppendLine("---");
            frontMatter.AppendLine();

            return frontMatter.ToString();
        }

        /// <summary>
        /// Fallback template for thread summary notes when the user has not provided their own version.
        /// Contains a DataviewJS script that summarises all emails sharing the same <c>threadId</c>.
        /// </summary>
        public string GetDefaultThreadNoteTemplate()
        {
            return @"---
title: ""{{title}}""
type: email-thread
threadId: ""{{threadId}}""
tags: [email-thread]
---

# {{title}}

```dataviewjs
// Get all emails with matching threadId from current folder
const threadId = ""{{threadId}}"";
const emails = dv.pages("""")
    .where(p => p.threadId === threadId && p.file.name !== dv.current().file.name)
    .sort(p => p.date, 'desc');

// Display thread summary
if (emails.length > 0) {
    const startDate = emails[emails.length-1].date;
    const latestDate = emails[0].date;
    const participants = new Set();
    emails.forEach(e => {
        // Handle from field
        if (e.from) {
            const fromName = String(e.from).match(/\[\[(.*?)\]\]/)?.[1];
            if (fromName) participants.add(fromName);
        }

        // Handle to field
        if (e.to) {
            const toList = Array.isArray(e.to) ? e.to : [e.to];
            toList.forEach(to => {
                const name = String(to).match(/\[\[(.*?)\]\]/)?.[1];
                if (name) participants.add(name);
            });
        }

        // Handle cc field
        if (e.cc) {
            const ccList = Array.isArray(e.cc) ? e.cc : [e.cc];
            ccList.forEach(cc => {
                const name = String(cc).match(/\[\[(.*?)\]\]/)?.[1];
                if (name) participants.add(name);
            });
        }
    });

    dv.header(2, 'Thread Summary');
    dv.list([
        `Started: ${startDate}`,
        `Latest: ${latestDate}`,
        `Messages: ${emails.length}`,
        `Participants: ${Array.from(participants).map(p => `[[${p}]]`).join(', ')}`
    ]);
}

// Display email timeline
dv.header(2, 'Email Timeline');
for (const email of emails) {
    dv.header(3, `${email.file.name} - ${email.date}`);
    dv.paragraph(`![[${email.file.name}]]`);
}
```";
        }
    }
} 