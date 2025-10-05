# SlingMD

Tools for Use with ObsidianMD - Seamlessly integrate your Outlook emails with Obsidian notes.

![SlingMD Logo](SlingMD_pixel.png)

## Overview

SlingMD is a powerful Outlook add-in that bridges the gap between your email communications and Obsidian notes. It allows you to easily export and manage your emails within your Obsidian knowledge base, helping you maintain a comprehensive personal knowledge management system.

## Features

- Export Outlook emails / appointments directly to Obsidian markdown format
- Preserve email / appointment metadata and formatting
- Create follow-up tasks in Obsidian notes and/or Outlook (email only)
- Seamless integration with Outlook's interface
- Easy-to-use ribbon interface
- Support for attachments and email threading
- Automatic email thread organization
- Thread summary pages with timeline views
- Automatic contact note creation with communication-history Dataview tables
- Customisable note title formatting (placeholders for {Subject}, {Sender}, {Date}) with max-length trimming
- Advanced subject clean-up engine using user-defined regex patterns
- Configurable default tags for notes and tasks
- Duplicate-email protection and safe file-naming, including chronological prefixes for threads
- User-overrideable markdown templates for both email and thread notes

## Installation

1. Go to the [Releases](https://github.com/dlboutwe/SlingMD/releases) for this repository
2. Download the latest version
3. **Important Security Step - Unblock the ZIP File**:
   - Right-click the downloaded ZIP file
   - Click "Properties"
   - At the bottom of the General tab, check the "Unblock" box
   - Click "Apply"
   - If you've already extracted the ZIP, delete the extracted folder first
   - Extract the ZIP file again after unblocking
   
   This step ensures all extracted files are trusted and in the same security zone, preventing potential issues.

4. Run the setup executable to install the Outlook add-in
5. Restart Outlook after installation
6. Enable the Sling Ribbon:
   - In Outlook, click "File" > "Options" > "Customize Ribbon"
   - In the left-hand dropdown menu, select "All Tabs"
   - Find "Sling Tab" in the left column
   - Click "Add >>" to add it to your ribbon
   - Click "OK" to save changes
7. The SlingMD ribbon will now appear in your Outlook interface

## System Requirements

- Microsoft Outlook (Office 365 or 2019+)
- Windows 10 or later
- Obsidian installed on your system

## Usage

1. Open Microsoft Outlook
2. Select any email or calendar appointment you want to save to Obsidian
3. In the Outlook ribbon menu, locate and click the "Sling" button from the Sling Ribbon
4. The selected item will be converted to Markdown format and saved to your configured Obsidian vault
5. If enabled, follow-up tasks will be created in Obsidian and/or Outlook

## Configuration

Before using SlingMD, you'll need to configure your Obsidian vault settings:

1. Click the "Settings" button in the Sling Ribbon
2. Configure the following options:
   - **Vault Settings**
   -    **Vault Name**: Enter the name of your Obsidian vault
   -    **Vault Base Path**: Set the path to your Obsidian vault folder (e.g., C:\Users\YourName\Documents\Notes)
   -    **Inbox Folder**: Specify the folder within your vault where emails should be saved (default: "Inbox")
   -    **Contacts Folder**: Where new contact notes will be stored (default: "Contacts")
   -    **Appointments Folder Path**: Specify the path to the folder withing your vault where emails shoudl be saved (ex: "Journal\Meeting Notes")
   - **General Settings**
   -    **Enable Contact Saving**: Toggle automatic creation of contact notes
   -    **Search Entire Vault For Contacts**: When enabled, SlingMD will look outside the contacts folder before creating a new contact note
   -    **Show countdown**: Toggle whether to show a countdown before launching Obsidian
   -    **Launch Obsidian**: Toggle whether Obsidian should automatically open after saving an email
   - **Timing Settings**
   -    **Delay (seconds)**: Set how long to wait before launching Obsidian (default: 1 second)
   -    **Due in Days**: Set the default number of days until tasks are due (0 = today, 1 = tomorrow, etc.)
   -    **Reminder Days**: Set the default reminder timing:
   -       - In relative mode: How many days before the due date
   -       - In absolute mode: How many days from today
   -    **Reminder Hour**: Set the default hour for task reminders (24-hour format)
   - **Subject Cleanup Patterns**: Configure patterns for cleaning up email subjects (e.g., removing "Re:", "[EXTERNAL]", etc.)
   - **Email Settings Tab**
   -    **Group Email Threads**: Toggle whether to automatically organize related emails into thread folders
   -    **Note Title Format / Max Length / Include Date**: Fine-tune how note titles are constructed
   -    **Move Date To Front In Thread**: When grouping emails, place the date at the beginning of the filename
   -    **Default Note Tags**: Tags automatically assigned to new email notes
   - **Appointment Settings Tab**
   -    **Note Title Format / Max Length**: Fine-tune how note titles are constructed
   -    **Save Attachments**: Toggle whether to save attachments. If attachments are present, this will create a folder with the same name as the markdown file and save all attachments to that folder. Links will be added into the note to launch the files.
   -    **Default Appointment Tags**: Tags automatically assigned to new appointment notes
   - **Task Settings Tab**
   -    **Create task in Outlook**: Toggle whether to create a follow-up task in Outlook
   -    **Create task in Obsidian note**: Toggle whether to create a follow-up task in the Obsidian note
   -    **Ask for dates**: Toggle whether to prompt for dates each time (shows the Task Options form)
   - **Development Settings Tab** (depreceated)
   -    **Show Development Settings**: Reveals additional debug options in the settings dialog
   -    **Show Thread Debug**: Pops up a diagnostic window listing every file that matches a conversationId


4. Click "Save" to apply your settings

Note: Make sure your Vault Base Path points to an existing Obsidian vault directory. If you haven't created a vault yet, please set one up in Obsidian first.

## Task Creation

When task creation is enabled, SlingMD can create follow-up tasks in two locations:

### Obsidian Tasks
- Created at the top of the note
- Uses Obsidian's task format with date metadata
- Links back to the email note
- Tagged with #FollowUp for easy tracking

### Outlook Tasks
- Creates a task in your Outlook task list
- Includes the email subject and content
- Configurable due date (0-30 days from creation)
- Configurable reminder time (0-23 hour)
- Option to prompt for due date and reminder time for each task

## Task Options

When "Ask for dates" is enabled, the Task Options form will appear when creating tasks. This form allows you to:

1. Set the due date using either:
   - Relative mode: Specify number of days from today
   - Absolute mode: Pick a specific date from a calendar

2. Set the reminder using either:
   - Relative mode: Specify number of days before the due date
   - Absolute mode: Pick a specific date from a calendar

3. Set the reminder hour (in 24-hour format)

The "Use Relative Dates" toggle switches between:
- Relative mode: Reminder is set relative to the due date (e.g., "remind me 2 days before it's due")
- Absolute mode: Reminder is set to a specific date (e.g., "remind me next Tuesday")

## Email Threading

When email threading is enabled (via the "Group Email Threads" setting), SlingMD will:

1. Automatically detect related emails using conversation topics and thread IDs
2. Create a dedicated folder for each email thread
3. Generate a thread summary note (0-threadname.md) containing:
   - Thread start and end dates
   - Number of messages
   - List of participants
   - Timeline view of all emails in the thread
4. Move all related emails into the thread folder
5. Update thread summary when new emails are added
6. Link emails to their thread summary for easy navigation

This organization helps keep related emails together and provides a clear overview of email conversations in your vault.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the terms included in the [LICENSE](LICENSE) file.

## Support

If you encounter any issues or have questions, please open an issue in the GitHub repository.

## Changelog

### Version 1.1.0.0
- Added functionality to create notes from calendar items (appointments)
-    Can Save attachments
-    Can bulk add all appointments for the day
- Updated Settings Form

### Version 1.0.0.44
- Added automatic email thread detection and organization
- Added thread summary pages with timeline views
- Added configurable subject cleanup patterns
- Added thread folder creation for related emails
- Added participant tracking in thread summaries
- Added dataview integration for thread visualization
- Improved email relationship detection
- Enhanced thread navigation with bidirectional links
- Fixed various bugs and improved stability

### Version 1.0.0.14
- Added ability to create follow-up tasks in Obsidian notes
- Added ability to create follow-up tasks in Outlook
- Added configurable due dates for tasks
- Added configurable reminder times for tasks
- Added option to prompt for due date and reminder time
- Added task options dialog for custom timing
- Updated settings interface with task configuration options

### Version 1.0.0.8
- Initial release
- Basic email to Obsidian note conversion
- Email metadata preservation
- Obsidian vault configuration
- Launch delay settings

## Disclaimer

THIS SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

The author is not responsible for any data loss, corruption, or other issues that may occur while using this software. Always ensure you have proper backups of your data before using any software that modifies your files.

---

Support the original creator on [Buy Me a Coffee](https://buymeacoffee.com/plainsprepper) 💻🧵🔥

