# SlingMD

Tools for Use with ObsidianMD - Seamlessly integrate your Outlook emails with Obsidian notes.

![SlingMD Logo](SlingMD_pixel.png)

## Overview

SlingMD is a powerful Outlook add-in that bridges the gap between your email communications and Obsidian notes. It allows you to easily export and manage your emails within your Obsidian knowledge base, helping you maintain a comprehensive personal knowledge management system.

## Features

- Export Outlook emails directly to Obsidian markdown format
- Preserve email metadata and formatting
- Seamless integration with Outlook's interface
- Easy-to-use ribbon interface
- Support for attachments and email threading

## Installation

1. Go to the [Releases](./Releases) folder in this repository
2. Download the latest version (currently `SlingMD.Outlook_1_0_0_8.zip`)
3. Extract the ZIP file to a location on your computer
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
2. Select any email you want to save to Obsidian
3. In the Outlook ribbon menu, locate and click the "Sling" button from the Sling Ribbon
4. The email will be converted to Markdown format and saved to your configured Obsidian vault

## Configuration

Before using SlingMD, you'll need to configure your Obsidian vault settings:

1. Click the "Settings" button in the Sling Ribbon
2. Configure the following options:
   - **Vault Name**: Enter the name of your Obsidian vault
   - **Vault Base Path**: Set the path to your Obsidian vault folder (e.g., C:\Users\YourName\Documents\Notes)
   - **Inbox Folder**: Specify the folder within your vault where emails should be saved (default: "Inbox")
   - **Launch Obsidian**: Toggle whether Obsidian should automatically open after saving an email
   - **Delay (seconds)**: Set how long to wait before launching Obsidian (default: 1 second)
   - **Show countdown**: Toggle whether to show a countdown before launching Obsidian

3. Click "Save" to apply your settings

Note: Make sure your Vault Base Path points to an existing Obsidian vault directory. If you haven't created a vault yet, please set one up in Obsidian first.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the terms included in the [LICENSE](LICENSE) file.

## Support

If you encounter any issues or have questions, please open an issue in the GitHub repository.
