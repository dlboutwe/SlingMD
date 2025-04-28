# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
SlingMD is a .NET Framework 4.7.2 C# Outlook add-in that enables exporting emails to Obsidian as markdown.

## Build Commands
- Build: `dotnet build SlingMD.sln --configuration Release`
- Publish: `dotnet publish SlingMD.Outlook\SlingMD.Outlook.csproj --configuration Release`
- Package release: `.\package-release.ps1`

## Code Style Guidelines
- **Naming**: PascalCase for classes/methods/properties, camelCase for variables/parameters, _camelCase for private fields
- **Structure**: Use service-oriented architecture with services in the Services/ folder
- **Types**: Use explicit typing over var, utilize strong typing
- **Imports**: System namespaces first, followed by third-party, then project-specific
- **Formatting**: Braces on new lines, 4-space indentation
- **Error Handling**: Use specific try-catch blocks, display errors via MessageBox.Show()
- **Interfaces**: Start with 'I' prefix
- **Services**: Use 'Service' suffix

## Organization
Maintain existing folder structure (Forms/, Helpers/, Models/, Services/, etc.)