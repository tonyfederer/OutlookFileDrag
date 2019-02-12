# Outlook File Drag

*Drag and drop Outlook items as files into any application*

## Overview

Outlook File Drag is an add-in for Outlook 2013 and 2016 that allows you to drag
and drop Outlook items (messages, attachments, contacts, tasks, appointments, 
meetings, etc) to applications that allow physical files to be dropped, such as
web browsers.

## How Does it Work?

When you try to drag and drop from Outlook, Outlook correctly identifies the 
format as virtual files (CFSTR_FILEDESCRIPTORW) since the files do not exist 
directly on disk.  Instead, they are contained in a PST file, OST file, or on 
an Exchange server.

However, many applications do not support this format, such as web browers and 
most .NET/Java applications.

To work around this issue, Outlook File Drag hooks the Outlook drag and drop
process and adds support for physical files (CF_HDROP).  When the receiving 
application asks for the physical files, the files are saved to a temp folder 
and those filenames are returned to the application.  The application processes
the files (such as uploading them).  Outlook File Drag deletes the temp files 
later in a cleanup process.

## Features

- Works with Chrome, Firefox, Internet Explorer, Edge, and other applications that accept files to be dropped
- Allows drag and drop into HTML5-based web applications
- Drag e-mails, attachments, contacts, calendar items, and more
- Drag multiple items at once
- Supports Unicode characters

## Installation

To install, run the installer that matches your Windows build:

- [Download for 64-bit Windows (Outlook 32-bit or 64-bit)](https://github.com/tonyfederer/OutlookFileDrag/releases/download/v1.0.10/OutlookFileDragSetup_x64.zip)
- [Download for 32-bit Windows](https://github.com/tonyfederer/OutlookFileDrag/releases/download/v1.0.10/OutlookFileDragSetup.zip)

After installing, restart Outlook for the add-in to take effect.

## Automated (Silent) Installation

For administrators, OutlookFileDrag supports automated (silent) installation and uninstallation using `msiexec` with command line parameters.

### Silent Installation

To silently install OutlookFileDrag, use this command:

`msiexec.exe /i <pathtomsi> /qn /log <pathtolog>`

- `<pathtomsi>`: Path to MSI file
- `<pathtolog>`: Path to log file (if folder is not specified, MSI path is used)

Example: 

`msiexec.exe C:\Install\OutlookFileDrag_x64.msi /qn /log C:\Logs\OutlookFileDragInstall.log`

After installing, restart Outlook for the add-in to take effect.

### Silent Uninstallation

To silently uninstall OutlookFileDrag, use this command:

`msiexec.exe /x <productcode> /qn /log <pathtolog>`

- `<productcode>` for 64-bit version: `{CF5F9043-967C-400D-B6D5-F41AF6AD83AE}`
- `<productcode>` for 32-bit version: `{7EA6E17B-8802-4E1F-9669-248670B31BFB}`
- `<logfile>`: Path to log file

Example:

`msiexec.exe /x {CF5F9043-967C-400D-B6D5-F41AF6AD83AE} /qn /log C:\Logs\OutlookFileDragUninstall.log`

## Acknowledgements

Outlook File Drag uses these open source projects:

- [Easyhook](https://easyhook.github.io/)
- [log4net](http://logging.apache.org/log4net/)

## Feedback/Contribute

You can view the source code, report issues, and contribute on [Github](https://github.com/tonyfederer/OutlookFileDrag).

## Donate

If you find this project useful, please consider donating.  Your donations are appreciated. =)

[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=BSAGCF5VAJLN2)

## Version History

### 1.0.10
- Fixed System.ArgumentException bug in ReadHGlobalIntoStream method when reading more than 4 KB introduced in version 1.0.8.

### 1.0.9
- If files were dropped and drop effect was "move", then override to "copy" so original item is not deleted

### 1.0.8
- Fixed releasing of unmanaged resources 
- Memory usage improvements
- Added more details to log file

### 1.0.5
- Fixed crash when dragging calendar items

### 1.0.4
- Added additional debug logging
- Fixed issue where STGMEDIUM was not being released after reading filenames
- Fixed issue that where reading filenames sometimes failed
- Fixed hooking process to allow starting and stopping hook without disposing and recreating hook

### 1.0.3
- Fixed issue that prevented dragging items from one group to another

### 1.0.2
- Fixed PathTooLong exception when temporary filename was longer than MAX_PATH

### 1.0.1
- Fixed issues with 64-bit Outlook
- Added self-signed certificate

### 1.0
- Initial Release

## Copyright

Outlook File Drag is copyright (c) 2018 by [Tony Federer](https://github.com/tonyfederer) and released under the MIT License.
