# Outlook File Drag

*Drag and drop Outlook items as files into any application*

## Overview

Outlook File Drag is a plugin for Outlook 2013 and 2016 that allows you to drag
and drop Outlook items (messages, attachments, etc) to applications that allow 
physical files to be dropped, such as web browsers.

## How Does it Work?

When you try to drag and drop from Outlook, Outlook correctly identifies the 
format as virtual files (CFSTR_FILEDESCRIPTORW), which many applications do 
not support, such as web browsers.

To work around this issue, Outlook File Drag hooks the Outlook drag and drop
process and adds support for physical files (CF_HDROP).  When the application 
asks for the physical files, the files are saved to a temp folder.

## Installation

To install, run the installer that matches your Windows build.  The 64-bit 
installer supports both 32-bit and 64-bit Outlook running on 64-bit Windows.
The 32-bit installer supports 32-bit Outlook running on 32-bit Windows.

## Known Issues

Outlook File Drag does not currently allow dragging and dropping items as 
files in calendar view, as this was causing Outlook to crash.  Hopefully this
will be fixed in a future release.

## Acknowledgements

Outlook File Drag uses these open source projects:

- [Easyhook](https://easyhook.github.io/)

## Version History

### 1.0
- Initial Release
