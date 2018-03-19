# Outlook File Drag

*Drag and drop Outlook items as files into any application*

## Overview

Outlook File Drag is an add-in for Outlook 2013 and 2016 that allows you to drag
and drop Outlook items (messages, attachments, etc) to applications that allow 
physical files to be dropped, such as web browsers.

## How Does it Work?

When you try to drag and drop from Outlook, Outlook correctly identifies the 
format as virtual files (CFSTR_FILEDESCRIPTORW) since the files do not exist 
directly on disk.  Instead, they are contained in a PST file, OST file, or on 
an Exchange server.

However, many applications do not support, such as web browers and most .NET/
Java applications.

To work around this issue, Outlook File Drag hooks the Outlook drag and drop
process and adds support for physical files (CF_HDROP).  When the application 
asks for the physical files, the files are saved to a temp folder.

## Installation

To install, run the installer that matches your Windows build:

- [Download for 64-bit Windows (Outlook 32-bit or 64-bit)](https://github.com/tonyfederer/OutlookFileDrag/files/1823357/OutlookFileDragSetup_x64.zip)
- [Download for 32-bit Windows](https://github.com/tonyfederer/OutlookFileDrag/files/1823356/OutlookFileDragSetup.zip)

After installing, restart Outlook for the add-in to take effect.

## Known Issues

Outlook File Drag does not currently allow dragging and dropping items as 
files in calendar view, as this was causing Outlook to crash.  Hopefully this
will be fixed in a future release.

## Acknowledgements

Outlook File Drag uses these open source projects:

- [Easyhook](https://easyhook.github.io/)
- [log4net](http://logging.apache.org/log4net/)

## Feedback/Contribute

You can view the source code, report issues, and contribute on [Github](https://github.com/tonyfederer/OutlookFileDrag).

## Version History

### [1.0](https://github.com/tonyfederer/OutlookFileDrag/releases/tag/v1.0)
- Initial Release

## Copyright

Outlook File Drag is copyright (c) 2018 by [Tony Federer](https://github.com/tonyfederer) and released under the MIT License.
