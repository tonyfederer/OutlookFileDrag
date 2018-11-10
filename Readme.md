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

- [Download for 64-bit Windows (Outlook 32-bit or 64-bit)](https://github.com/tonyfederer/OutlookFileDrag/files/2564500/OutlookFileDragSetup_x64.zip)
- [Download for 32-bit Windows](https://github.com/tonyfederer/OutlookFileDrag/files/2564499/OutlookFileDragSetup.zip)

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

## Donate

If you find this project useful, please consider donating.  Your donations are appreciated. =)

[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=BSAGCF5VAJLN2)

## Version History

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
