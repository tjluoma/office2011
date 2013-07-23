office2011
==========

A shell script to (download, if necessary) and install Office:Mac 2011 and any known updates without any user intervention.

## Before we begin… ##

1. Let me make it clear that this script is not officially endorsed (suggested, recommended, or any such words like that) by Microsoft. In fact I'm fairly sure Microsoft does not even know that I am alive. All trademarks are used solely for identification.
2. This script has been tested with me but comes with absolutely no guarantees. If it works, hurrah! If it breaks, you get to keep both pieces. Anything that happens as a result of you running this will be your responsibility. Not mine, certainly not Microsoft's, etc.
3. This script will not provide you with a free copy of Microsoft Office 2011 for Mac. All it will do is download the necessary pieces (if necessary) and install them. You will still have to add your Product Key to it manually within 30 days of installation.

## What it does ##

The `office2011.sh`  script will look for certain DMG files in a particular folder in your hard drive (by default it will use ~/Downloads/Office2011/ but it will use another folder if you change the `DIR="$HOME/Downloads/Office2011"` line near the top of the script.

Here are the current (as of 2013-07-23) files which it looks for:

1.	`X18-08827.dmg` - this is the base installer for Microsoft Office 2001 "SP2" (aka version 14.2.0). It must be installed before anything else. You can find out more about it at [heidoc.net][1]. If this file is not found it will be downloaded from [msft.digitalrivercontent.net][], Microsoft's official digital store.

2. `MERP_229.dmg` - This is 'Microsoft Error Reporting for Mac 2.2.9 Update.' If this file is not found it will be [downloaded from Microsoft][2].

3. `	AutoUpdate_236.dmg` - This is 'Microsoft AutoUpdate for Mac 2.3.6 Update'. If this file is not found it will be [downloaded from Microsoft][3].

4. 	`Office2011-1436Update_EN-US.dmg` - This is "Microsoft Office for Mac 2011 14.3.6 Update" . If this file is not found it will be [downloaded from Microsoft][4]. When Microsoft issues a new updater, this will be replaced with that one (assuming that it does not require this as a prerequisite.)

The script is smart enough to check whether minimum requirements are met, and will also check to make sure that it does not install something which has already been installed. It will also checksum the DMGs using `shasum -a 256` to make sure that the files have not been altered or tampered with.

## Intended Usage

Simply put, this script is intended to quickly let you download and install the latest version of Microsoft Office 2011 when setting up a new Mac.

If setting up more than one Mac, simply copy the ~/Downloads/Office2011/ folder from one computer to another after the installation is finished.

## How to break this script ##

1. If you have a version of Microsoft Office installed which is less than 14.2.0 you will need to download and install [Microsoft Office for Mac 2011 Service Pack 2](http://www.microsoft.com/en-us/download/details.aspx?id=29419).

2. There are probably other ways too.

## This script does a basic, standard installation ##

I make a list of all of the files which are created by this script. You can find it at either:

* [all-office-files.mmd][5] for the MultiMarkdown version (3.5 MB)
* [all-office-files.html][6] for the HTML version (3.8 MB)

Note that a few files will differ between installations.

The only known side effect of using this script is that the Microsoft Office apps are not
forcibly inserted into your Dock, as they would be if you used the GUI installer.
This is considered a feature.


<!-- Reference Links -->

[1]: http://www.heidoc.net/joomla/technology-science/microsoft/61-office-2011-for-mac-direct-download-links

[2]: http://www.microsoft.com/en-us/download/details.aspx?id=35382

[3]: http://www.microsoft.com/en-us/download/details.aspx?id=35381

[4]: http://www.microsoft.com/en-us/download/details.aspx?id=39634

[5]: https://raw.github.com/tjluoma/office2011/master/all-office-files.mmd

[6]: https://raw.github.com/tjluoma/office2011/master/all-office-files.html

[msft.digitalrivercontent.net]: http://msft.digitalrivercontent.net/mac/X18-08827.dmg
