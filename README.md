office2011
==========

A shell script to (download, if necessary) and install Office:Mac 2011 and any known updates without any user intervention.

[1]:	http://officecdn.microsoft.com/pr/MacOffice2011/en-us/MicrosoftOffice2011.dmg

[2]:	https://download.microsoft.com/download/2/9/C/29C00BAC-97AA-4126-A10E-98EE37B9CD37/Office2011-1461Update_EN-US.dmg



## Intended Usage

Simply put, this script is intended to quickly let you download and install the latest version of Microsoft Office 2011.

This is most often useful for when you are setting up a new Mac.

If you are setting up _more_ than one Mac, you can run this script on *one* Mac, and then copy the `~/Downloads/Office2011/` folder from one computer to another after the installation is finished. That way you will not need to re-download the files, which should save time and bandwidth.

Note that *all* downloads come from Microsoft’s official servers. You are not downloading them off GitHub, or my website, or any other random site. These are official distributions.

(_Nerdy Technical Security Details_: The script verifies the checksum of each download using `shasum -a 256` to verify that it has not been tampered with or changed. If it does not match, installation will immediately stop.)

## There are a few provisos, a couple of *quid pro quos*… ##

1. This script is not officially endorsed suggested, recommended, or any such words to that effect by Microsoft or any Microsoft affiliated business, entity, person, or persons. (In fact I'm fairly sure Microsoft does not even know that I am alive.)

2.  All trademarks are used solely for identification.

3. This script has been tested with me (and others) but comes with absolutely no guarantees. If it works, hurrah! If it breaks, you get to keep both pieces. **Anything that happens as a result of you running this will be your responsibility.** Not mine, certainly not Microsoft's, etc.

4. This script will ***not*** provide you with a free copy of Microsoft Office 2011. All it will do is download the necessary pieces and install them. You will still have to enter your Product Key within 30 days of installation.

## What it does ##

The `office2011.sh`  script will look for certain DMG files in a particular folder in your hard drive (by default it will use `~/Downloads/Office2011/` but it will use another folder if you change the `DIR="$HOME/Downloads/Office2011"` line near the top of the script.

If you want to simply download and install them yourself, here are the direct download links used by the script:

1. [MicrosoftOffice2011.dmg][1]

2. [Latest Office Microsoft Office for Mac 2011 Update][2]

The script is intended to be smart enough to check whether minimum requirements are met, and also to check to make sure that it does not install something which has already been installed.

## Q: “Does using this script give me _any_ different result than if I had installed Office manually using OS X’s Installer.app?” ##

A. Yes.

The only known side effect of using this script is that the Microsoft Office apps are not
automatically added your Dock, as they are if you use the GUI installer.

This is considered a feature.

### GUI Installer ###

For people who are not comfortable with the Terminal, a GUI installer is available.

However, because the installer application is an Automator application and not officially signed, you will see this warning if/when you try to launch it:

![OS X Gatekeeper Warning](images/01-Install-Office-2011-Cant-Be-Opened.png)

You _could_ disable Gatekeeper in OS X’s settings, but I consider that to be a **very bad idea,** because _exceptions_ can be made very easily using a _one time per app_ process, which I will explain here:.

Simply selecting the application that you want to open and then control-click (or right-click) to bring up this contextual menu:

![Control-Click context menu](images/02-Install-Office-2011-Control-Click-Open.png)

Choose “Open” and it will ask you to confirm your choice:

![Gatekeeper wants to make sure you’re sure](images/03-Install-Office-2011-Is-From-Unidentified-Developer.png)

Once you choose "Open" from the above the menu, the application will run, and you will not have to authorize it again on that computer.

The **“Install Office 2011”** application will then download the latest version of [office2011.sh][11] to `~/Downloads/office2011.sh` and run it to download and install Office 2011 files, if necessary. The downloads will be saved to **~/Downloads/Office2011/** and that folder can be moved to other Macs (as long as it is put at **~/Downloads/Office2011/**) and the installer can be re-run without having to download the DMGs.

Because you will need to enter your administrator password, it will prompt you (last time!) before it begins:

![Last check before we start for real.](images/04-Install-Office-2011-Will-Download-And-Install.png)

Finally you will see the usual authentication window asking for your name and password:

![Make sure to change “Admin” to your actual administrator username.](images/05-osascript-wants-your-password.png)

Once the script is downloaded and begins to run, a log file will be created in **~/Library/Logs/** (the standard OS X location for log files).

That log file will be opened in the standard **/Applications/Utilities/Console.app** application so you can watch the installation’s progress. (However, I recommend going away to do something fun instead, and just check back later.)

