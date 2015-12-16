#!/bin/zsh -f
# Install MSOffice for Mac and any/all available 'patches'.
#
# From:	Timothy J. Luoma
# Mail:	luomat at gmail dot com
# Date:	2013-01-31
# Last Updated: 2015-12-16 for Office 14.5.9


PATH=/usr/bin:/usr/sbin:/sbin:/bin:/usr/local/bin


	# This is the only thing you should have to change.
	# 	Where are the installation files located,
	# 	or to where do you want them downloaded?
	# 	(The script will attempt to create this
	# 	folder if it does not already exist.)

DIR="$HOME/Downloads/Office2011"



	# Set INDEX='yes' if you want to make an index of files on your hard drive before/after each install
INDEX='no'

####|####|####|####|####|####|####|####|####|####|####|####|####|####|####
#
#		A Note For Convenience
#
#
#	If you want to be able to run this script without needing to enter
#	your administrator (sudo) password multiple times, I recommend adding
#	this line to /etc/sudoers:
#
#		%admin ALL=NOPASSWD: /usr/sbin/installer
#
# 	Alternatively you can call this script under 'sudo' but that's risky
# 	for reasons you ought to understand before doing it. (Notably: if I
# 	made any mistakes here and they're running as 'sudo' they might be
# 	able to (unintentionally) do a lot more damage



####|####|####|####|####|####|####|####|####|####|####|####|####|####|####
#
#
#	You should not _have_ to change anything below here.
#
#	You are, of course, encouraged to look and see what I'm doing, and why
#	(which should be explained in the comments).
#
#
####|####|####|####|####|####|####|####|####|####|####|####|####|####|####

umask 022

NAME="$0:t:r"

zmodload zsh/datetime

TIME=$(strftime "%Y-%m-%d--%H.%M.%S" "$EPOCHSECONDS")

HOST=`hostname -s`

HOST="$HOST:l"

LOG="$HOME/Library/Logs/$NAME/$HOST.$TIME.log"

	[[ ! -d "$LOG:h" ]]	&& mkdir -p "$LOG:h"	# create the parent directory where the LOG will be stored, iff it doesn't exist
	[[ ! -e "$LOG" ]] 	&& touch "$LOG"			# 'touch' the logfile, iff it doesn't exist

chmod a+r "$LOG"

zmodload zsh/datetime # initialize a time module so we an avoid calling `date`

function timestamp { strftime "%Y-%m-%d @ %H:%M:%S" "$EPOCHSECONDS" }

function msg { echo "$NAME [`timestamp`]: $@" | tee -a "$LOG" }

function version { echo "$@" | awk -F. '{ printf("%d%03d%03d%03d\n", $1,$2,$3,$4); }'; }

msg "Starting"

open -F -n -a Console "$LOG"

	# This is the file that we'll look for to see a) if Office is
	# installed, and b) if so, what version.
	#
	# Note that this script is not designed to handle installations to
	# non-standard locations, so if you have installed Office 2011
	# somewhere else, expect this to break.
MSWORD_INFO_PLIST='/Applications/Microsoft Office 2011/Microsoft Word.app/Contents/Info.plist'

	# this is the directory where Mac OS X stores 'receipts' from pkgs which have already been installed.
	# You should not change this.
RECEIPTS_DIR='/private/var/db/receipts'





	# This function is meant to mount a DMG (the filename must be given as
	# an argument to this function) and return the mount point (i.e.
	# /Volumes/Whatever) if successful. If the DMG fails to mount, it will
	# return nothing but an error.
function mntdmg {

	DMG_TO_MOUNT="$@"

	if [ -r "$DMG_TO_MOUNT" -a -f "$DMG_TO_MOUNT" ]
	then

		MNTPNT=$(hdiutil attach -nobrowse -plist "$DMG_TO_MOUNT" 2>/dev/null \
				| fgrep -A 1 '<key>mount-point</key>' \
				| tail -1 \
				| sed 's#</string>.*##g ; s#.*<string>##g')

		[[ "$MNTPNT" = "" ]] && return 1

		echo "${MNTPNT}"

	else

		return 1

	fi
}


### We are 'verifying' DMGs against the expected size in bytes and the shasum
function verify_dmg {



	# this function requires an argument: the filename of a DMG you wish to verify
DMG_TO_VERIFY="$@"

if [[ "$DL_URL" == "" ]]
then
		msg "Cannot verify $DMG_TO_VERIFY, URL is empty."
		return 1

elif [[ "$SUM" == "" ]]
then
		msg "Cannot verify $DMG_TO_VERIFY, SUM is empty."
		return 1

elif [[ "$BYTES" == "" ]]
then
		msg "Cannot verify $DMG_TO_VERIFY, BYTES is empty."
		return 1
else

		zmodload zsh/stat

		FOUND_BYTES=$(zstat -L +size "$DMG_TO_VERIFY")

		if [[ "$FOUND_BYTES" == "$BYTES" ]]
		then
				msg "$DMG_TO_VERIFY has the correct byte size"
		else
				# size is wrong, but is it too small or too large?

				if [ "$FOUND_BYTES" -lt "$BYTES" ]
				then
						msg "$DMG_TO_VERIFY is smaller than expected. An attempt will be made to continue downloading the DMG from $DL_URL..."

						# File is smaller than expected, try continuing download
						curl \
							--progress-bar \
							--location --max-time 3600 --fail \
							--output "${DMG_TO_VERIFY}" --continue-at - \
							--user-agent "curl/7.21.7 (x86_64-apple-darwin10.8.0) libcurl/7.21.7 OpenSSL/1.0.0d zlib/1.2.5 libidn/1.22" \
							--referer ";auto" \
							"$DL_URL" 2>&1 | tee -a "$LOG" || return 1


				else
						# If we got here, then we have a file which is larger than it should be!

						msg "$DMG_TO_VERIFY was expected to be $BYTES bytes but is $FOUND_BYTES bytes instead! I recommend that you move $DMG_TO_VERIFY to the trash and try again."

						return 1


				fi
		fi
fi


if [ -e "$DMG_TO_VERIFY" ]
then
		# check the SUM

		msg "Verifying shasum 256 of $DMG_TO_VERIFY. This may take a few moments, please be patient:"

		ACTUAL_SUM=$(shasum -a 256 -p "$DMG_TO_VERIFY" | awk '{print $1}')

		if [ "$SUM" = "$ACTUAL_SUM" ]
		then
				msg "$DMG_TO_VERIFY checksum verified"
		else
				msg "$DMG_TO_VERIFY checksum FAILED: is not what was expected: $ACTUAL_SUM (actual)\n$SUM (expected)"
				return 1
		fi

fi

}


####|####|####|####|####|####|####|####|####|####|####|####|####|####|####
#
#		Setup and verify our working environment, make sure things are like we want them
#

	# Create the directory where we will be keeping the downloads, but only if it doesn't exist already
[[ -d "$DIR" ]] || mkdir -p "$DIR"

if [[ ! -d "$DIR" ]]
then
		msg "Failed to create DIR at: $DIR"
		exit 1
fi

	# make sure everyone can read and execute (open) the directory
chmod a+rx "$DIR"

	# Now we need to chdir into the directory where the DMGs are, or will be stored.
cd "$DIR"

if [ "$PWD" != "$DIR" ]
then
		msg "failed to chdir to $DIR (we are in $PWD). Giving up."
		exit 1
fi


####|####|####|####|####|####|####|####|####|####|####|####|####|####|####
#
#	THIS IS WHERE WE DEFINE WHICH DMGs WE ARE GOING TO TRY TO INSTALL!
#
#	Note that 'MicrosoftOffice2011.dmg' *MUST* be first.
#	This script will attempt to install them in the order they are listed.
#

for DMG in \
		MicrosoftOffice2011.dmg \
		Office2011-1455Update_EN-US.dmg
do

MIN_VERSION=''
THIS_IS_OFFICE_INSTALLER='no'

if [[ "${DMG}" = "launchword" ]]
then
		# This is a special case. If the user puts 'launchword' into the
		# list of DMGs (in the "for DMG in " list) then MSWord will launch
		# when it gets to that entry in the 'for' loop.

		# At some point in the past, Office would not install one of its
		# updates until at least one of the apps had been run. As of
		# 2013-07-22 that no longer seems to be true, but I'm keeping it
		# here for two reasons: 1) in case it is needed again and 2) in
		# case the user wants to make `launchword` the last of the "DMGs"
		# which will let the user say "Run this installer and launch Word
		# when all of the updates have been installed.

	if [ "$INDEX" = "yes" ]
	then

		msg "making index before Launching Word, please wait..."
		sudo /usr/bin/find -x / -type f -print > filelist.launchword.before.`timestamp`.txt
	fi

		# just in case the user has not been paying attention, we'll play a
		# sound three times to try and draw their attention now.
	SND='/System/Library/Sounds/Glass.aiff'

	tput bel
	tput bel
	[[ -e "$SND" ]] && afplay "$SND"
	[[ -e "$SND" ]] && afplay "$SND"

	msg "You must quit Microsoft Word before this script will continue."

	open -W -a 'Microsoft Word'


	if [ "$INDEX" = "yes" ]
	then

		msg "making index AFTER launching Word"
		sudo /usr/bin/find -x / -type f -print > filelist.launchword.after.`timestamp`.txt
	fi

else # if NOT 'launchword'

	# In this part of the 'for' loop we will configure a set of variables
	# based on the filename of the DMG that the user has indicated they
	# would like to install.
	#
	#	FULL_NAME 	= a 'friendly' name for the DMG so that user knows what it is. This will be used in status messages
	#
	#	MORE_INFO 	= an URL to read about the update. This URL should link to the download link
	#
	#	RECEIPT   	= this is the unique filename associated with this DMG.
	#		We will look in /private/var/db/receipts/ for files which match
	#		"$RECEIPT" plus either ".bom" or ".plist" as an indication that
	#		this DMG/pkg has already been installed. This allows us to re-run
	#		this script multiple times without worrying that we will try to
	#		install something which has already been installed, or install
	#		something old over something new.
	#
	#	DL_URL 		= This is the actual download URL for the DMG. The end of this URL should match the DMG listed in the 'for' loop.
	#
	#	BYTES		= This is the exact number of bytes that we expect to have when we finish downloading the DMG from DL_URL.
	#
	#	SUM			= This is the 'shasum' of the DMG, using the SHA 256 algorithm. Probably overkill (we could probably use shasum -a 1)
	#
	#	Note that that the script will not both computing the shasum unless the BYTES match exactly, since it would be guaranteed not to match.
	#
	#	MIN_VERSION	= (optional) This is the minimum version of Office which has to be installed before we can try to install this DMG/pkg
	#
	#	THIS_IS_OFFICE_INSTALLER = 	if this DMG contains the necessary files to do
	#								the initial installation of Microsoft Office,
	#								then set this to yes. Otherwise, do not set it.



case "${DMG}" in
	MicrosoftOffice2011.dmg)
		FULL_NAME='Office Installer'
		MORE_INFO=''
		RECEIPT='com.microsoft.office.all.slt_std.pkg.14.5.2'
		DL_URL='http://officecdn.microsoft.com/pr/MacOffice2011/en-us/MicrosoftOffice2011.dmg'
		BYTES='0b8ecf514e3afef8b2ec3ed05af8ccf9c1a7574108ef4f27202f1b17bdf15f22'
		SUM='1013640929'
		THIS_IS_OFFICE_INSTALLER='yes'
	;;

	Office2011-1459Update_EN-US.dmg)
		FULL_NAME='Microsoft Office for Mac 2011 14.5.9 Update'
		MORE_INFO='http://www.microsoft.com/en-us/download/details.aspx?id=50361'
		RECEIPT='com.microsoft.office.all.core.pkg.14.5.9.update'
		DL_URL='https://download.microsoft.com/download/4/3/B/43B9AC64-FC80-414F-ABB8-8876339C828A/Office2011-1459Update_EN-US.dmg'
		BYTES='118883795'
		SUM='9741ba1b7bfc3784b095bfaa6b475d927a037adfbd0fce75696b04ae47fd3d9c'
		MIN_VERSION='14.1.0'
	;;


esac
# End URL / SUM case/esac
#########################

####|####|####|####|####|####|####|####|####|####|####|####|####|####|####
#
#		This is where we start to do the actual processing of the DMGs listed in the 'for' loop above
#

msg "Starting $DMG"

if [ -e "$MSWORD_INFO_PLIST" ]
then
			# if we found Microsoft Word installed where we would expect it to be, set INSTALLED='yes'
			# and then fetch the 'CFBundleShortVersionString' from MSWord's 'Info.plist' file

		INSTALLED='yes'
		INSTALLED_VERSION=$(fgrep -A1 'CFBundleShortVersionString' "$MSWORD_INFO_PLIST" | tr -dc '[0-9].')

		if [ "$THIS_IS_OFFICE_INSTALLER" = "yes" ]
		then
					# if this DMG is the primary Office installer, but Office is already installed, skip it
				msg "cannot install $DMG because it is the Office Installer and Office is already installed!"

				continue
		fi

else
			# if MSWord is NOT installed, set the variables accordingly

		INSTALLED='no'
		INSTALLED_VERSION=""

		if [ "$THIS_IS_OFFICE_INSTALLER" != "yes" ]
		then
					# If the THIS_IS_OFFICE_INSTALLER has NOT been set to
					# 'yes' (i.e. we are not working with the primary
					# Office Installer) and Office 2011 is not installed,
					# then we have to skip this DMG

				msg "Cannot install $DMG until Office 2011 base installation is complete."
				continue
		fi

fi # if MSWORD_INFO_PLIST exists





	####|####|####|####|####|####|####|####|####|####|####|####|####|####|####
	#
	#		Check to make sure that we have not already installed this DMG/pkg
	#
if [ -s "${RECEIPTS_DIR}/${RECEIPT}.plist" -a  -s "${RECEIPTS_DIR}/${RECEIPT}.bom" ]
then
		msg "Skipping $DMG because it is already installed ${RECEIPTS_DIR}."

		continue
fi

	####|####|####|####|####|####|####|####|####|####|####|####|####|####|####
	#
	#		Check to make sure we meet the minimum requirements
	#






if [ "$MIN_VERSION" = "" ]
then
		msg "No MIN_VERSION set for $DMG"
else
		if [ $(version ${INSTALLED_VERSION}) -ge $(version ${MIN_VERSION}) ]
		then
				msg "Installed version ($INSTALLED_VERSION) meets minimum required version ($MIN_VERSION) for $DMG"
		else
				msg "Cannot install $DMG because installed version of Office ($INSTALLED_VERSION) does not meet minimum requirement of $MIN_VERSION"
				continue
		fi

fi
# Minimum version check


####|####|####|####|####|####|####|####|####|####|####|####|####|####|####
#
#		If we have gotten this far, we are ready to install the DMG. Do we need to download it or does it already exist?
#

if [[ ! -e "${DMG}" ]]
then

		msg "Fetching $DMG"

		CURL_HEADERS="$DL_URL:t:r.log"

		curl --progress-bar --location --fail --dump-header "$CURL_HEADERS" --remote-name "$DL_URL"  2>&1 | tr -d '\r' | tee -a "$LOG"

		fgrep -q 'HTTP/1.1 200 OK' "$CURL_HEADERS"

		if [ "$?" != "0" ]
		then
				msg "Did not get STATUS 200 when we attempted to download $DL_URL. Giving up on $DMG."
				continue
		fi

		msg "Downloaded: $DMG"

fi

	# If we get here and the DMG doesn't exist, something went wrong, so let's skip this one and try the next one
if [[ ! -e "${DMG}" ]]
then

		msg "FAILED to find $DMG in $PWD. Giving up and going on to next DMG."

		continue
fi

####|####|####|####|####|####|####|####|####|####|####|####|####|####|####
#
#		If we get here then the DMG exists, so now let's verify that it meets our expectations
#

msg "Verifying $DMG in $PWD..."

verify_dmg "$DMG"

if [ "$?" = "1" ]
then

		msg "verify_dmg returned 1, giving up on $DMG"
		continue
fi

msg "Installing ${DMG} at `timestamp`"

	# mntdmg defined above
MNT=$(mntdmg "${DMG}")

if [[ "$MNT" = "" ]]
then

		msg "MNT is empty for ${DMG}"
		continue

else

			# Note: If you are the only admin user of your Mac, I
			# recommend putting 'installer' into /etc/sudoers like so:
			#
			#		%admin ALL=NOPASSWD: /usr/sbin/installer
			#
			# otherwise you may be prompted to re-enter your admin password for each DMG
			#


		if [ "$INDEX" = "yes" ]
		then

			msg "Making index before installing $DMG, please wait."
				/usr/bin/find -x / -type f -print 2>/dev/null > "filelist.$DMG.before.`timestamp`.txt"
		fi

		msg "Installing $MNT:t"

		sudo installer -verbose -pkg ${MNT}/*pkg -target / 2>&1 | tee -a "$LOG"


		if [ "$INDEX" = "yes" ]
		then

			msg "Making index AFTER installing $DMG, please wait."
				/usr/bin/find -x / -type f -print 2>/dev/null > "filelist.$DMG.after.`timestamp`.txt"
		fi


		msg "${DMG} installed successfully at `timestamp`" | tee -a "$LOG"

			# We'll try, at least once, to unmount the DMG. It's not that
			# big of a deal if this fails, but it's nice to clean up if we can.
		diskutil unmount "$MNT"

fi # MNT = ""

fi # if/else 'launchword'

done # for / do loop



if (( $+commands[terminal-notifier] ))
then

			terminal-notifier \
				-sender com.microsoft.word \
				-message "Click to reveal in Finder" \
				-title "Office 2011 Installed" \
				-open "file:///Applications/Microsoft Office 2011"
else
			# IF terminal-notifier is not installed
			# BUT Growl.app is running
			# AND `growlnotify` exists,
			# THEN show a Growl notification
		ps cx | fgrep -q Growl && test -x `which growlnotify` && \
			growlnotify --sticky \
				--appIcon "Microsoft Word"  \
				--identifier "$NAME"  \
				--message "Switch to Finder to see Office apps folder."  \
				--title "Office 2011 Installed"
fi

	# The -g prevents Finder from "stealing focus"
	# but the window will be there when they switch to Finder
open -g -a Finder "/Applications/Microsoft Office 2011"

	# Run the updater to make sure we haven't missed anything
open -g -a "/Library/Application Support/Microsoft/MAU2.0/Microsoft AutoUpdate.app"

exit 0

#
#EOF

# n.b. - a side effect of using this script is that the Microsoft Office apps are not
# forcibly inserted into your Dock, as they would be if you used the GUI installer.
# This is considered a feature.
