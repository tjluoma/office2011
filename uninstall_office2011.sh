#!/usr/bin/env zsh
# Remove ALL files associated with Microsoft Office 2011. BEFORE USING, BE SURE YOU HAVE A VERIFIED BACKUP!
#
# From:	Timothy J. Luoma
# Mail:	luomat at gmail dot com
# Date:	2013-07-23

NAME="$0:t:r"

echo "$NAME: This script will remove ALL FILES associated with Microsoft Office 2011 from this computer, INCLUDING some files from $HOME."
echo -n " If you are ABSOLUTELY SURE you want to do this, type 'yesiamsure' (without the quote marks) and then press enter: "

read ANS

if [ "$ANS" != "yesiamsure" ]
then

	echo "	$NAME: did not receive 'yesiamsure' confirmation. Quitting."
	exit 0

fi

echo -n "	$NAME: OK, will start deleting in 10 seconds. Press Control+C to cancel."

sleep 10

echo "$NAME: ok, here they go:"

# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Arial
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Arial Bold Italic.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Arial Bold.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Arial Italic.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Arial.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Brush Script.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Microsoft/
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Times New Roman Bold Italic.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Times New Roman Bold.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Times New Roman Italic.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Times New Roman.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Verdana Bold Italic.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Verdana Bold.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Verdana Italic.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Verdana.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Wingdings 2.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Wingdings 3.ttf
# Not sure if this is safe to delete, so not deleting it: /Library/Fonts/Wingdings.ttf
# Not sure if this is safe to delete, so not deleting it: /System/Library/Caches/com.apple.Components2.SystemCache.Components
# Not sure if this is safe to delete, so not deleting it: /System/Library/Caches/com.apple.Components2.SystemCache.QuickTimeComponents

for i in \
"${HOME}/Documents/Microsoft User Data/Office 2011 Identities/" \
"${HOME}/Library/Application Support/Microsoft/Office/" \
"${HOME}/Library/Caches/Microsoft/Office/com.microsoft.Word.mcpstrings.1033.plist" \
"${HOME}/Library/FontCollections/Windows Office Compatible.collection" \
"${HOME}/Library/Preferences/ByHost/com.microsoft.registrationDB.4F233473-742B-5A06-91E9-C2AD52C74DEB.plist" \
"${HOME}/Library/Preferences/com.microsoft.autoupdate2.plist" \
"${HOME}/Library/Preferences/com.microsoft.error_reporting.plist" \
"${HOME}/Library/Preferences/com.microsoft.office.plist" \
"${HOME}/Library/Preferences/com.microsoft.office.setupassistant.plist" \
"${HOME}/Library/Preferences/com.microsoft.outlook.databasedaemon.plist" \
"${HOME}/Library/Preferences/com.microsoft.Word.plist" \
"${HOME}/Library/Saved Application State/com.microsoft.Word.savedState/data.data" \
"${HOME}/Library/Saved Application State/com.microsoft.Word.savedState/windows.plist" \
"${HOME}/Library/Saved Application State/com.microsoft.Word.savedState/window_1.data" \
"/Applications/Microsoft Messenger.app" \
"/Applications/Microsoft Office 2011" \
"/Applications/Remote Desktop Connection.app" \
"/Library/Application Support/.jgB9k+J4V3" \
"/Library/Application Support/Microsoft/MAU2.0/Microsoft AutoUpdate.app" \
"/Library/Application Support/Microsoft/MERP2.0/Microsoft Error Reporting.app" \
"/Library/Application Support/Microsoft/MERP2.0/Microsoft Ship Asserts.app" \
"/Library/Automator/Add Attachments to Outlook Messages.action" \
"/Library/Automator/Add Content to Word Documents.action" \
"/Library/Automator/Add Document Properties Page to Word Documents.action" \
"/Library/Automator/Add New Sheet to Workbooks.action" \
"/Library/Automator/Add Table of Contents to Word Documents.action" \
"/Library/Automator/Add Watermark to Word Documents.action" \
"/Library/Automator/Apply Animation to PowerPoint Slide Parts.action" \
"/Library/Automator/Apply Font Format Settings to Word Documents.action" \
"/Library/Automator/AutoFormat Data in Excel Workbooks.action" \
"/Library/Automator/Bring Word Documents to Front.action" \
"/Library/Automator/Close Excel Workbooks.action" \
"/Library/Automator/Close Outlook Items.action" \
"/Library/Automator/Close PowerPoint Presentations.action" \
"/Library/Automator/Close Word Documents.action" \
"/Library/Automator/Combine Excel Files.action" \
"/Library/Automator/Combine PowerPoint Presentations.action" \
"/Library/Automator/Combine Word Documents.action" \
"/Library/Automator/Convert Format of Excel Files.action" \
"/Library/Automator/Convert Format of PowerPoint Presentations.action" \
"/Library/Automator/Convert PowerPoint Presentations to Movies.action" \
"/Library/Automator/Convert Word Content Object to Text Object.caction" \
"/Library/Automator/Copy Excel Workbook Content to the Clipboard.action" \
"/Library/Automator/Copy PowerPoint Slides to the Clipboard.action" \
"/Library/Automator/Copy Word Document Content to the Clipboard.action" \
"/Library/Automator/Create List from Data in Workbook.action" \
"/Library/Automator/Create New Excel Workbook.action" \
"/Library/Automator/Create New Outlook Mail Message.action" \
"/Library/Automator/Create New PowerPoint Presentation.action" \
"/Library/Automator/Create New Word Document.action" \
"/Library/Automator/Create PowerPoint Picture Slide Shows.action" \
"/Library/Automator/Create Table from Data in Workbook.action" \
"/Library/Automator/Delete Outlook Items.action" \
"/Library/Automator/Find and Replace Text in Word Documents.action" \
"/Library/Automator/Flag Word Documents for Follow Up.action" \
"/Library/Automator/Forward Outlook Mail Messages.action" \
"/Library/Automator/Get Content from Word Documents.action" \
"/Library/Automator/Get Images from PowerPoint Slides.action" \
"/Library/Automator/Get Images from Word Documents.action" \
"/Library/Automator/Get Parent Presentations of Slides.action" \
"/Library/Automator/Get Parent Workbooks.action" \
"/Library/Automator/Get Selected Content from Excel Workbooks.action" \
"/Library/Automator/Get Selected Content from Word Documents.action" \
"/Library/Automator/Get Selected Outlook Items.action" \
"/Library/Automator/Get Selected Text from Outlook Items.action" \
"/Library/Automator/Get Text From Outlook Mail Messages.action" \
"/Library/Automator/Get Text from Word Documents.action" \
"/Library/Automator/Import Text Files to Excel Workbook.action" \
"/Library/Automator/Insert Captions into Word Documents.action" \
"/Library/Automator/Insert Content into Outlook Mail Messages.action" \
"/Library/Automator/Insert Content into Word Documents.action" \
"/Library/Automator/Insert New PowerPoint Slides.action" \
"/Library/Automator/Mark Outlook Mail Message as a To Do Item.action" \
"/Library/Automator/Office.definition" \
"/Library/Automator/Open Excel Workbooks.action" \
"/Library/Automator/Open Outlook Items.action" \
"/Library/Automator/Open PowerPoint Presentations.action" \
"/Library/Automator/Open Word Documents.action" \
"/Library/Automator/Paste Clipboard Content into Excel Workbooks.action" \
"/Library/Automator/Paste Clipboard Content into Outlook Items.action" \
"/Library/Automator/Paste Clipboard Content into PowerPoint Presentations.action" \
"/Library/Automator/Paste Clipboard Content into Word Documents.action" \
"/Library/Automator/Play PowerPoint Slide Shows.action" \
"/Library/Automator/Print Excel Workbooks.action" \
"/Library/Automator/Print Outlook Messages.action" \
"/Library/Automator/Print PowerPoint Presentations.action" \
"/Library/Automator/Print Word Documents.action" \
"/Library/Automator/Protect Word Documents.action" \
"/Library/Automator/Quit Excel.action/Contents" \
"/Library/Automator/Quit PowerPoint.action" \
"/Library/Automator/Quit Word.action" \
"/Library/Automator/Reply to Outlook Mail Messages.action" \
"/Library/Automator/Save Excel Workbooks.action" \
"/Library/Automator/Save Outlook Draft Messages.action" \
"/Library/Automator/Save Outlook Items as Files.action" \
"/Library/Automator/Save Outlook Messages as Files.action" \
"/Library/Automator/Save PowerPoint Presentations.action" \
"/Library/Automator/Save Word Documents.action" \
"/Library/Automator/Search Outlook Items.action" \
"/Library/Automator/Select Cells in Excel Workbooks.action" \
"/Library/Automator/Select PowerPoint Slides.action" \
"/Library/Automator/Send Outgoing Outlook Mail Messages.action" \
"/Library/Automator/Set Category of Outlook Items.action" \
"/Library/Automator/Set Document Settings.action" \
"/Library/Automator/Set Excel Workbook Properties.action/Contents/" \
"/Library/Automator/Set Footer for PowerPoint Slides.action" \
"/Library/Automator/Set Outlook Contact Properties.action" \
"/Library/Automator/Set PowerPoint Slide Layout.action" \
"/Library/Automator/Set PowerPoint Slide Transition Settings.action" \
"/Library/Automator/Set Security Options for Word Documents.action" \
"/Library/Automator/Set Text Case in Word Documents.action" \
"/Library/Automator/Set Word Document Properties.action" \
"/Library/Automator/Sort Data in Excel Workbooks.action" \
"/Library/Internet Plug-Ins/SharePointBrowserPlugin.plugin" \
"/Library/LaunchDaemons/com.microsoft.office.licensing.helper.plist" \
"/Library/Preferences/.s8chpZHsqJ" \
"/Library/Preferences/com.microsoft.office.licensing.plist" \
"/Library/PrivilegedHelperTools/com.microsoft.office.licensing.helper" \
"/private/tmp/111" \
"/private/tmp/Microsoft/dock" \
"/private/tmp/Microsoft/launch" \
"/private/var/db/receipts/com.microsoft" \
"/private/var/db/receipts/com.microsoft.mau.all.autoupdate.pkg.2.3.3.bom" \
"/private/var/db/receipts/com.microsoft.mau.all.autoupdate.pkg.2.3.3.plist" \
"/private/var/db/receipts/com.microsoft.mau.all.autoupdate.pkg.2.3.6.bom" \
"/private/var/db/receipts/com.microsoft.mau.all.autoupdate.pkg.2.3.6.plist" \
"/private/var/db/receipts/com.microsoft.merp.all.errorreporting.pkg.2.2.8.bom" \
"/private/var/db/receipts/com.microsoft.merp.all.errorreporting.pkg.2.2.8.plist" \
"/private/var/db/receipts/com.microsoft.merp.all.errorreporting.pkg.2.2.9.bom" \
"/private/var/db/receipts/com.microsoft.merp.all.errorreporting.pkg.2.2.9.plist" \
"/private/var/db/receipts/com.microsoft.msgr.all.messenger.pkg.8.0.1.bom" \
"/private/var/db/receipts/com.microsoft.msgr.all.messenger.pkg.8.0.1.plist" \
"/private/var/db/receipts/com.microsoft.office.all.automator.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.automator.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.automator.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.automator.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.clipart_search9.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.clipart_search9.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.core.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.core.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.core.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.core.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.dcc.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.dcc.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.dcc.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.dcc.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.dock.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.dock.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.equationeditor.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.equationeditor.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.equationeditor.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.equationeditor.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.excel.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.excel.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.excel.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.excel.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.fix_permissions.pkg.14.3.6.bom" \
"/private/var/db/receipts/com.microsoft.office.all.fix_permissions.pkg.14.3.6.plist" \
"/private/var/db/receipts/com.microsoft.office.all.fonts.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.fonts.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.fonts.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.fonts.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.graph.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.graph.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.graph.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.graph.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.launch.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.launch.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.licensing.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.licensing.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.ooxml.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.ooxml.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.ooxml.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.ooxml.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.outlook.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.outlook.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.outlook.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.outlook.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.powerpoint.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.powerpoint.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.powerpoint.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.powerpoint.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_brazilian.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_brazilian.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_brazilian.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_brazilian.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_catalan.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_catalan.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_catalan.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_catalan.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_czech.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_czech.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_czech.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_czech.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_danish.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_danish.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_danish.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_danish.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_dutch.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_dutch.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_dutch.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_dutch.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_english.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_english.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_english.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_english.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_finnish.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_finnish.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_finnish.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_finnish.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_french.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_french.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_french.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_french.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_german.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_german.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_german.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_german.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_italian.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_italian.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_italian.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_italian.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_italian_grammar.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_italian_grammar.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_japanese.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_japanese.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_japanese.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_japanese.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_norwegian.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_norwegian.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_norwegian.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_norwegian.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_polish.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_polish.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_polish.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_polish.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_portuguese.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_portuguese.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_portuguese.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_portuguese.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_russian.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_russian.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_russian.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_russian.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_spanish.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_spanish.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_spanish.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_spanish.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_swedish.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_swedish.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_swedish.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_swedish.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_turkish.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_turkish.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_turkish.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.proofing_turkish.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.quit.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.quit.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.quit.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.quit.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.required_home.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.required_home.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.setupasst.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.setupasst.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.sharepointbrowserplugin.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.sharepointbrowserplugin.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.sharepointbrowserplugin.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.sharepointbrowserplugin.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.slt_std.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.slt_std.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.vb.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.vb.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.vb.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.vb.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.all.word.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.all.word.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.all.word.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.all.word.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.automator_workflow.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.automator_workflow.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.automator_workflow.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.automator_workflow.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.clipart.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.clipart.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.clipart.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.clipart.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.core_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.core_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.core_resources.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.core_resources.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.core_themes.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.core_themes.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.dcc_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.dcc_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.dcc_resources.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.dcc_resources.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.equationeditor_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.equationeditor_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.equationeditor_resources.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.equationeditor_resources.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.excel_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.excel_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.excel_resources.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.excel_resources.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.excel_templates.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.excel_templates.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.excel_templates.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.excel_templates.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.excel_webqueries.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.excel_webqueries.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.flip4mac.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.flip4mac.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.fonts_fontcollection.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.fonts_fontcollection.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.graph_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.graph_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.graph_resources.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.graph_resources.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.langregister.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.langregister.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.langregister.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.langregister.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.outlook_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.outlook_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.outlook_resources.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.outlook_resources.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.outlook_scriptmenuitems.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.outlook_scriptmenuitems.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.outlook_scriptmenuitems.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.outlook_scriptmenuitems.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.powerpoint_guidedmethods.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.powerpoint_guidedmethods.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.powerpoint_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.powerpoint_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.powerpoint_resources.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.powerpoint_resources.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.powerpoint_templates.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.powerpoint_templates.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.query.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.query.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.query.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.query.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.readme.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.readme.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.readme.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.readme.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.required.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.required.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.setupasst_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.setupasst_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.sharepointbrowserplugin_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.sharepointbrowserplugin_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.sharepointbrowserplugin_resources.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.sharepointbrowserplugin_resources.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.silverlight.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.silverlight.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.solver.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.solver.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.solver.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.solver.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.sounds.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.sounds.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.vb_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.vb_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.word_resources.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.word_resources.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.word_resources.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.word_resources.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.office.en.word_templates.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.word_templates.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.word_wizards.pkg.14.2.0.bom" \
"/private/var/db/receipts/com.microsoft.office.en.word_wizards.pkg.14.2.0.plist" \
"/private/var/db/receipts/com.microsoft.office.en.word_wizards.pkg.14.3.6.update.bom" \
"/private/var/db/receipts/com.microsoft.office.en.word_wizards.pkg.14.3.6.update.plist" \
"/private/var/db/receipts/com.microsoft.rdc.all.rdc.pkg.2.1.1.bom" \
"/private/var/db/receipts/com.microsoft.rdc.all.rdc.pkg.2.1.1.bom" \
"/private/var/db/receipts/com.microsoft.rdc.all.rdc.pkg.2.1.1.plist"
do

	if [ -e "$i" ]
	then
			if [ -f "$i" ]
			then
					sudo rm -f && echo "$i" | tee -a "$NAME.files-removed.txt"

					EXIT="$?"

					if [ "$EXIT" != "0" ]
					then
							echo "$i did not remove cleanly (EXIT = $EXIT)" | tee -a "$NAME.file-remove-errors.txt"
					fi

			else
					sudo rm -rf "$i" && echo "$i" | tee -a "$NAME.directories-removed.txt"

					EXIT="$?"

					if [ "$EXIT" != "0" ]
					then
							echo "$i did not remove cleanly (EXIT = $EXIT)" | tee -a "$NAME.directories-remove-errors.txt"
					fi
			fi
	else
			echo "$i does not exist" | tee -a "$NAME.does-not-exist.txt"
	fi

done

# For some reason this seems necessary

noglob echo "$NAME: If you continue, all of your /private/var/db/receipts/com.microsoft.* files will be deleted. This is almost certainly a very stupid idea and you are STRONGLY encouraged NOT to do it unless you KNOW that you want to remove them all."

echo "$NAME: Press Control+C to cancel, or press Enter to continue: "

read ANSWER2

	# invalidate their `sudo`
sudo -K

echo "$NAME: OK, if you really want to do this you will have to enter your Administrator (sudo) password again:"

	# invoke `rm` with `sudo`
sudo rm -fv /private/var/db/receipts/com.microsoft.*


exit
#
#EOF
