-- Author: Andrew Afonso - https://github.com/andrewbluepiano
-- Description: Used to pull email headers from messages in trash, or spam, to see similarities to find ways to make good email filters. 
-- License Note: Use whatever you want if its open sourced, just credit here if you copy paste the whole thing. For use in licensed software we need to talk first.

-- Variable Setup
set headerList to {}
set outFileName to "extracted-headers.csv"
tell application "Finder" to set currPath to POSIX path of (container of (path to me) as string)
set outputFile to (currPath as string) & outFileName

set promptTextOne to "This script defaults to saving in the same directory it is run from. Select a new output directory?

Current output file:
" & outputFile as string

-- Let user choose custom output path if they so desire
set saveChoice to button returned of (display dialog promptTextOne buttons {"Choose New Output Folder", "Continue"} default button "Continue")
if saveChoice is "Choose New Output Folder" then
	set outputFile to (((POSIX path of (choose folder with prompt "Please select an output folder:")) as string) & outFileName)
end if


tell application "Mail"
	-- Variables
	set MBList to {}
	set acctNames to {}
	set assocAccts to {}
	set assocMB to {}
	
	-- Get Accounts, have user pick account.
	set availableAccounts to every account
	set acctCount to 1
	repeat with anAccount in availableAccounts
		tell anAccount
			set acctListName to ((acctCount as string) & " - " & name of anAccount)
			set end of acctNames to acctListName
			set end of assocAccts to {acctListName, anAccount}
			set acctCount to acctCount + 1
		end tell
	end repeat
	set targetAcctTxt to choose from list acctNames with prompt "Select target account"
	
	-- Translate choice from list to account object
	repeat with acctAssocItem in assocAccts
		if (first item of acctAssocItem as string) is (targetAcctTxt as string) then
			set targetAcct to second item of acctAssocItem
		end if
	end repeat
	
	
	-- Get Mailboxes for account, have user pick target mailbox (MB)
	set MBCount to 1
	tell targetAcct
		set availableMBObjects to every mailbox
		repeat with aMailbox in availableMBObjects
			set MBListName to ((MBCount as string) & " - " & name of aMailbox)
			set end of MBList to MBListName
			set end of assocMB to {MBListName, aMailbox}
			set MBCount to MBCount + 1
		end repeat
	end tell
	set targetMBTxt to choose from list MBList with prompt "Select target mailbox"
	
	-- Translate mailbox choice back to the mailbox object
	repeat with MBAssocItem in assocMB
		if (first item of MBAssocItem as string) is (targetMBTxt as string) then
			set targetMB to second item of MBAssocItem
		end if
	end repeat
	
	-- Iterate over messages, store their headers
	set outputFileOpen to open for access outputFile with write permission
	tell targetMB
		set allMsg to every message
		repeat with aMsg in allMsg
			tell aMsg
				set theheads to all headers as string
				set AppleScript's text item delimiters to the "
"
				set the item_list to every text item of theheads
				set AppleScript's text item delimiters to "\",\""
				set theheads to the item_list as string
				set AppleScript's text item delimiters to ""
				set theheads to "\"" & theheads & "\""
				write (theheads & "
") to outputFileOpen
			end tell
		end repeat
	end tell
	close access outputFileOpen
	
	
end tell


