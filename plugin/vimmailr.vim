" MAPI mail scripts for [g]vim
" Language:    English
" Author:      RW Fuller, f_uller_r_w@bell_south.net (remove underscores)
" Last Change: July 12, 2001 (see history)
" Version:     1.20
"
" Synopsis:
" This script contains commands allowing you to send mail via MAPI
" from within [g]vim using the companion DLL vimmailr.dll. Unlike
" the File->Send option in several Windows products, the contents
" of the file you are editing becomes the email message as opposed
" to being an attachment to an email message.
"
" Environment Variables:
" $MYEMAIL and $MYEMAILSIG - both are optional - see preparation.
"
" Preparation:
" 1. Copy vimmailr.dll to one of the usual locations on your path.
" 2. Source this file (you might want to put the command to source
"    the file in your [g]vimrc file).
" 3. Optional - you might also want to set the $MYEMAIL environment
"    variable in your [g]vimrc file to contain your email address.
"    If you do, it will be used when building the header (see _mmh)
"    Example - let $MYEMAIL="foo@bar.com"
" 4. Optional - you can also create a signature file containing a
"    signature that you want to have added to the message. Then set
"    the $MYEMAILSIG environment variable to point to the full path
"    name of the file. The file will be pulled in and appended to
"    the message when you execute the _mmh command.
"    Example - let $MYEMAILSIG="d:\\vim\\data\\emailsig.txt"
"    Note that TWO backslashes are required.
"
" Usage:
" 1. Create the file you want to send and edit away. Before sending,
"    you'll need to create a header. See below for header format.
"    The _mmh command can speed this up if you like (see _mmh below).
" 2. Optional - You may find it handy to do a :set filetype=mail
"    if the file extension you are using does not map to a mail
"    file type. This will give you the mail syntax highlighting.
" 3. Save the file
" 4. Execute the mapped command _vms (for "vim mail send").
"
" Header:
" The first 5 lines of the file, referred to as the "header" MUST
" be as follows:
" To: address1&somewhere.com; address2@somewhere.com
" Cc: address1@somewhere.com; address2@somewhere.com
" Bcc: address1@somewhere.com; address2@somewhere.com
" From: me@here.com (i.e. your email address)
" Subject: The subject
"
" From here on is free form message text...
"
" AttachFile[file_to_attach] (optional)
" AttachFile[another_file_to_attach] (optional)
"
" The details:
" -- The To:, Cc:, Bcc:, From:, and subject: markers are not case
"    sensitive. That is, TO:, to:, or To: will work but there
"    must be a space after the ':' character.
" -- There must be at least one email address on the TO line. If
"    there is more than one email address, they must be separated
"    by a semicolon.
" -- You do not need to include any email addresses for the Cc: or
"    Bcc lines but the Cc: and Bcc: lines must still exist either
"    way. Separate multiple email addresses by a semicolon.
" -- The contents of the "header" are not sent as part of the
"    message. Only the text beginning on the line under the
"    subject is sent.
" -- If you want to attach a file to the email message, add the
"    keyword AttachFile[filename] to the message (usually at the
"    end). The full path to the file to attach must be enclosed
"    in square brackets and the opening square bracket must
"    immediately follow AttachFile.
"
" Included Maps:
" _vms - "Vim Mail Send"
"        Execute this command to do the send
" _mmh - "Make Mail Header"
"        A simple helper to construct the header for you. After
"        creating the file containing the message you want to send,
"        you can execute this command and it will insert the 5 lines
"        of the header for you. You'll still need to fill in the
"        recipients address, cc/bcc if any, and subject but this
"        will ensure everything is in the correct format.
"        After it creates the header it will place you in insert mode
"        on the first line ready for you to type in the recipient
"        information. If you have assigned the $MYEMAIL variable
"        it will also fill in the from information for you. If you
"        have assigned the $MYEMAILSIG to point to a signature file,
"        it will pull that in too.
"
" The internals:
" The DLL, vimmailr.dll, has an exported function VimSendMail with the
" prototype char* VimSendMail(char*). The return value is the final
" status of the call and will be printed in the command line area
" after the call. The parameter to the call is the file name that you
" are sending. The first 5 lines of that file must conform to the
" header format (described above) but the rest is free form.
"
" History:
" v 1.00, June 10, 2001
" - Initial release
"
" v 1.10, June 17, 2001
" - Added support for AttachFile[fullpath] so you can include a file
"   attachment (a change to vimmailr.dll).
" - Added support for $MYEMAILSIG. The _mmh command will now pull in
"   an email signature file if set.
"
" v 1.20, July 12, 2001
" - Added support for bcc
" - Added support for multiple file attachments
" - UNIX style files (i.e. terminated in 0x0A only) did not work
"   properly. Altered the parsing code slightly to fix that.
"


" Internal helper to do a sanity check of the header and format.
function VimMailrCheckHeader()
	" Make sure the file exists on disk
	let rc = filereadable(expand('%'))
	if 0 == rc
		echohl errormsg
		echo "Please save the file first"
		echohl None
		return 0
	endif

	" To:
	let str = getline(1)
	let rc = match(str, "^[Tt][Oo]: ")
	if rc != 0
		echohl errormsg
		echo "TO line is not formatted correctly"
		echohl None
		return 0
	endif

	" Cc:
	let str = getline(2)
	let rc = match(str, "^[Cc][Cc]: ")
	if rc != 0
		echohl errormsg
		echo "CC line is not formatted correctly"
		echohl None
		return 0
	endif

	" Bcc:
	let str = getline(3)
	let rc = match(str, "^[Bb][Cc][Cc]: ")
	if rc != 0
		echohl errormsg
		echo "BCC line is not formatted correctly"
		echohl None
		return 0
	endif

	" From:
	let str = getline(4)
	let rc = match(str, "^[Ff][Rr][Oo][Mm]: ")
	if rc != 0
		echohl errormsg
		echo "FROM line is not formatted correctly"
		echohl None
		return 0
	endif

	" Subject:
	let str = getline(5)
	let rc = match(str, "^[Ss][Uu][Bb][Jj][Ee][Cc][Tt]: ")
	if rc != 0
		echohl errormsg
		echo "SUBJECT line is not formatted correctly"
		echohl None
		return 0
	endif

	return 1
endfunction


function VimMailrSend()
	" Header OK?
	let rc = VimMailrCheckHeader()
	if rc != 1
		return
	endif

	" Confirm they want to send
	let question = 'Mail file ' . expand('%') . '. Continue?'
	let rc = confirm(question, "&Yes\n&No", 1, "Question")
	if rc != 1
		return
	endif

	" go
	echo "Sending. Standby..."
	let strRet = libcall("vimmailr.dll", "VimSendMail", expand('%:p'))
	let strMsg = "Return value: " . strRet
	echo strMsg
endfunction


function VimMailrMakeHeader()
	" If they have $MYEMAIL set, use it for the FROM: line
	let sFrom = "From: " . $MYEMAIL
	let r = append(0, "To: ")
	let r = append(1, "Cc: ")
	let r = append(2, "Bcc: ")
	let r = append(3, sFrom)
	let r = append(4, "Subject: ")
	let r = append(5, "")

	" If they have a signature file, read it in
	if filereadable($MYEMAILSIG)
		execute ":$r" $MYEMAILSIG
	endif

	normal 1G
	startinsert!
endfunction


" Maps
map _vms  :call VimMailrSend()<cr>
map _mmh  :call VimMailrMakeHeader()<cr>

