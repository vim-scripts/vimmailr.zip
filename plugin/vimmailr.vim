" MAPI mail scripts for [g]vim
" Language:    English
" Author:      RW Fuller, r_w.fuller@ver_izon.net (remove underscores)
" Last Change: November 25, 2001 (see history)
" Version:     1.30
"
" Synopsis:
" This script contains commands that allow you to send mail (via MAPI, on
" Windows) from within [g]vim using the companion DLL, vimmailr.dll. The
" basic concept is this: Just about any text file can be sent so long as
" the proper "header" attributes are present in the file (see the section
" on the header below). But, the script does allow for several alternate
" ways to send a message including quickly sending a message from the vim
" command line where you need not deal with the header if you do not want
" to.
"
" Environment Variables:
" $MYEMAIL and $MYEMAILSIG - both are optional - see preparation.
"
" Initial Preparation:
" 1. Copy vimmailr.dll to one of the usual locations on your path.
" 2. Optional - set the $MYEMAIL environment variable in your [g]vimrc
"    file to contain your email address. If you do, it will be used when
"    building the header. Example - let $MYEMAIL="foo@bar.com"
" 3. Optional - you can also create a signature file containing a
"    signature that you want to have added to the message. Set the
"    $MYEMAILSIG environment variable to point to the full path
"    name of the file. The file will be pulled in and appended to
"    the message when you execute the command that builds the header.
"    Example - let $MYEMAILSIG="d:\\vim\\emailsig.txt"
"    Note that TWO backslashes are required.
" 4. Put the command to source this file in your [g]vimrc file.
"    It is not (yet) a plugin - sorry...
"
" Basic Usage:
" The most basic usage of the script is where you create a text file
" and compose your message. While composing the message, you need not
" concern yourself with the header until you are ready to send the
" message. At that time, you'd execute the command to build the
" header, save the file, then send it. The general steps are:
" 1. Create a new file and compose your message. The contents of the
"    file are free form. If you want to add an attachment, see below.
" 2. Execute the _vmh command to create the header. This will put
"    things in the proper place.
" 3. Fill in the recipient info and the subject.
" 4. Optional - You may find it handy to do a :set filetype=mail if
"    the file extension you are using does not map to a mail file type.
"    This will give you the mail syntax highlighting.
" 5. Save the file
" 6. Execute the mapped command _vms (for "vim mailer send").
"
" Simpler Usage:
" An even simpler way to send a message is by executing the _vmm
" command. This command allows you to quickly send a message without
" manually creating a file at all. That is, you will be prompted for
" the recipient information, the subject, and the message text all
" on the command line. Behind the scenes, a temp file conforming to
" the correct format is still being created but you need not create
" it yourself. This method is probably the quickest and easiest way
" to send a message but is only appropriate for a quick one or two
" line message. The steps are:
" 1. Execute the _vmm command.
" 2. You will be prompted for the recipient information, the subject
"    and the message text on the command line. Hit enter after each.
"    Be sure to separate addresses with semicolons.
" 3. After confirming that you want to send the message, it will send
"    it for you.
"
" Another Use:
" There may be times when you just want to quickly send a file to
" someone. There are a couple ways to do this: a) open the file, add
" a header, then send it or b) compose a message and add that file as
" an attachment then send it. Now there's a new quicker option; the
" _vmf command. This command is similar to the _vmm (the "quick
" message" command) in that you will be prompted for the information
" on the command line. The steps are:
" 1. Execute the _vmf (mail file) command.
" 2. You will be prompted for the recipient information, the subject,
"    a short optional message, and the path to the file to send on the
"    command line. Hit enter after each. Be sure to separate addresses
"    with semicolons. Be sure to enter the fully qualified path name
"    of the file. It is assumed that the current file is the file you
"    want to send. Enter a different path if that is not the case.
" 3. After confirming that you want to send the message, it will send
"    it for you.
" 
"
" Header Format:
" The first 5 lines of the file, referred to as the "header" MUST be
" as follows:
" To: address1&somewhere.com; address2@somewhere.com
" Cc: address1@somewhere.com; address2@somewhere.com
" Bcc: address1@somewhere.com; address2@somewhere.com
" From: me@here.com (i.e. your email address)
" Subject: The subject
"
" The rest of the message is free form text...
"
" If you want to attach a file, do this:
" AttachFile[file_to_attach] (optional)
" AttachFile[another_file_to_attach] (optional)
"
" The Details:
" -- The To:, Cc:, Bcc:, From:, and subject: markers are not case
"    sensitive. That is, TO:, to:, or To: will work but there
"    must be a space after the ':' character.
" -- There must be at least one email address on the TO: line. If
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
" _vms - "Vim Mailer Send"
"        Execute this command to do the send of a properly formatted
"        file.
" _vmh - "Vim Mailer Header"
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
"        it will pull that in too at the end of the file.
" _vmm - "Vim Mailer (quick) Message"
"        An easier way to send a message quickly without having to
"        concern yourself with manually creating a file first. Just
"        execute the command and it will prompt you for the info on
"        the command line.
" _vmf - "Vim Mailer mail file"
"        Provides a quick way of sending a file
"
"
" The DLL Internals:
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
" - Added support for $MYEMAILSIG. The _vmh command will now pull in
"   an email signature file if set.
"
" v 1.20, July 12, 2001
" - Added support for bcc
" - Added support for multiple file attachments
" - UNIX style files (i.e. terminated in 0x0A only) did not work
"   properly. Altered the parsing code slightly to fix that.
"
" v 1.30, November 25, 2001
" - Bullet proofed the parsing code in the DLL (again).
" - Now works with vim6 (but still not a plugin - you must source it).
" - Added the _vmm command. An even quicker way to send a message
"   entirely from the vim command line.
" - Added the _vmf command. A quick way of sending a file entirely
"   from the vim command line.
" - Includes mailit.cpp - a simple test app you can use to test the
"   dll outside of vim. Could be used to send messages in a batch file.
" - Changed the _mmh command to _vmh (for vim mailer header) to be
"   more consistent with the other commands in the script.



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


function VimMailrQuickMessage()
	" get the main recipient(s)
	let sTO = input("To: ", "")
	" must be at least 1
	if (strlen(sTO) < 1)
		echohl errormsg
		echo "Must be at least one recipient - aborting"
		echohl None
		return
	endif
	" get the other info
	let sCC = input("Cc: ", "")
	let sBCC = input("Bcc: ", "")
	let sFrom = input("From: ", $MYEMAIL)
	let sSubj = input("Subject: ", "")
	let sMsg = input("Message: ", "")

	" put it all in a temp file
	let sFile = tempname()
	execute 'new ' . sFile
	execute 'set filetype=mail'
	let r = append(0, "To: " . sTO)
	let r = append(1, "Cc: " . sCC)
	let r = append(2, "Bcc: " . sBCC)
	let r = append(3, "From: " . sFrom)
	let r = append(4, "Subject: " . sSubj)
	let r = append(5, "")
	let r = append(6, sMsg)

	" If they have a signature file, read it in
	if filereadable($MYEMAILSIG)
		execute ":$r" $MYEMAILSIG
	endif

	" save it
	execute 'write'
	" go
	call VimMailrSend()
endfunction


function VimMailrMailFile()
	" get the file to send first
	let sFile = input("File to send: ", expand('%:p'))
	let r = filereadable(sFile)
	if (0 == r)
		echohl errormsg
		echo "File is not readable - aborting"
		echohl None
		return
	endif

	" get the main recipient(s)
	let sTO = input("To: ", "")
	" must be at least 1
	if (strlen(sTO) < 1)
		echohl errormsg
		echo "Must be at least one recipient - aborting"
		echohl None
		return
	endif
	" get the other info
	let sCC = input("Cc: ", "")
	let sBCC = input("Bcc: ", "")
	let sFrom = input("From: ", $MYEMAIL)
	let sSubj = input("Subject: ", "")
	let sMsg = input("Message: ", "")

	" put it all in a temp file
	execute 'new ' . tempname()
	execute 'set filetype=mail'
	let r = append(0, "To: " . sTO)
	let r = append(1, "Cc: " . sCC)
	let r = append(2, "Bcc: " . sBCC)
	let r = append(3, "From: " . sFrom)
	let r = append(4, "Subject: " . sSubj)
	let r = append(5, "")
	let r = append(6, sMsg)
	let r = append(7, "AttachFile[" . sFile . "]")

	" If they have a signature file, read it in
	if filereadable($MYEMAILSIG)
		execute ":$r" $MYEMAILSIG
	endif

	" save it
	execute 'write'
	" go
	call VimMailrSend()
endfunction


" Maps
map _vms  :call VimMailrSend()<cr>
map _vmh  :call VimMailrMakeHeader()<cr>
map _vmm  :call VimMailrQuickMessage()<cr>
map _vmf  :call VimMailrMailFile()<cr>

