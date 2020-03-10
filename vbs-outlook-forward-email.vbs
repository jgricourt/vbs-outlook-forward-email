'
' Forward any new incoming email in the Outlook mailbox  ...
'
' Non interactive mode : $ vbs-outlook-forward-email.vbs
' Interactive mode : $ cscript vbs-outlook-forward-email.vbs
'

Const olFolderInbox = 6

'Init
Const targetEmail = "john.doe@gmail.com"
Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objInbox = objNamespace.GetDefaultFolder(olFolderInbox)
Set objItems = objInbox.Items

'Date start
DateStart = #2020-01-28 12:00:00#

'Loop thru all emails ...
For Each objItem in objItems

	'Forward only newer emails ...
	If objItem.Unread = True AND DateStart <= objItem.ReceivedTime Then
		
		'Forward email
		Set forwardItem = objItem.Forward
		forwardItem.Subject = "[" & objItem.SenderName  & "] " & forwardItem.Subject
        forwardItem.Recipients.Add targetEmail
        forwardItem.Send
		
		'Mark original email as read
		objItem.Unread = False

	End If
Next


