Private WithEvents Items As Outlook.Items
Private Sub Application_Startup()
  
  'Outlook staging
    Dim olNs As Outlook.NameSpace
    Dim Inbox  As Outlook.MAPIFolder

    Set olNs = GetNamespace("MAPI")
    
    'Default inbox
    Set Inbox = olNs.GetDefaultFolder(olFolderInbox)
    Set Items = Inbox.Items
 End Sub

Private Sub Items_ItemAdd(ByVal Items As Object)
 
'If new mail has a specific subject then run access script

'staging area
Dim Msg As Outlook.MailItem
Dim MessageInfo
Dim Result

' if the email is the offer letter then run
If TypeName(Items) = "MailItem" Then
    Set Msg = Items
     If Msg.Subject Like "*Offer Letter Accepted*" Then
       ExecuteDealRequest Items
       
       'creates a popup for the user to go and print the labels
        MessageInfo = "" & "Subject : " & Items.Subject & vbCrLf
        Result = MsgBox(MessageInfo, vbOKOnly, "New Employee Label Required")
       
  End If
End If

End Sub

'open access and runs an export
Sub ExecuteDealRequest(Items As Outlook.MailItem)
    
    'staging area
    Dim AccessApp As Object
    Set AccessApp = CreateObject("Access.Application")
    AccessApp.Visible = True
    
    'Opens the comings and goings database
    AccessApp.OpenCurrentDatabase "H:\Workforce Maintenance\Comings & Goings & Team Structure\comings and goings report.accdb"
    
    'Runs the label export
    AccessApp.DoCmd.RunSavedImportExport ("Export-NewComingsLabels")
    
    'closes the database
    AccessApp.CloseCurrentDatabase
    Set AccessApp = Nothing
End Sub