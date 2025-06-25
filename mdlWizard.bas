Attribute VB_Name = "mdlWizard"
'****************************
'Component: mdlWixard Module
'Author: Armen Shimoon
'Copyright: Shimoon Technologies 2001
'Email: a_shimoon@hotmail.com
'****************************



'**************************************************************
'*****************How to get this working for you**********************
'**Simply just change the URL variables in Sub Main() to where the information is stored
'**on a webserver. (any host will work. ie: geocities, angelfire, etc). Then the rest is done
'**for you. Oh yeah make sure you change the Version variable to the app version that
'**you want.
'*************************************************************
'**************What to put in the text files in the server*******************
'**VersionURL - your newest version number without any letters. (ie: 1.1)
'**MessageURL - the message you want displayed when users install the update
'**UpdateURL - the URL for the install file.
'**************************************************************


'You have permission to use this code whenever and wherever you wish




'Variables to save data to
Global nVer As Single
Global nMsg As String
Global nURL As String
Dim b() As Byte
Global Version As Single



'Now for the URL variables
Global VersionURL As String
Global MessageURL As String
Global UpdateURL As String




Sub Main()
'Empty out our variables
nVer = 0
nMsg = ""
nURL = ""


'Declare the URLs for getting data
VersionURL = "http://www.weylan.com/version.txt"
MessageURL = "http://www.weylan.com/message.txt"
UpdateURL = "http://www.weylan.com/update.txt"

'Show form
frmWizard.Show

End Sub


Function GetData()
'Incase server is not up, it will trap the error
On Error GoTo 10
'Now we will download the information
frmWizard.cmdNext2.Enabled = False
frmWizard.lblStat.Caption = "Connecting..."
nVer = frmWizard.net.OpenURL(VersionURL)
frmWizard.lblStat.Caption = "Getting version info..."
nMsg = frmWizard.net.OpenURL(MessageURL)
frmWizard.lblStat.Caption = "Getting update message..."
nURL = frmWizard.net.OpenURL(UpdateURL)
frmWizard.lblStat.Caption = "Processing..."

If nVer < Version Then
    b() = frmWizard.net.OpenURL(nURL, icByteArray)
    frmWizard.lblStat.Caption = "Downloading update file..."
Else
    frmWizard.lblStat.Caption = "Complete."
    frmWizard.cmdNext2.Enabled = True
    Exit Function
End If

frmWizard.lblStat.Caption = "Complete."
frmWizard.cmdNext2.Enabled = True
Exit Function

10: MsgBox "The following error occurred:" & vbCrLf & Err.Description, vbCritical, Err.Number

End Function

Function Puttofile()
On Error GoTo 20
Unload frmWizard


Exit Function
20: MsgBox "Fatal error, please try again.", vbCritical, "Error"
End Function
