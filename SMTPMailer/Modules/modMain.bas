Attribute VB_Name = "modMain"
Sub Main()
    Dim f As New frmMail
    
    Load f
    If Not f.Running Then
        f.SendEmail """Name of Sender"" <Sender@attbi.com>", """Recipient Name"" <recipient1@cs.com>" & vbCrLf & """Recipient2 Name"" <recipient2@aol.com>", "Hey!!!", "This is just a test..."
    End If
    Do While f.Running
        DoEvents
    Loop
    If Not f.Running Then
        f.SendEmail """Name of Another Sender"" <Sender2@hotmail.com>", """Recipient Name3"" <recipient3@cs.com>" & vbCrLf & """Recipient4 Name"" <recipient4@aol.com>", "Hi Everyone", "This is another test..."
    End If
    Do While f.Running
        DoEvents
    Loop
    Unload f
    
End Sub
