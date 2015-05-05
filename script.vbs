Sub SavesAndSendsEmail()

    If Dir("C:\Weekly Time Records", vbDirectory) = "" Then
        MkDir Path:="C:\Weekly Time Record"
        MsgBox "The Weekly Time Records folder doesn't exist. The directory has been created, and time record successfully saved to C:\Weekly Time Records"
    Else
        MsgBox "Time record successfully saved to C:\Weekly Time Records"
    End If

    Dim fname           As String
    Dim fpath           As String
       
    fpath = "C:\Weekly Time Records"
    fname = Sheets("Weekly time record").Range("D8").Text & "-" & Range("H4").Text
   
    ThisWorkbook.SaveAs Filename:=fpath & "\" & fname
   
    Dim OutApp          As Object
    Dim OutMail         As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
   
    With OutMail
        .To = "katb"
        .CC = ""
        .BCC = ""
        .Subject = ("Time record for week ending on ") & (Sheets("Weekly time record").Range("D8").Value & (" for ") & Range("H4").Text)
        .Body = "Please review the weekly time record attatched to this message. Thanks!"
        .Attachments.Add ActiveWorkbook.FullName
        .Display
        .Send

    End With

    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
       
        MsgBox "The time record has been successfully sent."

End Sub
