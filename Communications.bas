Attribute VB_Name = "Communications"
Option Explicit

Sub SendEmail(ByVal strFound As String, ByVal strNA As String, ByVal strNotFound As String)
    'The purpose of this procedure is to send email notices to interested parties
    'concerning the files just moved.
    
    
    Dim i As Integer
    
    '***********************************************************************************************
    'CREATE AND SEND THE EMAIL
    '***********************************************************************************************
    
    'Create the email object
    Dim myOlApp As Object
    Dim myItem As Object
    Dim myAttachments As Object
    Dim strBody As String
    Set myOlApp = CreateObject("Outlook.Application")
    Set myItem = myOlApp.CreateItem(olMailItem)

    
    'Add email addresses of the recipients
    For i = 1 To UBound(CurrentSettings)
        If CurrentSettings(i, 1) = "EML" Then
            If CurrentSettings(i, 2) = "ALL" Then
                If CurrentSettings(i, 3) = "ALL" Then
                    myItem.Recipients.Add CurrentSettings(i, 4)
                End If
            End If
        End If
    Next i
    'myItem.Recipients.Add "jduane@mdtadvisers.com"
    'myItem.Recipients.Add "tbeals@mdtadvisers.com"
    
    'Create subject of email and body of the email.
    'Add the email subject line
    myItem.Subject = "UMA Holdings Files For: " & frmMain.tbDate.Text
        
    'Begin composing body of the email
    strBody = "The following UMA holdings files were moved to Momentum:  " & vbCrLf
    strBody = strBody & strFound & vbCrLf
    strBody = strBody & "The following files were not moved because those strategies had no assets:  " & vbCrLf
    strBody = strBody & strNA & vbCrLf & vbCrLf
    strBody = strBody & "The following UMA holdings files were not found:  " & vbCrLf
    strBody = strBody & strNotFound & vbCrLf & vbCrLf
    
    'Finish composing body of email
    strBody = strBody & "If you have any questions please call "
    strBody = strBody & myUser.LongName & " at " & myUser.Telephone & "."
    myItem.Body = strBody
    

    
    'Send email
    myItem.Send
    
    Set myAttachments = Nothing
    Set myItem = Nothing
    Set myOlApp = Nothing


End Sub
