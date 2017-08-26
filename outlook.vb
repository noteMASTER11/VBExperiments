Sub Send_Mail()
Dim objOutlookApp As Object, objMail As Object
Dim sTo, sSubject, sBody, sAttachment, githubnick, ghacc As String

Application.ScreenUpdating = False
On Error Resume Next
Set objOutlookApp = GetObject(, "Outlook.Application")
Err.Clear
If objOutlookApp Is Nothing Then
Set objOutlookApp = CreateObject("Outlook.Application")
End If
objOutlookApp.Session.Logon
Set objMail = objOutlookApp.CreateItem(0)
If Err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
githubnick = ActiveCell.Offset(0, 1)
ghacc = ActiveCell.Offset(0, -1)

sTo = ActiveCell.Value 'Кому
sSubject = "Поиск разработчиков" 'Тема
sBody = "Привет, " & githubnick & ", как твои дела? Нашел тебя под ником " & ghacc & "." & vbNewLine & "2 строчка" & vbNewLine & "3 строчка"

With objMail
.To = sTo
.CC = ""
.BCC = ""
.Subject = sSubject
.Body = sBody
'.HTMLBody = sBody
.Attachments.Add sAttachment
.Display 'Send
End With

Set objOutlookApp = Nothing: Set objMail = Nothing
Application.ScreenUpdating = True

ActiveCell.Interior.Color = vbGreen
End Sub


Public Function CreateEmailMsg(cRecipients, _
                        Optional sSubject As String = "", _
                        Optional sBody As String = "", _
                        Optional cAttachments = Nothing) _
                        As Object

    Dim appOL As Object
    Set appOL = CreateObject("Outlook.Application")

    Dim msgNew As Object
    Set msgNew = appOL.CreateItem(0) 'olMailItem

    Dim sItem
    With msgNew
        'Message body
        .BodyFormat = 2 'olFormatHTML
        .HTMLBody = sBody

        'Recipients
        If TypeName(cRecipients) = "String" Then
            .Recipients.Add cRecipients
        ElseIf Not cRecipients Is Nothing Then
            For Each sItem In cRecipients
                .Recipients.Add sItem
            Next sItem
        End If

        'Subject
        .Subject = sSubject

        'Attachments
        If TypeName(cAttachments) = "String" Then
            .Attachments.Add cAttachments, 1
        ElseIf Not cAttachments Is Nothing Then
            For Each sItem In cAttachments
                .Attachments.Add sItem, 1
            Next sItem
        End If
     End With

    Set CreateEmailMsg = msgNew

End Function
