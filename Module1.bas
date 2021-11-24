Attribute VB_Name = "Module1"
Sub ForwardVasttrafik(Item As Outlook.MailItem)
    Dim MyEmailAddress As String
    Dim TicketType As String
    Dim MyName As String
    Dim MyStudentCardNumber As String
    Dim TargetMailAddress As String
    
    MyEmailAddress = "xxxxxx@student.chalmers.se"
    TicketType = "Zon A 90 minuter Vuxen, 34,00 kr"
    MyName = "xxx xxxx"
    MyStudentCardNumber = "1234 5678"
    TargetMailAddress = "ersattning@chalmersstudentkar.se"
    
    If TypeOf Item Is MailItem Then
    Item.To = MyEmailAddress
        If InStr(Item.Body, TicketType) Then
            Set myForward = Item.Forward
            myForward.Subject = ("Fw: ") & Item.Subject
            myForward.HTMLBody = "<HTML><BODY>My student union card number: " & MyStudentCardNumber & " <br \> My Name: " & MyName & " </BODY></HTML>" & myForward.HTMLBody
            myForward.Recipients.Add TargetMailAddress
            myForward.Recipients.ResolveAll
            myForward.Send
        End If
    End If
End Sub

