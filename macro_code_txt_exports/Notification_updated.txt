Sub SendNotif()

Dim Objet As String, Corps As String


Objet = "Notification: Suivi Anticipation à Jour"

Corps = "Bonjour" & vbNewLine & vbNewLine & "le suivi Anticipation est à jour " & Date & vbNewLine

Call Send_mail(RangeN("RngMailMTD").Value, "naima.beddouz@safrangroup.com", Objet, Corps)

End Sub
