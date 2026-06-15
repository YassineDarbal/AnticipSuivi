    Sub ActionItemList()
Dim D2D3 As Worksheet, AI As Worksheet, Data As Worksheet, CellFind As Range, Main As Worksheet, MeMo As Worksheet
Dim L As Long, K As Long, J As Long, col As Long, Sap As Worksheet, i As Long, Key As String
Dim Fichier, dic As Object, pl() As Variant, cle As Variant
Dim MyPassword
MyPassword = InputBox("Merci de saisie le mot de passe", "Mot de passe, accès au Macro", "********")
If Not MyPassword = "MTDA350" Then MsgBox "Mot de passe est incorrecte !", vbInformation, "Access": Exit Sub

Set Sap = ThisWorkbook.Sheets("SAP")
Set D2D3 = ThisWorkbook.Sheets("D2D3")
Set Main = ThisWorkbook.Sheets("Main")
Set MeMo = ThisWorkbook.Sheets("Como-Memo")
Set AI = ThisWorkbook.Sheets("AI")
Set dic = CreateObject("Scripting.Dictionary")


L = AI.Range("A" & Rows.Count).End(xlUp).Row
If L > 2 Then AI.Rows("2:" & L).Delete Shift:=xlUp
With Main
  T = .Range(.[A3], .Range("AA" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
With MeMo
  Y = .Range(.[A2], .Range("AA" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
Fichier = Application.GetOpenFilename("Excel (*.xls*), *.xls*", , "Sélection du fichier", , False)
If Fichier = False Then MsgBox "Problème lors de la séléction du fichier": Exit Sub
' --------------- " Extract Action item " ------------------'
Application.ScreenUpdating = False
    With Workbooks.Open(Fichier)
        With .Sheets(1)
        If .FilterMode = True Then .ShowAllData
        Application.DisplayAlerts = False
         col = .Range("A" & Rows.Count).End(xlUp).Row
         .Range("A2:A" & col).Copy
         AI.Range("B2").PasteSpecial xlPasteValues
         .Range("B2:B" & col).Copy
         AI.Range("A2").PasteSpecial xlPasteValues
         .Range("J2:J" & col).Copy
         AI.Range("C2").PasteSpecial xlPasteValues
         .Range("AA2:AA" & col).Copy
         AI.Range("E2").PasteSpecial xlPasteValues
        End With
     .Close savechanges:=False
    End With
With AI
    L = .Range("A" & Rows.Count).End(xlUp).Row
    .Columns("A:A").Replace What:=".*-", Replacement:="-", LookAt:=xlPart, SearchOrder:=xlByRows
    .Columns("E:E").Replace What:="3D*", Replacement:="3D A JOUR", LookAt:=xlPart, SearchOrder:=xlByRows
    For K = 2 To L
            For i = 1 To Len(.Cells(K, 1).Value)
                If Left(Right(.Cells(K, 1), i), 1) = "-" Then
                .Cells(K, 4) = Right(.Cells(K, 1).Value, i - 1)
                Exit For
                End If
            Next i
    If .Cells(K, 5) = "" Then .Cells(K, 5) = "-"
    If Not Len(.Cells(K, 1).Value) > 14 Then .Cells(K, 6) = Left(.Cells(K, 1).Value, 12)
    Next K
    .Columns("D:D").Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
    pl = Range(.[D2], .Range("E" & .[A65536].End(xlUp).Row)).Value
    J = 2
    For Each cle In pl
        dic(cle) = .Cells(J, 5)
        J = J + 1
    Next cle
    W = .Range(.[A2], .Range("F" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
With Sap
    For J = 2 To .Range("A" & Rows.Count).End(xlUp).Row
        .Cells(J, 5) = dic.Item(.Cells(J, 1).Value)
        If .Cells(J, 5) = "" Then .Cells(J, 5) = "-"
    Next J
End With
For J = LBound(T) To UBound(T)
'Si TR Del <110
If T(J, 11) < 110 Then
    Key = T(J, 2) & "-" & T(J, 4) & "-" & T(J, 12)
    Temp = ActionItems(W, Key, 1, 2, 3)
    T(J, 22) = IIf(Temp(1) = "", "-", Temp(1))
    T(J, 23) = IIf(Temp(0) = "", "-", Temp(0))
'Si TR Del >110
ElseIf T(J, 11) >= 110 Then
    Key = T(J, 2) & "-" & T(J, 4)
    Temp = ActionItems(W, Key, 6, 2, 3)
    T(J, 22) = IIf(Temp(1) = "", "-", Temp(1))
    T(J, 23) = IIf(Temp(0) = "", "-", Temp(0))
    If T(J, 22) <> "-" Then
        T(J, 27) = "JIG REQUEST"
    Else
        T(J, 27) = "-"
    End If
End If
Next J
Main.[A3].Resize(UBound(T), UBound(T, 2)) = T
For J = LBound(Y) To UBound(Y)
If Y(J, 11) >= 110 Then
    Key = Y(J, 2) & "-" & Y(J, 4)
    Temp = ActionItems(W, Key, 6, 2, 3)
    Y(J, 21) = IIf(Temp(1) = "", "-", Temp(1))
    Y(J, 22) = IIf(Temp(0) = "", "-", Temp(0))
End If
Next J
MeMo.[A2].Resize(UBound(Y), UBound(Y, 2)) = Y

L = AI.Range("A" & Rows.Count).End(xlUp).Row
If L > 2 Then AI.Rows("2:" & L).Delete Shift:=xlUp
Main.Activate
Application.ScreenUpdating = True
End Sub
Function ActionItems(T, ByVal Crit1$, Optional C1% = 1, Optional Res2% = 4, Optional Res1% = 5)

Dim i&, J&, K&, Temp(0 To 2), Boule As Boolean
  For i = LBound(T) To UBound(T)
    If T(i, C1) = Crit1 Then
        Temp(0) = T(i, Res1): Temp(1) = T(i, Res2): Temp(2) = " "
        Exit For
    End If
  Next i
ActionItems = Temp
End Function
