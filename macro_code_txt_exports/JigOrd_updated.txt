Sub indexinJigOderingLM()
Dim Confsh As Worksheet, L As Long, LsRw As Long, NbLigne As Long, NbLg As Long, Main As Worksheet, MeMo As Worksheet
Dim Last_Row As Long, McB As Worksheet
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Set Confsh = ThisWorkbook.Sheets("CONF")
Set Main = ThisWorkbook.Sheets("Main")
Set McB = ThisWorkbook.Sheets("MCB")
Set MeMo = ThisWorkbook.Sheets("Como-Memo")
With Confsh
If .FilterMode = True Then .ShowAllData
L = .Range("A" & Rows.Count).End(xlUp).Row
If L > 2 Then .Rows("2:" & L).Delete Shift:=xlUp
End With

Fichier = "S:\Temara\D16\SUPPORT-CONFIGURATION\RESTRICTED\SUIVI-CONF-PZ\JIG_ORDERING_LM\JIG_ORDERING_LM.xlsx"
'Application.GetOpenFilename("Excel (*.xls*), *.xls*", , "Sélection du fichier", , False)
If Fichier = False Then MsgBox "Problème lors de la séléction du fichier": Exit Sub
Application.ScreenUpdating = False
Application.DisplayAlerts = False
With Workbooks.Open(Fichier)
 With .Sheets("JIG calcul")
   If .FilterMode = True Then .ShowAllData
   NbLigne = .Range("A" & Rows.Count).End(xlUp).Row
    .Range("A2:B" & NbLigne).Copy
     Confsh.Cells(2, 1).PasteSpecial Paste:=xlPasteValues
    .Range("D2:D" & NbLigne).Copy
     Confsh.Cells(2, 3).PasteSpecial Paste:=xlPasteValues
    .Range("H2:H" & NbLigne).Copy
     Confsh.Cells(2, 4).PasteSpecial Paste:=xlPasteValues
    .Range("J2:K" & NbLigne).Copy
     Confsh.Cells(2, 5).PasteSpecial Paste:=xlPasteValues
 End With
 With .Sheets("Historic")
   If .FilterMode = True Then .ShowAllData
   LsRw = Confsh.Range("B" & Rows.Count).End(xlUp).Row + 1
   NbLg = .Range("B" & Rows.Count).End(xlUp).Row
    .Range("A2:A" & NbLg).Copy
     Confsh.Cells(LsRw, 2).PasteSpecial Paste:=xlPasteValues
    .Range("C2:C" & NbLg).Copy
     Confsh.Cells(LsRw, 3).PasteSpecial Paste:=xlPasteValues
    .Range("G2:G" & NbLg).Copy
     Confsh.Cells(LsRw, 4).PasteSpecial Paste:=xlPasteValues
    .Range("I2:J" & NbLg).Copy
     Confsh.Cells(LsRw, 5).PasteSpecial Paste:=xlPasteValues
 End With
 .Close 0
End With
With Confsh
.Columns("F:F").Replace What:="/", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
For i = 2 To .Range("B" & Rows.Count).End(xlUp).Row
    If .Cells(i, 2).Value = 5 Then .Cells(i, 2).Value = "R005"
    .Cells(i, 1).Value = "MSN" & .Cells(i, 2).Value & "/" & .Cells(i, 3).Value
Next i
.Columns("F:F").Replace What:=".*", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
V = .Range(.[A2], .Range("F" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
With Main
    Last_Row = IIf(.Range("A" & Rows.Count).Value = vbNullString, .Range("A" & Rows.Count).End(xlUp).Row, Rows.Count)
    Y = .Range(.[A3], .Range("K" & Last_Row)).Value
End With
With McB
    Last_Row = IIf(.Range("A" & Rows.Count).Value = vbNullString, .Range("A" & Rows.Count).End(xlUp).Row, Rows.Count)
    W = .Range(.[A2], .Range("K" & Last_Row)).Value
End With
With MeMo
    Last_Row = IIf(.Range("A" & Rows.Count).Value = vbNullString, .Range("A" & Rows.Count).End(xlUp).Row, Rows.Count)
    Z = .Range(.[A2], .Range("K" & Last_Row)).Value
End With
For J = LBound(Y) To UBound(Y)
    Temp = FindConf(V, Y(J, 2), Y(J, 3), 3, 2, 6, 4)
    Y(J, 4) = IIf(Temp(0) = "", "-", Temp(0))
    If Y(J, 4) = "NO CONF" Then Y(J, 4) = "-"
Next J
Main.[A3].Resize(UBound(Y), UBound(Y, 2)) = Y

For Lig = LBound(W) To UBound(W)
    Temp = FindConf(V, W(Lig, 2), W(Lig, 3), 3, 2, 6, 4)
    W(Lig, 4) = IIf(Temp(0) = "", "-", Temp(0))
    If W(Lig, 4) = "NO CONF" Then W(Lig, 4) = "-"
Next Lig
McB.[A2].Resize(UBound(W), UBound(W, 2)) = W

For K = LBound(Z) To UBound(Z)
    Temp = FindConf(V, Z(K, 2), Z(K, 3), 3, 2, 6, 4)
    Z(K, 4) = IIf(Temp(0) = "", "-", Temp(0))
    If Z(K, 4) = "NO CONF" Then Z(K, 4) = "-"
Next K
MeMo.[A2].Resize(UBound(Z), UBound(Z, 2)) = Z

Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done!"
ThisWorkbook.Sheets("DashBoard").Activate
End Sub
Function FindConf(T, ByVal Crit1$, ByVal Crit2$, Optional C1% = 1, Optional C2% = 4, Optional Res2% = 4, Optional Res1% = 5)

Dim i&, J&, K&, Temp(0 To 2), Boule As Boolean
  For i = LBound(T) To UBound(T)
    If T(i, C1) = Crit1 Then
     If T(i, C2) = Crit2 Then
            Temp(0) = T(i, Res2): Temp(1) = T(i, Res1): Temp(2) = " "
            Exit For
     End If
    End If
  Next i
FindConf = Temp
End Function

