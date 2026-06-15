Public Sub Main_Data_Echeancier()
Echeancer_Datas (False)
End Sub
Public Sub Plaquette_Data_Echeancier()
Echeancer_Datas (True)
End Sub

Sub Echeancer_Datas(IS_Only As Boolean)
Dim Main As Worksheet, Phaz As Long, i As Long, MeMo As Worksheet, McB As Worksheet, Qlp As Worksheet
Dim Fichier, Txt As String, J As Long, K As Long, L As Long, Lin As Long
Dim DispVb, VbDisp As Worksheet


Application.DisplayAlerts = False 'pour éliminer les Alerts
Set Main = ThisWorkbook.Sheets("Main")
Set MeMo = ThisWorkbook.Sheets("Como-Memo")
Set McB = ThisWorkbook.Sheets("MCB")
Set VbDisp = ThisWorkbook.Sheets("Dispatch")
Set Qlp = ThisWorkbook.Sheets("Plaquette")

Qlp.Unprotect
'Qlp.ShowAllData
L = VbDisp.Range("A" & Rows.Count).End(xlUp).Row
If L > 2 Then VbDisp.Rows("2:" & L).Delete Shift:=xlUp

'Dim MyPassword
'MyPassword = InputBox("Merci de saisie le mot de passe", "Mot de passe, accès au Macro", "********")
'If Not MyPassword = "MTDA350" Then MsgBox "Mot de passe est incorrecte !", vbInformation, "Access": Exit Sub
On Error Resume Next
With Main
If .FilterMode = True Then .ShowAllData
    T = .Range(.[A3], .Range("AI" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
With MeMo
If .FilterMode = True Then .ShowAllData
    Y = .Range(.[A2], .Range("AI" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
With McB
If .FilterMode = True Then .ShowAllData
    Z = .Range(.[A2], .Range("AB" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With



DispVb = "S:\Temara\D16\ENGINEERING\RESTRICTED\PREPA-A350XWB\SUIVI LANCEMENT-DIFF\Dispatch_A350.xlsx"
'Fichier = "S:\Temara\D16\PRODUCTION-CONTROL\SHARED\Echeancier\Echéancier_A350_Nouveau.xlsb"
Fichier = ThisWorkbook.Sheets("Link").Range("C5").Value
'Application.GetOpenFilename("Excel (*.xls*), *.xls*", , "Sélection du fichier", , False)
'If Fichier = False Then MsgBox "Problème lors de la séléction du fichier": Exit Sub
Application.ScreenUpdating = False
Application.DisplayAlerts = False

With Workbooks.Open(DispVb)
    With .Sheets("Liste VB")
        Lin = .Range("A" & Rows.Count).End(xlUp).Row
        .Range(.Cells(2, 1), .Cells(Lin, 3)).Copy VbDisp.Range("A2")
    End With
    .Close 0
End With
With VbDisp

If .FilterMode = True Then .ShowAllData
    W = .Range(.[A2], .Range("C" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
Application.ScreenUpdating = False
Application.DisplayAlerts = False
With Workbooks.Open(Fichier)
   
    Call Get_Plaquette_VB(.Sheets("Echéancier"), .Sheets("Macros"), Qlp)
      With Qlp
     If .FilterMode = True Then .ShowAllData
    TT = .Range(.[A3], .Range("u" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
    End With
    
    With .Sheets("Echéancier")
    If .FilterMode = True Then .ShowAllData
       V = .Range(.[A6], .Range("FQ" & .Range("A" & Rows.Count).End(xlUp).Row)).Value 'changement du FK ---->FQ
    End With
    
     For i = LBound(TT) To UBound(TT)
    Temps = CheckVbResp(W, TT(i, 2), 1, 3)
  
    TT(i, 15) = IIf(Temps(0) = "", "", Temps(0)) 'Responsable VB
         Temps = CheckVbResp(W, TT(i, 2), 1, 3)
        Temp = Extract(V, TT(i, 2), TT(i, 3), 5, 2, 168, 100, 52, 14, 61, 60, 43, 44, 45, 103, 166)
            TT(i, 16) = IIf(Temp(11) = "", "-", Temp(11))  'Fin Manuf
            TT(i, 17) = IIf(Temp(10) = "", "-", Temp(10))  'Fin Cheminement
            TT(i, 20) = IIf(Temp(8) = "", "-", Temp(8)) 'Date Exp
            TT(i, 18) = IIf(Temp(10) = "", "-", Temp(10)) 'Date Exp
            TT(i, 7) = IIf(Temp(4) = "", "-", Temp(4)) 'Progess
            TT(i, 8) = IIf(Temp(3) = "", "-", Temp(3)) 'Phase
            TT(i, 19) = IIf(Temp(3) = "", "-", Temp(3)) 'Phase
            TT(i, 9) = IIf(Temp(2) = "", "-", Temp(2)) 'TR CUT
            TT(i, 10) = IIf(Temp(1) = "", "-", Temp(1)) 'TR CUT Réel
            TT(i, 11) = IIf(Temp(0) = "", "-", Temp(0)) 'TR DEL
            TT(i, 21) = IIf(Temp(7) = "", "-", Temp(7))
    Next
    Qlp.[A3].Resize(UBound(TT), UBound(TT, 2)) = TT
  If IS_Only Then GoTo Exit_As
    For i = LBound(T) To UBound(T)
    Temps = CheckVbResp(W, T(i, 2), 1, 3)
    T(i, 24) = IIf(Temps(0) = "", "", Temps(0)) 'Responsable VB
     If T(i, 8) <> "Livré" Then
        Temp = Extract(V, T(i, 2), T(i, 3), 5, 2, 168, 8, 52, 14, 61, 60, 43, 44, 45, 103, 166)
        T(i, 26) = IIf(Temp(11) = "", "-", Temp(11)) 'Fin Manuf
        T(i, 30) = IIf(Temp(10) = "", "-", Temp(10)) 'Fin Cheminement
        T(i, 29) = IIf(Temp(8) = "", "-", Temp(8)) 'Date Exp
        T(i, 28) = IIf(Temp(7) = "", "-", Temp(7)) 'Taille
        'T(i, 4) = IIf(Temp(6) = "", "-", Temp(6)) 'index
        T(i, 6) = IIf(Temp(5) = "", "-", Temp(5)) 'GAP
        T(i, 7) = IIf(Temp(4) = "", "-", Temp(4)) 'Progess
        T(i, 8) = IIf(Temp(3) = "", "-", Temp(3)) 'Phase
        T(i, 9) = IIf(Temp(2) = "", "-", Temp(2)) 'TR CUT
        T(i, 10) = IIf(Temp(1) = "", "-", Temp(1)) 'TR CUT Réel
        T(i, 11) = IIf(Temp(0) = "", "-", Temp(0)) 'TR DEL
        If T(i, 10) = "0" Then T(i, 10) = "-"
        Txt = Left(T(i, 8), 3)
        If Txt = "00." Then T(i, 8) = "A Lancer/Lancement"
        If Txt = "01." Then T(i, 8) = "Kitting/Servi"
        If Txt = "02." Then T(i, 8) = "1ers Boûts"
        If Txt = "03." Then T(i, 8) = "Cheminement"
        If Txt = "04." Then T(i, 8) = "Contrôle Morpho"
        If Txt = "05." Then T(i, 8) = "2èmes Boûts"
        If Txt = "06." Then T(i, 8) = "Contrôle Final"
        If Txt = "07." Then T(i, 8) = "OSW"
        If Txt = "08." Then T(i, 8) = "Test"
        If Txt = "09." Then T(i, 8) = "Finition / Mur Qualité"
        If Txt = "10." Then T(i, 8) = "FAI + Emballage"
        If Txt = "11." Then T(i, 8) = "Fin Manuf"
        If Txt = "12." Then T(i, 8) = "Livré"
     End If
        T(i, 35) = IIf(Temp(3) = "", "-", Temp(3)) 'Phase
    Next i
  Main.[A3].Resize(UBound(T), UBound(T, 2)) = T
    For J = LBound(Y) To UBound(Y)
    Temps = CheckVbResp(W, Y(J, 2), 1, 3)
    Y(J, 23) = IIf(Temps(0) = "", "", Temps(0)) 'Responsable VB
     If Y(J, 8) <> "Livré" Then
        Temp = Extract(V, Y(J, 2), Y(J, 3), 5, 2, 168, 8, 52, 14, 61, 60, 43, 44, 45, 103, 166)
        Y(J, 25) = IIf(Temp(11) = "", "-", Temp(11)) 'Fin Manuf
        Y(J, 24) = IIf(Temp(10) = "", "-", Temp(10)) 'Fin Cheminement
        Y(J, 29) = IIf(Temp(8) = "", "-", Temp(8)) 'Date Exp
        Y(J, 6) = IIf(Temp(5) = "", "-", Temp(5)) 'GAP
        Y(J, 7) = IIf(Temp(4) = "", "-", Temp(4)) 'Progess
        Y(J, 8) = IIf(Temp(3) = "", "-", Temp(3)) 'Phase
        Y(J, 9) = IIf(Temp(2) = "", "-", Temp(2)) 'TR CUT
        Y(J, 10) = IIf(Temp(1) = "", "-", Temp(1)) 'TR CUT Réel
        Y(J, 11) = IIf(Temp(0) = "", "-", Temp(0)) 'TR DEL
        If Y(J, 10) = "0" Then Y(J, 10) = "-"
        Txt = Left(Y(J, 8), 3)
        If Txt = "00." Then Y(J, 8) = "A Lancer/Lancement"
        If Txt = "01." Then Y(J, 8) = "Kitting/Servi"
        If Txt = "02." Then Y(J, 8) = "1ers Boûts"
        If Txt = "03." Then Y(J, 8) = "Cheminement"
        If Txt = "04." Then Y(J, 8) = "Contrôle Morpho"
        If Txt = "05." Then Y(J, 8) = "2èmes Boûts"
        If Txt = "06." Then Y(J, 8) = "Contrôle Final"
        If Txt = "07." Then Y(J, 8) = "OSW"
        If Txt = "08." Then Y(J, 8) = "Test"
        If Txt = "09." Then Y(J, 8) = "Finition / Mur Qualité"
        If Txt = "10." Then Y(J, 8) = "FAI + Emballage"
        If Txt = "11." Then Y(J, 8) = "Fin Manuf"
        If Txt = "12." Then Y(J, 8) = "Livré"
     End If
         Y(J, 28) = IIf(Temp(3) = "", "-", Temp(3)) 'Phase reél
    Next J
  MeMo.[A2].Resize(UBound(Y), UBound(Y, 2)) = Y
    For K = LBound(Z) To UBound(Z)
    Temps = CheckVbResp(W, Z(K, 2), 1, 3)
    Z(K, 21) = IIf(Temps(0) = "", "", Temps(0)) 'Responsable VB
     If Y(J, 8) <> "Livré" Then
        Temp = Extract(V, Z(K, 2), Z(K, 3), 5, 2, 168, 8, 52, 14, 61, 60, 43, 44, 45, 103, 166)
        
        Z(K, 24) = IIf(Temp(11) = "", "-", Temp(11)) 'Fin Manuf
        Z(K, 23) = IIf(Temp(10) = "", "-", Temp(10)) 'Fin Cheminement
        Z(K, 25) = IIf(Temp(8) = "", "-", Temp(8)) 'Date Exp
        Z(K, 6) = IIf(Temp(5) = "", "-", Temp(5)) 'GAP
        Z(K, 7) = IIf(Temp(4) = "", "-", Temp(4)) 'Progess
        Z(K, 8) = IIf(Temp(3) = "", "-", Temp(3)) 'Phase
        Z(K, 9) = IIf(Temp(2) = "", "-", Temp(2)) 'TR CUT
        Z(K, 10) = IIf(Temp(1) = "", "-", Temp(1)) 'TR CUT Réel
        Z(K, 11) = IIf(Temp(0) = "", "-", Temp(0)) 'TR DEL
        If Z(K, 10) = "0" Then Z(K, 10) = "-"
        Txt = Left(Z(K, 8), 3)
        If Txt = "00." Then Z(K, 8) = "A Lancer/Lancement"
        If Txt = "01." Then Z(K, 8) = "Kitting/Servi"
        If Txt = "02." Then Z(K, 8) = "1ers Boûts"
        If Txt = "03." Then Z(K, 8) = "Cheminement"
        If Txt = "04." Then Z(K, 8) = "Contrôle Morpho"
        If Txt = "05." Then Z(K, 8) = "2èmes Boûts"
        If Txt = "06." Then Z(K, 8) = "Contrôle Final"
        If Txt = "07." Then Z(K, 8) = "OSW"
        If Txt = "08." Then Z(K, 8) = "Test"
        If Txt = "09." Then Z(K, 8) = "Finition / Mur Qualité"
        If Txt = "10." Then Z(K, 8) = "FAI + Emballage"
        If Txt = "11." Then Z(K, 8) = "Fin Manuf"
        If Txt = "12." Then Z(K, 8) = "Livré"
     End If
         Z(K, 28) = IIf(Temp(3) = "", "-", Temp(3)) 'Phase reél
    Next K
  McB.[A2].Resize(UBound(Z), UBound(Z, 2)) = Z
Exit_As:
.Close savechanges:=False
End With
Qlp.Protect AllowFiltering:=True, AllowUsingPivotTables:=True, Password:=m_Password_Sheets, UserInterfaceOnly:=True
MsgBox "Done!"
Application.ScreenUpdating = True
End Sub
Function Extract(T, ByVal Crit1, ByVal Crit2, Optional C1% = 1, Optional C2% = 4, Optional Res9% = 7, Optional Res8% = 6, Optional Res7% = 6, Optional Res6% = 7, _
Optional Res5% = 1, Optional Res4% = 2, Optional Res3% = 3, Optional Res2% = 4, Optional Res1% = 5, Optional Res10% = 8, _
Optional Res11% = 9, Optional Res12% = 12)

Dim i&, J&, K&, Temp(0 To 11), Boule As Boolean
  For i = LBound(T) To UBound(T)
    If T(i, C1) = Crit1 Then
      If T(i, C2) = Crit2 Then
        Temp(0) = T(i, Res1): Temp(1) = T(i, Res2): Temp(2) = T(i, Res3): Temp(3) = T(i, Res4): _
        Temp(4) = T(i, Res5): Temp(5) = T(i, Res6): Temp(6) = T(i, Res7): Temp(7) = T(i, Res8): Temp(8) = T(i, Res9): Temp(10) = T(i, Res10): Temp(11) = T(i, Res11): Temp(9) = T(i, Res12)

        Exit For
      End If
    End If
  Next i
Extract = Temp
End Function
Function CheckVbResp(W, ByVal Crit1$, Optional C1% = 1, Optional Res1% = 3)
Dim i&, J&, K&, Temps(0 To 1), Boule As Boolean
  For i = LBound(W) To UBound(W)
    If W(i, C1) = Crit1 Then
        Temps(0) = W(i, Res1): Temps(1) = " "
        Exit For
    End If
  Next i
CheckVbResp = Temps
End Function

Private Sub Get_Plaquette_VB(Ech_Sheet As Worksheet, Ech_Sheet_Macro As Worksheet, Main As Worksheet)
Dim Used_Range As Range
   Dim result_Range As Range
  Dim Row_In_Range As Range
  Dim Value_In_Range As Range
  With Ech_Sheet
   On Error Resume Next: .Columns.Ungroup
   On Error Resume Next: .Columns.Hidden = False
 'If .AutoFilter = True Then .ShowAllData
  
    Set Used_Range = .UsedRange

    .Rows("5:5").AutoFilter Field:=Ech_Sheet_Macro.Cells(45, 2).Value, Criteria1:="<>Livré"
'    .Rows("5:5").AutoFilter Field:=Ech_Sheet_Macro.Cells(15, 5).Value, Criteria1:="<0.03"
  

       Set result_Range = Used_Range.SpecialCells(Excel.XlCellType.xlCellTypeVisible)
      For Each Row_In_Range In result_Range.Rows
                If Row_In_Range.Row > 5 Then
                
                Dim Gap As String
                Dim VB As String
                
                Gap = Row_In_Range.Cells(1, Ech_Sheet_Macro.Cells(24, 2).Value).Value
                VB = Row_In_Range.Cells(1, Ech_Sheet_Macro.Cells(28, 2).Value).Value
                
                If UCase(Gap) <> "SMALL" And UCase(Gap) <> "DF" And UCase(Gap) <> "ST-DF" And UCase(Gap) <> "OJT" Then
                
                With Main
                Qlp.ShowAllData
                Set Value_In_Range = .Columns(1).Find(Row_In_Range.Cells(1, 1).Value)
                If Value_In_Range Is Nothing Then
                Dim Position_Row As Integer
                Position_Row = .Columns(1).Find("").Row
                .Cells(Position_Row, 1).Value = Row_In_Range.Cells(1, 1).Value
               
                .Cells(Position_Row, 3).Value = Row_In_Range.Cells(1, Ech_Sheet_Macro.Cells(27, 2).Value).Value
                .Cells(Position_Row, 2).Value = Row_In_Range.Cells(1, Ech_Sheet_Macro.Cells(28, 2).Value).Value
                .Cells(Position_Row, 4).Value = Get_Correct_Index(Row_In_Range.Cells(1, Ech_Sheet_Macro.Cells(66, 2).Value).Value, ".")
                .Cells(Position_Row, 5).Value = Get_Correct_Index(Row_In_Range.Cells(1, Ech_Sheet_Macro.Cells(153, 2).Value).Value, "/")
                .Cells(Position_Row, 6).Value = Row_In_Range.Cells(1, Ech_Sheet_Macro.Cells(24, 2).Value).Value
                .Cells(Position_Row, 12).Value = "Plaquette route"
                .Cells(Position_Row, 13).Value = "à preparer "
                .Cells(Position_Row, 14).Value = "-"
                With .Range("A" & Position_Row & ":s" & Position_Row)
                  .CurrentRegion.Borders.LineStyle = XlLineStyle.xlContinuous
                
                  .HorizontalAlignment = xlCenter
                  .VerticalAlignment = xlCenter
                End With
                End If
                End With
                End If
             '   MsgBox Row_In_Range.Cells(1, 1).Value
                End If
      Next
      .ShowAllData
  End With

End Sub
Private Function Get_Correct_Index(Val_ue As String, Sep As String)

    Dim List() As String
    List() = Split(Val_ue, Sep)
    Get_Correct_Index = List(0)
    If Get_Correct_Index Like ("*/*") Then Get_Correct_Index = Replace(Get_Correct_Index, "/", "-")
    Erase List
    
End Function


