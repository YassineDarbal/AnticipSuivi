Sub Extract_D2D3_MAJ()
Dim D2D3 As Worksheet, Data As Worksheet, CellFind As Range, Main As Worksheet, Liv As Worksheet, ColCheck As Range, q As Long, s As Long, LinMcb As Long
Dim L As Long, K As Long, J As Long, col As Range, i As Long, Sap As Worksheet, VB As String, MeMo As Worksheet
Dim MonDico1 As Object, MonDico2 As Object, MonDico3 As Object, MonDico4 As Object, Lin As Long, P As Long, r As Long, LinCm As Long, McB As Worksheet
Dim Fichier
Dim Y As Variant
Dim Stat As Long, Prog As Long, NGam As Long, Dat As Long, MSN As Long, Harness As Long, Typ As Long, RefDoc As Long, Install As Long
Dim Dht As Long, Cat As Long, LastStat As Long
Dim MyPassword

'MyPassword = InputBox("Merci de saisie le mot de passe", "Mot de passe, accès au Macro", "********")
'If Not MyPassword = "MTDA350" Then MsgBox "Mot de passe est incorrecte !", vbInformation, "Access": Exit Sub

s = Timer
Set Liv = ThisWorkbook.Sheets("Livré")
Set Sap = ThisWorkbook.Sheets("SAP")
Set Data = ThisWorkbook.Sheets("LM Data")
Set D2D3 = ThisWorkbook.Sheets("D2D3")
Set McB = ThisWorkbook.Sheets("MCB")
Set Main = ThisWorkbook.Sheets("Main")
Set MeMo = ThisWorkbook.Sheets("Como-Memo")
Set MonDico1 = CreateObject("Scripting.Dictionary")
Set MonDico2 = CreateObject("Scripting.Dictionary")
Set MonDico3 = CreateObject("Scripting.Dictionary")
Set MonDico4 = CreateObject("Scripting.Dictionary")


If Main.FilterMode = True Then Main.ShowAllData
L = D2D3.Range("A" & Rows.Count).End(xlUp).Row
If L > 2 Then D2D3.Rows("2:" & L).Delete Shift:=xlUp

Fichier = Application.GetOpenFilename("Excel (*.xlsx), *.xlsx", , "Sélection du fichier", , False)
If Fichier = False Then MsgBox "Problème lors de la séléction du fichier": Exit Sub
Application.ScreenUpdating = False
Application.DisplayAlerts = False


' --------------- " Extract Data D2D3 " ------------------'
    With Workbooks.Open(Fichier)
        With .Sheets("Rapport1")
        
        If .FilterMode = True Then .ShowAllData
            With .Columns("T:T")
            
             .Replace What:="*DQN ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
             .Replace What:="*DQN", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
             .Replace What:="*ANTICIPATION ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
             
             
            End With
            
            
            .Columns("C:C").Replace What:="Coordination MEMO", Replacement:="MEMO", LookAt:=xlPart, SearchOrder:=xlByRows
            .Columns("B:B").Replace What:="CM", Replacement:="MEMO", LookAt:=xlPart, SearchOrder:=xlByRows
            
            
            With .Columns("K:K")
             .Replace What:="00", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
            End With
            
            
            Dim Last_Row As Long
            
            Last_Row = IIf(.Range("A" & Rows.Count).Value = vbNullString, .Range("A" & Rows.Count).End(xlUp).Row, Rows.Count)
            Y = .Range(.[A2], .Range("AC" & Last_Row)).Value
            K = 2
            
            
            Set rng = .Range("A1:AZ1") 'Plage de recherche
            
            Set ColCheck = rng.Find("Last Status (French)", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then Stat = ColCheck.Column 'statut
            
            Set ColCheck = rng.Find("Harness program", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then Prog = ColCheck.Column 'Programme
            
            Set ColCheck = rng.Find("Jobcard Number", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then NGam = ColCheck.Column 'N° de gamme

            Set ColCheck = rng.Find("Creation date (DD/MM/YYYY)", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then Dat = ColCheck.Column 'Date de création

            Set ColCheck = rng.Find("Aircraft number", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then MSN = ColCheck.Column 'N° d'avion

            Set ColCheck = rng.Find("Harness", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then Harness = ColCheck.Column 'ID de l'Harnais

            Set ColCheck = rng.Find("Trim Major document_Type (French)", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then Typ = ColCheck.Column ' Type de Ref.Doc

            Set ColCheck = rng.Find("Reference", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then RefDoc = ColCheck.Column 'Evenement (DQN,RDR , MEMO ...)
            
            Set ColCheck = rng.Find("Installation site", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then Install = ColCheck.Column 'Site d'install BU
            
            Set ColCheck = rng.Find("Preparation Realtime", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then Dht = ColCheck.Column 'Temps de gestion et/ou préparation de la gamme
            
            Set ColCheck = rng.Find("Category", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then Cat = ColCheck.Column 'catigorie de la gamme
            
            Set ColCheck = rng.Find("Update date of last status (DD/MM/YYYY HH:MM:SS)", LookAt:=xlWhole)
            If Not ColCheck Is Nothing Then LastStat = ColCheck.Column 'Date/Heure de derniére MAJ
            
            For J = LBound(Y) To UBound(Y)
            
             If Y(J, Typ - 1) <> "DIFF" And Y(J, Typ - 1) <> "Concessions" And Y(J, Typ - 1) <> "NC" And Y(J, RefDoc) <> "CAIR CREATION" And _
             Y(J, RefDoc) <> "JIG INDEX MANUF" And Len(Y(J, RefDoc)) > 3 And InStr(1, Y(J, RefDoc), "Engineering Neutralization") = 0 And _
             InStr(1, Y(J, RefDoc), "OSW RELATED TO") = 0 Then
             
                Set CellFind = Data.Range("A1:A800").Find(Y(J, Harness) & "VB", LookAt:=xlWhole)
                If Not CellFind Is Nothing Then
                 If Y(J, MSN) > 300 Or Y(J, MSN) = 65 Or Y(J, MSN) = 71 Then
                  If Not Y(J, RefDoc) = "" And (InStr(1, Y(J, NGam), "C1ML") Or InStr(1, Y(J, NGam), "VMEL")) Then '
                    With D2D3
                     .Cells(K, 1) = Y(J, Stat) 'Statut
                     .Cells(K, 2) = Y(J, Prog) 'Prog. Client
                     .Cells(K, 3) = Y(J, NGam) 'N° Gamme
                     .Cells(K, 4) = Y(J, Dat) 'Date
                     .Cells(K, 5) = Y(J, MSN) 'N° avion
                     
                     'If Y(J, MSN) = 5 Then .Cells(K, 5).Value = "R005"
                     
                     .Cells(K, 6) = Y(J, Harness) & "VB" 'Harnais
                     .Cells(K, 7) = Y(J, Typ) 'Type
                     If InStr(1, .Cells(K, 3), "VMEL") Then .Cells(K, 7) = "MCB"
                     .Cells(K, 8) = Y(J, RefDoc) 'Ref. Doc.
                     .Cells(K, 9) = Y(J, Install) 'Site installation
                     .Cells(K, 10) = Y(J, NGam) & "/" & .Cells(K, 6) ' Key
                     .Cells(K, 12) = Y(J, Dht) 'DHT
                     .Cells(K, 13) = Y(J, Cat) 'Catigories
                     If Y(J, 27) = "" Then .Cells(K, 13) = "-"
                     If Y(J, 14) = "" Then .Cells(K, 12) = "-"
                     .Cells(K, 15) = Y(J, LastStat) 'Last Stat
                     K = K + 1
                    End With
                  End If
                 End If
                End If
                
                
             End If
            Next J
            '.[A2].Resize(UBound(Y), UBound(Y, 2)) = Y
        End With
     .Close savechanges:=False
    End With
Application.ScreenUpdating = False
    With D2D3
          Set col = .Range("J:J")
          For Each c In col.SpecialCells(xlCellTypeConstants, 23)
            If Not MonDico1.Exists(c.Value) Then MonDico1.Add c.Value, c.Address
          Next
    End With
    With Main
      For Each c In .Range("A:A").SpecialCells(xlCellTypeConstants, 23)
        If Not MonDico2.Exists(c.Value) Then MonDico2.Add c.Value, c.Address
      Next
    End With

    With MeMo
      For Each c In .Range("A:A").SpecialCells(xlCellTypeConstants, 23)
        If Not MonDico3.Exists(c.Value) Then MonDico3.Add c.Value, c.Address
      Next
    End With
    With McB
      For Each D In .Range("A:A").SpecialCells(xlCellTypeConstants, 23)
        If Not MonDico4.Exists(D.Value) Then MonDico4.Add D.Value, D.Address
      Next
    End With
    With D2D3
      For i = 2 To .Range("A" & Rows.Count).End(xlUp).Row
          If MonDico2.Exists(.Cells(i, 10).Value) Then
              .Cells(i, 11) = "OLD-DQN"
           Else
             .Cells(i, 11) = "NEW-DQN"
          End If
          If .Cells(i, 7).Value = "MEMO" Then
            If MonDico3.Exists(.Cells(i, 10).Value) Then
                .Cells(i, 11) = "OLD-CM"
             Else
               .Cells(i, 11) = "NEW-CM"
            End If
          End If
          If .Cells(i, 7).Value = "MCB" Then
            If MonDico4.Exists(.Cells(i, 10).Value) Then
                .Cells(i, 11) = "OLD-MCB"
             Else
               .Cells(i, 11) = "NEW-MCB"
            End If
          End If
      Next i

    End With
' --------------- " MAJ du Suivi " ------------------'
With D2D3
    Z = .Range(.[A2], .Range("O" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
With Liv
    W = .Range(.[A2], .Range("E" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
For L = LBound(Z) To UBound(Z)
    Temps = CheckDel(W, Z(L, 6), Z(L, 5), 2, 3, 4)
    Z(L, 14) = IIf(Temps(0) = "", "-", Temps(0))
    If Z(L, 14) = "" Then Z(L, 14) = "-"
Next L
D2D3.[A2].Resize(UBound(Z), UBound(Z, 2)) = Z
K = Main.Range("A" & Rows.Count).End(xlUp).Row + 1
E = MeMo.Range("A" & Rows.Count).End(xlUp).Row + 1
q = McB.Range("A" & Rows.Count).End(xlUp).Row + 1
 For J = LBound(Z) To UBound(Z)
                 If Z(J, 11) = "NEW-DQN" And Z(J, 14) <> "Livré" Then 'Start IF
                    With Sap
                           P = .Range("A" & Rows.Count).End(xlUp).Row + 1
                           Set CellFind = .Range("A:A").Find(Z(J, 8), LookAt:=xlWhole)
                           If CellFind Is Nothing Then .Cells(P, 1) = Z(J, 8)
                    End With
                    With Main
                        .Cells(K, 1) = Z(J, 10) 'KEY
                        .Cells(K, 2) = Z(J, 6) 'Harnais
                        .Cells(K, 3) = Z(J, 5) 'MSN
                        .Cells(K, 5) = Z(J, 2) 'Prog
                        .Cells(K, 12) = Z(J, 8) 'Ref.Doc
                        .Cells(K, 15) = Z(J, 3) 'N° Gamme
                        .Cells(K, 16) = Z(J, 1) 'Statut
                        .Cells(K, 17) = Z(J, 4) 'Date
                        .Cells(K, 18) = Z(J, 9) 'Site install
                        .Cells(K, 19) = Z(J, 12) 'Dht
                        .Cells(K, 20) = Z(J, 13) 'Gammes Stat
                        .Cells(K, 21) = "A préparer GAP"
                        If .Cells(K, 21) = "A préparer GAP" And .Cells(K, 16) = "D2 SR" Then .Cells(K, 21) = "SR"
                        .Cells(K, 12).Interior.ColorIndex = 27
                        .Cells(K, 15).Interior.ColorIndex = 27
                        .Cells(K, 16).Interior.ColorIndex = 27
                        .Cells(K, 34) = Z(J, 15) 'Last Stat
                            With .Cells(K, 21).Validation
                               .Delete
                               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                               Operator:=xlBetween, Formula1:= _
                               "A préparer GAP,A préparer OSW,Diffusée GAP,Diffusée OSW,Intégrée au LCMT,Intégrée CIRCE,Para,Intégrée 3D,Bloquée,SR,Annulée"
                            End With
                            With .Cells(K, 24).Validation
                               .Delete
                               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                               Operator:=xlBetween, Formula1:= _
                               "H.AKEL,L.ALAOUI,M.AIT OUCHEN,L.LAAKILI,N.BEDDOUZ,R.AIT OUCHEN,Z.ELKASMI,G.OUJIL ,H.KABBABE,I.ELJOUHRI,A.EL MOUDDEN,H.EL MENNOUNY,A.CAID,S.CHAREF,K.ZOUANAT"
                            End With
                      If Left(.Cells(K, 16), 2) = "D3" Then .Cells(K, 21) = "Para"
                     K = K + 1
                    End With
                 ElseIf Z(J, 11) = "OLD-DQN" Then '2nd IF
                        With Main
                           Set CellFind = .Range("A:A").Find(Z(J, 10), LookAt:=xlWhole)
                           If Not CellFind Is Nothing Then
                            r = CellFind.Row
                            .Cells(r, 16) = Z(J, 1) 'Statut
                            .Cells(r, 19) = Z(J, 12) 'Dht
                            .Cells(r, 18) = Z(J, 9)  'Site install
                            .Cells(r, 20) = Z(J, 13)  'Site install
                            .Cells(r, 34) = Z(J, 15) 'Last Stat
                            'If .Cells(R, 21) = "A préparer GAP" And .Cells(R, 16) = "D2 SR" Then .Cells(R, 21) = "SR"
                            .Cells(r, 12).Interior.ColorIndex = 0
                            .Cells(r, 15).Interior.ColorIndex = 0
                            .Cells(r, 16).Interior.ColorIndex = 0
                            
                           End If
                        End With
                 ' Lignes des CoMo/ MeMo
                 ElseIf Z(J, 11) = "NEW-CM" And Z(J, 14) <> "Livré" Then
                    With MeMo
                     .Cells(E, 1) = Z(J, 10) 'KEY
                     .Cells(E, 2) = Z(J, 6) 'Harnais
                     .Cells(E, 3) = Replace(CStr(Z(J, 5)), "R", "") 'MSN
                     .Cells(E, 5) = Z(J, 2) 'Prog
                     .Cells(E, 12) = Z(J, 8) 'Ref.Doc
                     .Cells(E, 15) = Z(J, 3) 'N° Gamme
                     .Cells(E, 16) = Z(J, 1) 'Statut
                     .Cells(E, 17) = Z(J, 4) 'Date
                     .Cells(E, 18) = Z(J, 9) 'Site install
                     .Cells(E, 19) = Z(J, 12) 'Dht
                     .Cells(E, 20) = "A préparer GAP" ' Statut par defaut
                    ' If .Cells(E, 20) = "A préparer GAP" And .Cells(E, 16) = "D2 SR" Then .Cells(E, 20) = "SR"
                     .Cells(E, 26) = Z(J, 15) 'Last Stat
                     .Cells(E, 12).Interior.ColorIndex = 27
                     .Cells(E, 15).Interior.ColorIndex = 27
                     .Cells(E, 16).Interior.ColorIndex = 27
                     With .Cells(E, 20).Validation
                        .Delete
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                        Operator:=xlBetween, Formula1:= _
                        "A préparer GAP,A préparer OSW,Diffusée GAP,Diffusée OSW,Intégrée au LCMT,Intégrée CIRCE,Para,Intégrée 3D,Bloquée,SR,Annulée"
                     End With
                        With .Cells(E, 23).Validation
                           .Delete
                           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                           Operator:=xlBetween, Formula1:= _
                           "H.AKEL,L.ALAOUI,M.AIT OUCHEN,L.LAAKILI,N.BEDDOUZ,R.AIT OUCHEN,Z.ELKASMI,G.OUJIL ,H.KABBABE,I.ELJOUHRI,A.EL MOUDDEN,H.EL MENNOUNY,A.CAID,S.CHAREF,K.ZOUANAT"
                        End With
                     E = E + 1
                    End With
                 ElseIf Z(J, 11) = "OLD-CM" Then
                        With MeMo
                           Set CellFind = .Range("A:A").Find(Z(J, 10), LookAt:=xlWhole)
                           If Not CellFind Is Nothing Then
                            r = CellFind.Row
                            .Cells(r, 16) = Z(J, 1) 'Statut D2D3
                            .Cells(r, 19) = Z(J, 12) 'Dht
                            .Cells(r, 18) = Z(J, 9)  'Site install
                            .Cells(r, 26) = Z(J, 15) 'Last Stat
                           ' If .Cells(R, 20) = "A préparer GAP" And .Cells(R, 16) = "D2 SR" Then .Cells(R, 20) = "SR"
                            .Cells(r, 12).Interior.ColorIndex = 0
                            .Cells(r, 15).Interior.ColorIndex = 0
                            .Cells(r, 16).Interior.ColorIndex = 0
                           End If
                        End With
                 ElseIf Z(J, 11) = "NEW-MCB" And Z(J, 14) <> "Livré" Then
                    With McB
                     .Cells(q, 1) = Z(J, 10) 'KEY
                     .Cells(q, 2) = Z(J, 6) 'Harnais
                     .Cells(q, 3) = Replace(Z(J, 5), "R", "") 'MSN
                     .Cells(q, 5) = Z(J, 2) 'Prog
                     .Cells(q, 12) = Z(J, 8) 'Ref.Doc
                     .Cells(q, 13) = Left(Right(Z(J, 3), 4), 3) 'Release
                     .Cells(q, 14) = Z(J, 3) 'N° Gamme
                     .Cells(q, 15) = Z(J, 1) 'Statut
                     .Cells(q, 16) = Z(J, 4) 'Date
                     .Cells(q, 17) = Z(J, 9) 'Site install
                     .Cells(q, 18) = Z(J, 12) 'Dht
                     .Cells(q, 20) = "A préparer GAP" ' Statut par defaut
                     .Cells(q, 26) = Z(J, 15) 'Last Stat
                     .Cells(q, 12).Interior.ColorIndex = 27
                     .Cells(q, 14).Interior.ColorIndex = 27
                     .Cells(q, 15).Interior.ColorIndex = 27
                     With .Cells(q, 20).Validation
                        .Delete
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                        Operator:=xlBetween, Formula1:= _
                        "A préparer GAP,A préparer OSW,Diffusée GAP,Diffusée OSW,Intégrée au LCMT,Intégrée CIRCE,Para,Intégrée 3D,Bloquée,SR,Annulée"
                     End With
                        With .Cells(q, 21).Validation
                           .Delete
                           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                           Operator:=xlBetween, Formula1:= _
                           "H.AKEL,L.ALAOUI,M.AIT OUCHEN,L.LAAKILI,N.BEDDOUZ,R.AIT OUCHEN,Z.ELKASMI,G.OUJIL ,H.KABBABE,I.ELJOUHRI,A.EL MOUDDEN,H.EL MENNOUNY,A.CAID,S.CHAREF,K.ZOUANAT"
                        End With
                     q = q + 1
                    End With
                 ElseIf Z(J, 11) = "OLD-MCB" Then
                        With McB
                           Set CellFind = .Range("A:A").Find(Z(J, 10), LookAt:=xlWhole)
                           If Not CellFind Is Nothing Then
                            s = CellFind.Row
                            .Cells(s, 15) = Z(J, 1) 'Statut D2D3
                            .Cells(s, 18) = Z(J, 12) 'Dht
                            .Cells(s, 17) = Z(J, 9)  'Site install
                            .Cells(s, 26) = Z(J, 15) 'Last Stat
                            .Cells(s, 12).Interior.ColorIndex = 0
                            .Cells(s, 14).Interior.ColorIndex = 0
                            .Cells(s, 15).Interior.ColorIndex = 0
                           End If
                        End With
                        
                 End If
 Next J
D2D3.[A2].Resize(UBound(Z), UBound(Z, 2)) = Z

With Main
  T = .Range(.[L3], .Range("AA" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
With Sap
  V = .Range(.[A2], .Range("E" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
End With
'For J = LBound(T) To UBound(T)
    'Temp = DqnStatus(V, T(J, 1), 1, 2, 4, 5)
    'T(J, 3) = IIf(Temp(2) = "", "-", Temp(2))
    'T(J, 2) = IIf(Temp(1) = "", "-", Temp(1))
    'T(J, 16) = IIf(Temp(0) = "", "-", Temp(0))
'Next J


Main.[L3].Resize(UBound(T), UBound(T, 2)) = T
Call NewOrOld_DQN
Lin = Application.WorksheetFunction.CountIfs(D2D3.Range("K:K"), "NEW-DQN", D2D3.Range("N:N"), "-")
LinCm = Application.WorksheetFunction.CountIfs(D2D3.Range("K:K"), "NEW-CM", D2D3.Range("N:N"), "-")
LinMcb = Application.WorksheetFunction.CountIfs(D2D3.Range("K:K"), "NEW-MCB", D2D3.Range("N:N"), "-")
L = D2D3.Range("A" & Rows.Count).End(xlUp).Row
If L > 2 Then D2D3.Rows("2:" & L).Delete Shift:=xlUp
Main.Activate

Application.ScreenUpdating = True
MsgBox "Analyse effectué en : " & Timer - s & " (Sec) " & vbCrLf & Lin & " DQN/RDR" & vbCrLf & LinCm & "  MEMO/COMO" & vbCrLf & LinMcb & "  MCB"
End Sub
Function DqnStatus(T, ByVal Crit1$, Optional C1% = 1, Optional Res3% = 4, Optional Res2% = 4, Optional Res1% = 5)

Dim i&, J&, K&, Temp(0 To 3), Boule As Boolean
  For i = LBound(T) To UBound(T)
    If T(i, C1) = Crit1 Then
        Temp(0) = T(i, Res1): Temp(1) = T(i, Res2): Temp(2) = T(i, Res3): Temp(3) = " "
        Exit For
    End If
  Next i
DqnStatus = Temp
End Function

Function CheckDel(W, ByVal Crit1$, ByVal Crit2$, Optional C1% = 1, Optional C2% = 1, Optional Res1% = 5)

Dim i&, J&, K&, Temps(0 To 1), Boule As Boolean
  For i = LBound(W) To UBound(W)
    If W(i, C1) = Crit1 Then
     If W(i, C2) = Crit2 Then
        Temps(0) = W(i, Res1): Temps(1) = " "
        Exit For
     End If
    End If
  Next i
CheckDel = Temps
End Function
Sub NewOrOld_DQN()
Dim col As Range, i As Long, Sap As Worksheet, VB As String, Main As Worksheet
Dim MonDico1 As Object, MonDico2 As Object, Lin As Long, P As Long, r As Long, LinCm As Long
Set Sap = Sheets("SAP")
Set Main = Sheets("Main")

Set MonDico1 = CreateObject("Scripting.Dictionary")
Set MonDico2 = CreateObject("Scripting.Dictionary")
Set MonDico3 = CreateObject("Scripting.Dictionary")

    With Sap
      If .FilterMode = True Then Main.ShowAllData
          Set col = .Range("A2:A" & .Range("A" & Rows.Count).End(xlUp).Row)
          For Each c In col.SpecialCells(xlCellTypeConstants, 23)
            If Not MonDico1.Exists(c.Value) Then MonDico1.Add c.Value, c.Address
          Next
    End With
    
    With Main
      For i = 3 To .Range("A" & Rows.Count).End(xlUp).Row
          If Not MonDico1.Exists(.Cells(i, 12).Value) Then
              .Cells(i, 12).Interior.ColorIndex = 28
              .Cells(i, 13).Interior.ColorIndex = 28
              .Cells(i, 14).Interior.ColorIndex = 28
           Else
             .Cells(i, 12).Interior.ColorIndex = 0
             .Cells(i, 13).Interior.ColorIndex = 0
             .Cells(i, 14).Interior.ColorIndex = 0
          End If
      Next i

    End With


End Sub







