Const Bswh$ = "Main" ' nom de la feuille où on met les résultats
Const NomGauche$ = "SOL" 'partie texte de nom de la feuille dans BDD
Sub GetindexMorpho()
Dim s As Long
  Dim TWay As String, WbDest As String, TDest As String, Tsours$
  Dim sol As String, Notfound As Boolean, i As Long, J As Long, DerL As Long
  Dim ExtSh As Worksheet, Insh As Worksheet, Disws As Worksheet
'Dim MyPassword
'
'MyPassword = InputBox("Merci de saisie le mot de passe", "Mot de passe, accès au Macro", "********")
'If Not MyPassword = "MTDA350" Then MsgBox "Mot de passe est incorrecte !", vbInformation, "Access": Exit Sub

s = Timer
Call CrSol

Tsours = "S:\Temara\D16\SUPPORT-CONFIGURATION\RESTRICTED\SUIVI-CONF-PZ\CONFIGURATION FOLLOW-UP-GLOBAL.xlsx"
WbDest = "ConfigurationDataBase.xlsx"
TWay = Environ("temp")
TDest = TWay & WbDest
On Error Resume Next
Kill TDest
On Error GoTo 0
FileCopy Source:=Tsours, Destination:=TDest
  Application.ScreenUpdating = False
    Set wb1 = ThisWorkbook
    Set wb2 = Workbooks.Open(TDest)
        With ThisWorkbook.Worksheets(Bswh)
           ' .Columns("G:G").NumberFormat = "@"
          Dim Last_Row As Long
          Last_Row = IIf(.Range("A" & Rows.Count).Value = vbNullString, .Range("A" & Rows.Count).End(xlUp).Row, Rows.Count)
          T = .Range(.[B3], .Range("K" & Last_Row)).Value
        End With
            With wb2
              For i = LBound(T) To UBound(T)
              If T(i, 2) = "-" Then
                 T(i, 2) = "-"
              ElseIf Len(T(i, 3)) = 3 Then
                With .Sheets(NomGauche & T(i, 3))
                    If .FilterMode = True Then
                        .ShowAllData
                    End If
          Last_Row = IIf(.Range("A" & Rows.Count).Value = vbNullString, .Range("A" & Rows.Count).End(xlUp).Row, Rows.Count)
                    V = .Range(.[A2], .Range("N" & Last_Row)).Value
                  End With
                    If T(i, 10) <= ThisWorkbook.Worksheets(Bswh).Cells(1, 2).Value Then
                      Tr = T(i, 10)
                    Else
                      Tr = ThisWorkbook.Worksheets(Bswh).Cells(1, 2).Value
                    End If
                  Temp = RechIndice(V, T(i, 1), Tr, 1, 7, 7, 10)
                  T(i, 3) = IIf(Temp(0) = "", "-", Temp(0)) 'index morpho colonne J dans bdd, et traitemment colonne 7
                  End If
              Next i
              .Close 0
            End With
  ThisWorkbook.Worksheets(Bswh).[B3].Resize(UBound(T), UBound(T, 2)) = T
  ThisWorkbook.Worksheets(Bswh).Columns("D:D").Replace What:="/", Replacement:="-", LookAt:=xlPart, SearchOrder:=xlByRows
  ThisWorkbook.Worksheets(Bswh).Columns("D:D").Replace What:=".*", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
 Application.ScreenUpdating = True
 MsgBox "Done in " & Timer - s & " Sec"
End Sub
Function RechIndice(T, ByVal Crit1$, ByVal Crit2&, Optional C1% = 1, Optional C2% = 4, _
                      Optional Res2% = 4, Optional Res1% = 5)
  'T est 1 tableau à 2 dimension, Crit1 est le nom du VB, crit2 est 1 nombre de traitement
  'C1 et C2 sont les colonnes sur lesquelles on veut effectuer le critère,
  'RES1 et res2 sont les numéros de colonne de résultat
Dim i&, J&, K&, Temp(0 To 2), Boule As Boolean
  For i = LBound(T) To UBound(T)
    If T(i, C1) = Crit1 Then
      If T(i, C2) = Crit2 Then
        Temp(0) = T(i, Res1): Temp(1) = T(i, Res2): Temp(2) = " "
        Exit For
      Else
        For J = 1 To Crit2
          For K = LBound(T) To UBound(T)
            If T(K, C1) = Crit1 Then
              If T(K, C2) = Crit2 - J Then
                Temp(0) = T(K, Res1): Temp(1) = T(K, Res2): Temp(2) = " "
                Boule = True
                Exit For
              End If
            End If
          Next K
          If Boule Then Exit For
        Next J
      End If
    End If
  Next i
RechIndice = Temp
End Function
Private Sub CrSol()
Dim Bswh As Worksheet, i As Long
Dim TWay As String, WbDest As String, TDest As String, Tsours$

Set Bswh = ThisWorkbook.Sheets("Main") ' nom de la feuille où on met les résultats

Tsours = "S:\Temara\D16\SUPPORT-CONFIGURATION\RESTRICTED\INPUT\01 - CONF REPORT\OLD Conf Report\01 - Conf Report Global\Conf Report Global_A350_DS-Only.xlsx"
WbDest = "-ConfReportGlobal.xlsx"
TWay = Environ("temp")
TDest = TWay & WbDest
On Error Resume Next
Kill TDest
On Error GoTo 0
FileCopy Source:=Tsours, Destination:=TDest

Application.ScreenUpdating = False
        With Bswh
          T = .Range(.[B3], .Range("K" & .Range("A" & Rows.Count).End(xlUp).Row)).Value
          '.Range("L2:L65000").NumberFormat = "@"
        End With
    Set wb2 = Workbooks.Open(TDest)
    With wb2
        With .Worksheets("Conf Report")
                    If .FilterMode = True Then
                        .ShowAllData
                    End If
            Dim Last_Row As Long
            Last_Row = IIf(.Range("A" & Rows.Count).Value = vbNullString, .Range("A" & Rows.Count).End(xlUp).Row, Rows.Count)
            .Columns("D:D").NumberFormat = "@"
           V = .Range(.[A2], .Range("I" & Last_Row)).Value

        End With
              For i = LBound(T) To UBound(T)
              If T(i, 3) = "-" Or T(i, 3) = "" Then
                    If T(i, 10) <= Bswh.Cells(1, 2).Value Then
                      Tr = T(i, 10)
                    Else
                      Tr = Bswh.Cells(1, 2).Value
                    End If
                    Avion = "|" & T(i, 2) & "|"
                        Temp = RechSol(V, T(i, 1), Tr, Avion, 2, 6, 8, 4, 5, 7)
                        T(i, 3) = IIf(Temp(3) = "", "-", Temp(3)) 'index morpho colonne J dans bdd, et traitemment colonne 7
                        'T(I, 8) = IIf(Temp(1) = "", "None!", Temp(1))
                        'T(I, 12) = IIf(Temp(0) = "", "None!", Temp(0))
                End If
                Next i
              .Close 0
     End With
    Bswh.[B3].Resize(UBound(T), UBound(T, 2)) = T
Application.ScreenUpdating = True
End Sub
Function RechSol(T, ByVal Crit1$, ByVal Crit2&, ByVal Crit3, Optional C1% = 1, Optional C2% = 4, Optional C3% = 5, _
                      Optional Res3% = 3, Optional Res2% = 4, Optional Res1% = 5)

Dim i&, J&, K&, Temp(0 To 3), Boule As Boolean
  For i = LBound(T) To UBound(T)
    If T(i, C1) = Crit1 Then
     If T(i, C2) = Crit2 Then
      If InStr(1, T(i, C3), Crit3) Then
            Temp(0) = T(i, Res1): Temp(1) = T(i, Res2): Temp(2) = " ": Temp(3) = T(i, Res3)
            Exit For
      End If
     End If
    End If
  Next i
RechSol = Temp
End Function


