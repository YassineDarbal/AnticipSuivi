Sub ArchiveLignes()

Dim File As String, K As Long, Main As Worksheet, Liv As Worksheet, i As Long, P As Long, MeMo As Worksheet

Set Main = ThisWorkbook.Sheets("Main")
Set Liv = ThisWorkbook.Sheets("Livré")
Set MeMo = ThisWorkbook.Sheets("Como-Memo")
Set Plaquette = ThisWorkbook.Sheets("Plaquette")

Dim MyPassword
MyPassword = InputBox("Merci de saisie le mot de passe", "Mot de passe, accès au Macro", "********")
If Not MyPassword = "A350" Then MsgBox "Mot de passe est incorrecte !", vbInformation, "Access": Exit Sub

File = ThisWorkbook.path & "\" & "ARCHIVE ANTICIPATION LM\Archive_Suivi_Anticipations_DQN-RDR.xlsx"
Application.ScreenUpdating = False
With Workbooks.Open(File)
With .Sheets("Main")
    For K = 3 To Main.Range("A" & Rows.Count).End(xlUp).Row
        If Main.Cells(K, 8) = "Livré" Then
        P = .Range("A" & Rows.Count).End(xlUp).Row + 1
        .Range("A" & P) = Now
        Main.Range(Main.Cells(K, 1), Main.Cells(K, 30)).Copy
        .Cells(P, 2).PasteSpecial Paste:=xlPasteValues
            With Liv
            i = .Range("A" & Rows.Count).End(xlUp).Row + 1
                .Cells(i, 1) = "MSN" & Main.Cells(K, 3) & "/" & Main.Cells(K, 2)
                .Cells(i, 2) = Main.Cells(K, 2)
                .Cells(i, 3) = Main.Cells(K, 3)
                .Cells(i, 4) = Main.Cells(K, 8)
                .Cells(i, 5) = Now
            End With
         Main.Rows(K).EntireRow.Delete
        K = K - 1
        End If
    Next K
End With
With .Sheets("Como-Memo")
    For K = 2 To MeMo.Range("A" & Rows.Count).End(xlUp).Row
        If MeMo.Cells(K, 8) = "Livré" Then
        P = .Range("A" & Rows.Count).End(xlUp).Row + 1
       
        MeMo.Range(MeMo.Cells(K, 1), MeMo.Cells(K, 28)).Copy
        .Cells(P, 1).PasteSpecial Paste:=xlPasteValues
            With Liv
            i = .Range("A" & Rows.Count).End(xlUp).Row + 1
                .Cells(i, 1) = "MSN" & MeMo.Cells(K, 3) & "/" & MeMo.Cells(K, 2)
                .Cells(i, 2) = MeMo.Cells(K, 2)
                .Cells(i, 3) = MeMo.Cells(K, 3)
                .Cells(i, 4) = MeMo.Cells(K, 8)
                .Cells(i, 5) = Now
            End With
        MeMo.Rows(K).EntireRow.Delete
        K = K - 1
        End If
    Next K
End With
' With .Sheets("Plaquette")

    ' For K = 3 To Plaquette.Range("A" & Rows.Count).End(xlUp).Row
        ' If Plaquette.Cells(K, 19) = "Livré" Then
       '  P = .Range("A" & Rows.Count).End(xlUp).Row + 1
        '.Range("A" & P) = Now
       '  Plaquette.Range(Plaquette.Cells(K, 1), Plaquette.Cells(K, 28)).Copy .Range("A" & P)
            '  With Liv
           '  I = .Range("A" & Rows.Count).End(xlUp).Row + 1
             '    .Cells(I, 1) = "MSN" & Plaquette.Cells(K, 3) & "/" & Plaquette.Cells(K, 2)
              '  .Cells(I, 2) = Plaquette.Cells(K, 2)
               ' .Cells(I, 3) = Plaquette.Cells(K, 3)
              '  .Cells(I, 4) = Plaquette.Cells(K, 8)
               ' .Cells(I, 5) = Now
            'End With
       ' Plaquette.Cells(K, 10).EntireRow.Delete
      '  K = K - 1
       ' End If
   ' Next K
'End With

.Close 1
End With
Application.ScreenUpdating = True

End Sub
