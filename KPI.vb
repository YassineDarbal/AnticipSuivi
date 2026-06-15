Sub DatasGraph_KPI()
Dim Main As Worksheet, Dash As Worksheet, Del As Range, Cat As Range
Dim Pic As Object, Graph As Worksheet, Phaz As Range, Gap As Range
'Dim MyPassword
'MyPassword = InputBox("Merci de saisie le mot de passe", "Mot de passe, accès au Macro", "********")
'If Not MyPassword = "MTDA350" Then MsgBox "Mot de passe est incorrecte !", vbInformation, "Access": Exit Sub

Set Graph = ThisWorkbook.Sheets("Graph")
Set Main = ThisWorkbook.Sheets("Main")
Set Dash = ThisWorkbook.Sheets("DashBoard")
Set Del = Main.Columns("K:K")
Set Cat = Main.Columns("T:T")
Set Phaz = Main.Columns("H:H")
Set Gap = Main.Columns("F:F")
Set GamSt = Main.Columns("U:U")

Application.ScreenUpdating = False
With Dash
'.ScrollArea = ""
.Rows("1:16").EntireRow.Hidden = False
For Each Pic In .Pictures
Pic.Delete
Next Pic
    For i = 3 To 8
        For col = 2 To 52
            .Cells(i, col) = Application.WorksheetFunction.CountIfs(Cat, .Cells(i, 1), Del, .Cells(2, col))
        Next col
    Next i
    For col = 55 To 67
        .Cells(3, col) = Application.WorksheetFunction.CountIfs(GamSt, .Cells(3, 54), Phaz, .Cells(2, col))
        .Cells(4, col) = Application.WorksheetFunction.CountIfs(GamSt, "Diffusée OSW", Phaz, .Cells(2, col)) + Application.WorksheetFunction.CountIfs(GamSt, "Diffusée GAP", Phaz, .Cells(2, col))
        .Cells(5, col) = Application.WorksheetFunction.CountIfs(GamSt, .Cells(5, 54), Phaz, .Cells(2, col))
        .Cells(6, col) = Application.WorksheetFunction.CountIfs(GamSt, .Cells(6, 54), Phaz, .Cells(2, col))
        .Cells(7, col) = Application.WorksheetFunction.CountIfs(GamSt, "Intégrée 3D", Phaz, .Cells(2, col)) + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée CIRCE", Phaz, .Cells(2, col))
        .Cells(8, col) = Application.WorksheetFunction.CountIfs(GamSt, "Non appliquée", Phaz, .Cells(2, col))
    Next col
    .Cells(3, 68) = Application.WorksheetFunction.Sum(Range(.Cells(3, 55), .Cells(3, 67)))
    .Cells(4, 68) = Application.WorksheetFunction.Sum(Range(.Cells(4, 55), .Cells(4, 67)))
    .Cells(5, 68) = Application.WorksheetFunction.Sum(Range(.Cells(5, 55), .Cells(5, 67)))
    .Cells(6, 68) = Application.WorksheetFunction.Sum(Range(.Cells(6, 55), .Cells(6, 67)))
    .Cells(7, 68) = Application.WorksheetFunction.Sum(Range(.Cells(7, 55), .Cells(7, 67)))
    .Cells(8, 68) = Application.WorksheetFunction.Sum(Range(.Cells(8, 55), .Cells(8, 67)))
    For col = 2 To 7
        .Cells(10, col) = Application.WorksheetFunction.CountIfs(GamSt, .Cells(10, 1), Gap, .Cells(9, col))
        .Cells(11, col) = Application.WorksheetFunction.CountIfs(GamSt, "Diffusée OSW", Gap, .Cells(9, col)) + Application.WorksheetFunction.CountIfs(GamSt, "Diffusée GAP", Gap, .Cells(9, col))
        .Cells(12, col) = Application.WorksheetFunction.CountIfs(GamSt, .Cells(12, 1), Gap, .Cells(9, col))
        .Cells(13, col) = Application.WorksheetFunction.CountIfs(GamSt, .Cells(13, 1), Gap, .Cells(9, col))
        .Cells(14, col) = Application.WorksheetFunction.CountIfs(GamSt, "Intégrée 3D", Gap, .Cells(9, col)) + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée CIRCE", Gap, .Cells(9, col)) + Application.WorksheetFunction.CountIfs(GamSt, "SR", Gap, .Cells(9, col))
        .Cells(15, col) = Application.WorksheetFunction.CountIfs(GamSt, "Non appliquée", Gap, .Cells(9, col))
    Next col
        .Cells(10, 7) = Application.WorksheetFunction.CountIfs(GamSt, .Cells(10, 1), Gap, "OSW-S11") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(10, 1), Gap, "OSW") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(10, 1), Gap, "OSW") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(10, 1), Gap, "OSW") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(10, 1), Gap, "OSW")
        .Cells(11, 7) = Application.WorksheetFunction.CountIfs(GamSt, "Diffusée OSW", Gap, "OSW-S11") + Application.WorksheetFunction.CountIfs(GamSt, "Diffusée OSW", Gap, "OSW") + Application.WorksheetFunction.CountIfs(GamSt, "Diffusée OSW", Gap, "OSW") + Application.WorksheetFunction.CountIfs(GamSt, "Diffusée OSW", Gap, "OSW") + Application.WorksheetFunction.CountIfs(GamSt, "Diffusée OSW", Gap, "OSW")
        .Cells(12, 7) = Application.WorksheetFunction.CountIfs(GamSt, .Cells(12, 1), Gap, "OSW-S11") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(12, 1), Gap, "OSW-S13") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(12, 1), Gap, "OSW-S15-A") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(12, 1), Gap, "OSW-S15-B") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(12, 1), Gap, "OSW-S16")
        .Cells(13, 7) = Application.WorksheetFunction.CountIfs(GamSt, .Cells(13, 1), Gap, "OSW-S11") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(13, 1), Gap, "OSW-S13") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(13, 1), Gap, "OSW-S15-A") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(13, 1), Gap, "OSW-S15-B") + Application.WorksheetFunction.CountIfs(GamSt, .Cells(13, 1), Gap, "OSW-S16")
        .Cells(14, 7) = Application.WorksheetFunction.CountIfs(GamSt, "Intégrée 3D", Gap, "OSW-S11") + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée CIRCE", Gap, "OSW-S11") + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée 3D", Gap, "OSW-S13") + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée CIRCE", Gap, "OSW-S13") + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée 3D", Gap, "OSW-S15-A") + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée CIRCE", Gap, "OSW-S15-A") _
        + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée 3D", Gap, "OSW-S15-B") + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée CIRCE", Gap, "OSW-S15-B") + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée 3D", Gap, "OSW-S16") + Application.WorksheetFunction.CountIfs(GamSt, "Intégrée CIRCE", Gap, "OSW-S16")
        .Cells(15, 7) = Application.WorksheetFunction.CountIfs(GamSt, "Non appliquée", Gap, "OSW-S11") + Application.WorksheetFunction.CountIfs(GamSt, "Non appliquée", Gap, "OSW-S13") + Application.WorksheetFunction.CountIfs(GamSt, "Non appliquée", Gap, "OSW-S15-A") + Application.WorksheetFunction.CountIfs(GamSt, "Non appliquée", Gap, "OSW-S15-B") + Application.WorksheetFunction.CountIfs(GamSt, "Non appliquée", Gap, "OSW-S16")
    For K = 10 To 14
        For col = 12 To 52
            .Cells(K, col) = Application.WorksheetFunction.CountIfs(Main.Columns("P:P"), .Cells(K, 11), Del, .Cells(9, col))
        Next col
    Next K
    End With
    Graph.ChartObjects("Graphique 1").Copy
    Dash.Range("A16").Select
    ActiveSheet.PasteSpecial Format:="Image (JPEG)"
    Graph.ChartObjects("Graphique 2").Copy
    Dash.Range("A35").Select
    ActiveSheet.PasteSpecial Format:="Image (JPEG)"
    Graph.ChartObjects("Graphique 3").Copy
    Dash.Range("A57").Select
    ActiveSheet.PasteSpecial Format:="Image (JPEG)"
    Graph.ChartObjects("Graphique 4").Copy
    Dash.Range("A79").Select
    ActiveSheet.PasteSpecial Format:="Image (JPEG)"
    Graph.ChartObjects("Graphique 5").Copy
    Dash.Range("A100").Select
    ActiveSheet.PasteSpecial Format:="Image (JPEG)"
    Dash.Range("A1").Select
    Dash.Rows("2:15").EntireRow.Hidden = True
    'Dash.ScrollArea = "$A$1:$AU$200"
Application.ScreenUpdating = True
End Sub
