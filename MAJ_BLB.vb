Public Sub Run()
DoEvents
Progress.Show vbModeless
DoEvents
End Sub

Public Sub Main_BLB()
 'On Error GoTo out
 Dim WB As Workbook
 Dim WB_BD As Workbook
 
 Set WB = ActiveWorkbook
 Dim AppExcel As Excel.Application
 Set AppExcel = New Excel.Application
 Set WB_BD = AppExcel.Workbooks.Open(WB.Sheets("DashBoard").Cells(1, 35).Value, ReadOnly = True)
 AppExcel.Visible = False

 Progress.L2 = "Main"
 Call Get_BLB_ST(WB, WB_BD, "Main", 36, 3, 2, 12, 2)
 Progress.L2 = "Como-Memo"
 Call Get_BLB_ST(WB, WB_BD, "Como-Memo", 30, 3, 2, 12, 1)




'out:
 WB_BD.Close (0)
 AppExcel.Quit
 Application.Visible = True

End Sub


Private Sub Get_BLB_ST(WB As Workbook, WB_BD As Workbook, N_Sheet As String, Pos1 As Integer, Pos2 As Integer, Pos3 As Integer, Pos4 As Integer, Row_P As Integer)

 
 Dim ws As Worksheet
 Dim RN As Range
 Dim Result As Range
 Dim R_ow As Range

 Set ws = WB.Sheets(N_Sheet)

On Error Resume Next: ws.Unprotect
 With ws

  If .FilterMode = True Then .ShowAllData
 
   Set RN = .UsedRange
   .Rows(Row_P & ":" & Row_P).AutoFilter Field:=Pos1, Criteria1:="="
   Set Result = RN.SpecialCells(Excel.XlCellType.xlCellTypeVisible)
  
   For Each R_ow In Result.Rows
    If R_ow.Row > Row_P Then
     Dim P_rog As Integer
On Error Resume Next:     P_rog = Math.Round((R_ow.Row / Result.Rows.Count * 100))
    
     DoEvents
     If R_ow.Cells(1, 1).Value <> "" Then
      Progress.L3 = "MSN" & R_ow.Cells(1, Pos2).Value & "/" & R_ow.Cells(1, Pos3).Value
      R_ow.Cells(1, Pos1).Value = Get_BLBST(WB_BD, "MSN" & R_ow.Cells(1, Pos2).Value & "/" & R_ow.Cells(1, Pos3).Value, R_ow.Cells(1, Pos3).Value & "/" & R_ow.Cells(1, Pos4).Value)
     Else
     Exit For
     End If
    End If
   Next
   If .FilterMode = True Then .ShowAllData
    
 End With


End Sub


Private Function Get_BLBST(WB_BD As Workbook, Key1 As String, Key2 As String)

 Dim ws As Worksheet
 Dim RN As Range
 Dim Result As Range
 Dim R_ow As Range

 Get_BLBST = ""
 Set ws = WB_BD.Sheets("DQN-COMO OK")
  With ws
   If .FilterMode = True Then .ShowAllData
 
    Set RN = .UsedRange
    .Rows("1:1").AutoFilter Field:=11, Criteria1:=Key1
    .Rows("1:1").AutoFilter Field:=10, Criteria1:=Key2
 
     Set Result = RN.SpecialCells(Excel.XlCellType.xlCellTypeVisible)
  
     For Each R_ow In Result.Rows
      If R_ow.Row > 1 Then
       If R_ow.Cells(1, 1).Value <> "" Then
        Get_BLBST = R_ow.Cells(1, 12).Value
       Else
       Exit For
       End If
      End If
     Next
  If .FilterMode = True Then .ShowAllData
 End With


End Function
