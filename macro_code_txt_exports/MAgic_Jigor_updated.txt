Sub MAgic_Jigor()
    
    
    Speed
    
    Dim derliglancement As Long
    Dim fichjigor As Workbook
    Dim fichmagic As Workbook
    Dim varJig As Variant
    Dim varMagic As Variant
    Dim varDiff As Variant
    Dim jigOrd As String
    Dim Versmagic As String
    Dim vbmsn As String
    
    Dim splitJig() As String
    Dim splitMagic() As String
    
    
    Set fichjigor = Workbooks.Open(ThisWorkbook.Sheets("Link").Range("C3").Value, False, isReadonly)
    Set fichmagic = Workbooks.Open(ThisWorkbook.Sheets("Link").Range("C4").Value, False, isReadonly)
    
    derliglancement = 0
    varJig = ""
    varMagic = ""
    varDiff = ""
    jigOrd = ""
    Versmagic = ""
    vbmsn = ""

    
    Call UnprotectSheet(ThisWorkbook.Sheets("Main"))
    derliglancement = Last_Row(ThisWorkbook.Sheets("Main"), 1)
    
    For i = 2 To derliglancement
        vbmsn = "MSN" & ThisWorkbook.Sheets("Main").Range("C" & i).Value & "/" & ThisWorkbook.Sheets("Main").Range("B" & i).Value

        


        varJig = Application.Match(vbmsn, fichjigor.Sheets(1).Columns(1), 0)

        If Not IsError(varJig) Then

            jigOrd = fichjigor.Sheets(1).Cells(varJig, 7).Value

            splitJig = Split(jigOrd, "-")


            If jigOrd <> "NO DATA" And jigOrd <> "NO CONF" Then

                ThisWorkbook.Sheets("Main").Range("D" & i).Value = splitJig(1) & "-" & splitJig(2)
            Else
                ThisWorkbook.Sheets("Main").Range("D" & i).Value = "-"
            End If
            
            varMagic = Application.Match(jigOrd, fichmagic.Sheets(1).Columns(5), 0)

            If Not IsError(varMagic) Then

                Versmagic = fichmagic.Sheets(1).Cells(varMagic, 7).Value

                splitMagic = Split(Versmagic, ".")

                ThisWorkbook.Sheets("Main").Range("D" & i).Value = splitJig(1) & "-" & splitJig(2) & "-" & splitMagic(0)

            End If

        End If

    Next
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    derliglancement = 0
    varJig = ""
    varMagic = ""
    varDiff = ""
    jigOrd = ""
    Versmagic = ""
    vbmsn = ""
    

    Call UnprotectSheet(ThisWorkbook.Sheets("Como-Memo"))
    derliglancement = Last_Row(ThisWorkbook.Sheets("Como-Memo"), 1)
    
    For i = 2 To derliglancement
        vbmsn = "MSN" & ThisWorkbook.Sheets("Como-Memo").Range("C" & i).Value & "/" & ThisWorkbook.Sheets("Como-Memo").Range("B" & i).Value

        


        varJig = Application.Match(vbmsn, fichjigor.Sheets(1).Columns(1), 0)

        If Not IsError(varJig) Then

            jigOrd = fichjigor.Sheets(1).Cells(varJig, 7).Value

            splitJig = Split(jigOrd, "-")


            If jigOrd <> "NO DATA" And jigOrd <> "NO CONF" Then

                ThisWorkbook.Sheets("Como-Memo").Range("D" & i).Value = splitJig(1) & "-" & splitJig(2)
            Else
                ThisWorkbook.Sheets("Como-Memo").Range("D" & i).Value = "-"
            End If


            varMagic = Application.Match(jigOrd, fichmagic.Sheets(1).Columns(5), 0)

            If Not IsError(varMagic) Then

                Versmagic = fichmagic.Sheets(1).Cells(varMagic, 7).Value

                splitMagic = Split(Versmagic, ".")

                ThisWorkbook.Sheets("Como-Memo").Range("D" & i).Value = splitJig(1) & "-" & splitJig(2) & "-" & splitMagic(0)

            End If

        End If

    Next
    
    
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    derliglancement = 0
    varJig = ""
    varMagic = ""
    varDiff = ""
    jigOrd = ""
    Versmagic = ""
    vbmsn = ""
    

    Call UnprotectSheet(ThisWorkbook.Sheets("MCB"))
    derliglancement = Last_Row(ThisWorkbook.Sheets("MCB"), 1)
    
    For i = 2 To derliglancement
        vbmsn = "MSN" & ThisWorkbook.Sheets("MCB").Range("C" & i).Value & "/" & ThisWorkbook.Sheets("MCB").Range("B" & i).Value

        


        varJig = Application.Match(vbmsn, fichjigor.Sheets(1).Columns(1), 0)

        If Not IsError(varJig) Then

            jigOrd = fichjigor.Sheets(1).Cells(varJig, 7).Value

            splitJig = Split(jigOrd, "-")


            If jigOrd <> "NO DATA" And jigOrd <> "NO CONF" Then

                ThisWorkbook.Sheets("MCB").Range("D" & i).Value = splitJig(1) & "-" & splitJig(2)
            Else
                ThisWorkbook.Sheets("MCB").Range("D" & i).Value = "-"
            End If
            
            varMagic = Application.Match(jigOrd, fichmagic.Sheets(1).Columns(5), 0)

            If Not IsError(varMagic) Then

                Versmagic = fichmagic.Sheets(1).Cells(varMagic, 7).Value

                splitMagic = Split(Versmagic, ".")

                ThisWorkbook.Sheets("MCB").Range("D" & i).Value = splitJig(1) & "-" & splitJig(2) & "-" & splitMagic(0)

            End If

        End If

    Next
    
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    derliglancement = 0
    varJig = ""
    varMagic = ""
    varDiff = ""
    jigOrd = ""
    Versmagic = ""
    vbmsn = ""
    

    Call UnprotectSheet(ThisWorkbook.Sheets("Plaquette"))
    derliglancement = Last_Row(ThisWorkbook.Sheets("Plaquette"), 1)
    
    For i = 2 To derliglancement
        vbmsn = "MSN" & ThisWorkbook.Sheets("Plaquette").Range("C" & i).Value & "/" & ThisWorkbook.Sheets("Plaquette").Range("B" & i).Value

        


        varJig = Application.Match(vbmsn, fichjigor.Sheets(1).Columns(1), 0)

        If Not IsError(varJig) Then

            jigOrd = fichjigor.Sheets(1).Cells(varJig, 7).Value

            splitJig = Split(jigOrd, "-")


            
            If jigOrd <> "NO DATA" And jigOrd <> "NO CONF" Then

                ThisWorkbook.Sheets("Plaquette").Range("D" & i).Value = splitJig(1) & "-" & splitJig(2)
            Else
                ThisWorkbook.Sheets("Plaquette").Range("D" & i).Value = "-"
            End If
            
            
            varMagic = Application.Match(jigOrd, fichmagic.Sheets(1).Columns(5), 0)

            If Not IsError(varMagic) Then

                Versmagic = fichmagic.Sheets(1).Cells(varMagic, 7).Value

                splitMagic = Split(Versmagic, ".")

                ThisWorkbook.Sheets("Plaquette").Range("D" & i).Value = splitJig(1) & "-" & splitJig(2) & "-" & splitMagic(0)

            End If

        End If

    Next
    
    Call protectSheet(ThisWorkbook.Sheets("Main"))
    Call protectSheet(ThisWorkbook.Sheets("MCB"))
    Call protectSheet(ThisWorkbook.Sheets("Plaquette"))
    Call protectSheet(ThisWorkbook.Sheets("Como-Memo"))
    
    fichjigor.Close (False)
    fichmagic.Close (False)
    Unspeed
    
    
End Sub

Sub enleveractiverfilter()
    
    
    
    If ThisWorkbook.Sheets("Main").FilterMode = True Then
        
        ThisWorkbook.Sheets("Main").Rows(2).AutoFilter
        
    Else
        
        ThisWorkbook.Sheets("Main").Rows(2).AutoFilter
        
    End If
    
    
    
    If ThisWorkbook.Sheets("Como-Memo").FilterMode = True Then
        
        ThisWorkbook.Sheets("Como-Memo").Rows(1).AutoFilter
        
    Else
        
        ThisWorkbook.Sheets("Como-Memo").Rows(1).AutoFilter
        
    End If
    
    
    If ThisWorkbook.Sheets("MCB").FilterMode = True Then
        
        ThisWorkbook.Sheets("MCB").Rows(1).AutoFilter
        
    Else
        
        ThisWorkbook.Sheets("MCB").Rows(1).AutoFilter
        
    End If
    
      If ThisWorkbook.Sheets("Plaquette").FilterMode = True Then
        
        ThisWorkbook.Sheets("Plaquette").Rows(2).AutoFilter
        
    Else
        
        ThisWorkbook.Sheets("Plaquette").Rows(2).AutoFilter
        
    End If
    
    
    
End Sub




