Sub extractD()
    
    Dim Rng As Range
    Dim dtRng As Range
    Dim c As Integer, cc As Integer, r As Integer, rr As Integer, r1 As Integer, c1 As Integer
    Dim kell As Range
    Dim NextRow As Integer, i As Integer
    
'
'    Cells.Find(What:="???", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
'    xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
'    , SearchFormat:=False).Activate

    'Columns("C:D").UnMerge
    'Columns("C:D").Delete
    
    Range("B1").Select
    
    
    NextRow = Cells(Rows.count, "b").End(xlUp).Row + 1
    
    Cells(NextRow + 2, 2) = "Municipality"
    Cells(NextRow + 2, 3) = "Fishery Type"
    Cells(NextRow + 2, 4) = "Qtty"
    Cells(NextRow + 2, 5) = "Area"
    Cells(NextRow + 2, 6) = "Active"
    Cells(NextRow + 2, 7) = "Inactive"
    Cells(NextRow + 2, 8) = "Production Volume"
    
    Range("A1").Select
    Cells.Find(What:="??????????????", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

    
    ActiveCell.UnMerge
    
    
    Selection.End(xlDown).Select
    Set kell = ActiveCell
    
    ActiveCell.Offset(, 1).Select
    
    r1 = kell.Offset(, 1).Row
    c1 = kell.Offset(, 1).Column
    
    r = Range(Selection, Selection.End(xlDown)).count - 1
    kell.Select
    
    c = Range(Selection, Selection.End(xlToRight)).count
    
    Set dtRng = Range(Cells(r1, c1), Cells(kell.Row + r - 1, c + 1))
    dtRng.Select
    
'    For rr = 1 To r
'
'        For cc = 1 To c Step 2
'
'            dtRng(rr, cc).Select
'
'            If dtRng(rr, cc) <> "" Then
'
'                NextRow = Cells(Rows.count, "b").End(xlUp).Row + 1
'
''mun
'                Cells(NextRow, 2) = Cells(ActiveCell.Row, 2)
''prod
'                Cells(NextRow, 3) = Cells(kell.Row - 2, ActiveCell.Column)
''area
'                Cells(NextRow, 4) = dtRng(rr, cc)
''projected
'                Cells(NextRow, 5) = dtRng(rr, cc + 1)
'
'            End If
'
'        Next cc
'
'    Next rr
    
    
    For rr = 1 To r
        
        For cc = 1 To c Step 5
            
            dtRng(rr, cc).Select
            
            If dtRng(rr, cc) <> "" Then
                
                NextRow = Cells(Rows.count, "b").End(xlUp).Row + 1
                
'mun
                Cells(NextRow, 2) = Cells(ActiveCell.Row, 2)
'Qtty
                Cells(NextRow, 3) = Cells(kell.Row - 1, ActiveCell.Column)
'Fishery Type
                Cells(NextRow, 4) = dtRng(rr, cc)
'Area
                Cells(NextRow, 5) = dtRng(rr, cc + 1)
'Active
                Cells(NextRow, 6) = dtRng(rr, cc + 2)
'Inactive
                Cells(NextRow, 7) = dtRng(rr, cc + 3)
'Prod vol
                Cells(NextRow, 8) = dtRng(rr, cc + 4)
                
                
            End If
            
        Next cc
            
    Next rr
    
    
    Range("a1").Select
            
    End Sub
        

