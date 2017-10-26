Sub FundUpdate()

Dim iRow As Integer
Dim notFound As Integer
notFound = 0

Workbooks("SDB_FUNDLIST.xlsm").Sheets("FundUpdate").Activate
For iRow = 1 To WorksheetFunction.CountA(Sheets("FundUpdate").Columns(5))

    'Searches SCS, RCS or LL columns in FundList'
    Workbooks("SDB_FUNDLIST.xlsm").Sheets("FundUpdate").Activate
    SCS = Sheets("FundUpdate").Cells(iRow + 1, 2).Value
    RCS = Sheets("FundUpdate").Cells(iRow + 1, 3).Value
    LL = Sheets("FundUpdate").Cells(iRow + 1, 4).Value
    Eng = Cells(iRow + 1, 5).Value
    NEng = Cells(iRow + 1, 6).Value
    NFR = Cells(iRow + 1, 7).Value
    AddHis = Cells(iRow + 1, 8).Value
    Status = Cells(iRow + 1, 9).Value
    
    'SCS'
    If SCS <> "" Then
    Sheets("SDB_FUNDLIST").Activate
    Range("B1").Select
    
    'Loop to find both the correct fund code and the correct fund name'
    
        Do While Not SCS = ActiveCell.Value
           
            ActiveCell.Offset(1, 0).Activate
            j = ActiveCell.Value
            If SCS = j Then
            ActiveCell.Offset(0, 4).Select
            
                If (Eng <> ActiveCell.Value) Then
                
                ActiveCell.Offset(0, -4).Select
                ActiveCell.Offset(1, 0).Select
                
                Else
                If Not NEng = "" Then
                ActiveCell.Value = NEng
                End If
                If Not NFR = "" Then
                ActiveCell.Offset(0, 1).Value = NFR
                End If
                If Not Status = "" Then
                ActiveCell.Offset(0, 5).Value = Status
                End If
                ActiveCell.Offset(0, 6).Value = AddHis + ActiveCell.Offset(0, 6).Value
                With Sheets("FundUpdate").Cells(iRow + 1, 2).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 15773696
                .TintAndShade = 0
                .PatternTintAndShade = 0
                End With

                Exit Do
                End If
                
            'Stops program from running ad infinitum if no funds exist'
            ElseIf j = "12345" Then
            notFound = notFound + 1
            Sheets("FundUpdate").Activate
            ActiveSheet.Cells(iRow + 1, 2).Select
            With Sheets("FundUpdate").Cells(iRow + 1, 2).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Exit Do
            End If
            
        Loop
    
    'RCS'
    ElseIf RCS <> "" Then
    
    Sheets("SDB_FUNDLIST").Activate
    Range("C1").Select
    
    'Loop to find both the correct fund code and the correct fund name'
    
        Do While Not RCS = ActiveCell.Value
           
            ActiveCell.Offset(1, 0).Activate
            j = ActiveCell.Value
            If RCS = j Then
            ActiveCell.Offset(0, 3).Select
            
                If (Eng <> ActiveCell.Value) Then
                
                ActiveCell.Offset(0, -3).Select
                ActiveCell.Offset(1, 0).Select
                
                Else
                If Not NEng = "" Then
                ActiveCell.Value = NEng
                End If
                If Not NFR = "" Then
                ActiveCell.Offset(0, 1).Value = NFR
                End If
                If Not Status = "" Then
                ActiveCell.Offset(0, 5).Value = Status
                End If
                ActiveCell.Offset(0, 6).Value = AddHis + ActiveCell.Offset(0, 6).Value
                With Sheets("FundUpdate").Cells(iRow + 1, 3).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 15773696
                .TintAndShade = 0
                .PatternTintAndShade = 0
                End With

                Exit Do
                End If
                
            'Stops program from running ad infinitum if no funds exist'
            ElseIf j = "12345" Then
            notFound = notFound + 1
            Sheets("FundUpdate").Activate
            ActiveSheet.Cells(iRow + 1, 3).Select
            With Sheets("FundUpdate").Cells(iRow + 1, 3).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Exit Do
            End If
            
        Loop
    
    'LL'
    ElseIf LL <> "" Then
    Sheets("SDB_FUNDLIST").Activate
    Range("D1").Select
    
    'Loop to find both the correct fund code and the correct fund name'
    
        Do While Not LL = ActiveCell.Value
           
            ActiveCell.Offset(1, 0).Activate
            j = ActiveCell.Value
            If LL = j Then
            ActiveCell.Offset(0, 2).Select
            
                If (Eng <> ActiveCell.Value) Then
                
                ActiveCell.Offset(0, -2).Select
                ActiveCell.Offset(1, 0).Select
                
                Else
                If Not NEng = "" Then
                ActiveCell.Value = NEng
                End If
                If Not NFR = "" Then
                ActiveCell.Offset(0, 1).Value = NFR
                End If
                If Not Status = "" Then
                ActiveCell.Offset(0, 5).Value = Status
                End If
                ActiveCell.Offset(0, 6).Value = AddHis + ActiveCell.Offset(0, 6).Value
                With Sheets("FundUpdate").Cells(iRow + 1, 4).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 15773696
                .TintAndShade = 0
                .PatternTintAndShade = 0
                End With

                Exit Do
                End If
                
            'Stops program from running ad infinitum if no funds exist'
            ElseIf j = "12345" Then
            notFound = notFound + 1
            Sheets("FundUpdate").Activate
            ActiveSheet.Cells(iRow + 1, 4).Select
            With Sheets("FundUpdate").Cells(iRow + 1, 4).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Exit Do
            End If
            
        Loop
        
    End If
    
Next iRow
MsgBox "Search Complete. Found: " & iRow - 2 - notFound & " out of " & iRow - 2
End Sub
