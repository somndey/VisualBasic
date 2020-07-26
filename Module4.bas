Attribute VB_Name = "Module4"
Sub Validate_Output()
    
    WorkSheetProtection (False)
    Application.ScreenUpdating = False
    Sheets("Input").Range("H5").Value = Format(Now, "mm/dd/yyyy HH:mm:ss")
    Sheets("Input").Range("B15:J32000").Interior.Color = xlNone
    'Sort
    Sort_Row
    
    
    'OverallCheckIsNOTGood will validate data entered
    If OverallCheckIsNOTGood(15, 32000, True) Then
        Call WorkSheetProtection(False)
        Sheets("Input").Range("I5").Value = Format(Now, "mm/dd/yyyy HH:mm:ss")
        Call MsgBox("Highlighted Rows have Invalid GL Code Or Amount!", vbOKOnly, "Warning")
        WorkSheetProtection (True)
        Exit Sub
    End If

    With Application
        .Calculation = xlAutomatic
        .MaxChange = 0.001
    End With
    ActiveWorkbook.PrecisionAsDisplayed = False
    Application.Calculate  '  Added application object
    
    Sheets("Input").Range("I5").Value = Format(Now, "mm/dd/yyyy HH:mm:ss")
    Application.ScreenUpdating = True
    
    Create_Template
    Call MsgBox("Validation Successfull. Click on Validate and Create Output!", vbOKOnly, "Warning")
    WorkSheetProtection (True)
    
    Exit Sub
    
End Sub
Sub Create_Template()
'Subroutine creates the output files based on the file name passed to it
    Dim myFile As String, rng As Range, rng2 As Range, cellValue As Variant, i As Integer, j As Integer, k As Integer, l As Integer
    Dim fileExplorer As FileDialog
    Dim fileSaveName As Variant
    Dim Desc As String
    Dim sResult As Variant
    
    Set rng = Sheets("Input").Range("B15:K32000")
    Set rng2 = Sheets("Trintech Template").Range("A1:L32000")
    
    Sheets("Trintech Template").Range("A2:M32000").ClearContents
        
    'Open myFile For Output As #1
    k = 2
    l = 15
    
    For i = 1 To rng.Rows.Count
    'On Error GoTo err
    If Sheets("Input").Range("K" & l).Value <> "A" Then
    Exit For
    End If
    
    'Move Cells from Input Sheet to Trintech Template Sheet
    Sheets("Trintech Template").Range("H" & k).Value = Sheets("Input").Range("G" & l).Value
    Sheets("Trintech Template").Range("I" & k).Value = Sheets("Input").Range("H" & l).Value
    Sheets("Trintech Template").Range("J" & k).Value = Sheets("Input").Range("I" & l).Value
    Sheets("Trintech Template").Range("K" & k).Value = Sheets("Input").Range("J" & l).Value
    Sheets("Trintech Template").Range("M" & k).Value = Sheets("Input").Range("K" & l).Value
    Sheets("Trintech Template").Range("B" & k).Value = Sheets("Input").Range("B" & l).Value
    Sheets("Trintech Template").Range("C" & k).Value = Sheets("Input").Range("C" & l).Value
    Sheets("Trintech Template").Range("D" & k).Value = Sheets("Input").Range("D" & l).Value
    Sheets("Trintech Template").Range("F" & k).Value = Sheets("Input").Range("E" & l).Value
    
    'Remove Special Characters from Description Field
    Desc = Sheets("Input").Range("F" & l).Value
    Sheets("Trintech Template").Range("G" & k).Value = cleanString(Desc)
    
    'Sheets("Trintech Template").Range("G" & k).Value = Sheets("Input").Range("F" & l).Value
    
    On Error Resume Next
    Err.Clear
    
    sResult = Application.WorksheetFunction.VLookup _
                (Sheets("Input").Range("D" & l).Value, Sheets("Account_Names").Range("A:B"), 2, False)
    
    ' Check if value found
    If Err.Number = 0 Then
        Debug.Print "Found item. The value is " & sResult
        Sheets("Trintech Template").Range("E" & k).Value = sResult
    Else
        Debug.Print "Could not find value: " & glcode
        
    End If
    
    k = k + 1
    l = l + 1
    
    Next i
        
    Sheets("Input").Range("I5").Value = Format(Now, "mm/dd/yyyy HH:mm:ss")
    Application.ScreenUpdating = True

    
End Sub

Sub Sort_Row()

'Sort First Colum

    ActiveWorkbook.Worksheets("Input").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Input").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("K:K"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Input").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

