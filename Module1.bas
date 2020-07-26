Attribute VB_Name = "Module1"
'Macro processes the spreadsheet and creates the output files for upload to Lawson
'Updated range from 1512 to 3000
Sub Create_Output()
    
    WorkSheetProtection (False)
    Application.ScreenUpdating = False
    Sheets("Input").Range("H5").Value = Format(Now, "mm/dd/yyyy HH:mm:ss")
    Sheets("Input").Range("B15:J32000").Interior.Color = xlNone
    
    'Sort
    Sort_Row
    
    'OverallCheckIsNOTGood will validate data entered and populate info to Invoice and Distrib worksheets
    If OverallCheckIsNOTGood(15, 32000, True) Then
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
   
    
    Call Create_Files
    
    WorkSheetProtection (True)
    
    Exit Sub
    
End Sub
Sub Create_Files()
'Subroutine creates the output files based on the file name passed to it
    Dim myFile As String, rng As Range, rng2 As Range, cellValue As Variant, i As Integer, j As Integer, k As Integer, l As Integer
    Dim fileExplorer As FileDialog
    Dim fileSaveName As Variant
    Dim Desc As String
    Dim sResult As Variant
    
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=InitialName, _
    fileFilter:="Text Files (*.txt), *.txt")
 
    If fileSaveName = False Then
        Call MsgBox("Process canceled! You have to click OK on the previous window!", vbOKOnly, "Warning")
        Exit Sub
    End If
 
    Sheets("Input").Range("G8").Value = fileSaveName
    Sheets("Input").Range("G8").Font.FontStyle = "Bold"
    Sheets("Input").Range("G8").Font.Color = vbRed
    
    Set rng = Sheets("Input").Range("B15:K32000")
    Set rng2 = Sheets("Trintech Template").Range("A1:L32000")
    
    Sheets("Trintech Template").Range("M2:L32000").ClearContents
        
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
    
    Open fileSaveName For Output As #1
    
    i = 1
    
    For i = 1 To rng2.Rows.Count
        If Sheets("Trintech Template").Range("M" & i).Value <> "A" Then
           Exit For
        End If
        For j = 1 To rng2.Columns.Count
            cellValue = rng2.Cells(i, j).Value
            If j = rng2.Columns.Count Then
                Print #1, cellValue; vbTab
            Else
                Print #1, (cellValue & vbTab);
            End If
        Next j
    Next i
    Close #1
    
    Sheets("Input").Range("I5").Value = Format(Now, "mm/dd/yyyy HH:mm:ss")
    Application.ScreenUpdating = True
    
    Call MsgBox("Journal Batch Upload created successfully!", vbOKOnly, "Warning")
    
End Sub
Sub Clear_Sheet()
    
    WorkSheetProtection (False)
    Sheets("Trintech Template").Range("A2:L32000").ClearContents
    Sheets("Input").Range("B15:J32000").ClearContents
    Sheets("Input").Range("B15:J32000").Interior.Color = xlNone
    WorkSheetProtection (True)
    
End Sub


'In the case that you not only want to exclude a list of special characters, but to exclude all characters that are not letters or numbers, I would suggest that you use a char type comparison approach.

'For each character in the String, I would check if the unicode character is between "A" and "Z", between "a" and "z" or between "0" and "9". This is the vba code:

Function cleanString(text As String) As String
    Dim output As String
    Dim c 'since char type does not exist in vba, we have to use variant type.
    For i = 1 To Len(text)
        c = Mid(text, i, 1) 'Select the character at the i position
        If (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Or (c >= "A" And c <= "Z") Then
            output = output & c 'add the character to your output.
        Else
            output = output & " " 'add the replacement character (space) to your output
        End If
    Next
    cleanString = output
End Function
