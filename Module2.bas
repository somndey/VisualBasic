Attribute VB_Name = "Module2"
Function OverallCheckIsNOTGood(StartRow, EndRow As Long, CopyRows As Boolean) As Boolean
Dim i, j, k As Long
Dim glcode As String
Dim sRes As Variant


Sheets("Input").Calculate
OverallCheckIsNOTGood = False


For i = StartRow To EndRow  'Row count
   If Range("K" & i).Value <> "A" Then Exit For
   glcode = Trim(Range("B" & i).Value) & Trim(Range("C" & i).Value) & Trim(Range("D" & i).Value)
   
   'Amount and Unit fields auto corrections
   If (Not IsNumeric(Range("I" & i).Value)) Then
       Range("I" & i).ClearContents
   End If
   
   If (Not IsNumeric(Range("J" & i).Value)) Then
       Range("J" & i).ClearContents
   End If
   
   
   'Validate Amount and Units Field
   If (Range("I" & i).Value = "") Then
       If (Range("J" & i).Value = "") Then
           Range("I" & i & ":J" & i).Interior.Color = 65535
           OverallCheckIsNOTGood = True
       Else
           Range("I" & i & ":J" & i).Interior.Color = xlNone
       End If
   Else
       Range("I" & i & ":J" & i).Interior.Color = xlNone
   End If

   On Error Resume Next
   Err.Clear
   sRes = Application.WorksheetFunction.VLookup _
                (glcode, Sheets("Accounts").Range("A:A"), 1, False)
  
    ' Check if value foundsss
    If Err.Number = 0 Then
        Debug.Print "Found item. The value is " & sRes
        Range("B" & i & ":D" & i).Interior.Color = xlNone
    Else
        Debug.Print "Could not find value: " & glcode
        Range("B" & i & ":D" & i).Interior.Color = 65535
        OverallCheckIsNOTGood = True
    End If
    
Next i


End Function



