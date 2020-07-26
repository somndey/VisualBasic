Attribute VB_Name = "Module3"

Sub WorkSheetProtection(strSecureStatus As Boolean)
'Subroutine used to protect/unprotect the first 3 worksheets in the workbook
'if strSecureStatus is TRUE, worksheets are protected

Dim intCounter As Integer ' counter
Const strPassword = "FAST" 'Password used to protect worksheets
    
For intCounter = 1 To 1
    If strSecureStatus Then
        Worksheets(intCounter).Protect Password:=strPassword
        ActiveSheet.EnableSelection = xlUnlockedCells
        ActiveSheet.Protect AllowFiltering:=True
    Else
        Worksheets(intCounter).Unprotect Password:=strPassword
    End If
Next intCounter

End Sub
