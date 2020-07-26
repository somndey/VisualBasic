Attribute VB_Name = "Module5"
Sub DisableCut()
 With Application
  '~~> Disable RigthClick on SheetCells which also gives you the option to cut
  .CommandBars("Cell").Enabled = False
 
  '~~> Disable Cut button
  .CommandBars("Standard").Controls.Item("Cut").Enabled = False
 
  '~~> Disable Cut button
  .CommandBars("Edit").Controls.Item("Cut").Enabled = False
 
  '~~> Divert Ctrl + X = Cut
  .OnKey "^x", "CutDisabled"
 
  '~~> Divert "Delete" KEY
 ' .OnKey ("{Delete}"), "CutDisabled"
 
  '~~> Disable Cell drag & Drop
  .CellDragAndDrop = False
 End With
End Sub


Sub EnableCut()
 With Application
  .CommandBars("Edit").Controls.Item("Cut").Enabled = True
  .CommandBars("Standard").Controls.Item("Cut").Enabled = True
  .CommandBars("Cell").Enabled = True
  .OnKey "^x"
  .OnKey "{Delete}"
  .CellDragAndDrop = True
 End With
End Sub

Sub CutDisabled()
 Application.CutCopyMode = False
End Sub
Sub DeactivateIt()
     
    With Application.CommandBars("Worksheet Menu Bar")
        .Controls("&Edit").Enabled = False
        .Controls("&Window").Visible = False
        With .Controls("&File")
            .Controls("&Print...").Enabled = False
            .Controls("&Print Preview").Enabled = False
        End With
    End With
    Application.CommandBars("Drawing").Enabled = False
    Application.CommandBars("Standard").Controls("&Save").Enabled = False
     
End Sub
 
Sub ActivateIt()
     
    With Application.CommandBars("Worksheet Menu Bar")
        .Controls("&Edit").Enabled = True
        .Controls("&Window").Visible = True
        With .Controls("&File")
            .Controls("&Print...").Enabled = True
            .Controls("&Print Preview").Enabled = True
        End With
    End With
    Application.CommandBars("Drawing").Enabled = True
    Application.CommandBars("Standard").Controls("&Save").Enabled = True
     
End Sub
