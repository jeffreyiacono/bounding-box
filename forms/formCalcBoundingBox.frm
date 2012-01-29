VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCalcBoundingBox 
   Caption         =   "Calculate Bounding Box"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   OleObjectBlob   =   "formCalcBoundingBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formCalcBoundingBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_close_form_Click()
  Call UnloadForm
End Sub
Private Sub btn_do_calc_Click()
On Error GoTo fail
  Dim rngX As Range
  Dim rngY As Range
  Dim rngIDs As Range
    
  ' check if a range was specified
  If Me.inputRangeX.Value = "" Or Me.inputRangeY.Value = "" Then
    MsgBox "Please specify the X values range and the Y values range!", vbOKOnly + vbExclamation, "ERROR"
  Else
    ' set the ranges of X and Y vals
    Set rngX = Range(Me.inputRangeX.Value)
    Set rngY = Range(Me.inputRangeY.Value)
    ' hack to add in IDs - assumed to be in A1
    Set rngIDs = rngX.Offset(0, Int(rngX.Column - 1) * -1)
        
    ' check if the same number of cells have been specified
    If rngX.Cells.Count <> rngY.Cells.Count Then
      MsgBox "The X values range and Y values range must have the same number of values!", vbOKOnly + vbExclamation, "ERROR"
    Else
      Call findBoundingBox(rngX, rngY, rngIDs)
      Call UnloadForm
    End If
  End If
Exit Sub
fail:
  MsgBox "Cannot calculate the bounding box at this time due to the following error: " & Err.Description, vbOKOnly + vbExclamation, "ERROR"
  Call UnloadForm
End Sub
Private Sub clearRngInputs()
  Me.inputRangeX.Value = Null
  Me.inputRangeY.Value = Null
End Sub
Private Sub UnloadForm()
  Unload Me
End Sub


