Attribute VB_Name = "boundingBox"
Option Explicit

Type Coordinate
  x As Double
  y As Double
  id As Long
End Type
Public Sub findBoundingBox(rangeOfCoordinatesX As Range, rangeOfCoordinatesY As Range, rangeOfCoordinatesIDs As Range)
On Error GoTo fail

  Dim sortedCoordinates() As Coordinate
  Dim boundingCoordinates() As Coordinate
  Dim anchor As Coordinate
  Dim flagged As Coordinate
    
  Dim wsTemp As Worksheet
    
  Dim calcIsFinished As Boolean
    
  Dim i As Long
  Dim startRow As Long
  Dim startCol As Long
  Dim startingIndex As Long
  Dim outputOffset As Long
    
  Dim numOfRows As Long
    
  Application.ScreenUpdating = False
    
  outputOffset = 4
    
  ' Add new worksheet to do the sorting
  Set wsTemp = ThisWorkbook.Worksheets.Add
    
  startRow = 1
  startCol = 1
    
  ' Add X, Y values of range to new worksheet, set the number of rows
  For i = 0 To rangeOfCoordinatesX.Cells.Count - 1 Step 1
    wsTemp.Cells(startRow + i, startCol).Value = rangeOfCoordinatesX.Cells(startRow + i, 1).Value
    wsTemp.Cells(startRow + i, startCol + 1).Value = rangeOfCoordinatesY.Cells(startRow + i, 1).Value
    ' Add the ID column
    wsTemp.Cells(startRow + i, startCol + 2).Value = rangeOfCoordinatesIDs.Cells(startRow + i, 1).Value
    numOfRows = i
  Next i

  ' Sort new X, Y pairs based on X column
  wsTemp.Range(wsTemp.Cells(startRow, startCol), wsTemp.Cells(numOfRows + 1, startCol + 2)).Sort _
    Key1:=wsTemp.Cells(startRow, startCol), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
        
  ' resize the sortCoordinates Array
  ReDim sortedCoordinates(0 To numOfRows)
        
  ' Load X, Y coordinates into array
  For i = 0 To numOfRows Step 1
    sortedCoordinates(i).x = wsTemp.Cells(startRow + i, startCol).Value
    sortedCoordinates(i).y = wsTemp.Cells(startRow + i, startCol + 1).Value
    ' load the IDs
    sortedCoordinates(i).id = wsTemp.Cells(startRow + i, startCol + 2).Value
  Next i
    
  ' Calculate the HighLine
  ' Set the first anchor to the current min (first point), set flagged point to the first point / anchor
    
  ' Set anchor to first point
  anchor.x = sortedCoordinates(0).x
  anchor.y = sortedCoordinates(0).y
  ' set first anchor id
  anchor.id = sortedCoordinates(0).id
    
  ' resize the bounding array and add this first anchor
  ReDim boundingCoordinates(0)
  ' Add first point to the bounding coorindates
  boundingCoordinates(0).x = anchor.x
  boundingCoordinates(0).y = anchor.y
  ' Add first ID
  boundingCoordinates(0).id = anchor.id
       
  ' Set end point as Flagged
  flagged.x = sortedCoordinates(UBound(sortedCoordinates)).x
  flagged.y = sortedCoordinates(UBound(sortedCoordinates)).y
  ' set first flagged ID
  flagged.id = sortedCoordinates(UBound(sortedCoordinates)).id
    
  startingIndex = 1
  calcIsFinished = False

  Do While Not calcIsFinished
    For i = startingIndex To UBound(sortedCoordinates) Step 1
      ' Calculate the slope between anchor point and current point and compare to anchor and flagged point
      ' If greater, flag current point AND if not the same x and y points (remove dups)
      If (anchor.x <> sortedCoordinates(i).x And anchor.y <> sortedCoordinates(i).y) And slope(anchor.x, anchor.y, sortedCoordinates(i).x, sortedCoordinates(i).y) > _
        slope(anchor.x, anchor.y, flagged.x, flagged.y) Then
        ' set flag to these new points
        flagged.x = sortedCoordinates(i).x
        flagged.y = sortedCoordinates(i).y
        ' flag ID
        flagged.id = sortedCoordinates(i).id
        startingIndex = i + 1
      End If
            
      ' Check if end has been reached ...
      If i >= UBound(sortedCoordinates) Then
        ' Check if flagged point is the End Point
        If flagged.x = sortedCoordinates(UBound(sortedCoordinates)).x And flagged.y = sortedCoordinates(UBound(sortedCoordinates)).y Then
          ' If so, we are done, just want to add this last End Point
          calcIsFinished = True
        End If
            
        ' Resize the Bounded Coordinates Array and add Flagged Point
        ReDim Preserve boundingCoordinates(0 To UBound(boundingCoordinates) + 1)
        boundingCoordinates(UBound(boundingCoordinates)).x = flagged.x
        boundingCoordinates(UBound(boundingCoordinates)).y = flagged.y
        ' add ID
        boundingCoordinates(UBound(boundingCoordinates)).id = flagged.id
        ' Set Anchor to previously Flagged point
        anchor.x = flagged.x
        anchor.y = flagged.y
        anchor.id = flagged.id
        ' Set Flagged to the End Point
        flagged.x = sortedCoordinates(UBound(sortedCoordinates)).x
        flagged.y = sortedCoordinates(UBound(sortedCoordinates)).y
        flagged.id = sortedCoordinates(UBound(sortedCoordinates)).id
      End If
    Next i
  Loop

  ' Calculate the Low Line bounding box
  ' Set the first anchor to the current max (first point), set flagged point to the first point / anchor
    
  ' Set Anchor to the End Point (most right point)
  anchor.x = sortedCoordinates(UBound(sortedCoordinates)).x
  anchor.y = sortedCoordinates(UBound(sortedCoordinates)).y
  anchor.id = sortedCoordinates(UBound(sortedCoordinates)).id
        
  ' resize the bounding array and add this first anchor
  ReDim Preserve boundingCoordinates(0 To UBound(boundingCoordinates) + 1)
  ' Add the Anchor Point to the Bounding Coordinates
  boundingCoordinates(UBound(boundingCoordinates)).x = anchor.x
  boundingCoordinates(UBound(boundingCoordinates)).y = anchor.y
  boundingCoordinates(UBound(boundingCoordinates)).id = anchor.id
    
  ' Set Flagged Point to the First Point (most left)
  flagged.x = sortedCoordinates(0).x
  flagged.y = sortedCoordinates(0).y
  ' set ID
  flagged.id = sortedCoordinates(0).id
    
  startingIndex = UBound(sortedCoordinates) - 1
  calcIsFinished = False

  Do While Not calcIsFinished
    For i = startingIndex To 0 Step -1
      ' Calculate the slope between anchor point and current point and compare to anchor and flagged point
      ' AND if not the same x and y points (remove dups)
        If (anchor.x <> sortedCoordinates(i).x And anchor.y <> sortedCoordinates(i).y) And slope(anchor.x, anchor.y, sortedCoordinates(i).x, sortedCoordinates(i).y) > slope(anchor.x, anchor.y, flagged.x, flagged.y) Then
          ' set flag to these new points
          flagged.x = sortedCoordinates(i).x
          flagged.y = sortedCoordinates(i).y
          ' flag ID
          flagged.id = sortedCoordinates(i).id
          'startingIndex = Application.WorksheetFunction.Max(i - 1, 0)
          startingIndex = i - 1
        End If
            
        ' Check if end has been reached ...
        If i <= 0 Then
          ' Check if flagged point is the Start (most left) Point
          If flagged.x = sortedCoordinates(0).x And flagged.y = sortedCoordinates(0).y Then
            ' If so, we are done, just want to add this last End Point
            calcIsFinished = True
          End If
            
          ' Resize the Bounded Coordinates Array and add the new coorindate
          ReDim Preserve boundingCoordinates(0 To UBound(boundingCoordinates) + 1)
          boundingCoordinates(UBound(boundingCoordinates)).x = flagged.x
          boundingCoordinates(UBound(boundingCoordinates)).y = flagged.y
          boundingCoordinates(UBound(boundingCoordinates)).id = flagged.id
          ' Set Anchor to previously Flagged point
          anchor.x = flagged.x
          anchor.y = flagged.y
          anchor.id = flagged.id
          ' Set Flagged to the End Point
          flagged.x = sortedCoordinates(0).x
          flagged.y = sortedCoordinates(0).y
          ' set ID
          flagged.id = sortedCoordinates(0).id
        End If
      Next i
  Loop
    
  ' Output the final bounding box coordinates
  For i = 0 To UBound(boundingCoordinates) Step 1
    wsTemp.Cells(startRow + i, startCol + outputOffset).Value = boundingCoordinates(i).x
    wsTemp.Cells(startRow + i, startCol + outputOffset + 1).Value = boundingCoordinates(i).y
    wsTemp.Cells(startRow + i, startCol + outputOffset + 2).Value = boundingCoordinates(i).id
    numOfRows = i
  Next i
    
cleanUp:
  Set wsTemp = Nothing
  Application.ScreenUpdating = True
Exit Sub

fail:
  Application.ScreenUpdating = True
  MsgBox "Error: " & Err.Description, vbOKOnly, "Error"
  GoTo cleanUp
End Sub
Private Function slope(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
  If x1 = x2 Then
    slope = 0
  Else
    slope = (y2 - y1) / (x2 - x1)
  End If
End Function
Public Sub showForm()
On Error GoTo fail
  formCalcBoundingBox.Show
Exit Sub
fail:
  MsgBox "Error: " & Err.Description, vbOKOnly, "Error"
End Sub
