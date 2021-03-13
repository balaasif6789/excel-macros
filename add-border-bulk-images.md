Sub AddPictureBorders()<br>
  Dim objShape As Shape<br>
  Dim objInLineShape As InlineShape<br>
  Dim objDoc As Document<br>
 <br>
  Set objDoc = ActiveDocument<br>
 <br>
  With objDoc<br>
    For Each objInLineShape In .InlineShapes<br>
      With objInLineShape.Line<br>
        .Style = msoLineSingle<br>
        .ForeColor.RGB = RGB(0, 0, 0)<br>
      End With<br>
    Next<br>
    For Each objShape In .Shapes<br>
      objShape.Fill.Solid<br>
      With objShape.Line<br>
        .Style = msoLineSingle<br>
        .ForeColor.RGB = RGB(0, 0, 0)<br>
      End With<br>
    Next<br>
  End With<br>
End Sub<br>
