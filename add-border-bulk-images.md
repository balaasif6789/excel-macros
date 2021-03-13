<pre>
Sub AddPictureBorders()
  Dim objShape As Shape
  Dim objInLineShape As InlineShape
  Dim objDoc As Document
 
  Set objDoc = ActiveDocument
 
  With objDoc
    For Each objInLineShape In .InlineShapes
      With objInLineShape.Line
        .Style = msoLineSingle
        .ForeColor.RGB = RGB(0, 0, 0)
      End With
    Next
    For Each objShape In .Shapes
      objShape.Fill.Solid
      With objShape.Line
        .Style = msoLineSingle
        .ForeColor.RGB = RGB(0, 0, 0)
      End With
    Next
  End With
End Sub


Ref link : https://www.datanumen.com/blogs/3-quick-methods-add-borders-pictures-word-document/


<pre>
