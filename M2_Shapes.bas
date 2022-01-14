Attribute VB_Name = "M2_Shapes"
Sub PlaceShapes()

    zmaxcolumn = 0
    zmaxrow = 0

    For kx = 1 To 200
        If Range("ValidRange").Item(kx) <> "-" Then
            ShapeX = Range("ShapeXRange").Item(kx)
            ShapeY = Range("ShapeYRange").Item(kx)
            ShapeWidth = Range("ShapeWidthRange").Item(kx)
            ShapeHeight = Range("ShapeHeightRange").Item(kx)
            ShapeType = Range("ShapeTypeRange").Item(kx)
        
            Sheet3.Shapes.AddShape(ShapeType, ShapeX, ShapeY, ShapeWidth, ShapeHeight).Name = "ShapeIndex" & kx
            
            With Sheet3.Shapes("ShapeIndex" & kx)
                .TextFrame2.TextRange.Characters.Text = Range("ShapeTextRange").Item(kx)
                .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .TextFrame2.VerticalAnchor = msoAnchorMiddle
                .TextFrame2.MarginLeft = 2.8
                .TextFrame2.MarginRight = 2.8
                .TextFrame2.MarginTop = 0
                .TextFrame2.MarginBottom = 0
'                .TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
'                .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
                .TextFrame2.TextRange.Font.Size = Range("ShapeFontSize").Value
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Range("ShapeFontColor").Interior.Color
                .Fill.ForeColor.RGB = Range("ColorsRange").Item(Range("ShapeColorRange").Item(kx)).Interior.Color
                .Fill.Transparency = 0.1
                .Placement = xlMove
            End With
            
            'Find right border
            If Sheet3.Shapes("ShapeIndex" & kx).BottomRightCell.Column > zmaxcolumn Then
                zmaxcolumn = Sheet3.Shapes("ShapeIndex" & kx).BottomRightCell.Column
            End If
            
            'Find bottom border
            If Sheet3.Shapes("ShapeIndex" & kx).BottomRightCell.Row > zmaxrow Then
                zmaxrow = Sheet3.Shapes("ShapeIndex" & kx).BottomRightCell.Row
            End If
            
        End If
    Next kx
    
    Range("MaxColumn").Value = zmaxcolumn
    Range("MaxRow").Value = zmaxrow

End Sub


