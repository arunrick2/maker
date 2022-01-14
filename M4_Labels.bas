Attribute VB_Name = "M4_Labels"
Sub PlaceLabels()

    Dim sline As Shape, fromsp As Shape, tosp As Shape

    For kx = 1 To 200
        For mx = 1 To 3
            'Is this mx blank? If yes, skip to next mx.
            If Len(Range("LabelsRange").Item(kx, mx)) > 0 And Range("LabelsRange").Item(kx, mx) <> "-" Then
                'Make definitions
                Set sline = Sheet3.Shapes("ArrowIndex" & kx & "-" & mx)
                Set fromsp = sline.ConnectorFormat.BeginConnectedShape
                Set tosp = sline.ConnectorFormat.EndConnectedShape
                fromnode = sline.ConnectorFormat.BeginConnectionSite
                tonode = sline.ConnectorFormat.EndConnectionSite
                
                'First let's find the specs of the arrow on which we will place our label:
                'Arrow specs: slinefromx / slinefromy /
                If fromnode = Range("ConvertedNodesRange").Item(kx, 1) Then
                    slinefromx = fromsp.Left + fromsp.Width / 2
                    slinefromy = fromsp.Top
                ElseIf fromnode = Range("ConvertedNodesRange").Item(kx, 2) Then
                    slinefromx = fromsp.Left
                    slinefromy = fromsp.Top + fromsp.Height / 2
                ElseIf fromnode = Range("ConvertedNodesRange").Item(kx, 3) Then
                    slinefromx = fromsp.Left + fromsp.Width / 2
                    slinefromy = fromsp.Top + fromsp.Height
                ElseIf fromnode = Range("ConvertedNodesRange").Item(kx, 4) Then
                    slinefromx = fromsp.Left + fromsp.Width
                    slinefromy = fromsp.Top + fromsp.Height / 2
                End If
'                'Arrow specs: slinetox / slinetoy
                If tonode = Range("ConvertedNodesRange").Item(Range("DependentsIndexRange").Item(kx, mx), 1) Then
                    slinetox = tosp.Left + tosp.Width / 2
                    slinetoy = tosp.Top
                ElseIf tonode = Range("ConvertedNodesRange").Item(Range("DependentsIndexRange").Item(kx, mx), 2) Then
                    slinetox = tosp.Left
                    slinetoy = tosp.Top + tosp.Height / 2
                ElseIf tonode = Range("ConvertedNodesRange").Item(Range("DependentsIndexRange").Item(kx, mx), 3) Then
                    slinetox = tosp.Left + tosp.Width / 2
                    slinetoy = tosp.Top + tosp.Height
                ElseIf tonode = Range("ConvertedNodesRange").Item(Range("DependentsIndexRange").Item(kx, mx), 4) Then
                    slinetox = tosp.Left + tosp.Width
                    slinetoy = tosp.Top + tosp.Height / 2
                End If
                
                'Arrow specs are done. Now, let's find label specs
                'But, we should first create our label to see its width according to its text.
                Sheet3.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 20, 30, 15).Name = "LabelIndex" & kx & "-" & mx
                'Format the textbox
                With Sheet3.Shapes("LabelIndex" & kx & "-" & mx)
                    .TextFrame2.TextRange.Characters.Text = Range("LabelsRange").Item(kx, mx)
                    .TextFrame2.MarginLeft = 2
                    .TextFrame2.MarginRight = 2
                    .TextFrame2.MarginTop = 0
                    .TextFrame2.MarginBottom = 0
                    .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                    .TextFrame2.WordWrap = msoFalse
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                    .TextFrame2.VerticalAnchor = msoAnchorMiddle
                    .TextFrame2.TextRange.Font.Size = Range("ShapeFontSize").Value
                    .Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .Line.Visible = msoFalse
                    .Placement = xlMove
                End With
                'Write the label specs. These are easy: labheight / labwidth
                labwidth = Sheet3.Shapes("LabelIndex" & kx & "-" & mx).Width
                labheight = Sheet3.Shapes("LabelIndex" & kx & "-" & mx).Height
                
                'Now, troublesome specs: lableft / labtop
                If sline.Adjustments.Count = 0 Then
                    'CASE 1 (Adjusment count is zero and straight connector. Simple calculation)
                    If sline.ConnectorFormat.Type = msoConnectorStraight Then
                        lableft = slinefromx + (slinetox - slinefromx) / 2
                        labtop = slinefromy + (slinetoy - slinefromy) / 2
                    'CASE 2 (Adjusment count is zero and elbow connector. A bit more complicated calculation)
                    Else
                        'Begin node is 1 or 3
                        If fromnode = Range("ConvertedNodesRange").Item(kx, 1) Or fromnode = Range("ConvertedNodesRange").Item(kx, 3) Then
                            If Abs(slinetoy - slinefromy) > Abs(slinetox - slinefromx) Then
                                lableft = slinefromx
                                labtop = slinefromy + (slinetoy - slinefromy) / 2
                            Else
                                labtop = slinetoy
                                lableft = slinefromx + (slinetox - slinefromx) / 2
                            End If
                        'Begin node is 2 or 4
                        Else
                            If Abs(slinetox - slinefromx) > Abs(slinetoy - slinefromy) Then
                                labtop = slinefromy
                                lableft = slinefromx + (slinetox - slinefromx) / 2
                            Else
                                lableft = slinetox
                                labtop = slinefromy + (slinetoy - slinefromy) / 2
                            End If
                        End If
                    End If
                'CASE 3 (Adjustment count is 1. Problematic calculation and has exceptions)
                ElseIf sline.Adjustments.Count = 1 Then
                    If fromnode = Range("ConvertedNodesRange").Item(kx, 2) And _
                    tonode = Range("ConvertedNodesRange").Item(Range("DependentsIndexRange").Item(kx, mx), 2) And _
                    slinefromx = slinetox Then
                        lableft = slinefromx - 18
                        labtop = slinefromy + (slinetoy - slinefromy) / 2
                    ElseIf fromnode = Range("ConvertedNodesRange").Item(kx, 4) And _
                    tonode = Range("ConvertedNodesRange").Item(Range("DependentsIndexRange").Item(kx, mx), 4) And _
                    slinefromx = slinetox Then
                        lableft = slinefromx + 18
                        labtop = slinefromy + (slinetoy - slinefromy) / 2
                    ElseIf fromnode = Range("ConvertedNodesRange").Item(kx, 2) Or fromnode = Range("ConvertedNodesRange").Item(kx, 4) Then
                        lableft = slinefromx - (slinefromx - slinetox) * sline.Adjustments.Item(1)
                        labtop = slinefromy + (slinetoy - slinefromy) / 2
                    ElseIf fromnode = Range("ConvertedNodesRange").Item(kx, 3) And _
                    tonode = Range("ConvertedNodesRange").Item(Range("DependentsIndexRange").Item(kx, mx), 3) And _
                    slinefromy = slinetoy Then
                        labtop = slinefromy + 18
                        lableft = slinefromx + (slinetox - slinefromx) / 2
                    ElseIf fromnode = Range("ConvertedNodesRange").Item(kx, 1) And _
                    tonode = Range("ConvertedNodesRange").Item(Range("DependentsIndexRange").Item(kx, mx), 1) And _
                    slinefromy = slinetoy Then
                        labtop = slinefromy - 18
                        lableft = slinefromx + (slinetox - slinefromx) / 2
                    ElseIf fromnode = Range("ConvertedNodesRange").Item(kx, 3) Or fromnode = Range("ConvertedNodesRange").Item(kx, 1) Then
                        labtop = slinefromy - (slinefromy - slinetoy) * sline.Adjustments.Item(1)
                        lableft = slinefromx + (slinetox - slinefromx) / 2
                    End If
                ElseIf sline.Adjustments.Count > 1 Then 'Bu olmamasý gereken bir senaryo zaten
                    lableft = slinefromx + (slinetox - slinefromx) / 2
                    labtop = slinefromy + (slinetoy - slinefromy) / 2
                End If
                'We should make a final small adjustment to center-align the label.
                lableft = lableft - labwidth / 2
                labtop = labtop - labheight / 2
                'Now we know where to move the label which we have created a few pharagraphs ago.
                Sheet3.Shapes("LabelIndex" & kx & "-" & mx).Left = lableft
                Sheet3.Shapes("LabelIndex" & kx & "-" & mx).Top = labtop
            End If
        Next mx
    Next kx
    
End Sub


