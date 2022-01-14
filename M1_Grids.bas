Attribute VB_Name = "M1_Grids"
'Make text comparisons Case Insensitive
Option Compare Text
Sub ApplyColor(rng As Range)
    Const Limit As Integer = 25
    For Each c In rng
        'If c.Value > Limit Then
            c.Interior.ColorIndex = 27
        'End If
    Next c
End Sub
Sub printAllRanges()

'ThisWorkbook.Names("myNamedRange").RefersToRange(1,1)
End Sub
Sub DefineGrids()

    'Clear GridX and GridY ranges
    Range("GridXRange").ClearContents
    Dim rng As Range
    Set rng = Range("GridXRange")
    'ApplyColor rng
    'rng.Rows.Count
'    Range("GridXRange").Select
    Range("GridYRange").ClearContents
    Range("ProcessOrderRange").ClearContents
    Range("DirectPrecedentRange").ClearContents
    'CommandButton1_Click
    'Find the start point. Define its grids. Find next ID.
    sx = Range("StartShapeIndex").Value
    Range("GridXRange").Item(sx) = Range("StartGridX")
    Range("GridYRange").Item(sx) = Range("StartGridY")
    Range("ProcessOrderRange").Item(1) = sx
    
    For kx = 1 To 200
        If Range("ProcessOrderRange").Item(kx) <> "" Then
            gthis = Range("ProcessOrderRange").Item(kx)
            For mx = 1 To 3
                'Is this mx blank? If yes, skip to next mx.
                If IsNumeric(Range("DependentsIndexRange").Item(gthis, mx)) Then
                    'Ok, this mx is not blank. Then, let's define gnext and check it.
                    gnext = Range("DependentsIndexRange").Item(gthis, mx)
                    'Now, is the grid of this gnext blank? If not blank, this means it is being used. Skip to next mx.
                    If Range("GridCombinedRange").Item(gnext) = "-" Then
                        'Good. Now we can add this to the Process Order. Define its direct precedent.
                        Range("ProcessOrderRange").Item(Range("ProcessIndex")) = gnext
                        Range("DirectPrecedentRange").Item(gnext) = gthis
                        'Now, define its grid
                        Select Case Range("DependentsDirectionRange").Item(gthis, mx)
                            Case "Right"
                                Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) + 1
                                Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis)
                            Case "Below"
                                Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis)
                                Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) + 1
                            Case "Left"
                                Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) - 1
                                Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis)
                            Case "Top"
                                Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis)
                                Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) - 1
                            Case "Below-Right"
                                Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) + 1
                                Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) + 1
                            Case "Top-Right"
                                Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) + 1
                                Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) - 1
                            Case "Below-Left"
                                Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) - 1
                                Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) + 1
                            Case "Top-Left"
                                Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) - 1
                                Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) - 1
                        End Select
                    End If
                 End If
            Next mx
            
            'If mx is finished > we still don't have the Next Process Order Item > but we have items waiting for it, we add it.
            If Range("ProcessOrderRange").Item(kx + 1) = "" And Range("ProcessOrderFallback").Value <> 99999 Then
                gnext = Range("ProcessOrderFallback").Value
                Range("NextIDSourceRange").Item(gthis) = Range("IDSourceRange").Item(gnext)
                'Good. Now we can add this to the Process Order. Define its direct precedent.
                Range("ProcessOrderRange").Item(Range("ProcessIndex")) = gnext
                Range("DirectPrecedentRange").Item(gnext) = gthis
                'Now, define its grid
                Select Case Range("DependentsDirectionRange").Item(gthis, 1)
                    Case "Right"
                        Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) + 1
                        Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis)
                    Case "Below"
                        Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis)
                        Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) + 1
                    Case "Left"
                        Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) - 1
                        Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis)
                    Case "Top"
                        Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis)
                        Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) - 1
                    Case "Below-Right"
                        Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) + 1
                        Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) + 1
                    Case "Top-Right"
                        Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) + 1
                        Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) - 1
                    Case "Below-Left"
                        Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) - 1
                        Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) + 1
                    Case "Top-Left"
                        Range("GridXRange").Item(gnext) = Range("GridXRange").Item(gthis) - 1
                        Range("GridYRange").Item(gnext) = Range("GridYRange").Item(gthis) - 1
                End Select
            End If
            
        End If
    Next kx

End Sub


