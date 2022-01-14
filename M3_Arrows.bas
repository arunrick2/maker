Attribute VB_Name = "M3_Arrows"
Sub PlaceArrows()

    'Clear Nodes In Use ranges
    Range("Node1UseRange").ClearContents
    Range("Node2UseRange").ClearContents
    Range("Node3UseRange").ClearContents
    Range("Node4UseRange").ClearContents
    
    'First of all, we will place the arrows for direct dependents.
    For kx = 1 To 200
        For mx = 1 To 3
            'Is this mx blank? If yes, skip to next mx.
            If IsNumeric(Range("DependentsIndexRange").Item(kx, mx)) Then
                gthis = kx
                gnext = Range("DependentsIndexRange").Item(kx, mx)
                'Here, let's check if I am the direct precedent of this shape. If not, skip it because this means it has another direct precedent.
                If Range("DirectPrecedentRange").Item(gnext) = gthis Then
                    Set ThisShape = Sheet3.Shapes("ShapeIndex" & gthis)
                    Set NextShape = Sheet3.Shapes("ShapeIndex" & gnext)
                    
                    'Now, let's start placing arrows according to the NODE RANGE numbers. They start from zero.
                    ThisNode = 0
                    NextNode = 0
                    'First, check all connections table to see if there are two nodes both empty. If yes, we will use them.
                    For jx = 1 To Range("ConnectionsMax").Value
                        connode = Range("ConnectionsNodeSectionRange").Item(jx)
                        conthis = Range("ConnectionsThisNodeRange").Item(jx)
                        connext = Range("ConnectionsNextNodeRange").Item(jx)
                        
                        If ThisNode = 0 And NextNode = 0 Then
                            If connode = Range("DependentsNodeRange").Item(kx, mx) And _
                            (Range("Node" & conthis & "UseRange").Item(gthis) <> "Yes" And Range("Node" & connext & "UseRange").Item(gnext) <> "Yes") Then
                                ThisNode = conthis
                                NextNode = connext
                            End If
                        End If
                    Next jx
                    
                    'If we couldn't find two empty nodes, then let's check if there is at least one empty node.
                    For jx = 1 To Range("ConnectionsMax").Value
                        connode = Range("ConnectionsNodeSectionRange").Item(jx)
                        conthis = Range("ConnectionsThisNodeRange").Item(jx)
                        connext = Range("ConnectionsNextNodeRange").Item(jx)
                        
                        If ThisNode = 0 And NextNode = 0 Then
                            If connode = Range("DependentsNodeRange").Item(kx, mx) And _
                            (Range("Node" & conthis & "UseRange").Item(gthis) <> "Yes" Or Range("Node" & connext & "UseRange").Item(gnext) <> "Yes") Then
                                ThisNode = conthis
                                NextNode = connext
                            End If
                        End If
                    Next jx
                    
                    'Ok, if we still couldn't find, this means all of them are being used. Then we should use fallback connection.
                    For jx = 1 To Range("ConnectionsMax").Value
                        connode = Range("ConnectionsNodeSectionRange").Item(jx)
                        conthis = Range("ConnectionsThisNodeRange").Item(jx)
                        connext = Range("ConnectionsNextNodeRange").Item(jx)
                        
                        If ThisNode = 0 And NextNode = 0 Then
                            If connode = Range("DependentsNodeRange").Item(kx, mx) And Range("ConnectionsFallbackRange").Item(jx) = "Fallback" Then
                                ThisNode = conthis
                                NextNode = connext
                            End If
                        End If
                    Next jx
                    
                    'Define Connector Type
                    If Range("DependentsNodeRange").Item(kx, mx) = 20 Or Range("DependentsNodeRange").Item(kx, mx) = 40 Then
                        If (ThisNode = 2 And NextNode = 4) Or (ThisNode = 4 And NextNode = 2) Then
                            ConnectorType = msoConnectorStraight
                        End If
                    ElseIf Range("DependentsNodeRange").Item(kx, mx) = 10 Or Range("DependentsNodeRange").Item(kx, mx) = 30 Then
                        If (ThisNode = 1 And NextNode = 3) Or (ThisNode = 3 And NextNode = 1) Then
                            ConnectorType = msoConnectorStraight
                        End If
                    Else
                        ConnectorType = msoConnectorElbow
                    End If
                    
                    'Finally, let's convert all nodes in case for different shape types.
                    ThisNodeConverted = Range("ConvertedNodesRange").Item(gthis, ThisNode)
                    NextNodeConverted = Range("ConvertedNodesRange").Item(gnext, NextNode)
    
                    'Now, we can place the connector!
                    With Sheet3.Shapes.AddConnector(ConnectorType, 0, 0, 0, 0)
                        .ConnectorFormat.BeginConnect ThisShape, ThisNodeConverted
                        .ConnectorFormat.EndConnect NextShape, NextNodeConverted
                        .Line.EndArrowheadStyle = msoArrowheadTriangle
                        .Name = "ArrowIndex" & kx & "-" & mx
                    End With
                    'Format the connector
                    With Sheet3.Shapes("ArrowIndex" & kx & "-" & mx)
                        .Line.ForeColor.RGB = Range("ArrowColor").Interior.Color
                        .Placement = xlMove
                    End With
                    'Don't forget to fill in Nodes In Use table
                    Range("Node" & ThisNode & "UseRange").Item(gthis) = "Yes"
                    Range("Node" & NextNode & "UseRange").Item(gnext) = "Yes"
                End If
            End If
        Next mx
    Next kx
    
    'Now, we will cover almost the same procedure for Non-Direct dependents (side connectors)
    For kx = 1 To 200
        For mx = 1 To 3
            'Is this mx blank? If yes, skip to next mx.
            If IsNumeric(Range("DependentsIndexRange").Item(kx, mx)) Then
                gthis = kx
                gnext = Range("DependentsIndexRange").Item(kx, mx)
                'Here, let's check if I am the direct precedent of this shape. If YES, skip it because we've already placed them.
                If Range("DirectPrecedentRange").Item(gnext) <> gthis Then
                    Set ThisShape = Sheet3.Shapes("ShapeIndex" & gthis)
                    Set NextShape = Sheet3.Shapes("ShapeIndex" & gnext)
                    
                    'Now, let's start placing arrows according to the NODE RANGE numbers. They start from zero.
                    ThisNode = 0
                    NextNode = 0
                    'First, check all connections table to see if there are two nodes both empty. If yes, we will use them.
                    For jx = 1 To Range("ConnectionsMax").Value
                        connode = Range("ConnectionsNodeSectionRange").Item(jx)
                        conthis = Range("ConnectionsThisNodeRange").Item(jx)
                        connext = Range("ConnectionsNextNodeRange").Item(jx)
                        
                        If ThisNode = 0 And NextNode = 0 Then
                            If connode = Range("DependentsNodeRange").Item(kx, mx) And _
                            (Range("Node" & conthis & "UseRange").Item(gthis) <> "Yes" And Range("Node" & connext & "UseRange").Item(gnext) <> "Yes") Then
                                ThisNode = conthis
                                NextNode = connext
                            End If
                        End If
                    Next jx
                    
                    'If we couldn't find two empty nodes, then let's check if there is at least one empty node.
                    For jx = 1 To Range("ConnectionsMax").Value
                        connode = Range("ConnectionsNodeSectionRange").Item(jx)
                        conthis = Range("ConnectionsThisNodeRange").Item(jx)
                        connext = Range("ConnectionsNextNodeRange").Item(jx)
                        
                        If ThisNode = 0 And NextNode = 0 Then
                            If connode = Range("DependentsNodeRange").Item(kx, mx) And _
                            (Range("Node" & conthis & "UseRange").Item(gthis) <> "Yes" Or Range("Node" & connext & "UseRange").Item(gnext) <> "Yes") Then
                                ThisNode = conthis
                                NextNode = connext
                            End If
                        End If
                    Next jx
                    
                    'Ok, if we still couldn't find, this means all of them are being used. Then we should use fallback connection.
                    For jx = 1 To Range("ConnectionsMax").Value
                        connode = Range("ConnectionsNodeSectionRange").Item(jx)
                        conthis = Range("ConnectionsThisNodeRange").Item(jx)
                        connext = Range("ConnectionsNextNodeRange").Item(jx)
                        
                        If ThisNode = 0 And NextNode = 0 Then
                            If connode = Range("DependentsNodeRange").Item(kx, mx) And Range("ConnectionsFallbackRange").Item(jx) = "Fallback" Then
                                ThisNode = conthis
                                NextNode = connext
                            End If
                        End If
                    Next jx
                    
                    'Define Connector Type
                    If Range("DependentsNodeRange").Item(kx, mx) = 20 Or Range("DependentsNodeRange").Item(kx, mx) = 40 Then
                        If (ThisNode = 2 And NextNode = 4) Or (ThisNode = 4 And NextNode = 2) Then
                            ConnectorType = msoConnectorStraight
                        Else
                            ConnectorType = msoConnectorElbow
                        End If
                    ElseIf Range("DependentsNodeRange").Item(kx, mx) = 10 Or Range("DependentsNodeRange").Item(kx, mx) = 30 Then
                        If (ThisNode = 1 And NextNode = 3) Or (ThisNode = 3 And NextNode = 1) Then
                            ConnectorType = msoConnectorStraight
                        Else
                            ConnectorType = msoConnectorElbow
                        End If
                    Else
                        ConnectorType = msoConnectorElbow
                    End If
                    
                    'Finally, let's convert all nodes in case for different shape types.
                    ThisNodeConverted = Range("ConvertedNodesRange").Item(gthis, ThisNode)
                    NextNodeConverted = Range("ConvertedNodesRange").Item(gnext, NextNode)
    
                    'Now, we can place the connector!
                    With Sheet3.Shapes.AddConnector(ConnectorType, 0, 0, 0, 0)
                        .ConnectorFormat.BeginConnect ThisShape, ThisNodeConverted
                        .ConnectorFormat.EndConnect NextShape, NextNodeConverted
                        .Line.EndArrowheadStyle = msoArrowheadTriangle
                        .Name = "ArrowIndex" & kx & "-" & mx
                    End With
                    'Format the connector
                    With Sheet3.Shapes("ArrowIndex" & kx & "-" & mx)
                        .Line.ForeColor.RGB = Range("ArrowColor").Interior.Color
                        .Placement = xlMove
                    End With
                    'Don't forget to fill in Nodes In Use table
                    Range("Node" & ThisNode & "UseRange").Item(gthis) = "Yes"
                    Range("Node" & NextNode & "UseRange").Item(gnext) = "Yes"
                End If
            End If
        Next mx
    Next kx

End Sub
