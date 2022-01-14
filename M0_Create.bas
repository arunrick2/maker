Attribute VB_Name = "M0_Create"
Sub CreateFlowchart()
Attribute CreateFlowchart.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False
    
    If Not InStr(Now, "1/12/2022") > 0 Then
        On Error GoTo Message
    End If
    
    Sheet3.Activate

    Dim xshp As Shape
    For Each xshp In Sheet3.Shapes
        If xshp.Name <> "mainicon" And xshp.Name <> "somekalogo" And xshp.Name <> "backtomenu" And xshp.Name <> "exportPDF" Then
            xshp.Delete
        End If
    Next xshp
    
    'Call necessary macros
    Call DefineGrids
    Call PlaceShapes
    Call PlaceArrows
    Call PlaceLabels
    
    'Prepare chart area
    Range("BaseChartArea").Interior.Color = xlNone
    Range("BaseChartArea").Borders.LineStyle = xlNone
    Range("BaseChartArea").Interior.Color = RGB(242, 242, 242)
    Range("FinalChartArea").Interior.Color = xlNone
    Range("FinalChartArea").BorderAround (xlContinuous)
    
    Sheet3.Shapes("somekalogo").Left = Range("FinalChartArea").Width + Range("A4").Width - Sheet3.Shapes("somekalogo").Width
    
    'Fix Next ID Formula Range and finalize stuff
    Range("NextIDFormula").Copy Range("NextIDSourceRange")
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Sheet3.Range("A4").Activate
    
    If Range("ErrorCount").Value > 0 Then
        MsgBox ("Flowchart generated but not perfectly! Check Dashboard to see the issues.")
    End If
    
    Exit Sub
    
Message:
    MsgBox ("Can not generate Flowchart! Check Dashboard to see the issues with your data")
    
    'Fix Next ID Formula Range and finalize stuff
    Range("NextIDFormula").Copy Range("NextIDSourceRange")
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Sheet4.Activate

End Sub


