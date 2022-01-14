Attribute VB_Name = "M5_Others"
Sub ExportPDF()

    #If Mac Then
        MsgBox "Please go to: File -> Saveas -> PDF and Export to PDF"
    #Else
        Sheets("Flowchart").ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
    #End If

End Sub

Sub ClearData()

    If MsgBox("This will erase existing data, want to continue?", vbYesNo) = vbNo Then Exit Sub

    Range("DataTable").ClearContents

End Sub


Sub InsOff()

    Sheet4.Shapes("yellownotes1").Visible = False
    Sheet4.Shapes("yellownotes2").Visible = False
    Sheet4.Shapes("yellownotes3").Visible = False

End Sub

Sub InsOn()

    Sheet4.Shapes("yellownotes1").Visible = True
    Sheet4.Shapes("yellownotes2").Visible = True
    Sheet4.Shapes("yellownotes3").Visible = True

End Sub
