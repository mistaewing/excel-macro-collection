'Resizes all charts in the active worksheet to match the size of the active chart
Sub ResizeCharts()
    Dim i As Integer
    On Error GoTo Last
    iWidth = ActiveChart.Parent.Width
    iHeight = ActiveChart.Parent.Height
    For i = 1 To ActiveSheet.ChartObjects.Count
        With ActiveSheet.ChartObjects(i)
            .Width = iWidth
            .Height = iHeight
        End With
    Next i
    Last: Exit Sub
End Sub

'Exports selected chart as JPG
Sub SaveChartAsJpg()
    Dim mChart as Chart
    Set mChart = ActiveChart

    On Error GoTo Last

    If mChart.HasTitle Then
        mFileName = mChart.ChartTitle.Text & ".jpg"
    Else
        mFileName = "Chart_" & mChart.Parent.Index & ".jpg"
    End If
  
    mChart.Export FileName:= ActiveWorkbook.Path & "\" & mFileName, FilterName:= "JPG"

    Last: Exit Sub
End Sub

'Exports selected chart as PNG
Sub SaveChartAsPng()
    Dim mChart as Chart
    Set mChart = ActiveChart

    On Error GoTo Last

    If mChart.HasTitle Then
        mFileName = mChart.ChartTitle.Text & ".png"
    Else
        mFileName = "Chart_" & mChart.Parent.Index & ".png"
    End If
  
    mChart.Export FileName:= ActiveWorkbook.Path & "\" & mFileName, FilterName:= "PNG"

    Last: Exit Sub
End Sub

'Exports selected chart as BMP
Sub SaveChartAsBmp()
    Dim mChart as Chart
    Set mChart = ActiveChart

    On Error GoTo Last

    If mChart.HasTitle Then
        mFileName = mChart.ChartTitle.Text & ".bmp"
    Else
        mFileName = "Chart_" & mChart.Parent.Index & ".bmp"
    End If
  
    mChart.Export FileName:= ActiveWorkbook.Path & "\" & mFileName, FilterName:= "BMP"

    Last: Exit Sub
End Sub