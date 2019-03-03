'Save selected range as png
Sub SaveRangeAsPng()
    Set sheet = ActiveSheet
    splitStr = Split(ActiveWorkbook.Name, ".")
    output = ActiveWorkbook.Path & "\" & splitStr(0) & "_" & ActiveSheet.Name & "_Selection.png"

    zoom_coef = 100 / sheet.Parent.Windows(1).Zoom
    Set area = Selection
    area.CopyPicture xlPrinter
    Set chartobj = sheet.ChartObjects.Add(0, 0, area.Width * zoom_coef, area.Height * zoom_coef)
    chartobj.Chart.Paste
    chartobj.Chart.Export output, "png"
    chartobj.Delete
End Sub

'Save selected range as jpg
Sub SaveRangeAsJpg()
    Set sheet = ActiveSheet
    splitStr = Split(ActiveWorkbook.Name, ".")
    output = ActiveWorkbook.Path & "\" & splitStr(0) & "_" & ActiveSheet.Name & "_Selection.jpg"

    zoom_coef = 100 / sheet.Parent.Windows(1).Zoom
    Set area = Selection
    area.CopyPicture xlPrinter
    Set chartobj = sheet.ChartObjects.Add(0, 0, area.Width * zoom_coef, area.Height * zoom_coef)
    chartobj.Chart.Paste
    chartobj.Chart.Export output, "jpg"
    chartobj.Delete
End Sub

'Save selected range as bmp
Sub SaveRangeAsBmp()
    Set sheet = ActiveSheet
    splitStr = Split(ActiveWorkbook.Name, ".")
    output = ActiveWorkbook.Path & "\" & splitStr(0) & "_" & ActiveSheet.Name & "_Selection.bmp"

    zoom_coef = 100 / sheet.Parent.Windows(1).Zoom
    Set area = Selection
    area.CopyPicture xlPrinter
    Set chartobj = sheet.ChartObjects.Add(0, 0, area.Width * zoom_coef, area.Height * zoom_coef)
    chartobj.Chart.Paste
    chartobj.Chart.Export output, "bmp"
    chartobj.Delete
End Sub

'Save selected range as pdf
Sub SaveRangeAsPdf()
    splitStr = Split(ActiveWorkbook.Name, ".")
    Selection.ExportAsFixedFormat _ 
        Type:=xlTypePDF, _
        Filename:= ActiveWorkbook.Path & "\" & splitStr(0) & "_" & ActiveSheet.Name & "_Selection.pdf", _
        Quality:= xlQualityStandard, _
        IgnorePrintAreas:= True
End Sub

'Saves the acive worksheet as pdf
Sub SaveWorksheetAsPdf()
    ActiveSheet.ExportAsFixedFormat _ 
        Type:= xlTypePDF, _
        FileName:= ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".pdf", _
        Quality:= xlQualityStandard, _
        IgnorePrintAreas:= False
End Sub

'Saves each worksheet in a separate pdf
Sub SaveWorksheetsAsPdf()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.ExportAsFixedFormat _
            Type:= xlTypePDF, _
            FileName:= ActiveWorkbook.Path & "\" & ws.Name & ".pdf", _
            Quality:= xlQualityStandard, _
            IgnorePrintAreas:= False     
    Next ws
End Sub

'Creates a backup of the active workbook
Sub CreateBackUp()
    splitStr = Split(ActiveWorkbook.Name, ".")
    ActiveWorkbook.SaveCopyAs Filename:=ActiveWorkbook.Path & _
        "\" & splitStr(0) & "_backup_" & Format(Now, "yy-mm-dd-hh-MM") & "." & splitStr(1)
End Sub

'Copies active sheet to new workbook
Sub CopyWorksheetToNewWorkbook()
    ActiveSheet.Copy
End Sub