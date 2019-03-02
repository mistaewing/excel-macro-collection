Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel


Public Class Macros

    Private Sub Macros_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub AddSerialNumbers_Click(sender As Object, e As RibbonControlEventArgs) Handles AddSerialNumbers.Click
        Dim ActiveCell As Range
        ActiveCell = Globals.ThisAddIn.Application.ActiveCell
        Dim iVal As Integer
        iVal = InputBox("Enter Value", "Enter Serial Numbers")
        For i As Integer = 1 To iVal
            ActiveCell.Value = i
            ActiveCell = ActiveCell.Offset(1, 0)
        Next
    End Sub

    Private Sub ConvertToUpperCase_Click(sender As Object, e As RibbonControlEventArgs) Handles ConvertToUpperCase.Click
        Dim Rng As Range
        Dim App As Application
        App = Globals.ThisAddIn.Application
        Dim mText As String
        For Each Rng In App.Selection
            If App.WorksheetFunction.IsText(Rng) Then
                mText = UCase(Rng.Value)
                If mText = Rng.Value Then
                    Rng.Value = LCase(Rng.Value)
                Else
                    Rng.Value = UCase(Rng.Value)
                End If
            End If
        Next
    End Sub

    Private Sub ToggleSentenceCase_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleSentenceCase.Click
        Dim Rng As Range
        Dim App As Application
        App = Globals.ThisAddIn.Application
        For Each Rng In App.Selection
            If App.WorksheetFunction.IsText(Rng) Then
                If Rng.Value = LCase(Rng.Value) Then
                    Rng.Value = UCase(Left(Rng.Value, 1)) & LCase(Right(Rng.Value, Len(Rng.Value) - 1))
                Else
                    Rng.Value = LCase(Rng.Value)
                End If
            End If
        Next Rng
    End Sub

    Private Sub InsertMultipleColumns_Click(sender As Object, e As RibbonControlEventArgs) Handles InsertMultipleColumns.Click
        Dim i As Integer
        Dim j As Integer
        Dim App As Application
        App = Globals.ThisAddIn.Application
        On Error Resume Next
        i = InputBox("Enter number of columns to insert", "Insert Columns")
        For j = 1 To i
            App.Selection.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
        Next j
    End Sub

    Private Sub AddMultipleRows_Click(sender As Object, e As RibbonControlEventArgs) Handles AddMultipleRows.Click
        Dim i As Integer
        Dim j As Integer
        Dim App As Application
        App = Globals.ThisAddIn.Application
        On Error Resume Next
        i = InputBox("Enter number of rows to insert", "Insert Rows")
        For j = 1 To i
            App.Selection.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
        Next j
    End Sub

    Private Sub AutofitColumns_Click(sender As Object, e As RibbonControlEventArgs) Handles AutofitColumns.Click
        Dim cell As Range
        Dim App As Application
        App = Globals.ThisAddIn.Application
        cell = App.ActiveCell
        App.ScreenUpdating = False
        App.Cells.Select()
        App.Cells.EntireColumn.AutoFit()
        cell.Select()
        App.ScreenUpdating = True
    End Sub

    Private Sub AutofitRows_Click(sender As Object, e As RibbonControlEventArgs) Handles AutofitRows.Click
        Dim cell As Range
        Dim App As Application
        App = Globals.ThisAddIn.Application
        cell = App.ActiveCell
        App.ScreenUpdating = False
        App.Cells.Select()
        App.Cells.EntireRow.AutoFit()
        cell.Select()
        App.ScreenUpdating = True
    End Sub

    Private Sub EqualizeColumns_Click(sender As Object, e As RibbonControlEventArgs) Handles EqualizeColumns.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application
        Dim iColumns As Range
        Dim sum_size As Double
        Dim size As Double
        iColumns = App.Selection.Columns
        On Error Resume Next
        sum_size = 0
        For Each col In iColumns
            sum_size = sum_size + col.ColumnWidth
        Next
        size = sum_size / iColumns.Count
        iColumns.ColumnWidth = size
    End Sub

    Private Sub ToggleTitleCase_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleTitleCase.Click
        Dim Rng As Range
        Dim App As Application
        App = Globals.ThisAddIn.Application
        For Each Rng In App.Selection
            If App.WorksheetFunction.IsText(Rng) Then
                If Rng.Value = LCase(Rng.Value) Then
                    Rng.Value = App.WorksheetFunction.Proper(Rng.Value)
                Else
                    Rng.Value = LCase(Rng.Value)
                End If
            End If
        Next
    End Sub

    Private Sub ResizeCharts_Click(sender As Object, e As RibbonControlEventArgs) Handles ResizeCharts.Click
        Dim i, iWidth, iHeight As Integer
        Dim App As Application
        App = Globals.ThisAddIn.Application
        iWidth = App.ActiveChart.Parent.Width
        iHeight = App.ActiveChart.Parent.Height
        If iWidth = 0 Or iHeight = 0 Then
            Exit Sub
        End If
        For i = 1 To App.ActiveSheet.ChartObjects.Count
            With App.ActiveSheet.ChartObjects(i)
                .Width = iWidth
                .Height = iHeight
            End With
        Next i
    End Sub

    Private Sub ExportChartAsPng_Click(sender As Object, e As RibbonControlEventArgs) Handles ExportChartAsPng.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application
        Dim mChart As Chart
        mChart = App.ActiveChart
        Dim mFileName As String

        If mChart.HasTitle Then
            mFileName = mChart.ChartTitle.Text & ".png"
        Else
            mFileName = "Chart_" & mChart.Parent.Index & ".png"
        End If
        On Error Resume Next
        mChart.Export(Filename:=App.ActiveWorkbook.Path & "\" & mFileName, FilterName:="PNG")
    End Sub

    Private Sub ExportChartAsJpg_Click(sender As Object, e As RibbonControlEventArgs) Handles ExportChartAsJpg.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application
        Dim mChart As Chart
        mChart = App.ActiveChart
        Dim mFileName As String

        If mChart.HasTitle Then
            mFileName = mChart.ChartTitle.Text & ".jpg"
        Else
            mFileName = "Chart_" & mChart.Parent.Index & ".jpg"
        End If
        On Error Resume Next
        mChart.Export(Filename:=App.ActiveWorkbook.Path & "\" & mFileName, FilterName:="JPG")
    End Sub

    Private Sub ExportChartAsBmp_Click(sender As Object, e As RibbonControlEventArgs) Handles ExportChartAsBmp.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application
        Dim mChart As Chart
        mChart = App.ActiveChart
        Dim mFileName As String

        If mChart.HasTitle Then
            mFileName = mChart.ChartTitle.Text & ".bmp"
        Else
            mFileName = "Chart_" & mChart.Parent.Index & ".bmp"
        End If
        On Error Resume Next
        mChart.Export(Filename:=App.ActiveWorkbook.Path & "\" & mFileName, FilterName:="BMP")
    End Sub

    Private Sub Backup_Click(sender As Object, e As RibbonControlEventArgs) Handles Backup.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application
        Dim splitStr As String()
        splitStr = Split(App.ActiveWorkbook.Name, ".")
        App.ActiveWorkbook.SaveCopyAs(Filename:=App.ActiveWorkbook.Path &
            "\" & splitStr(0) & "_backup_" & Format(Now, "yyyy-MM-dd_HH-mm") & "." & splitStr(1))
    End Sub

    Private Sub CopyToNew_Click(sender As Object, e As RibbonControlEventArgs) Handles CopyToNew.Click
        Globals.ThisAddIn.Application.ActiveSheet.Copy()
    End Sub

    Private Sub ExportRange_Click(sender As Object, e As RibbonControlEventArgs) Handles ExportRange.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application
        Dim splitStr As String()
        splitStr = Split(App.ActiveWorkbook.Name, ".")
        App.Selection.ExportAsFixedFormat(
        Type:=XlFixedFormatType.xlTypePDF,
            Filename:=App.ActiveWorkbook.Path & "\" & splitStr(0) & "_" & App.ActiveSheet.Name & "_Selection" & ".pdf",
            Quality:=XlFixedFormatQuality.xlQualityStandard,
            IgnorePrintAreas:=True)
    End Sub

    Private Sub ExportSheet_Click(sender As Object, e As RibbonControlEventArgs) Handles ExportSheet.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application
        App.ActiveSheet.ExportAsFixedFormat(
        Type:=XlFixedFormatType.xlTypePDF,
        FileName:=App.ActiveWorkbook.Path & "\" & App.ActiveSheet.Name & ".pdf",
        Quality:=XlFixedFormatQuality.xlQualityStandard,
        IgnorePrintAreas:=False)
    End Sub

    Private Sub ExportAllSheets_Click(sender As Object, e As RibbonControlEventArgs) Handles ExportAllSheets.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application
        Dim ws As Worksheet
        For Each ws In App.Worksheets
            ws.ExportAsFixedFormat(
            Type:=XlFixedFormatType.xlTypePDF,
                Filename:=App.ActiveWorkbook.Path & "\" & ws.Name & ".pdf",
                Quality:=XlFixedFormatQuality.xlQualityStandard,
                IgnorePrintAreas:=False)
        Next ws
    End Sub

    Private Sub SaveRangeAsPng_Click(sender As Object, e As RibbonControlEventArgs) Handles SaveRangeAsPng.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application

        Dim output As String
        Dim splitStr As String()
        Dim zoom_coef As Double
        Dim area As Range
        Dim chartobj As ChartObject

        splitStr = Split(App.ActiveWorkbook.Name, ".")
        output = App.ActiveWorkbook.Path & "\" & splitStr(0) & "_" & App.ActiveSheet.Name & "_Selection.png"
        zoom_coef = 100 / App.ActiveSheet.Parent.Windows(1).Zoom
        area = App.Selection
        App.ScreenUpdating = False
        area.CopyPicture(XlPictureAppearance.xlPrinter, XlCopyPictureFormat.xlPicture)
        chartobj = App.ActiveSheet.ChartObjects.Add(0, 0, area.Width * zoom_coef, area.Height * zoom_coef)
        chartobj.Chart.Paste()
        chartobj.Chart.Export(output, "png")
        chartobj.Delete()
        area.Select()
        App.ScreenUpdating = True
    End Sub

    Private Sub SaveRangeAsJpg_Click(sender As Object, e As RibbonControlEventArgs) Handles SaveRangeAsJpg.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application

        Dim output As String
        Dim splitStr As String()
        Dim zoom_coef As Double
        Dim area As Range
        Dim chartobj As ChartObject

        splitStr = Split(App.ActiveWorkbook.Name, ".")
        output = App.ActiveWorkbook.Path & "\" & splitStr(0) & "_" & App.ActiveSheet.Name & "_Selection.jpg"
        zoom_coef = 100 / App.ActiveSheet.Parent.Windows(1).Zoom
        area = App.Selection
        App.ScreenUpdating = False
        area.CopyPicture(XlPictureAppearance.xlPrinter, XlCopyPictureFormat.xlPicture)
        chartobj = App.ActiveSheet.ChartObjects.Add(0, 0, area.Width * zoom_coef, area.Height * zoom_coef)
        chartobj.Chart.Paste()
        chartobj.Chart.Export(output, "jpg")
        chartobj.Delete()
        area.Select()
        App.ScreenUpdating = True
    End Sub

    Private Sub SaveRangeAsBmp_Click(sender As Object, e As RibbonControlEventArgs) Handles SaveRangeAsBmp.Click
        Dim App As Application
        App = Globals.ThisAddIn.Application

        Dim output As String
        Dim splitStr As String()
        Dim zoom_coef As Double
        Dim area As Range
        Dim chartobj As ChartObject

        splitStr = Split(App.ActiveWorkbook.Name, ".")
        output = App.ActiveWorkbook.Path & "\" & splitStr(0) & "_" & App.ActiveSheet.Name & "_Selection.bmp"
        zoom_coef = 100 / App.ActiveSheet.Parent.Windows(1).Zoom
        area = App.Selection
        App.ScreenUpdating = False
        area.CopyPicture(XlPictureAppearance.xlPrinter, XlCopyPictureFormat.xlPicture)
        chartobj = App.ActiveSheet.ChartObjects.Add(0, 0, area.Width * zoom_coef, area.Height * zoom_coef)
        chartobj.Chart.Paste()
        chartobj.Chart.Export(output, "bmp")
        chartobj.Delete()
        area.Select()
        App.ScreenUpdating = True
    End Sub
End Class
