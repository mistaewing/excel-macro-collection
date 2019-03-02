Partial Class Macros
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Macros))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.AddSerialNumbers = Me.Factory.CreateRibbonButton
        Me.ConvertToUpperCase = Me.Factory.CreateRibbonButton
        Me.ToggleSentenceCase = Me.Factory.CreateRibbonButton
        Me.ToggleTitleCase = Me.Factory.CreateRibbonButton
        Me.InsertMultipleColumns = Me.Factory.CreateRibbonButton
        Me.AddMultipleRows = Me.Factory.CreateRibbonButton
        Me.AutofitColumns = Me.Factory.CreateRibbonButton
        Me.AutofitRows = Me.Factory.CreateRibbonButton
        Me.EqualizeColumns = Me.Factory.CreateRibbonButton
        Me.ResizeCharts = Me.Factory.CreateRibbonButton
        Me.Menu1 = Me.Factory.CreateRibbonMenu
        Me.ExportChartAsPng = Me.Factory.CreateRibbonButton
        Me.ExportChartAsJpg = Me.Factory.CreateRibbonButton
        Me.ExportChartAsBmp = Me.Factory.CreateRibbonButton
        Me.Backup = Me.Factory.CreateRibbonButton
        Me.CopyToNew = Me.Factory.CreateRibbonButton
        Me.Menu2 = Me.Factory.CreateRibbonMenu
        Me.ExportRange = Me.Factory.CreateRibbonButton
        Me.ExportSheet = Me.Factory.CreateRibbonButton
        Me.ExportAllSheets = Me.Factory.CreateRibbonButton
        Me.SaveRangeAsPng = Me.Factory.CreateRibbonButton
        Me.SaveRangeAsJpg = Me.Factory.CreateRibbonButton
        Me.SaveRangeAsBmp = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Label = "MACROS"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.AddSerialNumbers)
        Me.Group1.Items.Add(Me.ConvertToUpperCase)
        Me.Group1.Items.Add(Me.ToggleSentenceCase)
        Me.Group1.Items.Add(Me.ToggleTitleCase)
        Me.Group1.Label = "Text"
        Me.Group1.Name = "Group1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.InsertMultipleColumns)
        Me.Group2.Items.Add(Me.AddMultipleRows)
        Me.Group2.Items.Add(Me.AutofitColumns)
        Me.Group2.Items.Add(Me.AutofitRows)
        Me.Group2.Items.Add(Me.EqualizeColumns)
        Me.Group2.Label = "Layout"
        Me.Group2.Name = "Group2"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.ResizeCharts)
        Me.Group3.Items.Add(Me.Menu1)
        Me.Group3.Label = "Chart"
        Me.Group3.Name = "Group3"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Backup)
        Me.Group4.Items.Add(Me.CopyToNew)
        Me.Group4.Items.Add(Me.Menu2)
        Me.Group4.Items.Add(Me.ExportRange)
        Me.Group4.Items.Add(Me.ExportSheet)
        Me.Group4.Items.Add(Me.ExportAllSheets)
        Me.Group4.Label = "Save and Export"
        Me.Group4.Name = "Group4"
        '
        'AddSerialNumbers
        '
        Me.AddSerialNumbers.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.AddSerialNumbers.Image = Global.MacroCollection.My.Resources.Resources.counter
        Me.AddSerialNumbers.Label = "Add Series of Numbers"
        Me.AddSerialNumbers.Name = "AddSerialNumbers"
        Me.AddSerialNumbers.ScreenTip = "Add Series of Numbers"
        Me.AddSerialNumbers.ShowImage = True
        Me.AddSerialNumbers.SuperTip = "Adds a series of numbers from 1 to a number specified by the user."
        '
        'ConvertToUpperCase
        '
        Me.ConvertToUpperCase.Image = Global.MacroCollection.My.Resources.Resources.edit_uppercase
        Me.ConvertToUpperCase.Label = "Toggle Upper Case"
        Me.ConvertToUpperCase.Name = "ConvertToUpperCase"
        Me.ConvertToUpperCase.ScreenTip = "Toggle Upper Case"
        Me.ConvertToUpperCase.ShowImage = True
        Me.ConvertToUpperCase.SuperTip = "Toggles the text in the selected cells to upper case or lower case."
        '
        'ToggleSentenceCase
        '
        Me.ToggleSentenceCase.Image = Global.MacroCollection.My.Resources.Resources.edit_small_caps
        Me.ToggleSentenceCase.Label = "Toggle Sentence Case"
        Me.ToggleSentenceCase.Name = "ToggleSentenceCase"
        Me.ToggleSentenceCase.ScreenTip = "Toggle Sentence Case"
        Me.ToggleSentenceCase.ShowImage = True
        Me.ToggleSentenceCase.SuperTip = "Changes the text in the selected cells so that the first characters will be upper" &
    " case."
        '
        'ToggleTitleCase
        '
        Me.ToggleTitleCase.Image = Global.MacroCollection.My.Resources.Resources.edit_all_caps
        Me.ToggleTitleCase.Label = "Toggle Title Case"
        Me.ToggleTitleCase.Name = "ToggleTitleCase"
        Me.ToggleTitleCase.ScreenTip = "Toggle Title Case"
        Me.ToggleTitleCase.ShowImage = True
        Me.ToggleTitleCase.SuperTip = "Changes the text in the selected range so that the first letter of every word is " &
    "uppercase."
        '
        'InsertMultipleColumns
        '
        Me.InsertMultipleColumns.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.InsertMultipleColumns.Image = Global.MacroCollection.My.Resources.Resources.AddColumnToLeft_16x
        Me.InsertMultipleColumns.Label = "Insert Columns"
        Me.InsertMultipleColumns.Name = "InsertMultipleColumns"
        Me.InsertMultipleColumns.ScreenTip = "Inserts Columns"
        Me.InsertMultipleColumns.ShowImage = True
        Me.InsertMultipleColumns.SuperTip = "Inserts given number of columns before the selected column."
        '
        'AddMultipleRows
        '
        Me.AddMultipleRows.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.AddMultipleRows.Image = Global.MacroCollection.My.Resources.Resources.AddRowToAbove_16x
        Me.AddMultipleRows.Label = "Insert Rows"
        Me.AddMultipleRows.Name = "AddMultipleRows"
        Me.AddMultipleRows.ScreenTip = "Insert Rows"
        Me.AddMultipleRows.ShowImage = True
        Me.AddMultipleRows.SuperTip = "Inserts given number of rows before the selected row."
        '
        'AutofitColumns
        '
        Me.AutofitColumns.Image = Global.MacroCollection.My.Resources.Resources.AutosizeStretch_16x
        Me.AutofitColumns.Label = "Fit Columns"
        Me.AutofitColumns.Name = "AutofitColumns"
        Me.AutofitColumns.ScreenTip = "Fit Columns"
        Me.AutofitColumns.ShowImage = True
        Me.AutofitColumns.SuperTip = "Fits all columns on the worksheet."
        '
        'AutofitRows
        '
        Me.AutofitRows.Image = Global.MacroCollection.My.Resources.Resources.MakeSameHeight_16x
        Me.AutofitRows.Label = "Fit Rows"
        Me.AutofitRows.Name = "AutofitRows"
        Me.AutofitRows.ScreenTip = "Fit Rows"
        Me.AutofitRows.ShowImage = True
        Me.AutofitRows.SuperTip = "Fits all rows on the worksheet."
        '
        'EqualizeColumns
        '
        Me.EqualizeColumns.Image = Global.MacroCollection.My.Resources.Resources.AutosizeColumn_16x
        Me.EqualizeColumns.Label = "Equalize Columns"
        Me.EqualizeColumns.Name = "EqualizeColumns"
        Me.EqualizeColumns.ScreenTip = "Equalize Columns"
        Me.EqualizeColumns.ShowImage = True
        Me.EqualizeColumns.SuperTip = "Equalizes the column width in the selected range while the overall width of the r" &
    "ange does not change."
        '
        'ResizeCharts
        '
        Me.ResizeCharts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ResizeCharts.Image = CType(resources.GetObject("ResizeCharts.Image"), System.Drawing.Image)
        Me.ResizeCharts.Label = "Resize Charts"
        Me.ResizeCharts.Name = "ResizeCharts"
        Me.ResizeCharts.ScreenTip = "Resize Charts"
        Me.ResizeCharts.ShowImage = True
        Me.ResizeCharts.SuperTip = "Resizes all charts on the worksheet to the same size as the selected chart."
        '
        'Menu1
        '
        Me.Menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu1.Image = Global.MacroCollection.My.Resources.Resources.image
        Me.Menu1.Items.Add(Me.ExportChartAsPng)
        Me.Menu1.Items.Add(Me.ExportChartAsJpg)
        Me.Menu1.Items.Add(Me.ExportChartAsBmp)
        Me.Menu1.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu1.Label = "Save Chart"
        Me.Menu1.Name = "Menu1"
        Me.Menu1.ScreenTip = "Save Chart As Image"
        Me.Menu1.ShowImage = True
        '
        'ExportChartAsPng
        '
        Me.ExportChartAsPng.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ExportChartAsPng.Image = Global.MacroCollection.My.Resources.Resources.icon_png
        Me.ExportChartAsPng.Label = "Save as PNG"
        Me.ExportChartAsPng.Name = "ExportChartAsPng"
        Me.ExportChartAsPng.ScreenTip = "Save as PNG"
        Me.ExportChartAsPng.ShowImage = True
        Me.ExportChartAsPng.SuperTip = "Saves selected chart as a PNG file."
        '
        'ExportChartAsJpg
        '
        Me.ExportChartAsJpg.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ExportChartAsJpg.Image = Global.MacroCollection.My.Resources.Resources.icon_jpg
        Me.ExportChartAsJpg.Label = "Save as JPG"
        Me.ExportChartAsJpg.Name = "ExportChartAsJpg"
        Me.ExportChartAsJpg.ScreenTip = "Save as JPG"
        Me.ExportChartAsJpg.ShowImage = True
        Me.ExportChartAsJpg.SuperTip = "Saves selected chart as a JPG file."
        '
        'ExportChartAsBmp
        '
        Me.ExportChartAsBmp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ExportChartAsBmp.Image = Global.MacroCollection.My.Resources.Resources.icon_bmp
        Me.ExportChartAsBmp.Label = "Save as BMP"
        Me.ExportChartAsBmp.Name = "ExportChartAsBmp"
        Me.ExportChartAsBmp.ScreenTip = "Save as BMP"
        Me.ExportChartAsBmp.ShowImage = True
        Me.ExportChartAsBmp.SuperTip = "Saves selected chart as a BMP file."
        '
        'Backup
        '
        Me.Backup.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Backup.Image = Global.MacroCollection.My.Resources.Resources.if_backup_383184
        Me.Backup.Label = "Backup"
        Me.Backup.Name = "Backup"
        Me.Backup.ScreenTip = "Create Backup"
        Me.Backup.ShowImage = True
        Me.Backup.SuperTip = "Creates a backup of the whole workbook in the current directory."
        '
        'CopyToNew
        '
        Me.CopyToNew.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.CopyToNew.Image = Global.MacroCollection.My.Resources.Resources.if_BT_copy_905652
        Me.CopyToNew.Label = "Copy to New"
        Me.CopyToNew.Name = "CopyToNew"
        Me.CopyToNew.ScreenTip = "Copy Worksheet to New Workbook"
        Me.CopyToNew.ShowImage = True
        Me.CopyToNew.SuperTip = "Copies the current worksheet to a new workbook."
        '
        'Menu2
        '
        Me.Menu2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu2.Image = Global.MacroCollection.My.Resources.Resources.image
        Me.Menu2.Items.Add(Me.SaveRangeAsPng)
        Me.Menu2.Items.Add(Me.SaveRangeAsJpg)
        Me.Menu2.Items.Add(Me.SaveRangeAsBmp)
        Me.Menu2.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu2.Label = "Save Range"
        Me.Menu2.Name = "Menu2"
        Me.Menu2.ShowImage = True
        '
        'ExportRange
        '
        Me.ExportRange.Image = CType(resources.GetObject("ExportRange.Image"), System.Drawing.Image)
        Me.ExportRange.Label = "Export Range"
        Me.ExportRange.Name = "ExportRange"
        Me.ExportRange.ScreenTip = "Export Range as PDF"
        Me.ExportRange.ShowImage = True
        Me.ExportRange.SuperTip = "Exports the selected range as a pdf file."
        '
        'ExportSheet
        '
        Me.ExportSheet.Image = CType(resources.GetObject("ExportSheet.Image"), System.Drawing.Image)
        Me.ExportSheet.Label = "Export Worksheet"
        Me.ExportSheet.Name = "ExportSheet"
        Me.ExportSheet.ScreenTip = "Export Worksheet as PDF"
        Me.ExportSheet.ShowImage = True
        Me.ExportSheet.SuperTip = "Exports the current worksheet as a pdf file."
        '
        'ExportAllSheets
        '
        Me.ExportAllSheets.Image = CType(resources.GetObject("ExportAllSheets.Image"), System.Drawing.Image)
        Me.ExportAllSheets.Label = "Export All Worksheets"
        Me.ExportAllSheets.Name = "ExportAllSheets"
        Me.ExportAllSheets.ScreenTip = "Export all worksheets as pdf"
        Me.ExportAllSheets.ShowImage = True
        Me.ExportAllSheets.SuperTip = "Exports all worksheets as pdf files."
        '
        'SaveRangeAsPng
        '
        Me.SaveRangeAsPng.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SaveRangeAsPng.Image = Global.MacroCollection.My.Resources.Resources.icon_png
        Me.SaveRangeAsPng.Label = "Save as PNG"
        Me.SaveRangeAsPng.Name = "SaveRangeAsPng"
        Me.SaveRangeAsPng.ScreenTip = "Save range as PNG"
        Me.SaveRangeAsPng.ShowImage = True
        Me.SaveRangeAsPng.SuperTip = "Saves selected range as a png file."
        '
        'SaveRangeAsJpg
        '
        Me.SaveRangeAsJpg.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SaveRangeAsJpg.Image = Global.MacroCollection.My.Resources.Resources.icon_jpg
        Me.SaveRangeAsJpg.Label = "Save as JPG"
        Me.SaveRangeAsJpg.Name = "SaveRangeAsJpg"
        Me.SaveRangeAsJpg.ScreenTip = "Save range as JPG"
        Me.SaveRangeAsJpg.ShowImage = True
        Me.SaveRangeAsJpg.SuperTip = "Saves selected range as a JPG file."
        '
        'SaveRangeAsBmp
        '
        Me.SaveRangeAsBmp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SaveRangeAsBmp.Image = Global.MacroCollection.My.Resources.Resources.icon_bmp
        Me.SaveRangeAsBmp.Label = "Save as BMP"
        Me.SaveRangeAsBmp.Name = "SaveRangeAsBmp"
        Me.SaveRangeAsBmp.ScreenTip = "Save range as BMP"
        Me.SaveRangeAsBmp.ShowImage = True
        Me.SaveRangeAsBmp.SuperTip = "Saves selected range as a BMP file."
        '
        'Macros
        '
        Me.Name = "Macros"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AddSerialNumbers As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ConvertToUpperCase As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ToggleSentenceCase As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents InsertMultipleColumns As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AddMultipleRows As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AutofitColumns As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AutofitRows As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EqualizeColumns As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ToggleTitleCase As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ResizeCharts As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents ExportChartAsPng As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExportChartAsJpg As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExportChartAsBmp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Backup As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CopyToNew As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExportRange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExportSheet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExportAllSheets As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents SaveRangeAsPng As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SaveRangeAsJpg As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SaveRangeAsBmp As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Macros() As Macros
        Get
            Return Me.GetRibbon(Of Macros)()
        End Get
    End Property
End Class
