Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Threading
Imports Newtonsoft.Json
Imports System.Xml

Public Class Form1
    Private excelApp As Application
    Private excelFilePath As String
    Private watcher As FileSystemWatcher
    Private debounceTimer As Timer
    Private lastEventTime As DateTime = DateTime.MinValue
    Private debounceInterval As Integer = 1000 ' milliseconds

    Public Class ExcelConfig
        Public Property ExcelFilePath As String
        Public Property SheetName As String
        Public Property CellReference As String
    End Class

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize the Excel application object
        excelApp = New Application()

        ' Load configuration if exists
        Dim config As ExcelConfig = LoadConfig()
        If config IsNot Nothing Then
            excelFilePath = config.ExcelFilePath
            PopulateSheetNames()
            ComboBoxSheets.SelectedItem = config.SheetName
            TextBoxCellReference.Text = config.CellReference
            InitializeFileWatcher()
        End If
    End Sub

    Private Sub ButtonSelectFile_Click(sender As Object, e As EventArgs) Handles ButtonSelectFile.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel Files|*.xls;*.xlsx"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            excelFilePath = openFileDialog.FileName
            PopulateSheetNames()

            ' Initialize and configure the FileSystemWatcher
            InitializeFileWatcher()

            ' Save configuration
            If ComboBoxSheets.SelectedItem IsNot Nothing AndAlso Not String.IsNullOrEmpty(TextBoxCellReference.Text) Then
                SaveConfig(excelFilePath, ComboBoxSheets.SelectedItem.ToString(), TextBoxCellReference.Text)
            End If
        End If
    End Sub

    Private Sub InitializeFileWatcher()
        If watcher IsNot Nothing Then
            RemoveHandler watcher.Changed, AddressOf OnExcelFileChanged
            RemoveHandler watcher.Renamed, AddressOf OnExcelFileChanged
            RemoveHandler watcher.Created, AddressOf OnExcelFileChanged
            RemoveHandler watcher.Deleted, AddressOf OnExcelFileChanged
            watcher.EnableRaisingEvents = False
            watcher.Dispose()
        End If

        watcher = New FileSystemWatcher(Path.GetDirectoryName(excelFilePath))
        watcher.Filter = Path.GetFileName(excelFilePath)
        watcher.NotifyFilter = NotifyFilters.LastWrite Or NotifyFilters.FileName Or NotifyFilters.Size
        AddHandler watcher.Changed, AddressOf OnExcelFileChanged
        AddHandler watcher.Renamed, AddressOf OnExcelFileChanged
        AddHandler watcher.Created, AddressOf OnExcelFileChanged
        AddHandler watcher.Deleted, AddressOf OnExcelFileChanged
        watcher.EnableRaisingEvents = True

        ' Logging to verify watcher is set up
        Debug.WriteLine("FileSystemWatcher initialized for: " & watcher.Path & "\" & watcher.Filter)
    End Sub

    Private Sub PopulateSheetNames()
        ComboBoxSheets.Items.Clear()

        ' Open the workbook, populate sheet names, then close it
        Dim workbook As Workbook = excelApp.Workbooks.Open(excelFilePath)
        For Each sheet As Worksheet In workbook.Sheets
            ComboBoxSheets.Items.Add(sheet.Name)
        Next
        workbook.Close(False)
        Marshal.ReleaseComObject(workbook)

        If ComboBoxSheets.Items.Count > 0 Then
            ComboBoxSheets.SelectedIndex = 0
        End If
    End Sub

    Private Sub ButtonGetCellValue_Click(sender As Object, e As EventArgs) Handles ButtonGetCellValue.Click
        UpdateCellValue()
    End Sub

    Private Sub OnExcelFileChanged(source As Object, e As FileSystemEventArgs)
        ' Logging the event
        Debug.WriteLine("FileSystemWatcher triggered: " & e.ChangeType.ToString())

        ' Check the time interval since the last event
        Dim currentTime As DateTime = DateTime.Now
        If (currentTime - lastEventTime).TotalMilliseconds < debounceInterval Then
            ' Ignore the event if it occurs within the debounce interval
            Return
        End If
        lastEventTime = currentTime

        ' Reset the debounce timer
        If debounceTimer IsNot Nothing Then
            debounceTimer.Change(500, Timeout.Infinite)
        Else
            debounceTimer = New Timer(AddressOf DebouncedUpdateCellValue, Nothing, 500, Timeout.Infinite)
        End If
    End Sub

    Private Sub DebouncedUpdateCellValue(state As Object)
        ' Invoke on the UI thread to safely update the UI
        Me.Invoke(New MethodInvoker(Sub() UpdateCellValue()))
    End Sub

    Private Sub UpdateCellValue()
        If ComboBoxSheets.SelectedItem IsNot Nothing Then
            Dim selectedSheetName As String = ComboBoxSheets.SelectedItem.ToString()
            Dim cellReference As String = TextBoxCellReference.Text

            If Not String.IsNullOrEmpty(selectedSheetName) AndAlso Not String.IsNullOrEmpty(cellReference) Then
                ' Open the workbook, read the cell value, then close it
                Dim workbook As Workbook = excelApp.Workbooks.Open(excelFilePath)
                Dim sheet As Worksheet = CType(workbook.Sheets(selectedSheetName), Worksheet)
                Dim cell As Range = sheet.Range(cellReference)

                ' Display the updated cell value in the label
                LabelCellValue.Text = If(cell.Value IsNot Nothing, cell.Value.ToString(), "Empty")

                ' Close the workbook and release resources
                workbook.Close(False)
                Marshal.ReleaseComObject(sheet)
                Marshal.ReleaseComObject(cell)
                Marshal.ReleaseComObject(workbook)
            End If
        End If
    End Sub

    Private Sub SaveConfig(filePath As String, sheetName As String, cellReference As String)
        Dim config As New ExcelConfig With {
            .ExcelFilePath = filePath,
            .SheetName = sheetName,
            .CellReference = cellReference
        }
        Dim json As String = JsonConvert.SerializeObject(config, Newtonsoft.Json.Formatting.Indented)
        File.WriteAllText("config.json", json)
    End Sub

    Private Function LoadConfig() As ExcelConfig
        If File.Exists("config.json") Then
            Dim json As String = File.ReadAllText("config.json")
            Return JsonConvert.DeserializeObject(Of ExcelConfig)(json)
        End If
        Return Nothing
    End Function

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Clean up resources
        If excelApp IsNot Nothing Then
            excelApp.Quit()
            Marshal.ReleaseComObject(excelApp)
        End If

        If watcher IsNot Nothing Then
            watcher.EnableRaisingEvents = False
            watcher.Dispose()
        End If

        If debounceTimer IsNot Nothing Then
            debounceTimer.Dispose()
        End If
    End Sub
End Class
