Imports System.IO
Imports System.Reflection
Imports TesterYieldReport.clsReport


Public Class frmMain
    Dim objEPROs As New List(Of clsEPRO)
    Dim objReport As New clsReport

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Select Tester type EPRO 
        cbTesterType.SelectedIndex = 0

        'Files information
        With ListView1
            .View = View.Details
            .AllowColumnReorder = True
            .FullRowSelect = True
            .AllowDrop = True
            .Columns.Add("File name", 100, HorizontalAlignment.Left)
            .Columns.Add("Tested", 50, HorizontalAlignment.Right)
            .Columns.Add("Passed", 50, HorizontalAlignment.Right)
            .Columns.Add("Failed", 50, HorizontalAlignment.Right)
            .Columns.Add("Yield(%)", 50, HorizontalAlignment.Right)
            .Columns.Add("Status", 300, HorizontalAlignment.Left)
        End With

        'lot infomation
        With lvLot
            .View = View.Details
            .AllowColumnReorder = True
            .FullRowSelect = True
            .Columns.Add("Lot number", 100, HorizontalAlignment.Left)
            .Columns.Add("file(s)", 50, HorizontalAlignment.Center)
            .Columns.Add("Tested", 50, HorizontalAlignment.Right)
            .Columns.Add("Passed", 50, HorizontalAlignment.Right)
            .Columns.Add("Failed", 50, HorizontalAlignment.Right)
            .Columns.Add("Yield(%)", 50, HorizontalAlignment.Right)
            .Columns.Add("File list", 400, HorizontalAlignment.Left)
        End With

        Dim versionNumber As Version
        versionNumber = Assembly.GetExecutingAssembly().GetName().Version
        Me.Text = Me.Text & " version : " & versionNumber.ToString

    End Sub

    Sub generate_report_file(vFileName As String)
        Me.Cursor = Cursors.WaitCursor
        Dim vShortFileName As String
        vShortFileName = My.Computer.FileSystem.GetFileInfo(vFileName).Name

        Dim selectedValue As clsEPRO
        selectedValue = objEPROs.Find(Function(EPRO) EPRO.keyName = vShortFileName)

        If Not selectedValue Is Nothing Then
            MsgBox("File " & vShortFileName & " already exist in list", MsgBoxStyle.Exclamation,
                    "File already exist")
            Exit Sub
        End If

        Dim vSuccessFile As Integer = 0
        Dim vErrorFile As Integer = 0

        ' Do work, example
        Dim objEpro As New clsEPRO(vFileName)
        'ListBox1.Items.Add(file.Name & ":" & objEpro.message)

        Dim newItem As ListViewItem = New ListViewItem(vShortFileName)

        If objEpro.completed Then
            vSuccessFile = vSuccessFile + 1
            newItem.SubItems.Add(objEpro.Tested)
            newItem.SubItems.Add(objEpro.Passed)
            newItem.SubItems.Add(objEpro.Failed)
            newItem.SubItems.Add(Format(objEpro.Yield, "0.00"))


        Else
            vErrorFile = vErrorFile + 1
            newItem.SubItems.Add("")
            newItem.SubItems.Add("")
            newItem.SubItems.Add("")
            newItem.SubItems.Add("")
        End If
        newItem.SubItems.Add(objEpro.message)


        ListView1.Items.Add(newItem)
        objEPROs.Add(objEpro)



        ToolStripStatusLabel1.Text = "Total " & objEPROs.Count.ToString & " file(s) "
        Me.Cursor = Cursors.Default
    End Sub

    Sub summary_report_object()
        objReport.Lots.Clear()
        With objReport
            .Tested = 0
            .Passed = 0
            .Failed = 0
            For Each iepro In objEPROs
                .Tested = .Tested + iepro.Tested
                .Passed = .Passed + iepro.Passed
                .Failed = .Failed + iepro.Failed
                'Update lot information
                Dim objReportLots As Lot
                objReportLots = .Lots.Find(Function(lot) lot.Name = iepro.LotNumber)


                If objReportLots Is Nothing Then
                    'MsgBox("Not exist")
                    .Lots.Add(New Lot With {.Name = iepro.LotNumber,
                                            .Tested = iepro.Tested,
                                            .Passed = iepro.Passed,
                                            .Failed = iepro.Failed})

                Else
                    'MsgBox("exist")
                    With objReportLots
                        .Tested = .Tested + iepro.Tested
                        .Passed = .Passed + iepro.Passed
                        .Failed = .Failed + iepro.Failed
                    End With
                End If

                objReportLots = .Lots.Find(Function(lot) lot.Name = iepro.LotNumber)
                objReportLots.Files.Add(iepro)

            Next

            'Update Lot summary listview
            lvLot.Items.Clear()
            For Each i In objReport.Lots
                Dim newLotItem As ListViewItem = New ListViewItem(i.Name)
                newLotItem.SubItems.Add(i.Files.Count)
                newLotItem.SubItems.Add(i.Tested)
                newLotItem.SubItems.Add(i.Passed)
                newLotItem.SubItems.Add(i.Failed)
                newLotItem.SubItems.Add(Format(i.Yield, "0.00"))
                Dim vFileListStr As String = ""
                For Each f In i.Files
                    vFileListStr = vFileListStr & f.keyName & ","
                Next
                'remove last comma(,)
                vFileListStr = vFileListStr.Trim.Remove(vFileListStr.Length - 1)
                'Add Files to sub item
                newLotItem.SubItems.Add(vFileListStr)
                'Add Sub item to listview
                lvLot.Items.Add(newLotItem)
            Next

            '--------------------

            tssSummary.Text = "Total Tested : " & .Tested.ToString &
                           " Passed : " & .Passed.ToString &
                           " Failed : " & .Failed.ToString &
                           "(Yield : " & Format((.Passed / .Tested) * 100, "0.00") & "%)"
        End With
    End Sub

    Sub generate_report_folder(vPath As String)
        Me.Cursor = Cursors.WaitCursor

        objEPROs.Clear()
        Dim dinfo As New DirectoryInfo(vPath)
        'Get the files based on .txt extension
        Dim files As FileInfo() = dinfo.GetFiles("*.sum")




        Dim vSuccessFile As Integer = 0
        Dim vErrorFile As Integer = 0

        For Each file As FileInfo In files
            ' Do work, example
            Dim objEpro As New clsEPRO(file.FullName)
            'ListBox1.Items.Add(file.Name & ":" & objEpro.message)
            Dim newItem As ListViewItem = New ListViewItem(file.Name)

            If objEpro.completed Then
                vSuccessFile = vSuccessFile + 1
                newItem.SubItems.Add(objEpro.Tested)
                newItem.SubItems.Add(objEpro.Passed)
                newItem.SubItems.Add(objEpro.Failed)
                newItem.SubItems.Add(Format(objEpro.Yield, "0.00"))
            Else
                vErrorFile = vErrorFile + 1
                newItem.SubItems.Add("")
                newItem.SubItems.Add("")
                newItem.SubItems.Add("")
                newItem.SubItems.Add("")
            End If
            newItem.SubItems.Add(objEpro.message)


            ListView1.Items.Add(newItem)
            objEPROs.Add(objEpro)

        Next
        'update summary report object
        summary_report_object()
        Me.Cursor = Cursors.Default
        ToolStripStatusLabel1.Text = "Total " & objEPROs.Count.ToString & " file(s) " &
            ",Successful " & vSuccessFile.ToString & " , Error " & vErrorFile.ToString & " file(s)"
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            tbFolder.Text = FolderBrowserDialog1.SelectedPath
            generate_report_folder(FolderBrowserDialog1.SelectedPath)
            'Properties.Settings.Default.LastSelectedFolder = FolderBrowserDialog1.SelectedPath.ToString()
            'Properties.Settings.Default.Save()
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ListView1.Items.Clear()
        objEPROs.Clear()
        summary_report_object()
        ToolStripStatusLabel1.Text = ""
        tssSummary.Text = ""
    End Sub

    Private Sub ListView1_DragDrop(sender As Object, e As DragEventArgs) Handles ListView1.DragDrop

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim droppedObjects() As String
            ' Assign the files to an array.
            droppedObjects = e.Data.GetData(DataFormats.FileDrop)

            If Directory.Exists(droppedObjects(0)) Then
                'Return FileSystemObject.Directory
                If droppedObjects.Length > 1 Then
                    MsgBox("Not allow multiple folder",
                           MsgBoxStyle.Critical, "Not allow mutilple folder")
                    Exit Sub
                End If
                generate_report_folder(droppedObjects(0))
            ElseIf File.Exists(droppedObjects(0)) Then
                For Each file In droppedObjects
                    generate_report_file(file)
                Next



            Else
                MsgBox("Invalid file or directory",
                       MsgBoxStyle.Critical, "Invalid file or directory")
            End If

            summary_report_object()

        End If
    End Sub

    Private Sub ListView1_DragEnter(sender As Object, e As DragEventArgs) Handles ListView1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If tbFolder.Text = "" Then
            MsgBox("Please enter folder name", MsgBoxStyle.Exclamation, "Folder is blank")
            Exit Sub
        End If

        If Not Directory.Exists(tbFolder.Text) Then
            MsgBox("Folder doesn't exist", MsgBoxStyle.Exclamation, "Folder doesn't exist")
            Exit Sub
        End If

        generate_report_folder(tbFolder.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For Each listItem As ListViewItem In ListView1.SelectedItems
            listItem.Remove()
            'remove from object list
            objEPROs.RemoveAll(Function(objEPRO) objEPRO.keyName = listItem.Text)

        Next
        summary_report_object()
        ToolStripStatusLabel1.Text = "Total " & objEPROs.Count.ToString & " file(s) "

    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Me.Cursor = Cursors.WaitCursor
        objReport.ExportToExcel("d:\test_epro.xlsx")
        Me.Cursor = Cursors.Default
    End Sub
End Class
