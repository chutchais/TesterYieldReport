Imports Microsoft.Office.Interop

Public Class clsReport

    Dim vLots As New List(Of Lot)

    Public Property Tested As Integer
    Public Property Passed As Integer
    Public Property Failed As Integer
    Public ReadOnly Property Lots() As List(Of Lot)
        Get
            Return vLots
        End Get
    End Property



    Public Class Lot

        Dim vFiles As New List(Of clsEPRO)

        Public Property Name As String
        Public Property Tested As Integer
        Public Property Passed As Integer
        Public Property Failed As Integer

        Public ReadOnly Property Files() As List(Of clsEPRO)
            Get
                Return vFiles
            End Get
        End Property

        Public ReadOnly Property Yield() As Decimal
            Get
                Try
                    Return (Passed / Tested) * 100
                Catch ex As Exception
                    Return 0
                End Try

            End Get
        End Property
    End Class


    Public Sub ExportToExcel(ByVal filepath As String)
        Dim strFileName As String = filepath

        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If

        'Create Excel Book
        Dim _excel As New Excel.Application
        Dim wBook As Excel.Workbook
        Dim wSheet As Excel.Worksheet

        wBook = _excel.Workbooks.Add()
        wSheet = wBook.ActiveSheet()

        'Dim dt As System.Data.DataTable = dtTemp
        ' Dim dc As System.Data.DataColumn
        'Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0


        'Loop for Lot in Report object
        Dim vFirstDatable As Boolean = True
        For Each l In Lots
            'Loop for file(EPRO in Lot)
            For Each oEPRO In l.Files
                Dim vDt As New DataTable
                vDt = oEPRO.DataTable
                'Fill in Excel
                'Make column
                If vFirstDatable Then
                    For Each dc In vDt.Columns
                        colIndex = colIndex + 1
                        wSheet.Cells(1, colIndex) = dc.ColumnName
                    Next
                    vFirstDatable = False
                    rowIndex = rowIndex + 1
                End If
                'Start to fill row to excel
                colIndex = 0
                For Each dr In vDt.Rows
                    For Each dc In vDt.Columns
                        colIndex = colIndex + 1
                        wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                    Next
                    rowIndex = rowIndex + 1
                Next
                '------------

            Next
        Next

        wSheet.Columns.AutoFit()
        wBook.SaveAs(strFileName)

        ReleaseObject(wSheet)
        wBook.Close(False)
        ReleaseObject(wBook)
        _excel.Quit()
        ReleaseObject(_excel)
        GC.Collect()
        System.Diagnostics.Process.Start(filepath)
    End Sub

    Sub ReleaseObject(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub
End Class
