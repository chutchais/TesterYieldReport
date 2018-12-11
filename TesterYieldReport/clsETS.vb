Imports System.IO

Public Class clsETS
    Public vFileName As String = ""
    Dim vLotNumber As String = ""
    Dim vSeqNumber As String = ""
    Dim vTestName As String = ""
    Dim vTester As String = ""
    Dim vOperatorName As String = ""
    Dim vProgramRevision As String = ""
    Dim vStartDate As String = ""
    Dim vStopDate As String = ""
    Dim vTested As Integer = 0
    Dim vPassed As Integer = 0
    Dim vFailed As Integer = 0

    Dim vSummaryName As String = ""
    Dim vAssyNumber As String = ""


    Dim vTemperature As String = ""

    Dim vHandlerId As String = ""
    Dim vProgramFileName As String = ""
    Dim vReportDate As String = ""
    Dim vReportTime As String = ""
    Dim vProgramName As String = ""
    Dim vSystemId As String = ""

    'Dim vIBs As New List(Of IB)
    'Dim vDBs As New List(Of DB)
    'Dim vBINs As New List(Of BIN)

    Dim vSfwr As New List(Of Bin)
    Dim vHdwr As New List(Of Bin)
    Dim vKeyName As String = "" 'File name (short)

    Dim vCompleted As Boolean
    Dim vMessage As String


    Dim vTotalUnitTested_Line_Started As Boolean
    Dim vSfwr_Line_Started As Boolean
    Dim vHdwr_Line_Started As Boolean

    Public Sub New(ByVal fileName As String)
        vFileName = fileName
        vKeyName = My.Computer.FileSystem.GetFileInfo(vFileName).Name
        process_file()
    End Sub

    Public ReadOnly Property keyName() As String
        Get
            Return vKeyName
        End Get
    End Property

    Public ReadOnly Property DataTable() As DataTable
        Get
            Return convert_to_datatable()
        End Get
    End Property

    'Status 
    Public ReadOnly Property completed() As Boolean
        Get
            Return vCompleted
        End Get
    End Property
    Public ReadOnly Property message() As String
        Get
            Return vMessage
        End Get
    End Property



    Public Property FileName() As String
        Get
            Return vFileName
        End Get
        Set(ByVal value As String)
            vFileName = value
        End Set
    End Property

    'Public ReadOnly Property IBs() As List(Of IB)
    '    Get
    '        Return vIBs
    '    End Get
    'End Property
    'Public ReadOnly Property DBs() As List(Of DB)
    '    Get
    '        Return vDBs
    '    End Get
    'End Property
    'Public ReadOnly Property BINs() As List(Of BIN)
    '    Get
    '        Return vBINs
    '    End Get
    'End Property

    'Public ReadOnly Property SummaryName() As String
    '    Get
    '        Return vSummaryName
    '    End Get
    'End Property

    'Public ReadOnly Property AssyNumber() As String
    '    Get
    '        Return vLotNumber
    '    End Get
    'End Property

    Public ReadOnly Property LotNumber() As String
        Get
            Return vLotNumber
        End Get
    End Property

    Public ReadOnly Property SeqNumber() As String
        Get
            Return vLotNumber
        End Get
    End Property

    Public ReadOnly Property ReportStartDate() As String
        Get
            Return vStartDate
        End Get
    End Property

    Public ReadOnly Property ReportStopDate() As String
        Get
            Return vStopDate
        End Get
    End Property



    Public ReadOnly Property OperatorName() As String
        Get
            Return vOperatorName
        End Get
    End Property
    'Public ReadOnly Property Temperature() As String
    '    Get
    '        Return vTemperature
    '    End Get
    'End Property
    Public ReadOnly Property Tester() As String
        Get
            Return vTester
        End Get
    End Property

    Public ReadOnly Property ProgramRevision() As String
        Get
            Return vProgramRevision
        End Get
    End Property
    'Public ReadOnly Property HandlerId() As String
    '    Get
    '        Return vHandlerId
    '    End Get
    'End Property
    'Public ReadOnly Property ProgramFileName() As String
    '    Get
    '        Return vProgramFileName
    '    End Get
    'End Property
    'Public ReadOnly Property PatternName() As String
    '    Get
    '        Return ""
    '    End Get
    'End Property
    'Public ReadOnly Property ReportDate() As String
    '    Get
    '        Return vReportDate
    '    End Get
    'End Property
    'Public ReadOnly Property ReportTime() As String
    '    Get
    '        Return vReportTime
    '    End Get
    'End Property
    'Public ReadOnly Property SystemId() As String
    '    Get
    '        Return vSystemId
    '    End Get
    'End Property
    Public ReadOnly Property ProgramName() As String
        Get
            Return vProgramName
        End Get
    End Property
    Public ReadOnly Property Tested() As Integer
        Get
            Return vTested
        End Get
    End Property
    Public ReadOnly Property Passed() As Integer
        Get
            Return vPassed
        End Get
    End Property
    Public ReadOnly Property Failed() As Integer
        Get
            Return vTested - vPassed
        End Get
    End Property
    Public ReadOnly Property Yield() As Decimal
        Get
            Try
                Return (vPassed / vTested) * 100
            Catch ex As Exception
                Return 0
            End Try

        End Get
    End Property

    Overridable Sub process_file()
        Dim line_number As Integer = 0
        Dim line As String = ""
        Try
            Dim list As New List(Of String)

            ' Open file.txt with the Using statement.
            Using r As StreamReader = New StreamReader(vFileName)
                ' Store contents in this String.


                ' Read first line.
                line = r.ReadLine
                ' Loop over each line in file, While list is Not Nothing.
                Do While (Not line Is Nothing)
                    line_number = line_number + 1
                    ' Add this line to list.
                    list.Add(line)
                    'recognize data
                    recognize_data(line)
                    ' Read in the next line.
                    line = r.ReadLine
                Loop
            End Using
            vCompleted = True
            vMessage = "Successful"
        Catch ex As Exception
            vCompleted = False
            vMessage = "Error on line " & line_number.ToString & ":" & line & vbCrLf &
                        ex.Message
        End Try

    End Sub

    Sub recognize_data(line_data As String)
        If line_data = "" Then
            Exit Sub
        End If

        'DUTs Tested
        If line_data.Contains("DUTs Tested") Then
            vTotalUnitTested_Line_Started = True
            Exit Sub
        End If
        'Sfwr Line started
        If line_data.Contains("Sfwr   Bin") Then
            vSfwr_Line_Started = True
            vHdwr_Line_Started = False
            Exit Sub
        End If
        'Hdwr Line Started
        If line_data.Contains("Hdwr   Bin") Then
            vHdwr_Line_Started = True
            vSfwr_Line_Started = False
            Exit Sub
        End If

        If line_data.Contains("%") And vTotalUnitTested_Line_Started Then
            Dim vTestSplited As String()
            vTestSplited = line_data.Split("|")
            vTested = Val(vTestSplited(0))
            vPassed = Val(vTestSplited(1))
            vFailed = Val(vTestSplited(2))

            vTotalUnitTested_Line_Started = False
            Exit Sub
        End If

        If line_data.Contains("%") And vSfwr_Line_Started Then
            Dim vLineSplited As String()
            Dim vBinNumber As String = ""
            Dim vBinType As String = ""
            Dim vBinYield As Decimal = 0
            Dim vBinCount As Integer = 0
            Dim vBinDesc As String = ""
            vLineSplited = line_data.Split(" ")
            'find Bin number
            If vLineSplited(4) <> "" Then
                vBinNumber = "SW_Bin" & vLineSplited(4)
                vBinType = vLineSplited(10)
            End If
            If vLineSplited(5) <> "" Then
                vBinNumber = "SW_Bin" & vLineSplited(5)
                vBinType = vLineSplited(11)
            End If

            'Find Bin description
            For i = 12 To vLineSplited.Length - 1 - 10
                If vLineSplited(i) <> "" Then
                    vBinDesc = vBinDesc & " " & vLineSplited(i)
                End If

            Next
            vBinDesc = vBinDesc.Trim
            'Find Bin failed Count

            If vLineSplited(vLineSplited.Length - 1 - 9) <> "" Then
                vBinCount = Val(vLineSplited(vLineSplited.Length - 1 - 9))
            Else
                vBinCount = Val(vLineSplited(vLineSplited.Length - 1 - 8))
            End If
            'find Bin Yield
            vBinYield = Val(vLineSplited(vLineSplited.Length - 1).Replace("%", ""))

            'Add to Software List
            vSfwr.Add(New Bin With {.Name = vBinNumber, .Description = vBinDesc,
                                       .Count = vBinCount, .Yield = vBinYield})
            Exit Sub
        End If


        If line_data.Contains("%") And vHdwr_Line_Started Then
            Dim vLineSplited As String()
            Dim vBinNumber As String = ""
            Dim vBinType As String = ""
            Dim vBinYield As Decimal = 0
            Dim vBinCount As Integer = 0
            Dim vBinDesc As String = ""
            vLineSplited = line_data.Split(" ")
            'find Bin number
            If vLineSplited(4) <> "" Then
                vBinNumber = "HW_Bin" & vLineSplited(4)
                vBinType = vLineSplited(10)
            End If
            If vLineSplited(5) <> "" Then
                vBinNumber = "HW_Bin" & vLineSplited(5)
                vBinType = vLineSplited(11)
            End If

            'Find Bin description
            For i = 12 To vLineSplited.Length - 1 - 10
                If vLineSplited(i) <> "" Then
                    vBinDesc = vBinDesc & " " & vLineSplited(i)
                End If

            Next
            vBinDesc = vBinDesc.Trim
            'Find Bin failed Count

            If vLineSplited(vLineSplited.Length - 1 - 9) <> "" Then
                vBinCount = Val(vLineSplited(vLineSplited.Length - 1 - 9))
            Else
                vBinCount = Val(vLineSplited(vLineSplited.Length - 1 - 8))
            End If
            'find Bin Yield
            vBinYield = Val(vLineSplited(vLineSplited.Length - 1).Replace("%", ""))

            'Add to Software List
            vHdwr.Add(New Bin With {.Name = vBinNumber, .Description = vBinDesc,
                                       .Count = vBinCount, .Yield = vBinYield})
            Exit Sub
        End If

        If line_data.Contains("Test Name:") Then
            vTestName = line_data.Replace("Test Name:", "").Trim
            Exit Sub
        End If




        If line_data.Contains("Report for Lot:") Then
            Dim vLotNumberStr As String = ""
            vLotNumberStr = line_data.Replace("Report for Lot:", "").Trim
            If vLotNumberStr.Split("_").Length = 3 Then
                vAssyNumber = vLotNumberStr.Split("_")(0)
                vLotNumber = vLotNumberStr.Split("_")(1)
                vSeqNumber = vLotNumberStr.Split("_")(2)
            Else
                vAssyNumber = vLotNumberStr.Split("_")(0)
                vLotNumber = vLotNumberStr.Split("_")(0)
                vSeqNumber = vLotNumberStr.Split("_")(1)
            End If
            Exit Sub
        End If

        If line_data.Contains("Temperature:") Then
            vTemperature = line_data.Replace("Temperature:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Data Collected by Station:") Then
            vTester = line_data.Replace("Data Collected by Station:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Handler ID:") Then
            vHandlerId = line_data.Replace("Handler ID:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Data Collected by Operator:") Then
            vOperatorName = line_data.Replace("Data Collected by Operator:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Program Revision:") Then
            vProgramRevision = line_data.Replace("Program Revision:", "").Trim
            Exit Sub
        End If

        'If line_data.Contains("Program Name:") Then
        '    vProgramFileName = line_data.Replace("Program Name:", "").Trim
        '    Exit Sub
        'End If

        If line_data.Contains("Data Collection Start Date:") Then
            vStartDate = line_data.Replace("Data Collection Start Date:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Data Collection Stop Date:") Then
            vStopDate = line_data.Replace("Data Collection Stop Date:", "").Trim
            Exit Sub
        End If
        '

        'Start IB data
        If line_data.Contains("IB ") Then
            Dim str_splited() As String
            Dim vIBName As String = ""
            Dim vCount As Integer = 0
            Dim vYield As Decimal = 0
            str_splited = line_data.Split(" ")

            'Name
            If str_splited(1) <> "" Then
                vIBName = "IB" & str_splited(1).Trim
            End If
            If str_splited(2) <> "" Then
                vIBName = "IB" & str_splited(2).Trim
            End If

            'Count
            If str_splited(15) <> "" Then
                vCount = Val(str_splited(15))
            End If
            If str_splited(16) <> "" Then
                vCount = Val(str_splited(16))
            End If
            If str_splited(17) <> "" Then
                vCount = Val(str_splited(17))
            End If
            If str_splited(18) <> "" Then
                vCount = Val(str_splited(18))
            End If
            If str_splited(19) <> "" Then
                vCount = Val(str_splited(19))
            End If
            'Yield
            vYield = Val(str_splited(str_splited.Length - 1).Replace("%", ""))
            'Add to List
            'vIBs.Add(New IB With {.Name = vIBName, .Count = vCount, .Yield = vYield})
            Exit Sub
        End If

        'Start DB data
        If line_data.Contains("DB ") Then
            Dim str_splited() As String
            Dim vName As String = ""
            Dim vCount As Integer = 0
            Dim vYield As Decimal = 0
            str_splited = line_data.Split(" ")

            'Name
            If str_splited(1) <> "" Then
                vName = "DB" & str_splited(1).Trim
            End If
            If str_splited(2) <> "" Then
                vName = "DB" & str_splited(2).Trim
            End If

            'Count
            If str_splited(15) <> "" Then
                vCount = Val(str_splited(15))
            End If
            If str_splited(16) <> "" Then
                vCount = Val(str_splited(16))
            End If
            If str_splited(17) <> "" Then
                vCount = Val(str_splited(17))
            End If
            If str_splited(18) <> "" Then
                vCount = Val(str_splited(18))
            End If
            If str_splited(19) <> "" Then
                vCount = Val(str_splited(19))
            End If
            'Yield
            vYield = Val(str_splited(str_splited.Length - 1).Replace("%", ""))
            'Add to List
            'vDBs.Add(New DB With {.Name = vName, .Count = vCount, .Yield = vYield})
            Exit Sub
        End If

        'Start INTERFACE BINS data
        If line_data.Contains("Data Bin ") Or line_data.Contains("Bin  ") Then
            Dim str_splited() As String
            Dim vName As String = ""
            Dim vDescription As String = ""
            str_splited = line_data.Split(" ")

            'Name
            If str_splited(0) = "Bin" Then
                If str_splited(2) <> "" Then
                    vName = "BIN" & str_splited(2).Trim.Replace(":", "")
                End If
                If str_splited(2) <> "" Then
                    vName = "BIN" & str_splited(2).Trim.Replace(":", "")
                End If

                'Description
                For i = 3 To str_splited.Length - 1
                    vDescription = vDescription & " " & str_splited(i) & " "
                Next
            Else
                If str_splited(4) <> "" Then
                    vName = "BIN" & str_splited(4).Trim.Replace(":", "")
                End If
                If str_splited(4) <> "" Then
                    vName = "BIN" & str_splited(4).Trim.Replace(":", "")
                End If

                'Description
                For i = 5 To str_splited.Length - 1
                    vDescription = vDescription & " " & str_splited(i) & " "
                Next
            End If



            vDescription = vDescription.Trim

            'Add to List
            'vBINs.Add(New BIN With {.Name = vName, .Description = vDescription})
            Exit Sub
        End If

    End Sub

    Public Class Bin
        Public Property Name As String
        Public Property Description As String
        Public Property Count As Integer
        Public Property Yield As Decimal
    End Class

    Function convert_to_datatable(Optional SW_col_count As Integer = 32,
                                  Optional HW_col_count As Integer = 7) As DataTable
        Try
            Dim dtFile As New DataTable
            Dim i As Integer
            'Make Table column
            With dtFile
                .Columns.Add("lot", GetType(String))
                .Columns.Add("seq", GetType(String))
                .Columns.Add("testname", GetType(String))
                .Columns.Add("operator", GetType(String))
                .Columns.Add("tester", GetType(String))
                .Columns.Add("rev", GetType(String))
                .Columns.Add("start_date", GetType(String))
                .Columns.Add("stop_date", GetType(String))
                .Columns.Add("tested", GetType(String))
                .Columns.Add("passed", GetType(String))
                .Columns.Add("failed", GetType(String))
                .Columns.Add("yield", GetType(String))
                'Add Software column
                For i = 1 To SW_col_count
                    .Columns.Add("SW_Bin" & Str(i).Trim, GetType(String))
                Next
                'Add Hardware column
                For i = 1 To HW_col_count
                    .Columns.Add("HW_Bin" & Str(i).Trim, GetType(String))
                Next

            End With
            'Fill data
            Dim iRow As DataRow = dtFile.NewRow()
            With iRow
                iRow("lot") = vLotNumber
                iRow("seq") = vSeqNumber
                iRow("testname") = vTestName
                iRow("operator") = vOperatorName
                iRow("tester") = vTester
                iRow("rev") = vProgramRevision
                iRow("start_date") = vStartDate
                iRow("stop_date") = vStopDate
                iRow("tested") = vTested
                iRow("passed") = vPassed
                iRow("failed") = vTested - vPassed
                iRow("yield") = Yield

                'Dim selectedValue As clsEPRO
                'selectedValue = objEPROs.Find(Function(EPRO) EPRO.keyName = vShortFileName)
                Dim colName As String = ""
                Dim objSwBin As Bin
                Dim objHwBin As Bin

                'Fill IB column
                For i = 1 To SW_col_count
                    colName = "SW_Bin" & Str(i).Trim
                    objSwBin = vSfwr.Find(Function(sw) sw.Name = colName)
                    If objSwBin Is Nothing Then
                        iRow("SW_Bin" & Str(i).Trim) = ""
                    Else
                        iRow("SW_Bin" & Str(i).Trim) = objSwBin.Count
                    End If

                Next
                'Fill DB column

                For i = 1 To HW_col_count
                    colName = "HW_Bin" & Str(i).Trim
                    objHwBin = vHdwr.Find(Function(hw) hw.Name = colName)
                    If objHwBin Is Nothing Then
                        iRow("HW_Bin" & Str(i).Trim) = ""
                    Else
                        iRow("HW_Bin" & Str(i).Trim) = objHwBin.Count
                    End If
                Next

            End With
            dtFile.Rows.Add(iRow)
            Return dtFile
        Catch ex As Exception
            Return Nothing
        End Try

    End Function
End Class
