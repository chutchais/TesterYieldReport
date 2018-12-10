
Imports System.IO

Public Class clsEPRO
    Public vFileName As String = ""
    Dim vSummaryName As String = ""
    Dim vAssyNumber As String = ""
    Dim vLotNumber As String = ""
    Dim vSeqNumber As String = ""
    Dim vOperatorName As String = ""
    Dim vTemperature As String = ""
    Dim vTester As String = ""
    Dim vHandlerId As String = ""
    Dim vProgramFileName As String = ""
    Dim vReportDate As String = ""
    Dim vReportTime As String = ""
    Dim vProgramName As String = ""
    Dim vSystemId As String = ""
    Dim vTested As Integer = 0
    Dim vPassed As Integer = 0
    Dim vFailed As Integer = 0
    Dim vIBs As New List(Of IB)
    Dim vDBs As New List(Of DB)
    Dim vBINs As New List(Of BIN)
    Dim vKeyName As String = ""

    Dim vCompleted As Boolean
    Dim vMessage As String


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

    Public ReadOnly Property IBs() As List(Of IB)
        Get
            Return vIBs
        End Get
    End Property
    Public ReadOnly Property DBs() As List(Of DB)
        Get
            Return vDBs
        End Get
    End Property
    Public ReadOnly Property BINs() As List(Of BIN)
        Get
            Return vBINs
        End Get
    End Property

    Public ReadOnly Property SummaryName() As String
        Get
            Return vSummaryName
        End Get
    End Property

    Public ReadOnly Property AssyNumber() As String
        Get
            Return vLotNumber
        End Get
    End Property

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

    Public ReadOnly Property OperatorName() As String
        Get
            Return vOperatorName
        End Get
    End Property
    Public ReadOnly Property Temperature() As String
        Get
            Return vTemperature
        End Get
    End Property
    Public ReadOnly Property Tester() As String
        Get
            Return vTester
        End Get
    End Property
    Public ReadOnly Property HandlerId() As String
        Get
            Return vHandlerId
        End Get
    End Property
    Public ReadOnly Property ProgramFileName() As String
        Get
            Return vProgramFileName
        End Get
    End Property
    Public ReadOnly Property PatternName() As String
        Get
            Return ""
        End Get
    End Property
    Public ReadOnly Property ReportDate() As String
        Get
            Return vReportDate
        End Get
    End Property
    Public ReadOnly Property ReportTime() As String
        Get
            Return vReportTime
        End Get
    End Property
    Public ReadOnly Property SystemId() As String
        Get
            Return vSystemId
        End Get
    End Property
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

        '        Lot Number: S4UBH370001_TS00182921-001_F2
        '    Operator:  20358
        ' Temperature: 25 C
        '    Tester #: 1
        '  Handler ID: SRM#2
        'Program Name: W :  \Tip Files\P5119D0D.TIP

        If line_data.Contains("Summary Name") Then
            vSummaryName = line_data.Replace("Summary Name:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Lot Number:") Then
            Dim vLotNumberStr As String = ""
            vLotNumberStr = line_data.Replace("Lot Number:", "").Trim
            If vLotNumberStr.Split("_").Length = 3 Then
                vAssyNumber = vLotNumberStr.Split("_")(0)
                vLotNumber = vLotNumberStr.Split("_")(1)
                vSeqNumber = vLotNumberStr.Split("_")(2)
            Else
                vAssyNumber = vLotNumberStr.Split("_")(0)
                vLotNumber = vLotNumberStr.Split("_")(0)
                vSeqNumber = ""
            End If
            Exit Sub
        End If

        If line_data.Contains("Temperature:") Then
            vTemperature = line_data.Replace("Temperature:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Tester #:") Then
            vTester = line_data.Replace("Tester #:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Handler ID:") Then
            vHandlerId = line_data.Replace("Handler ID:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Operator:") Then
            vOperatorName = line_data.Replace("Operator:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Program Name:") Then
            vProgramFileName = line_data.Replace("Program Name:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Date:") Then
            vReportDate = line_data.Replace("Date:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Time:") Then
            vReportTime = line_data.Replace("Time:", "").Trim
            Exit Sub
        End If
        '
        If line_data.Contains("PROGRAM NAME:") Then
            vProgramName = line_data.Replace("PROGRAM NAME:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("SYSTEM ID #") Then
            vSystemId = line_data.Replace("SYSTEM ID", "").Trim
            Exit Sub
        End If

        If line_data.Contains("UNITS TESTED") Then
            Dim str_splited() As String
            str_splited = line_data.Split(" ")

            'vTested = IIf(str_splited(7) <> "", Val(str_splited(7)), Val(str_splited(8)))
            For i = 4 To str_splited.Length - 2
                If str_splited(i) <> "" Then
                    vTested = Val(str_splited(i))
                    Exit For
                End If
            Next

            Exit Sub
        End If

        If line_data.Contains("UNITS PASSED") Then
            Dim str_splited() As String
            str_splited = line_data.Split(" ")
            'vPassed = IIf(str_splited(7) <> "", Val(str_splited(7)), Val(str_splited(8)))
            For i = 4 To str_splited.Length - 2
                If str_splited(i) <> "" Then
                    vPassed = Val(str_splited(i))
                    Exit For
                End If
            Next
            Exit Sub
        End If

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
            vIBs.Add(New IB With {.Name = vIBName, .Count = vCount, .Yield = vYield})
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
            vDBs.Add(New DB With {.Name = vName, .Count = vCount, .Yield = vYield})
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
            vBINs.Add(New BIN With {.Name = vName, .Description = vDescription})
            Exit Sub
        End If

    End Sub

    Public Class IB
        Public Property Name As String
        Public Property Count As Integer
        Public Property Yield As Decimal
    End Class

    Public Class DB
        Public Property Name As String
        Public Property Count As Integer
        Public Property Yield As Decimal
    End Class

    Public Class BIN
        Public Property Name As String
        Public Property Description As String
    End Class


    Function convert_to_datatable(Optional IB_col_count As Integer = 7,
                                  Optional DB_col_count As Integer = 32,
                                  Optional BIN_col_count As Integer = 32) As DataTable
        Try
            Dim dtEPRO As New DataTable
            Dim i As Integer
            'Make Table column
            With dtEPRO
                .Columns.Add("assy", GetType(String))
                .Columns.Add("lot", GetType(String))
                .Columns.Add("seq", GetType(String))
                .Columns.Add("operator", GetType(String))
                .Columns.Add("temperature", GetType(String))
                .Columns.Add("tester", GetType(String))
                .Columns.Add("handler", GetType(String))
                .Columns.Add("filename", GetType(String))
                .Columns.Add("date", GetType(String))
                .Columns.Add("time", GetType(String))
                .Columns.Add("system_id", GetType(String))
                .Columns.Add("program_name", GetType(String))
                .Columns.Add("tested", GetType(String))
                .Columns.Add("passed", GetType(String))
                .Columns.Add("failed", GetType(String))
                .Columns.Add("yield", GetType(String))
                'Add IB column
                For i = 1 To IB_col_count
                    .Columns.Add("IB" & Str(i).Trim, GetType(String))
                Next
                'Add DB column
                For i = 1 To DB_col_count
                    .Columns.Add("DB" & Str(i).Trim, GetType(String))
                Next
                'Add BIN column
                For i = 1 To BIN_col_count
                    .Columns.Add("BIN" & Format(i, "00"), GetType(String))
                Next
            End With
            'Fill data
            Dim iRow As DataRow = dtEPRO.NewRow()
            With iRow
                iRow("assy") = vAssyNumber
                iRow("lot") = vLotNumber
                iRow("seq") = vSeqNumber
                iRow("operator") = vOperatorName
                iRow("temperature") = vTemperature
                iRow("tester") = vTester
                iRow("handler") = vHandlerId
                iRow("filename") = vFileName
                iRow("date") = vReportDate
                iRow("time") = vReportTime
                iRow("system_id") = vSystemId
                iRow("program_name") = vProgramName
                iRow("tested") = vTested
                iRow("passed") = vPassed
                iRow("failed") = vTested - vPassed
                iRow("yield") = Yield

                'Dim selectedValue As clsEPRO
                'selectedValue = objEPROs.Find(Function(EPRO) EPRO.keyName = vShortFileName)
                Dim colName As String = ""
                Dim objIB As IB
                Dim objDB As DB
                Dim objBIN As BIN
                'Fill IB column
                For i = 1 To IB_col_count
                    colName = "IB" & Str(i).Trim
                    objIB = vIBs.Find(Function(ib) ib.Name = colName)
                    If objIB Is Nothing Then
                        iRow("IB" & Str(i).Trim) = ""
                    Else
                        iRow("IB" & Str(i).Trim) = objIB.Count
                    End If

                Next
                'Fill DB column

                For i = 1 To DB_col_count
                    colName = "DB" & Str(i).Trim
                    objDB = vDBs.Find(Function(db) db.Name = colName)
                    If objDB Is Nothing Then
                        iRow("DB" & Str(i).Trim) = ""
                    Else
                        iRow("DB" & Str(i).Trim) = objDB.Count
                    End If
                Next
                'Fill BIN column
                For i = 1 To BIN_col_count
                    colName = "BIN" & Format(i, "00")
                    objBIN = vBINs.Find(Function(bin) bin.Name = colName)
                    If objBIN Is Nothing Then
                        iRow("BIN" & Format(i, "00")) = ""
                    Else
                        iRow("BIN" & Format(i, "00")) = objBIN.Description
                    End If
                Next
            End With
            dtEPRO.Rows.Add(iRow)
            Return dtEPRO
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    'Public Class EPRO
    '    Inherits testLog

    '    Public Sub New(ByVal fileName As String)
    '        MyBase.New(fileName)
    '    End Sub
    'End Class


End Class
