Imports System.IO
Imports System.Text.RegularExpressions

Public Class clsMAV
    Dim vKeyName As String = ""
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
    Dim vSystemId As String = ""
    Dim vTested As Integer = 0
    Dim vPassed As Integer = 0
    Dim vFailed As Integer = 0
    Dim vIBs As New List(Of Result)
    Dim vDBs As New List(Of Result)

    Dim vStartDate As String = ""
    Dim vStopDate As String = ""
    Dim vTestFlow As String = ""
    Dim vTestType As String = ""
    Dim vDeviceName As String = ""
    Dim vLotSize As String = ""
    Dim vTestCount As String = ""
    Dim vProgramName As String = ""
    Dim vProgramRevision As String = ""

    Dim vCompleted As Boolean
    Dim vMessage As String

    Dim vSiteTotal As Integer
    Dim vLotResult_line_start As Boolean
    Dim vLotResults As New List(Of Result)


    Public Sub New(ByVal fileName As String)
        vFileName = fileName
        vKeyName = My.Computer.FileSystem.GetFileInfo(vFileName).Name
        'Assy,Lot and Seq are in file name
        Dim strFileName As String() = vKeyName.Split("_")
        vAssyNumber = strFileName(0)
        vTestType = strFileName(0).Substring(0, 2) 'First 2 digits
        If strFileName.Length >= 3 Then
            vLotNumber = strFileName(1)
            vSeqNumber = strFileName(2).Replace(".txt", "")
        End If
        process_file()
    End Sub

    Public ReadOnly Property TestFlow() As String
        Get
            Return vTestFlow
        End Get
    End Property
    Public ReadOnly Property TestType() As String
        Get
            Return vTestType
        End Get
    End Property
    Public ReadOnly Property Devicename() As String
        Get
            Return vDeviceName
        End Get
    End Property
    Public ReadOnly Property LotSise() As String
        Get
            Return vLotSize
        End Get
    End Property
    Public ReadOnly Property TestCount() As String
        Get
            Return vTestCount
        End Get
    End Property
    Public ReadOnly Property ProgramName() As String
        Get
            Return vProgramName
        End Get
    End Property
    Public ReadOnly Property ProgramRevision() As String
        Get
            Return vProgramRevision
        End Get
    End Property


    Public ReadOnly Property keyName() As String
        Get
            Return vKeyName
        End Get
    End Property

    Public ReadOnly Property Results() As List(Of Result)
        Get
            Return vLotResults
        End Get
    End Property

    Public ReadOnly Property DataTable() As DataTable
        Get
            Return convert_to_datatable()
        End Get
    End Property
    Public ReadOnly Property IBs() As List(Of Result)
        Get
            Return vIBs
        End Get
    End Property
    Public ReadOnly Property DBs() As List(Of Result)
        Get
            Return vDBs
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
    Public ReadOnly Property StartDate() As String
        Get
            Return vStartDate
        End Get
    End Property
    Public ReadOnly Property StopDate() As String
        Get
            Return vStopDate
        End Get
    End Property
    Public ReadOnly Property SystemId() As String
        Get
            Return vSystemId
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
        'vLotResult_line_start
        If line_data.Contains("%LOT RESULTS") Then
            vLotResult_line_start = True
            Dim strLotResult As String()
            strLotResult = line_data.Split(" ")
            Dim strSiteFindResult As String()
            strSiteFindResult = Array.FindAll(strLotResult, Function(x) x Like "SITE*")
            vSiteTotal = strSiteFindResult.Length
            Exit Sub
        End If
        If line_data.Contains("%INTERFACE BINS") Then
            vLotResult_line_start = False
        End If

        If vLotResult_line_start Then
            Dim strLotResult As String()
            Dim ixYield As Integer = 0
            Dim i As Integer = 0
            strLotResult = line_data.Split(" ")
            'Unit Tested
            If strLotResult(1) = "TESTED" Then
                For i = 4 To 17 'To support 4,3,2,1 digit
                    If strLotResult(i) <> "" Then
                        vTested = Val(strLotResult(i))
                        ixYield = Array.FindIndex(strLotResult, Function(x) x.Contains("%"))
                        If ixYield = -1 Then
                            ixYield = i
                        End If
                        Exit For
                    End If
                Next
                Dim iStart As Integer = 0
                Dim iSite As Integer = 1
                Dim objSites As New List(Of SITE)
                For iStart = ixYield + 1 To strLotResult.Length - 1
                    If strLotResult(iStart) <> "" Then
                        Dim zz As String = "SITE" & iSite.ToString
                        Dim objSite As New SITE With {.Name = zz, .Count = Val(strLotResult(iStart))}
                        objSites.Add(objSite)
                        iSite = iSite + 1
                    End If
                Next
                Dim objResult As New Result With {.Name = "TESTED",
                                                .Count = vTested,
                                                .Sites = objSites}

                vLotResults.Add(objResult)

            End If

            If strLotResult(1) = "PASSED" Then
                For i = 4 To 17 'To support 4,3,2,1 digit
                    If strLotResult(i) <> "" Then
                        vPassed = Val(strLotResult(i))
                        ixYield = Array.FindIndex(strLotResult, Function(x) x.Contains("%"))
                        If ixYield = -1 Then
                            ixYield = i
                        End If
                        Exit For
                    End If
                Next
                Dim iStart As Integer = 0
                Dim iSite As Integer = 1
                Dim objSites As New List(Of SITE)
                For iStart = ixYield + 1 To strLotResult.Length - 1
                    If strLotResult(iStart) <> "" Then
                        Dim zz As String = "SITE" & iSite.ToString
                        Dim objSite As New SITE With {.Name = zz, .Count = Val(strLotResult(iStart))}
                        objSites.Add(objSite)
                        iSite = iSite + 1
                    End If
                Next
                Dim objResult As New Result With {.Name = "PASSED",
                                                .Count = vTested,
                                                .Sites = objSites}

                vLotResults.Add(objResult)
            End If

            If strLotResult(1) = "FAILED" Then
                For i = 4 To 17 'To support 4,3,2,1 digit
                    If strLotResult(i) <> "" Then
                        vFailed = Val(strLotResult(i))
                        ixYield = Array.FindIndex(strLotResult, Function(x) x.Contains("%"))
                        If ixYield = -1 Then
                            ixYield = i
                        End If
                        Exit For
                    End If
                Next
                Dim iStart As Integer = 0
                Dim iSite As Integer = 1
                Dim objSites As New List(Of SITE)
                For iStart = ixYield + 1 To strLotResult.Length - 1
                    If strLotResult(iStart) <> "" Then
                        Dim zz As String = "SITE" & iSite.ToString
                        Dim objSite As New SITE With {.Name = zz, .Count = Val(strLotResult(iStart))}
                        objSites.Add(objSite)
                        iSite = iSite + 1
                    End If
                Next
                Dim objResult As New Result With {.Name = "FAILED",
                                                .Count = vTested,
                                                .Sites = objSites}

                vLotResults.Add(objResult)
            End If

            If strLotResult(1) = "FAIL" Then
                Dim vParName As String = strLotResult(0) & " " & strLotResult(1)
                For i = 4 To 17 'To support 4,3,2,1 digit
                    If strLotResult(i) <> "" Then
                        vFailed = Val(strLotResult(i))
                        ixYield = Array.FindIndex(strLotResult, Function(x) x.Contains("%"))
                        If ixYield = -1 Then
                            ixYield = i
                        End If
                        Exit For
                    End If
                Next
                Dim iStart As Integer = 0
                Dim iSite As Integer = 1
                Dim objSites As New List(Of SITE)
                For iStart = ixYield + 1 To strLotResult.Length - 1
                    If strLotResult(iStart) <> "" Then
                        Dim zz As String = "SITE" & iSite.ToString
                        Dim objSite As New SITE With {.Name = zz, .Count = Val(strLotResult(iStart))}
                        objSites.Add(objSite)
                        iSite = iSite + 1
                    End If
                Next
                Dim objResult As New Result With {.Name = vParName,
                                                .Count = vFailed,
                                                .Sites = objSites}

                vLotResults.Add(objResult)
            End If

            'IB or DB
            If strLotResult(0) = "Bin" Or strLotResult(0) = "IB" Or strLotResult(0) = "DB" Then
                Dim vParName As String = IIf(strLotResult(0) = "IB" Or strLotResult(0) = "Bin", "BIN", "DB") & strLotResult(1) & strLotResult(2) & strLotResult(3)

                'If strLotResult(0) = "Bin" Then
                '    vParName = "BIN" & strLotResult(1) & strLotResult(2) & strLotResult(3)
                'End If



                For i = 4 To 17 'To support 4,3,2,1 digit
                    If strLotResult(i) <> "" Then
                        vFailed = Val(strLotResult(i))
                        ixYield = Array.FindIndex(strLotResult, Function(x) x.Contains("%"))
                        If ixYield = -1 Then
                            ixYield = i
                        End If
                        Exit For
                    End If
                Next
                Dim iStart As Integer = 0
                Dim iSite As Integer = 1
                Dim objSites As New List(Of SITE)
                For iStart = ixYield + 1 To strLotResult.Length - 1
                    If strLotResult(iStart) <> "" Then
                        Dim zz As String = "SITE" & iSite.ToString
                        Dim objSite As New SITE With {.Name = zz, .Count = Val(strLotResult(iStart))}
                        objSites.Add(objSite)
                        iSite = iSite + 1
                    End If
                Next
                Dim objResult As New Result With {.Name = vParName,
                                                .Count = vFailed,
                                                .Sites = objSites}

                'vLotResults.Add(objResult)
                If strLotResult(0) = "IB" Or strLotResult(0) = "Bin" Then
                    vIBs.Add(objResult)
                Else
                    vDBs.Add(objResult)
                End If
            End If


            Exit Sub
            'vLotResults.Add()
        End If

        If line_data.Contains("Lot Number:") Then
            vLotNumber = line_data.Replace("Lot Number:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Start Time:") Then
            vStartDate = line_data.Replace("Start Time:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("End Time:") Then
            vStopDate = line_data.Replace("End Time:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Test Flow:") Then
            vTestFlow = line_data.Replace("Test Flow:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Temperature:") Then
            vTemperature = line_data.Replace("Temperature:", "").Trim
            Dim cleanString As String = Regex.Replace(vTemperature, "[^A-Za-z0-9\-/]", "")
            vTemperature = cleanString
            Exit Sub
        End If

        If line_data.Contains("Device Name:") Then
            vDeviceName = line_data.Replace("Device Name:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Tester ID:") Then
            vTester = line_data.Replace("Tester ID:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Handler ID:") Then
            vHandlerId = line_data.Replace("Handler ID:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Operator ID:") Then
            vOperatorName = line_data.Replace("Operator ID:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Program Name:") Then
            vProgramFileName = line_data.Replace("Program Name:", "").Trim
            Exit Sub
        End If

        If line_data.Contains("Test Count:") Then
            vTestCount = line_data.Replace("Test Count:", "").Trim
            Exit Sub
        End If




    End Sub

    Public Class IB
        Public Property Name As String
        Public Property Count As Integer
        Public Property Yield As Decimal
        Public Property Sites As List(Of SITE)
    End Class

    Public Class DB
        Public Property Name As String
        Public Property Count As Integer
        Public Property Yield As Decimal
        Public Property Sites As List(Of SITE)
    End Class

    Public Class Result
        Public Property Name As String
        Public Property Count As Integer
        Public Property Yield As Decimal
        Public Property Sites As List(Of SITE)
    End Class

    Public Class SITE
        Public Property Name As String
        Public Property Count As Integer
    End Class

    Public Class BIN
        Public Property Name As String
        Public Property Description As String
    End Class


    Function convert_to_datatable(Optional Site_col_count As Integer = 8,
                                 Optional Bin_col_count As Integer = 8,
                                  Optional DB_col_count As Integer = 50
                                  ) As DataTable
        Try
            Dim dtEPRO As New DataTable
            Dim i As Integer
            'Make Table column
            With dtEPRO
                .Columns.Add("assy", GetType(String))
                .Columns.Add("lot", GetType(String))
                .Columns.Add("seq", GetType(String))
                .Columns.Add("start_time", GetType(String))
                .Columns.Add("end_time", GetType(String))
                .Columns.Add("test_flow", GetType(String))
                .Columns.Add("test_type", GetType(String))
                .Columns.Add("operator", GetType(String))
                .Columns.Add("device_name", GetType(String))
                .Columns.Add("lot_size", GetType(String))
                .Columns.Add("test_count", GetType(String))
                .Columns.Add("temperature", GetType(String))
                .Columns.Add("program_name", GetType(String))
                .Columns.Add("program_rev", GetType(String))
                .Columns.Add("tester", GetType(String))
                .Columns.Add("handler", GetType(String))
                .Columns.Add("tested", GetType(String))
                .Columns.Add("passed", GetType(String))
                .Columns.Add("failed", GetType(String))
                .Columns.Add("yield", GetType(String))
                'Add IB column
                For i = 1 To Site_col_count
                    .Columns.Add("Test_SITE" & Str(i).Trim, GetType(String))
                    .Columns.Add("Pass_SITE" & Str(i).Trim, GetType(String))
                    .Columns.Add("Fail_SITE" & Str(i).Trim, GetType(String))
                Next
                'Add DB column
                For i = 1 To Bin_col_count
                    .Columns.Add("BIN" & Str(i).Trim, GetType(String))
                Next
                'Add BIN column
                For i = 1 To DB_col_count
                    .Columns.Add("DB" & Str(i).Trim, GetType(String))
                Next
            End With
            'Fill data
            Dim iRow As DataRow = dtEPRO.NewRow()
            With iRow
                iRow("assy") = vAssyNumber
                iRow("lot") = vLotNumber
                iRow("seq") = vSeqNumber
                iRow("start_time") = vStartDate
                iRow("end_time") = vStopDate
                iRow("test_flow") = vTestFlow
                iRow("test_type") = vTestType
                iRow("device_name") = vDeviceName
                iRow("operator") = vOperatorName
                iRow("lot_size") = vLotSize
                iRow("test_count") = vTestCount
                iRow("temperature") = vTemperature
                iRow("program_name") = vProgramName
                iRow("program_rev") = vProgramRevision
                iRow("tester") = vTester
                iRow("handler") = vHandlerId
                iRow("tested") = vTested
                iRow("passed") = vPassed
                iRow("failed") = vTested - vPassed
                iRow("yield") = Yield

                'Dim selectedValue As clsEPRO
                'selectedValue = objEPROs.Find(Function(EPRO) EPRO.keyName = vShortFileName)
                Dim colName As String = ""
                Dim colTestName As String = ""
                Dim colPassName As String = ""
                Dim colFailName As String = ""
                Dim objTested As Result
                Dim objPassed As Result
                Dim objFailed As Result
                'Dim objIB As IB
                'Dim objDB As DB
                'Dim objBIN As BIN
                'Fill IB column
                objTested = vLotResults.Find(Function(ib) ib.Name = "TESTED")
                objPassed = vLotResults.Find(Function(ib) ib.Name = "PASSED")
                objFailed = vLotResults.Find(Function(ib) ib.Name = "FAILED")

                For Each objSite In objTested.Sites
                    colName = "Test_" & objSite.Name
                    iRow(colName) = objSite.Count
                Next
                For Each objSite In objPassed.Sites
                    colName = "Pass_" & objSite.Name
                    iRow(colName) = objSite.Count
                Next

                If Not objFailed Is Nothing Then
                    For Each objSite In objFailed.Sites
                        colName = "Fail_" & objSite.Name
                        iRow(colName) = objSite.Count
                    Next
                End If



                'fill BIN column
                For Each objBin In vIBs
                    colName = objBin.Name
                    iRow(colName) = objBin.Count
                Next

                'fill DB column
                For Each objBin In vDBs
                    colName = objBin.Name
                    'Verify Column name exist, if No create new col.
                    If Not dtEPRO.Columns.Contains(colName) Then
                        dtEPRO.Columns.Add(colName, GetType(String))
                    End If
                    iRow(colName) = objBin.Count
                Next

            End With
            dtEPRO.Rows.Add(iRow)
            Return dtEPRO
        Catch ex As Exception
            Return Nothing
        End Try

    End Function
End Class
