Imports System.IO
Imports System.Text.RegularExpressions

Public Class clsTMT
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
    Dim vMostFailBin As Integer = 0
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
    Dim vNextSerial As String = ""

    Dim vCompleted As Boolean
    Dim vMessage As String

    Dim vSiteTotal As Integer
    Dim vLotResult_line_start As Boolean
    Dim vLotResults As New List(Of Result)

    Dim vSwBin_line_start As Boolean
    Dim vSwBinSite_line_start As Boolean
    Dim vSwBins As New List(Of Result)
    Dim vHwBins As New List(Of Result)



    Public Sub New(ByVal fileName As String)
        vFileName = fileName
        vKeyName = My.Computer.FileSystem.GetFileInfo(vFileName).Name
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
        'vSwBin_line_start
        If line_data.Contains("SW Bins") Then
            vSwBin_line_start = True
            Exit Sub
        End If
        If line_data.Contains("SW Site") Then
            Dim strLotResult As String()
            strLotResult = line_data.Split(" ")
            Dim strSiteFindResult As String()
            strSiteFindResult = Array.FindAll(strLotResult, Function(x) x Like "Site")
            vSiteTotal = strSiteFindResult.Length
            vSwBin_line_start = False
            vSwBinSite_line_start = True
            Exit Sub
        End If

        If vSwBin_line_start Then
            Dim strSwBinArry As String()
            Dim ixYield As Integer = 0
            Dim i As Integer = 0
            Dim ixBinYield As Integer = 0

            strSwBinArry = line_data.Split(" ")
            ixBinYield = Array.FindIndex(strSwBinArry, Function(x) x.Contains("%"))
            Dim strPercentCount As String() = Array.FindAll(strSwBinArry, Function(x) x.Contains("%"))

            'if 2 means there is HW bin detail exist
            Dim strHwBinName As String = ""
            Dim strHwBinCount As String = ""
            Dim strHwBinYield As String = ""

            If strPercentCount.Length = 2 Then
                For y = ixBinYield + 1 To strSwBinArry.Length - 1
                    If strSwBinArry(y) <> "" Then
                        If strHwBinName = "" Then
                            strHwBinName = "HW_Bin" & strSwBinArry(y)
                            Continue For
                        End If
                        If IsNumeric(strSwBinArry(y)) Then
                            If strHwBinCount = "" Then
                                strHwBinCount = strSwBinArry(y)
                                Continue For
                            End If
                            If strHwBinYield = "" Then
                                strHwBinYield = strSwBinArry(y)
                                Continue For
                            End If
                        End If
                    End If
                Next
                'Add to Result
                Dim objHWResult As New Result With {.Name = strHwBinName,
                                            .Description = "",
                                            .Yield = Val(strHwBinYield),
                                            .Count = Val(strHwBinCount)}
                vHwBins.Add(objHWResult)
            End If


            Dim strBinNumber As String = Regex.Replace(strSwBinArry(0), "[^A-Za-z0-9\-/]", "")
            Dim strBinName As String = "SW_Bin" & strBinNumber
            Dim strBinYield As String = strSwBinArry(ixBinYield - 1)

            Dim strBinDescriotion As String = ""
            For i = 3 To ixBinYield
                strBinDescriotion = strBinDescriotion & strSwBinArry(i)
                If strSwBinArry(i + 1) = "" Then
                    Exit For
                End If
            Next

            'Find Number of Bin
            Dim strBinCount As String = "0"
            For x = i + 1 To ixBinYield - 2
                If strSwBinArry(x) <> "" Then
                    If IsNumeric(strSwBinArry(x)) Then
                        strBinCount = strSwBinArry(x)
                        Exit For
                    End If

                End If
            Next
            Dim objResult As New Result With {.Name = strBinName,
                                            .Description = strBinDescriotion,
                                            .Yield = Val(strBinYield),
                                            .Count = Val(strBinCount)}

            vSwBins.Add(objResult)


            Exit Sub
            'vLotResults.Add()
        End If

        If vSwBinSite_line_start Then
            Dim strSwBinArry As String()
            Dim ixYield As Integer = 0
            Dim i As Integer = 0
            Dim ixBinYield As Integer = 0

            strSwBinArry = line_data.Split(" ")

            Dim strBinDescriotion As String = ""
            Dim strBinNumber As String = ""
            Dim strBinName As String = ""
            Dim strBinSiteCount As String = ""
            Dim ixInit As Integer = 1

            strBinNumber = Regex.Replace(strSwBinArry(0), "[^A-Za-z0-9\-/]", "")
            strBinName = "SW_Bin" & strBinNumber
            Dim objSwbin As Result = vSwBins.Find(Function(x) x.Name = strBinName)
            Dim objSites As New List(Of SITE)
            For i = 1 To vSiteTotal
                ixBinYield = Array.IndexOf(strSwBinArry, "%", ixBinYield + 1)

                For x = ixInit To ixBinYield - 2
                    If strSwBinArry(x) <> "" Then
                        If IsNumeric(strSwBinArry(x)) Then
                            strBinSiteCount = strSwBinArry(x)
                            Exit For
                        End If
                    End If
                    ixInit = ixBinYield + 1
                Next
                'Create Site data
                Dim objSite As New SITE With {.Name = strBinName & "_Site" & i.ToString,
                                               .Count = Val(strBinSiteCount)}
                objSites.Add(objSite)

            Next
            'Add Site to Bin
            objSwbin.Sites = objSites

        End If

        If line_data.Contains("Lot ID") Then
            'Assy,Lot and Seq are in file name
            Dim ixColon As Integer = 0
            Dim strArray As String() = line_data.Split(" ")
            ixColon = Array.FindIndex(strArray, Function(x) x.Contains(":"))

            Dim strAssyData As String = ""
            'strAssyData = IIf(strArray(ixColon + 1) <> "", strArray(ixColon + 1), strArray(ixColon + 2))
            For i = ixColon + 1 To strArray.Length
                If strArray(i) <> "" Then
                    strAssyData = strArray(i)
                    Exit For
                End If
            Next
            vAssyNumber = strAssyData.Split("_")(0)
            'vTestType = strFileName(0).Substring(0, 2) 'First 2 digits
            If strAssyData.Split("_").Length >= 3 Then
                vLotNumber = strAssyData.Split("_")(1)
                vSeqNumber = strAssyData.Split("_")(strAssyData.Split("_").Length - 1).Substring(0, 2)
            End If
            Exit Sub
        End If


        If line_data.Contains("Computer") Then
            Dim ixTester As Integer = 0
            Dim strArray As String() = line_data.Split(" ")
            ixTester = Array.FindIndex(strArray, Function(x) x.Contains(":"))
            vTester = strArray(ixTester + 1)
            Exit Sub
        End If

        If line_data.Contains("Handler") Then
            Dim ixHandler As Integer = 0
            Dim strArray As String() = line_data.Split(" ")
            ixHandler = Array.FindIndex(strArray, Function(x) x.Contains(":"))
            vHandlerId = strArray(ixHandler + 1)
            Exit Sub
        End If

        If line_data.Contains("Operator") Then
            Dim ixOperator As Integer = 0
            Dim strArray As String() = line_data.Split(" ")
            ixOperator = Array.FindIndex(strArray, Function(x) x.Contains(":"))
            vOperatorName = strArray(ixOperator + 1)

            vMostFailBin = Val(strArray(strArray.Length - 1))
            Exit Sub
        End If

        If line_data.Contains("Test Program") Then
            Dim ixProgram As Integer = 0
            Dim strArray As String() = line_data.Split(" ")
            ixProgram = Array.FindIndex(strArray, Function(x) x.Contains(":"))
            vProgramFileName = strArray(ixProgram + 1)

            'Get Total Tested
            vTested = Val(strArray(strArray.Length - 1))
            Exit Sub
        End If

        If line_data.Contains("Autocorrelation") Then
            Dim ixAuto As Integer = 0
            Dim strArray As String() = line_data.Split(" ")

            'Get Total Passed
            vNextSerial = Val(strArray(strArray.Length - 1))

            Exit Sub
        End If

        If line_data.Contains("Version") Then
            Dim ixVersion As Integer = 0
            Dim strArray As String() = line_data.Split(" ")
            ixVersion = Array.FindIndex(strArray, Function(x) x.Contains(":"))
            vProgramRevision = strArray(ixVersion + 1)

            'Get Total Passed
            vPassed = Val(strArray(strArray.Length - 1))
            vFailed = vTested - vPassed
            Exit Sub
        End If

        'Get Start and Stop date
        If line_data.Contains("/") Then
            Dim strArray As String() = line_data.Split("/")
            If strArray.Length <> 2 Then
                Exit Sub
            End If
            If Not IsDate(strArray(0)) Then
                Exit Sub
            End If

            vStartDate = strArray(0)
            vStopDate = strArray(1)

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
        Public Property Description As String
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


    Function convert_to_datatable(Optional hwbin_col_count As Integer = 8,
                                  Optional swbin_col_count As Integer = 32
                                 ) As DataTable
        Try
            Dim dtEPRO As New DataTable
            Dim i As Integer
            'Make Table column
            With dtEPRO
                .Columns.Add("assy", GetType(String))
                .Columns.Add("lot", GetType(String))
                .Columns.Add("seq", GetType(String))
                .Columns.Add("operator", GetType(String))
                .Columns.Add("tester", GetType(String))
                .Columns.Add("handler", GetType(String))
                .Columns.Add("program_name", GetType(String))
                .Columns.Add("start_time", GetType(String))
                .Columns.Add("end_time", GetType(String))
                .Columns.Add("tested", GetType(String))
                .Columns.Add("passed", GetType(String))
                .Columns.Add("failed", GetType(String))
                .Columns.Add("yield", GetType(String))
                .Columns.Add("most_fail_bin", GetType(String))
                .Columns.Add("next_serial", GetType(String))


                'Add HW column
                For i = 0 To hwbin_col_count
                    .Columns.Add("HW_Bin" & Str(i).Trim, GetType(String))
                Next

                'Add SW column
                For i = 1 To swbin_col_count
                    .Columns.Add("SW_Bin" & Str(i).Trim, GetType(String))
                Next

                'Add SW item column
                For i = 1 To swbin_col_count
                    .Columns.Add("Bin" & Str(i).Trim & "_Item", GetType(String))
                Next

            End With
            'Fill data
            Dim iRow As DataRow = dtEPRO.NewRow()
            With iRow
                iRow("assy") = vAssyNumber
                iRow("lot") = vLotNumber
                iRow("seq") = vSeqNumber
                iRow("operator") = vOperatorName
                iRow("tester") = vTester
                iRow("handler") = vHandlerId
                iRow("program_name") = vProgramFileName
                iRow("start_time") = vStartDate
                iRow("end_time") = vStopDate
                iRow("tested") = vTested
                iRow("passed") = vPassed
                iRow("failed") = vTested - vPassed
                iRow("yield") = Yield
                iRow("most_fail_bin") = vMostFailBin
                iRow("next_serial") = vNextSerial

                Dim colName As String = ""
                'Add HW column
                For Each hw In vHwBins
                    colName = hw.Name
                    iRow(colName) = hw.Count
                Next
                'Add SW column
                For Each sw In vSwBins
                    colName = sw.Name
                    iRow(colName) = sw.Count
                Next

                'Add SW item column
                For i = 1 To swbin_col_count
                    colName = "Bin" & Str(i).Trim & "_Item"
                    Dim objSw As Result = vSwBins.Find(Function(sw) sw.Name = "SW_Bin" & Str(i).Trim)
                    If Not objSw Is Nothing Then
                        iRow(colName) = objSw.Description
                    Else
                        iRow(colName) = ""
                    End If
                Next


                ''fill SW Bin Site column (Automatic create column)
                For Each sw In vSwBins
                    If sw.Sites Is Nothing Then
                        Continue For

                    End If
                    For Each st In sw.Sites
                        colName = st.Name
                        'Verify Column name exist, if No create new col.
                        If Not dtEPRO.Columns.Contains(colName) Then
                            dtEPRO.Columns.Add(colName, GetType(String))
                        End If
                        iRow(colName) = st.Count
                    Next

                Next




                ''Dim selectedValue As clsEPRO
                ''selectedValue = objEPROs.Find(Function(EPRO) EPRO.keyName = vShortFileName)
                'Dim colName As String = ""
                'Dim colTestName As String = ""
                'Dim colPassName As String = ""
                'Dim colFailName As String = ""
                'Dim objTested As Result
                'Dim objPassed As Result
                'Dim objFailed As Result
                ''Dim objIB As IB
                ''Dim objDB As DB
                ''Dim objBIN As BIN
                ''Fill IB column
                'objTested = vLotResults.Find(Function(ib) ib.Name = "TESTED")
                'objPassed = vLotResults.Find(Function(ib) ib.Name = "PASSED")
                'objFailed = vLotResults.Find(Function(ib) ib.Name = "FAILED")

                'For Each objSite In objTested.Sites
                '    colName = "Test_" & objSite.Name
                '    iRow(colName) = objSite.Count
                'Next
                'For Each objSite In objPassed.Sites
                '    colName = "Pass_" & objSite.Name
                '    iRow(colName) = objSite.Count
                'Next

                'If Not objFailed Is Nothing Then
                '    For Each objSite In objFailed.Sites
                '        colName = "Fail_" & objSite.Name
                '        iRow(colName) = objSite.Count
                '    Next
                'End If



                ''fill BIN column
                'For Each objBin In vIBs
                '    colName = objBin.Name
                '    iRow(colName) = objBin.Count
                'Next

                ''fill DB column
                'For Each objBin In vDBs
                '    colName = objBin.Name
                '    'Verify Column name exist, if No create new col.
                '    If Not dtEPRO.Columns.Contains(colName) Then
                '        dtEPRO.Columns.Add(colName, GetType(String))
                '    End If
                '    iRow(colName) = objBin.Count
                'Next

            End With
            dtEPRO.Rows.Add(iRow)
            Return dtEPRO
        Catch ex As Exception
            Return Nothing
        End Try

    End Function
End Class
