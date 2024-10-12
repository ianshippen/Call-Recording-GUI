Imports System.Data.SqlClient

Module HandleRequestParameterisedSQL
    Dim myFieldNames() As String = {"time", "date", "internalNumber", "externalNumber", "direction", "duration", "callid", "action"}

    Public Function NewHandleRequestParameterisedSQL(ByRef p_dataSet As DataSet, ByVal resultOffset As Integer, ByRef p_startDate As String, ByRef p_endDate As String, ByRef p_extension As String, ByRef p_destination As String, ByVal p_callId As Integer, ByRef p_callType As String, ByRef p_startTimeText As String, ByRef p_endTimeText As String, ByRef p_securityData As String, ByRef p_loginExtension As String, ByVal p_useLegacyTable As Boolean) As Integer
        ' Get the total number of records matched
        Dim myTable As New DataTable
        Dim recordingsPath As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.recordingsPath))
        Dim maxRecordsToReturn As Integer = CInt(callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.maxRecordsToReturn)))
        Dim timeStampName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.timeStampName))
        Dim xrefDestination As Boolean = callRecordingInterfaceConfigDictionary.GetBooleanItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.xrefDestination))
        Dim xrefTransferredCalls As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.xrefTransferredCalls))
        Dim cdrTableName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.cdrTableName))
        Dim callRecordingsTableName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.callRecordingTable))
        Dim archiveEnabled As Boolean = callRecordingInterfaceConfigDictionary.GetBooleanItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.archiveEnabled))
        Dim recordsMatched As Integer
        Dim loadedOK As Boolean = False
        Dim myDataList As New List(Of String)
        Dim myInData As String = ""
        Dim useExternalNumberFunction As Boolean = True
        Dim externalNumberText As String = "ExternalNumber"
        Dim myFix As String = " COLLATE DATABASE_DEFAULT"
        Dim infrontFix As Boolean = True    ' To allow transferred calls to show up in supervisor login mode
        Dim swyxDatabaseName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.swyxDatabaseName))
        Dim mainNumberToReplace As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.mainNumberToReplace))
        Dim locateXferFromOutgoingCall As Boolean = callRecordingInterfaceConfigDictionary.GetBooleanItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.locateXferFromOutgoingCall))
        Dim myParms As New ParameterisedSQLParmClass

        If useExternalNumberFunction Then externalNumberText = "dbo.MapExternalNumber(" & externalNumberText & ")"

        If p_useLegacyTable Then callRecordingsTableName = "LegacyCallRecordingsTable"

        ' Parse security data
        ParseSecurityData(p_securityData, myDataList)

        ' *** SQL Query A ***
        If myDataList.Count > 0 Then
            Dim myName As String = p_extension
            Dim oldDatabaseName As String = callRecordingInterfaceConfigDictionary.GetItem("databaseName")

            myInData = "("

            For i = 0 To myDataList.Count - 1
                If i > 0 Then myInData &= ", "

                myInData &= WrapInSingleQuotes(SingleQuoteCheck(myDataList(i)))
            Next

            myInData &= ")"

            callRecordingInterfaceConfigDictionary.SetItem("databaseName", swyxDatabaseName)

            If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), "select name from users as a left join internalnumbers as b on a.userid = b.userid where b.number in " & myInData, myTable) Then
                If myTable.Rows.Count > 0 Then
                    If myTable.Columns.Count > 0 Then
                        myInData = myInData.Substring(0, myInData.Length - 1)

                        For i = 0 To myTable.Rows.Count - 1
                            If Not myTable.Rows(i).Item(0) Is DBNull.Value Then myInData &= ", " & WrapInSingleQuotes(SingleQuoteCheck(myTable.Rows(i).Item(0).ToString))
                        Next

                        myInData &= ")"
                    End If
                End If
            End If

            callRecordingInterfaceConfigDictionary.SetItem("databaseName", oldDatabaseName)
            myTable.Rows.Clear()
            myTable.Columns.Clear()
        End If
        ' *** End of SQL Query A ***

        ' Get the total number of available records
        ' *** SQL Query B ***
        Dim mySQLStatement As New SQLStatementClass

        mySQLStatement.SetPrimaryTable(callRecordingsTableName)
        mySQLStatement.AddSelectString("COUNT(*)", "")

        ' Should be able to handle Blind SQL Injection as is due to local processing of date fields
        AddDateCondition(p_startDate, p_endDate, mySQLStatement, timeStampName)

        If xrefDestination Then mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, cdrTableName & " AS B", "A.CallId", "B.CallId and convert(date, timestamp) = convert(date, startTime)")

        ' Change for CDR with CallId = 0
        'If xrefTransferredCalls Then mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, cdrTableName & " AS C", "B.TransferredToCallId", "CallId")
        If xrefTransferredCalls Then mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, cdrTableName & " AS C", "B.TransferredToCallId", "C.CallId and C.CallId > 0 and convert(date, b.startTime) = convert(date, c.startTime)")

        ' If we are mapping the main output number to individual DDIs, join to the Swyx Users and Internal Numbers tables to map the caller's name to their extension
        If mainNumberToReplace <> "" Then
            Dim usersTable As String = "[" & swyxDatabaseName & "].[dbo].[Users]"
            Dim internalNumbersTable As String = "[" & swyxDatabaseName & "].[dbo].[InternalNumbers]"

            mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, usersTable & " AS D", "B.OriginationName", "Name")
            mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, internalNumbersTable & " AS E", "D.UserId", "E.UserId")
        End If

        ' Are we filtering by extension ?
        If p_extension.Length > 0 Then
            Dim myExtension As String = p_extension
            Dim myOperator As String = "="
            Dim myInternalNumberMatchString As String = ""

            If myExtension.Contains("*") Then
                myOperator = "LIKE"
                myExtension = myExtension.Replace("*", "%")
            End If

            ' Possible Blind SQL Injection here with myExtension ..
            ' --> myInternalNumberMatchString = "(InternalNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & ")"
            myInternalNumberMatchString = "(InternalNumber " & myOperator & " " & myParms.CreateStringParm(myExtension) & ")"

            If mainNumberToReplace <> "" Then
                ' --> myInternalNumberMatchString = "(" & myInternalNumberMatchString & " OR (InternalNumber = " & WrapInSingleQuotes(mainNumberToReplace) & " AND Number " & myOperator & " " & WrapInSingleQuotes(myExtension) & "))"
                myInternalNumberMatchString = "(" & myInternalNumberMatchString & " OR (InternalNumber = " & WrapInSingleQuotes(mainNumberToReplace) & " AND Number " & myOperator & " " & myParms.CreateStringParm(myExtension) & "))"
            End If

            If xrefDestination Then
                Dim myName As String = p_extension
                Dim oldDatabaseName As String = callRecordingInterfaceConfigDictionary.GetItem("databaseName")
                Dim mySql As String = "select name, c.number as [publicNumber] from users as a"
                Dim publicNumbers As New List(Of String)

                mySql &= " left join internalnumbers as b on a.userid = b.userid"
                mySql &= " left join publicnumbers as c on b.internalNumberId = c.internalNumberId"
                mySql &= " where b.number = "

                callRecordingInterfaceConfigDictionary.SetItem("databaseName", swyxDatabaseName)

                ' *** SQL Query C ***
                mySql &= "@p_extension"

                Dim mySqlConnection1 As New SqlConnection(CreateConnectionString(callRecordingInterfaceConfigDictionary, True))
                Dim mySqlCommand1 As New SqlCommand(mySql, mySqlConnection1)
                Dim myReader1 As SqlDataReader = Nothing

                mySqlCommand1.Parameters.AddWithValue("@p_extension", p_extension)
                mySqlCommand1.Connection.Open()
                myReader1 = mySqlCommand1.ExecuteReader()
                myTable.Load(myReader1)
                mySqlCommand1.Connection.Close()
                ' *** End Of SQL Query C ***

                ' Get the username and all public numbers for this internal number
                If myTable.Rows.Count > 0 Then
                    If myTable.Columns.Count > 0 Then
                        If Not myTable.Rows(0).Item(0) Is DBNull.Value Then myName = myTable.Rows(0).Item(0).ToString

                        For i = 0 To myTable.Rows.Count - 1
                            With myTable.Rows(i)
                                If Not .Item(1) Is DBNull.Value Then
                                    If Not .Item(1) = "" Then publicNumbers.Add(.Item(1))
                                End If
                            End With
                        Next
                    End If
                End If

                callRecordingInterfaceConfigDictionary.SetItem("databaseName", oldDatabaseName)
                myTable = New DataTable

                Dim publicNumberString As String = ""

                If publicNumbers.Count > 0 Then
                    If publicNumbers.Count = 1 Then
                        publicNumberString = " = " & WrapInSingleQuotes(publicNumbers(0))
                    Else
                        For i = 0 To publicNumbers.Count - 1
                            If i > 0 Then publicNumberString &= ", "

                            publicNumberString &= WrapInSingleQuotes(publicNumbers(i))
                        Next

                        publicNumberString = " IN (" & publicNumberString & ")"
                    End If
                End If

                Dim publicNumberString1 As String = ""
                Dim publicNumberString2 As String = ""
                Dim publicNumberString3 As String = ""

                If publicNumberString <> "" Then
                    publicNumberString1 = " OR InternalNumber " & publicNumberString
                    publicNumberString2 = " OR B.DestinationNumber " & publicNumberString
                    publicNumberString3 = " OR C.DestinationNumber " & publicNumberString
                End If

                If xrefTransferredCalls Then
                    ' --> mySQLStatement.AddCondition(myInternalNumberMatchString & publicNumberString1 & " OR (callDirection = 'In' AND ((B.DestinationNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & publicNumberString2 & ") OR (C.DestinationNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & publicNumberString3 & ") OR (B.DestinationNumber = " & WrapInSingleQuotes(SingleQuoteCheck(myName)) & ")))")
                    mySQLStatement.AddCondition(myInternalNumberMatchString & publicNumberString1 & " OR (callDirection = 'In' AND ((B.DestinationNumber " & myOperator & " " & myParms.CreateStringParm(myExtension) & publicNumberString2 & ") OR (C.DestinationNumber " & myOperator & " " & myParms.CreateStringParm(myExtension) & publicNumberString3 & ") OR (B.DestinationNumber = " & myParms.CreateStringParm(myName) & ")))")
                Else
                    ' --> mySQLStatement.AddCondition(myInternalNumberMatchString & publicNumberString1 & " OR (callDirection = 'In' AND (DestinationNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & publicNumberString2 & "))")
                    mySQLStatement.AddCondition(myInternalNumberMatchString & publicNumberString1 & " OR (callDirection = 'In' AND (DestinationNumber " & myOperator & " " & myParms.CreateStringParm(myExtension) & publicNumberString2 & "))")
                End If
            Else
                mySQLStatement.AddCondition(myInternalNumberMatchString)
            End If
        End If

        If p_destination.Length > 0 Then
            Dim myExternalNumber As String = p_destination
            Dim myOperator As String = "="

            If myExternalNumber.Contains("*") Then
                myOperator = "LIKE"
                myExternalNumber = myExternalNumber.Replace("*", "%")
            End If

            ' --> mySQLStatement.AddCondition(externalNumberText & " " & myOperator & " " & WrapInSingleQuotes(myExternalNumber))
            mySQLStatement.AddCondition(externalNumberText & " " & myOperator & " " & myParms.CreateStringParm(myExternalNumber))
        End If

        If myInData <> "" Then
            If xrefDestination Then
                If infrontFix Then
                    If xrefTransferredCalls Then
                        mySQLStatement.AddCondition("(callDirection = 'Out' AND InternalNumber IN " & myInData & ") OR (callDirection = 'In' AND (InternalNumber IN " & myInData & " OR C.DestinationNumber " & myFix & " IN " & myInData & " OR B.DestinationNumber " & myFix & " IN " & myInData & "))")
                    Else
                        mySQLStatement.AddCondition("(callDirection = 'Out' AND InternalNumber IN " & myInData & ") OR (callDirection = 'In' AND (InternalNumber IN " & myInData & " OR DestinationNumber " & myFix & " IN " & myInData & "))")
                    End If
                Else
                    If xrefTransferredCalls Then
                        mySQLStatement.AddCondition("(callDirection = 'Out' AND InternalNumber IN " & myInData & ") OR (callDirection = 'In' AND COALESCE(CASE WHEN IsNumeric(C.DestinationNumber) = 1 THEN C.DestinationNumber ELSE NULL END, CASE WHEN IsNumeric(B.DestinationNumber) = 1 THEN B.DestinationNumber ELSE NULL END, InternalNumber)" & myFix & " IN " & myInData & ")")
                    Else
                        mySQLStatement.AddCondition("(callDirection = 'Out' AND InternalNumber IN " & myInData & ") OR (callDirection = 'In' AND COALESCE(CASE WHEN IsNumeric(DestinationNumber) = 1 THEN DestinationNumber ELSE NULL END, InternalNumber)" & myFix & " IN " & myInData & ")")
                    End If
                End If
            Else
                mySQLStatement.AddCondition("InternalNumber IN " & myInData)
            End If
        End If

        If p_callId > 0 Then mySQLStatement.AddCondition("A.CallId = " & p_callId)

        Select Case p_callType
            Case "" ' This is when the call id passed in as part of the URL for call recording access from a drilldown report

            Case "Both Ways"
                ' Do nothing

            Case "Incoming"
                mySQLStatement.AddCondition("CallDirection = 'In'")

            Case "Outgoing"
                mySQLStatement.AddCondition("CallDirection = 'Out'")

            Case Else
                mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, "CallOptionsTable AS D", "CallId", "CallId")
                mySQLStatement.AddCondition("CallDescription = " & WrapInSingleQuotes(p_callType))
        End Select

        If p_startTimeText <> "" Then
            Dim myHour As Integer = CInt(p_startTimeText.Substring(0, 2))
            Dim myMinutes As Integer = CInt(p_startTimeText.Substring(3, 2))
            Dim asMinutes As Integer = (myHour * 60) + myMinutes

            If asMinutes > 0 Then mySQLStatement.AddCondition("(60 * DATEPART(hh, " & timeStampName & ")) + DATEPART(mi, " & timeStampName & ") >= " & asMinutes)
        End If

        If p_endTimeText <> "" Then
            Dim myHour As Integer = CInt(p_endTimeText.Substring(0, 2))
            Dim myMinutes As Integer = CInt(p_endTimeText.Substring(3, 2))
            Dim asMinutes As Integer = (myHour * 60) + myMinutes

            If asMinutes < ((23 * 60) + 59) Then mySQLStatement.AddCondition("(60 * DATEPART(hh, " & timeStampName & ")) + DATEPART(mi, " & timeStampName & ") <= " & asMinutes)
        End If

        Dim mySqlConnection As New SqlConnection(CreateConnectionString(callRecordingInterfaceConfigDictionary, True))
        Dim mySqlCommand As New SqlCommand(mySQLStatement.GetSQLStatement, mySqlConnection)
        Dim myReader As SqlDataReader = Nothing

        myParms.AddValues(mySqlCommand)
        mySqlCommand.Connection.Open()
        myReader = mySqlCommand.ExecuteReader()
        myTable.Load(myReader)
        mySqlCommand.Connection.Close()
        ' *** End Of SQL Query B ***

        ' Determine the total number of matches in the database
        If myTable.Columns.Count > 0 Then
            If myTable.Rows.Count > 0 Then
                If Not myTable.Rows(0).Item(0) Is DBNull.Value Then
                    Dim startRow As Integer = resultOffset + 1
                    Dim endRow As Integer = startRow + maxRecordsToReturn - 1
                    Dim x As New SQLStatementClass

                    recordsMatched = myTable.Rows(0).Item(0)

                    mySQLStatement.ClearSelectStrings()
                    mySQLStatement.AddSelectString("filename", "")
                    mySQLStatement.AddSelectString("a.callid", "")
                    mySQLStatement.AddSelectString("callDirection", "")
                    mySQLStatement.AddSelectString("internalNumber", "")
                    mySQLStatement.AddSelectString(externalNumberText, "externalNumber")
                    mySQLStatement.AddSelectString("timestamp", "")
                    mySQLStatement.AddSelectString("filelength", "")
                    mySQLStatement.AddSelectString("status", "")

                    If xrefDestination Then mySQLStatement.AddSelectString("B.DestinationNumber", "")
                    If xrefTransferredCalls Then mySQLStatement.AddSelectString("C.DestinationNumber", "XferDest")

                    If mainNumberToReplace <> "" Then mySQLStatement.AddSelectString("E.Number", "swyxInternalNumber")

                    mySQLStatement.AddSelectString("ROW_NUMBER() OVER (ORDER BY timeStamp DESC)", "rowNum")

                    x.AddSelectString("filename", "")
                    x.AddSelectString("callid", "")
                    x.AddSelectString("callDirection", "")
                    x.AddSelectString("internalNumber", "")
                    x.AddSelectString("externalNumber", "externalNumber")
                    x.AddSelectString("timestamp", "")
                    x.AddSelectString("filelength", "")
                    x.AddSelectString("status", "")

                    If xrefDestination Then x.AddSelectString("DestinationNumber", "")
                    If xrefTransferredCalls Then x.AddSelectString("XferDest", "")

                    If mainNumberToReplace <> "" Then x.AddSelectString("swyxInternalNumber", "")

                    x.SetPrimaryTable("(" & mySQLStatement.GetSQLStatement & ") t")
                    x.AddCondition("rowNum BETWEEN " & startRow & " AND " & endRow)
                    ' Dim x As String = "SELECT * FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY timeStamp DESC) AS [rowNum] FROM callRecordingsTable) t WHERE rowNum BETWEEN " & startRow & " AND " & endRow

                    myTable.Rows.Clear()
                    myTable.Columns.Clear()

                    ' Load the windowed records
                    ' *** SQL Query D ***
                    mySqlConnection = New SqlConnection(CreateConnectionString(callRecordingInterfaceConfigDictionary, True))
                    mySqlCommand = New SqlCommand(x.GetSQLStatement, mySqlConnection)
                    myParms.AddValues(mySqlCommand)
                    mySqlCommand.Connection.Open()
                    myReader = mySqlCommand.ExecuteReader
                    myTable.Load(myReader)
                    mySqlCommand.Connection.Close()
                    'If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), x.GetSQLStatement, myTable) Then loadedOK = True
                    ' *** End of SQL Query D ***
                End If
            End If
        End If

        ' Display the windowed records
        CreateTable(p_dataSet, "Records", myFieldNames)

        Dim altFilenames As New Dictionary(Of Integer, String)

        If locateXferFromOutgoingCall Then
            Dim myList As String = ""
            Dim myCommand As String = ""
            Dim xferTable As New DataTable

            For i = 0 To myTable.Rows.Count - 1
                Dim callid As String = SafeDBRead(myTable, i, "callId")
                Dim callIdAsInt As Integer = -1

                If callid = "" Then callid = "-1"

                callIdAsInt = CInt(callid)

                If callIdAsInt >= 0 Then
                    If myList <> "" Then myList &= ", "

                    myList &= callid
                End If
            Next

            If myList.Length > 0 Then
                myList = "(" & myList & ")"

                myCommand = "select a.callid, d.filename from " & callRecordingsTableName & " as a"
                myCommand &= " left join " & cdrTableName & " as b on a.callid = b.callid and convert(date, a.timestamp) = convert(date, b.starttime)"
                myCommand &= " left join " & cdrTableName & " as c on b.transferredToCallId = c.callid and b.transferredCallId1 <> b.callid"
                myCommand &= " and b.transferredToCallId <> 0"
                myCommand &= " and convert(date, b.starttime) = convert(date, c.starttime)"
                myCommand &= " left join " & callRecordingsTableName & " as d on c.transferredCallId1 = d.callId and convert(date, c.starttime) = convert(date, d.timestamp)"
                myCommand &= "where a.CallId IN " & myList

                'myCommand = "select a.callid, (select filename from " & callRecordingsTableName & " as c where c.callid = b.transferredcallid1) from "
                'myCommand &= cdrTableName & " as a left join " & cdrTableName & " as b on a.TransferredToCallId = b.CallId and b.TransferredCallId1 <> a.CallId "
                'myCommand &= "where a.CallId IN " & myList

                If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), myCommand, xferTable) Then
                    For i = 0 To xferTable.Rows.Count - 1
                        With xferTable.Rows(i)
                            If .Item(0) IsNot DBNull.Value Then
                                If IsInteger(.Item(0)) Then
                                    Dim myCallId As Integer = CInt(.Item(0))

                                    If .Item(1) IsNot DBNull.Value Then
                                        If Not altFilenames.ContainsKey(myCallId) Then altFilenames.Add(myCallId, .Item(1))
                                    End If
                                End If
                            End If
                        End With
                    Next
                End If
            End If
        End If

        For i = 0 To myTable.Rows.Count - 1
            Dim dateAndTime As New DateAndTimeClass
            Dim direction, internalNumber, externalNumber, callid, dateStamp, timeStamp, displayedInternalNumber, filename As String
            Dim altFileName As String = ""
            Dim thisFileIsArchived As Boolean = False
            Dim myFileLength As Integer = 0
            Dim myStatus As Integer = 0

            direction = SafeDBRead(myTable, i, "callDirection")
            internalNumber = SafeDBRead(myTable, i, "internalNumber")
            externalNumber = SafeDBRead(myTable, i, "externalNumber")
            callid = SafeDBRead(myTable, i, "callId")
            dateStamp = ""
            timeStamp = ""
            filename = SafeDBRead(myTable, i, "filename")

            If callid = "" Then callid = "-1"

            If altFilenames.ContainsKey(CInt(callid)) Then altFileName = altFilenames.Item(CInt(callid))

            If Not myTable.Rows(i).Item("filelength") Is DBNull.Value Then myFileLength = CInt(myTable.Rows(i).Item("filelength"))
            If Not myTable.Rows(i).Item("status") Is DBNull.Value Then myStatus = CInt(myTable.Rows(i).Item("status"))

            displayedInternalNumber = internalNumber

            If myStatus = 1 And archiveEnabled Then thisFileIsArchived = True

            If StrComp(direction, "In", CompareMethod.Text) = 0 Then
                If xrefDestination Then
                    If Not myTable.Rows(i).Item("DestinationNumber") Is DBNull.Value Then
                        Dim myDest As String = myTable.Rows(i).Item("DestinationNumber")

                        If myDest <> internalNumber Then displayedInternalNumber &= " >> " & myDest
                    End If

                    If xrefTransferredCalls Then
                        If Not myTable.Rows(i).Item("XferDest") Is DBNull.Value Then
                            Dim myDest As String = myTable.Rows(i).Item("XferDest")

                            If myDest <> internalNumber Then displayedInternalNumber &= " >> " & myDest
                        End If
                    End If
                End If
            Else
                If mainNumberToReplace <> "" Then
                    If Not myTable.Rows(i).Item("swyxInternalNumber") Is DBNull.Value Then
                        Dim myOrig As String = myTable.Rows(i).Item("swyxInternalNumber")

                        If myOrig <> "" Then
                            If myOrig <> internalNumber Then displayedInternalNumber &= " << " & myOrig
                        End If
                    End If
                End If
            End If

            If Not myTable.Rows(i).Item("timeStamp") Is DBNull.Value Then
                Dim dt As DateTime = myTable.Rows(i).Item("timestamp")

                dateStamp = dt.ToString("yyyyMMdd")
                timeStamp = dt.ToString("HHmmss")
                'dateAndTime.SetFromDBReadString(myTable.Rows(i).Item("timeStamp"))
                'dateStamp = dateAndTime.AsCallRecordingFilenameDate
                ' timeStamp = dateAndTime.AsCallRecordingFilenameTime
            End If

            filename = GetGUIDFileName(filename)

            If locateXferFromOutgoingCall Then
                AddRow(p_dataSet, thisFileIsArchived, dateStamp, timeStamp, MapFilename(filename, recordingsPath), myFileLength, displayedInternalNumber, externalNumber, direction, callid, p_loginExtension, filename, MapFilename(altFileName, recordingsPath))
            Else
                AddRow(p_dataSet, thisFileIsArchived, dateStamp, timeStamp, MapFilename(filename, recordingsPath), myFileLength, displayedInternalNumber, externalNumber, direction, callid, p_loginExtension, filename)
            End If
        Next

        Return recordsMatched
    End Function

    Private Function GetGUIDFileName(ByRef p_filename) As String
        Dim guidFileName As String = p_filename
        Dim mySql As String = "select guidFileName from CallRecordingsGUIDTable where filename = " & WrapInSingleQuotes(p_filename)
        Dim myTable As New DataTable

        If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), mySql, myTable) Then
            If myTable.Rows.Count > 0 Then
                With myTable.Rows(0)
                    If Not .Item(0) Is DBNull.Value Then
                        guidFileName = .Item(0)
                    End If
                End With
            End If
        End If

        Return guidFileName
    End Function
End Module
