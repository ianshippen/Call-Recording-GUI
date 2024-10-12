Imports System.Web

Module HandleRequest
    Public myFieldNames() As String = {"time", "date", "internalNumber", "externalNumber", "direction", "duration", "callid", "action"}
    Public Const FOR_ALDERMORE As Boolean = True
    Public Const FOR_NEW_ALDERMORE As Boolean = False ' This is for the new Aldermore server implementation
    Public Const USE_LONG_FILENAMES As Boolean = False  ' This was going to be used as an option for Aldermore secure recording vault which was not implemented
    Public Const USE_PARAMETERISED_SQL As Boolean = False
    Public Const XFER_DEST_STRING As String = "XferDest"
    Public Const DOUBLE_XFER_DEST_STRING As String = "DoubleXferDest"
    Public Const doubleXrefTransferredCall As Boolean = True ' For Romec
    Private Const INFRONT_FIX As Boolean = True    ' To allow transferred calls to show up in supervisor login mode

    ' If no start date then empty string. Specified date gets converted from DD/MM/YYYY to YYYYMMDD by ParseDate() 
    ' If no end date then empty string. Specified date gets converted from DD/MM/YYYY to YYYYMMDD by ParseDate()
    ' Default start time is "00:00"
    ' Default end time is "23:59"
    Public Function NewHandleRequest(ByRef p_dataSet As DataSet, ByVal resultOffset As Integer, ByRef p_startDate As String, ByRef p_endDate As String, ByRef p_extension As String, ByRef p_destination As String, ByVal p_callId As Integer, ByRef p_callType As String, ByRef p_startTimeText As String, ByRef p_endTimeText As String, ByRef p_securityData As String, ByRef p_loginExtension As String, ByVal p_useLegacyTable As Boolean) As Integer

        If USE_PARAMETERISED_SQL Then Return NewHandleRequestParameterisedSQL(p_dataSet, resultOffset, p_startDate, p_endDate, p_extension, p_destination, p_callId, p_callType, p_startTimeText, p_endTimeText, p_securityData, p_loginExtension, p_useLegacyTable)

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
        Dim swyxDatabaseName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.swyxDatabaseName))
        Dim mainNumberToReplace As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.mainNumberToReplace))
        Dim locateXferFromOutgoingCall As Boolean = callRecordingInterfaceConfigDictionary.GetBooleanItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.locateXferFromOutgoingCall))

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

        If doubleXrefTransferredCall Then mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, cdrTableName & " AS D", "C.TransferredToCallId", "D.CallId and D.CallId > 0 and convert(date, C.startTime) = convert(date, D.startTime)")

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
            myInternalNumberMatchString = "(InternalNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & ")"

            If mainNumberToReplace <> "" Then
                myInternalNumberMatchString = "(" & myInternalNumberMatchString & " OR (InternalNumber = " & WrapInSingleQuotes(mainNumberToReplace) & " AND Number " & myOperator & " " & WrapInSingleQuotes(myExtension) & "))"
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
                mySql &= WrapInSingleQuotes(p_extension)
                FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), mySql, myTable)
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
                myTable.Rows.Clear()
                myTable.Columns.Clear()

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
                Dim publicNumberString4 As String = ""

                If publicNumberString <> "" Then
                    publicNumberString1 = " OR InternalNumber " & publicNumberString
                    publicNumberString2 = " OR B.DestinationNumber " & publicNumberString
                    publicNumberString3 = " OR C.DestinationNumber " & publicNumberString
                    publicNumberString4 = " OR D.DestinationNumber " & publicNumberString
                End If

                If xrefTransferredCalls Then
                    If doubleXrefTransferredCall Then
                        mySQLStatement.AddCondition(myInternalNumberMatchString & publicNumberString1 & " OR (callDirection = 'In' AND ((B.DestinationNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & publicNumberString2 & ") OR (C.DestinationNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & publicNumberString3 & ") OR (D.DestinationNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & publicNumberString4 & ") OR (B.DestinationNumber = " & WrapInSingleQuotes(SingleQuoteCheck(myName)) & ")))")
                    Else
                        mySQLStatement.AddCondition(myInternalNumberMatchString & publicNumberString1 & " OR (callDirection = 'In' AND ((B.DestinationNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & publicNumberString2 & ") OR (C.DestinationNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & publicNumberString3 & ") OR (B.DestinationNumber = " & WrapInSingleQuotes(SingleQuoteCheck(myName)) & ")))")
                    End If
                Else
                    mySQLStatement.AddCondition(myInternalNumberMatchString & publicNumberString1 & " OR (callDirection = 'In' AND (DestinationNumber " & myOperator & " " & WrapInSingleQuotes(myExtension) & publicNumberString2 & "))")
                End If
            Else
                mySQLStatement.AddCondition(myInternalNumberMatchString)
            End If
        End If

        If p_destination.Length > 0 Then
            Dim x As String = p_destination
            Dim myOperator As String = "="

            If x.Contains("*") Then
                myOperator = "LIKE"
                x = x.Replace("*", "%")
            End If

            mySQLStatement.AddCondition(externalNumberText & " " & myOperator & " " & WrapInSingleQuotes(x))
        End If

        If myInData <> "" Then
            If xrefDestination Then
                If INFRONT_FIX Then
                    If xrefTransferredCalls Then
                        If doubleXrefTransferredCall Then
                            mySQLStatement.AddCondition("(callDirection = 'Out' AND InternalNumber IN " & myInData & ") OR (callDirection = 'In' AND (InternalNumber IN " & myInData & " OR D.DestinationNumber " & myFix & " IN " & myInData & " OR C.DestinationNumber " & myFix & " IN " & myInData & " OR B.DestinationNumber " & myFix & " IN " & myInData & "))")
                        Else
                            mySQLStatement.AddCondition("(callDirection = 'Out' AND InternalNumber IN " & myInData & ") OR (callDirection = 'In' AND (InternalNumber IN " & myInData & " OR C.DestinationNumber " & myFix & " IN " & myInData & " OR B.DestinationNumber " & myFix & " IN " & myInData & "))")
                        End If
                    Else
                        mySQLStatement.AddCondition("(callDirection = 'Out' AND InternalNumber IN " & myInData & ") OR (callDirection = 'In' AND (InternalNumber IN " & myInData & " OR DestinationNumber " & myFix & " IN " & myInData & "))")
                    End If
                Else
                    If xrefTransferredCalls Then
                        If doubleXrefTransferredCall Then
                            mySQLStatement.AddCondition("(callDirection = 'Out' AND InternalNumber IN " & myInData & ") OR (callDirection = 'In' AND COALESCE(CASE WHEN IsNumeric(D.DestinationNumber) = 1 THEN C.DestinationNumber ELSE NULL END, CASE WHEN IsNumeric(C.DestinationNumber) = 1 THEN C.DestinationNumber ELSE NULL END, CASE WHEN IsNumeric(B.DestinationNumber) = 1 THEN B.DestinationNumber ELSE NULL END, InternalNumber)" & myFix & " IN " & myInData & ")")
                        Else
                            mySQLStatement.AddCondition("(callDirection = 'Out' AND InternalNumber IN " & myInData & ") OR (callDirection = 'In' AND COALESCE(CASE WHEN IsNumeric(C.DestinationNumber) = 1 THEN C.DestinationNumber ELSE NULL END, CASE WHEN IsNumeric(B.DestinationNumber) = 1 THEN B.DestinationNumber ELSE NULL END, InternalNumber)" & myFix & " IN " & myInData & ")")
                        End If
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
                ' ** Joining as D will clash with the double transfer join so change to E
                mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, "CallOptionsTable AS E", "CallId", "CallId")
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
        ' *** End Of SQL Query B ***

        ' Determine the total number of matches in the database
        If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), mySQLStatement.GetSQLStatement, myTable) Then
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
                        If xrefTransferredCalls Then mySQLStatement.AddSelectString("C.DestinationNumber", XFER_DEST_STRING)
                        If doubleXrefTransferredCall Then mySQLStatement.AddSelectString("D.DestinationNumber", DOUBLE_XFER_DEST_STRING)

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
                        If xrefTransferredCalls Then x.AddSelectString(XFER_DEST_STRING, "")
                        If doubleXrefTransferredCall Then x.AddSelectString(DOUBLE_XFER_DEST_STRING, "")

                        If mainNumberToReplace <> "" Then x.AddSelectString("swyxInternalNumber", "")

                        x.SetPrimaryTable("(" & mySQLStatement.GetSQLStatement & ") t")
                        x.AddCondition("rowNum BETWEEN " & startRow & " AND " & endRow)
                        ' Dim x As String = "SELECT * FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY timeStamp DESC) AS [rowNum] FROM callRecordingsTable) t WHERE rowNum BETWEEN " & startRow & " AND " & endRow

                        myTable.Rows.Clear()
                        myTable.Columns.Clear()

                        ' Load the windowed records
                        If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), x.GetSQLStatement, myTable) Then loadedOK = True
                    End If
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
                        If doubleXrefTransferredCall Then
                            If Not myTable.Rows(i).Item(DOUBLE_XFER_DEST_STRING) Is DBNull.Value Then
                                Dim myDest As String = myTable.Rows(i).Item(DOUBLE_XFER_DEST_STRING)

                                If myDest <> internalNumber Then displayedInternalNumber &= " >> " & myDest
                            Else
                                If Not myTable.Rows(i).Item(XFER_DEST_STRING) Is DBNull.Value Then
                                    Dim myDest As String = myTable.Rows(i).Item(XFER_DEST_STRING)

                                    If myDest <> internalNumber Then displayedInternalNumber &= " >> " & myDest
                                End If
                            End If
                        Else
                            If Not myTable.Rows(i).Item(XFER_DEST_STRING) Is DBNull.Value Then
                                Dim myDest As String = myTable.Rows(i).Item(XFER_DEST_STRING)

                                If myDest <> internalNumber Then displayedInternalNumber &= " >> " & myDest
                            End If
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

            If locateXferFromOutgoingCall Then
                AddRow(p_dataSet, thisFileIsArchived, dateStamp, timeStamp, MapFilename(filename, recordingsPath), myFileLength, displayedInternalNumber, externalNumber, direction, callid, p_loginExtension, filename, MapFilename(altFileName, recordingsPath))
            Else
                AddRow(p_dataSet, thisFileIsArchived, dateStamp, timeStamp, MapFilename(filename, recordingsPath), myFileLength, displayedInternalNumber, externalNumber, direction, callid, p_loginExtension, filename)
            End If
        Next

        Return recordsMatched
    End Function

    Public Sub AddRow(ByRef p_dataSet As DataSet, ByVal p_thisFileIsArchived As Boolean, ByRef p_dateStamp As String, ByRef p_timeStamp As String, ByRef p_filename As String, ByVal p_fileLength As Integer, ByRef p_displayedInternalNumber As String, ByRef p_externalNumber As String, ByRef p_direction As String, ByVal p_callid As Integer, ByRef p_loginExtension As String, ByRef p_fullFilename As String, Optional ByRef p_altFileName As String = "")
        Dim extra As String = ""
        Dim myPath As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.virtualPath))
        Dim useStructuredFilesystem As Boolean = callRecordingInterfaceConfigDictionary.GetBooleanItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.useStructuredFileSystem))
        Dim thisFileArchived As Boolean = False
        Dim myDateAndTime As New DateAndTimeClass
        Dim myFields(myFieldNames.Length - 1) As String
        Dim duration As String = "--"
        Dim myIndex As Integer = 0
        Dim recordingRate As Integer = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.recordingRateBytesPerSecond))

        If p_thisFileIsArchived Then
            Dim archiveVirtualPath As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.archiveVirtualPath))

            If archiveVirtualPath.Length > 0 Then myPath = archiveVirtualPath

            thisFileArchived = True
        End If

        If useStructuredFilesystem Or thisFileArchived Then
            extra = p_dateStamp.Substring(0, 4) & "/" & p_dateStamp.Substring(4, 2) & "/" & p_dateStamp.Substring(6, 2) & "/"
        Else
            If myIndex > 0 Then
                extra = BackSlashToForwardSlash(p_filename.Substring(0, myIndex))
            End If
        End If

        Try
            p_timeStamp = p_timeStamp.Substring(0, 2) & ":" & p_timeStamp.Substring(2, 2) & ":" & p_timeStamp.Substring(4, 2)
            p_dateStamp = p_dateStamp.Substring(6, 2) & "/" & p_dateStamp.Substring(4, 2) & "/" & p_dateStamp.Substring(0, 4)
        Catch e As Exception
            LogError("Substring failure on filename " & p_filename & " for timeStamp=" & p_timeStamp & " and datestamp=" & p_dateStamp)
        End Try

        p_dateStamp = FormatDateTime(p_dateStamp, 1)

        If p_fileLength >= 0 Then
            myDateAndTime.SetFromSeconds((p_fileLength - 48) \ recordingRate)
            duration = myDateAndTime.AsTime
        End If

        myFields(0) = p_timeStamp
        myFields(1) = p_dateStamp
        myFields(2) = p_displayedInternalNumber
        myFields(3) = p_externalNumber
        myFields(4) = MapDirection(p_direction)
        myFields(5) = duration

        If p_callid >= 0 Then myFields(6) = p_callid

        Dim myURL As String = ""

        ' Option for Aldermore secure call recording vault
        If USE_LONG_FILENAMES Then
            myURL = p_fullFilename
        Else
            myURL = myPath & extra & PlusToSpace(System.Web.HttpUtility.UrlEncode(p_filename)) & ".wav"
        End If

        'myFields(7) = "<a href=" & WrapInQuotes(x) & " onclick=" & WrapInQuotes("window.open ('http://localhost/SMSHandler/SMSHandler.aspx?client=" & f_loginExtension.Text & "-" & x & "&billingStatus=2')") & ">Listen</a>"

        ' To use this set Download Text to "download" in the Options settings within the app for recordings
        Dim myDownloadText As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.downloadText))

        If FOR_ALDERMORE Then
            If FOR_NEW_ALDERMORE Then
                myFields(7) = "<a href=" & WrapInQuotes("https://STPV-SWYX-REA01/SecureRecordingAccess/SecureRecordingAccess.aspx?userName=" & p_loginExtension & "&requestedRecording=" & myURL) & ">Access Recording</a>"
            Else
                ' myFields(7) = "<a href=" & WrapInQuotes("http://z2swma01/SecureRecordingAccess/SecureRecordingAccess.aspx?userName=" & p_loginExtension & "&requestedRecording=" & myURL) & ">Access Recording</a>"
                Dim myArgs As String = "userName=" & p_loginExtension & "&requestedRecording=" & myURL

                'myArgs = Encrypt(myArgs)
                myFields(7) = "<a href=" & WrapInQuotes("http://localhost/SecureRecordingAccess/SecureRecordingAccess.aspx?" & myArgs) & ">Access Recording</a>"
            End If
        Else
            Dim myField = "<a href=" & WrapInQuotes(myURL) & myDownloadText & ">Listen</a>"

            If p_altFileName <> "" Then
                Dim myAltURL As String = myPath & extra & PlusToSpace(System.Web.HttpUtility.UrlEncode(p_altFileName)) & ".wav"

                myField &= " >> <a href=" & WrapInQuotes(myAltURL) & myDownloadText & ">Listen</a>"
            End If

            myFields(7) = myField
        End If

        AddRow(p_dataSet, "Records", myFields)
    End Sub

    Sub AddRow(ByRef p_dataSet As DataSet, ByVal p_tableName As String, ByRef p_fields As String())
        p_dataSet.Tables(p_tableName).Rows.Add(p_fields)
    End Sub

    Public Function SafeDBRead(ByRef p_dataTable As DataTable, ByVal p_rowIndex As Integer, ByRef p_item As String) As String
        Dim result As String = ""

        If Not p_dataTable.Rows(p_rowIndex).Item(p_item) Is DBNull.Value Then result = p_dataTable.Rows(p_rowIndex).Item(p_item)

        Return result
    End Function

    Public Function PlusToSpace(ByRef p As String) As String
        Dim result As String = p

        If p.Contains("+") Then
            Dim i As Integer

            result = ""

            For i = 0 To p.Length - 1
                If p(i) = "+" Then
                    result &= " "
                Else
                    result &= p(i)
                End If
            Next
        End If

        Return result
    End Function

    Public Sub CreateTable(ByRef p_dataSet As DataSet, ByRef p_tableName As String, ByRef p_fieldNames As String())
        Dim myDataTable As New DataTable(p_tableName)
        Dim i As Integer

        For i = 0 To p_fieldNames.Length - 1
            myDataTable.Columns.Add(New DataColumn(p_fieldNames(i)))
        Next

        myDataTable.AcceptChanges()
        p_dataSet.Tables.Add(myDataTable)
    End Sub

    Public Function MapDirection(ByRef p_direction As String) As String
        Dim result As String = "Outgoing"

        If StrComp(p_direction, "In", CompareMethod.Text) = 0 Then result = "Incoming"

        Return result
    End Function

    Public Sub AddDateCondition(ByRef p_startDate As String, ByRef p_endDate As String, ByRef p_sqlStatement As SQLStatementClass, ByRef p_timeStampFieldName As String)
        Dim useNewDateSearch As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.useNewDateSearch))
        Dim useStartDate As Boolean = False
        Dim useEndDate As Boolean = False

        If p_startDate.Length = 8 Then useStartDate = True
        If p_endDate.Length = 8 Then useEndDate = True

        If useNewDateSearch Then
            If useStartDate Then
                If GetYear(p_startDate) = Now.Year Then
                    If GetMonth(p_startDate) = Now.Month Then
                        If GetDay(p_startDate) = Now.Day Then
                            p_sqlStatement.AddCondition("DATEPART(year, " & p_timeStampFieldName & ") = " & GetYear(p_startDate) & " AND DATEPART(month, " & p_timeStampFieldName & ") = " & GetMonth(p_startDate) & " AND DATEPART(day, " & p_timeStampFieldName & ") = " & GetDay(p_startDate))
                        Else
                            p_sqlStatement.AddCondition("DATEPART(year, " & p_timeStampFieldName & ") = " & GetYear(p_startDate) & " AND DATEPART(month, " & p_timeStampFieldName & ") = " & GetMonth(p_startDate) & " AND DATEPART(day, " & p_timeStampFieldName & ") >= " & GetDay(p_startDate))
                        End If
                    Else
                        p_sqlStatement.AddCondition("DATEPART(year, " & p_timeStampFieldName & ") = " & GetYear(p_startDate) & " AND (DATEPART(month, " & p_timeStampFieldName & ") > " & GetMonth(p_startDate) & " OR (DATEPART(month, " & p_timeStampFieldName & ") = " & GetMonth(p_startDate) & " AND DATEPART(day, " & p_timeStampFieldName & ") >= " & GetDay(p_startDate) & "))")
                    End If
                Else
                    p_sqlStatement.AddCondition("DATEPART(year, " & p_timeStampFieldName & ") > " & GetYear(p_startDate) & " OR (DATEPART(year, " & p_timeStampFieldName & ") = " & GetYear(p_startDate) & " AND (DATEPART(month, " & p_timeStampFieldName & ") > " & GetMonth(p_startDate) & " OR (DATEPART(month, " & p_timeStampFieldName & ") = " & GetMonth(p_startDate) & " AND DATEPART(day, " & p_timeStampFieldName & ") >= " & GetDay(p_startDate) & ")))")
                End If
            End If

            If useEndDate Then
                ' Just filtering on end date
                If Not (GetYear(p_endDate) = Now.Year And GetMonth(p_endDate) = Now.Month And GetDay(p_endDate) = Now.Day) Then
                    p_sqlStatement.AddCondition("DATEPART(year, " & p_timeStampFieldName & ") < " & GetYear(p_endDate) & " OR (DATEPART(year, " & p_timeStampFieldName & ") = " & GetYear(p_endDate) & " AND (DATEPART(month, " & p_timeStampFieldName & ") < " & GetMonth(p_endDate) & " OR (DATEPART(month, " & p_timeStampFieldName & ") = " & GetMonth(p_endDate) & " AND DATEPART(day, " & p_timeStampFieldName & ") <= " & GetDay(p_endDate) & ")))")
                End If
            End If
        Else
            If useStartDate Then
                p_sqlStatement.AddCondition(p_timeStampFieldName & " >= '" & ConvertDateToTimeStamp(p_startDate) & " 00:00:00'")
            End If

            If useEndDate Then
                p_sqlStatement.AddCondition(p_timeStampFieldName & " <= '" & ConvertDateToTimeStamp(p_endDate) & " 23:59:59'")
            End If
        End If
    End Sub

    Private Function GetYear(ByRef p As String) As Integer
        Return CInt(p.Substring(0, 4))
    End Function

    Private Function GetMonth(ByRef p As String) As Integer
        Return CInt(p.Substring(4, 2))
    End Function

    Private Function GetDay(ByRef p As String) As Integer
        Return CInt(p.Substring(6, 2))
    End Function

    Function ConvertDateToTimeStamp(ByRef p As String) As String
        Dim result As String = ""

        If p.Length = 8 Then result = p.Substring(0, 4) & "-" & p.Substring(4, 2) & "-" & p.Substring(6, 2)

        Return result
    End Function

    Public Function MapFilename(ByRef p_filename As String, ByRef p_recordingsPath As String) As String
        Dim result As String = ""

        If p_filename <> "" Then
            Dim lastIndex As Integer = p_filename.LastIndexOf("\")
            Dim pathWithSlash As String = p_recordingsPath.ToLower & "\"

            If p_filename.ToLower.StartsWith(pathWithSlash) Then
                result = p_filename.Substring(pathWithSlash.Length, p_filename.Length - (pathWithSlash.Length + ".wav".Length))
            Else
                result = p_filename.Substring(lastIndex + 1, p_filename.Length - (lastIndex + 1 + ".wav".Length))
            End If
        End If

        Return result
    End Function

    Public Function BackSlashToForwardSlash(ByRef p As String) As String
        Return p.Replace("\", "/")
    End Function
End Module
