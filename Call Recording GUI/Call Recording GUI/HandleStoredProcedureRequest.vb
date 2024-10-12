Module HandleStoredProcedureRequest
    Public Function NewHandleRequestStoredProcedure(ByRef p_dataSet As DataSet, ByVal resultOffset As Integer, ByRef p_startDate As String, ByRef p_endDate As String, ByRef p_extension As String, ByRef p_destination As String, ByVal p_callId As Integer, ByRef p_callType As String, ByRef p_startTimeText As String, ByRef p_endTimeText As String, ByRef p_securityData As String, ByRef p_loginExtension As String, ByVal p_useLegacyTable As Boolean) As Integer
        ' Get the total number of available records
        ' *** SQL Query B ***
        Dim maxRecordsToReturn As Integer = CInt(callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.maxRecordsToReturn)))
        Dim myTable As New DataTable
        Dim myStartTime As String = ""
        Dim myEndTime As String = ""
        Dim recordsMatched As Integer = 0
        Dim loadedOK As Boolean = False
        Dim archiveEnabled As Boolean = callRecordingInterfaceConfigDictionary.GetBooleanItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.archiveEnabled))
        Dim xrefDestination As Boolean = callRecordingInterfaceConfigDictionary.GetBooleanItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.xrefDestination))
        Dim xrefTransferredCalls As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.xrefTransferredCalls))
        Dim recordingsPath As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.recordingsPath))
        Dim mainNumberToReplace As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.mainNumberToReplace))
        Dim locateXferFromOutgoingCall As Boolean = callRecordingInterfaceConfigDictionary.GetBooleanItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.locateXferFromOutgoingCall))

        If p_startDate.Length = 8 Then myStartTime = ConvertDateToTimeStamp(p_startDate) & " 00:00:00"
        If p_endDate.Length = 8 Then myEndTime = ConvertDateToTimeStamp(p_endDate) & " 23:59:59"

        Dim mySql As String = "EXEC SP_CALL_RECORDING_ACCESS 1, " & WrapInSingleQuotes(myStartTime) & ", " & WrapInSingleQuotes(myEndTime) & ", " & p_callId & ", " & WrapInSingleQuotes(p_extension) & ", " & WrapInSingleQuotes(p_destination) & ", " & WrapInSingleQuotes(p_callType) & ", " & WrapInSingleQuotes(p_startTimeText) & ", " & WrapInSingleQuotes(p_endTimeText) & ", " & p_useLegacyTable.ToString & ", NULL, NULL"
        ' *** End Of SQL Query B ***

        ' Determine the total number of matches in the database
        If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), mySql, myTable) Then
            If myTable.Columns.Count > 0 Then
                If myTable.Rows.Count > 0 Then
                    If Not myTable.Rows(0).Item(0) Is DBNull.Value Then
                        Dim startRow As Integer = resultOffset + 1
                        Dim endRow As Integer = startRow + maxRecordsToReturn - 1
                        Dim x As New SQLStatementClass

                        recordsMatched = myTable.Rows(0).Item(0)
                        mySql = "EXEC SP_CALL_RECORDING_ACCESS 2, " & WrapInSingleQuotes(myStartTime) & ", " & WrapInSingleQuotes(myEndTime) & ", " & p_callId & ", " & WrapInSingleQuotes(p_extension) & ", " & WrapInSingleQuotes(p_destination) & ", " & WrapInSingleQuotes(p_callType) & ", " & WrapInSingleQuotes(p_startTimeText) & ", " & WrapInSingleQuotes(p_endTimeText) & ", " & p_useLegacyTable.ToString & ", " & startRow & ", " & endRow

                        'If mainNumberToReplace <> "" Then mySQLStatement.AddSelectString("E.Number", "swyxInternalNumber")
                        'If mainNumberToReplace <> "" Then x.AddSelectString("swyxInternalNumber", "")

                        myTable.Rows.Clear()
                        myTable.Columns.Clear()

                        ' Load the windowed records
                        If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), mySql, myTable) Then loadedOK = True
                    End If
                End If
            End If
        End If

        ' Display the windowed records
        CreateTable(p_dataSet, "Records", myFieldNames)

        Dim altFilenames As New Dictionary(Of Integer, String)

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
End Module
