Module NewHandler
    Public Sub NewHandleRequest1(ByVal resultOffset As Integer, Optional ByVal p_callId As Integer = 0)
        Dim mySQLStatement As New SQLStatementClass
        Dim myTable As New DataTable
        Dim callRecordingsTableName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.callRecordingTable))

        mySQLStatement.SetPrimaryTable(callRecordingsTableName)

        ' Get the total number of available records
        mySQLStatement.AddSelectString("COUNT(*)", "")

        If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), mySQLStatement.GetSQLStatement, myTable) Then
            If myTable.Columns.Count > 0 Then
                If myTable.Rows.Count > 0 Then
                    Dim totalMatchedRecords As Integer = 0

                    If Not myTable.Rows(0).Item(0) Is DBNull.Value Then
                        Dim maxRecordsToReturn As Integer = CInt(callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.maxRecordsToReturn)))
                        Dim startRow As Integer = resultOffset + 1
                        Dim endRow As Integer = startRow + maxRecordsToReturn - 1

                        totalMatchedRecords = myTable.Rows(0).Item(0)

                        Dim x As String = "SELECT * from (SELECT *, ROW_NUMBER() OVER (ORDER BY timeStamp DESC) AS [rowNum] FROM " & callRecordingsTableName & ") t where rowNum between " & startRow & " AND " & endRow

                        myTable.Rows.Clear()
                        myTable.Columns.Clear()

                        If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), x, myTable) Then
                            If myTable.Columns.Count > 0 Then
                                Dim recordsReturned As Integer = myTable.Rows.Count
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
End Module
