Imports System.Xml

Partial Public Class _Default
    Inherits System.Web.UI.Page
    Const USERNAME As String = "Admin"
    Const CALLID_STRING_LENGTH As Integer = 10
    Const USE_NEW_HANDLER As Boolean = True
    Const ALLOW_REQUEST_PARMS As Boolean = True
    Const USE_LEGACY_DATA As Boolean = False
    Const USE_STORED_PROCEDURE As Boolean = True

    Private myFieldNames() As String = {"time", "date", "internalNumber", "externalNumber", "direction", "duration", "callid", "action"}
    Dim usingNewInterface As Boolean = False
    Dim myAspect As New InterfaceAspectClass
    Dim timeSpans() As String = {"This Hour", "Previous Hour", "Today", "Yesterday", "Last 7 Days", "Last Week (Mon - Sun)", "Last 30 Days", "Last Month", "Custom Date/Time"}

    Sub TestButtonPressed(ByVal Source As Object, ByVal e As EventArgs)
        Dim response As HttpResponse = System.Web.HttpContext.Current.Response

        response.ContentType = "audio/wav"
        'response.ContentType = "application/octet-stream"
        response.WriteFile("c:\inetpub\wwwroot\test.wav")
        response.Flush()
    End Sub

    Public Function GetAdminPassword() As String
        Return Decrypt(callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.adminPassword)))
    End Function

    Sub LoginPressed(ByVal Source As Object, ByVal e As EventArgs)
        myAspect.changePasswordResponseLabel.Text = ""

        If CheckLogin() Then
            If UsingNewSecurity(True) Then
                Dim requestPasswordChange As Boolean = False
                Dim myConnectionString As String = GetDatabaseConfigStringForSecurity()

                ' Do we have to change passwords every X days ?
                If GetUserMustChangePasswordDays() > 0 Then
                    Dim myTable As New DataTable
                    Dim myDays As Integer = 0

                    If FillTableFromCommand(myConnectionString, "select datediff(d, lastPasswordChange, getdate()) from " & USER_SECURITY_TABLE_NAME & " where name = " & WrapInSingleQuotes(SingleQuoteCheck(myAspect.userNameTextBox.Text)), myTable) Then
                        If myTable.Rows.Count > 0 Then
                            With myTable.Rows(0)
                                If .Item(0) IsNot DBNull.Value Then
                                    If IsInteger(.Item(0)) Then myDays = CInt(.Item(0))

                                    If myDays > GetUserMustChangePasswordDays() Then requestPasswordChange = True
                                End If
                            End With
                        End If
                    End If
                End If

                ' Do we need to change the password upon first login ?
                If GetUserMustChangePasswordOnFirstLogin() Then
                    Dim myTable As New DataTable
                    Dim numberOfLogins As Integer = 0

                    If FillTableFromCommand(myConnectionString, "select lastPasswordChange from " & Security.USER_SECURITY_TABLE_NAME & " where name = " & WrapInSingleQuotes(SingleQuoteCheck(myAspect.userNameTextBox.Text)), myTable) Then
                        If myTable.Rows.Count > 0 Then
                            With myTable.Rows(0)
                                If .Item(0) Is DBNull.Value Then requestPasswordChange = True
                            End With
                        End If
                    End If
                End If

                If requestPasswordChange Then
                    ShowChangePasswordRows(True)
                Else
                    DoLoginOK(True)
                End If
            Else
                DoLoginOK(True)
            End If
        End If

        ' myAspect.userNameTextBox.Text = ""
        myAspect.passwordTextBox.Text = ""
        LoadTimeDropDowns()
        LoadCallTypes()
    End Sub

    Function CheckLogin() As Boolean
        Dim result As Boolean = False

        CallRecordingInterfaceInitialiseConfig()
        LoadAdminData(GetAdminDataFilename())
        f_data.Text = ""

        'Dim security As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.security))

        If UsingNewSecurity(True) Then
            ' Check the master admin account
            If ((myAspect.userNameTextBox.Text = USERNAME) Or (USERNAME.Length = 0)) And ((myAspect.passwordTextBox.Text = GetAdminPassword()) Or (GetAdminPassword.Length = 0)) Then
                result = True
            End If

            If Not result Then
                ' Check all admin accounts
                For i = 0 To securityList.GetNumberOfUserNames - 1
                    If securityList.GetUserNameFromIndex(i) = myAspect.userNameTextBox.Text Then
                        If (securityList.GetAttributesFromIndex(i) And RECORDINGS_ADMINISTRATOR_MASK) > 0 Then
                            Dim myPassword As String = securityList.GetRecordingsPasswordFromIndex(i)

                            If myPassword = "" Then myPassword = securityList.GetPasswordFromIndex(i)

                            If myAspect.passwordTextBox.Text = myPassword Then
                                If Not securityList.GetExpiredFromIndex(i) Then result = True

                                Exit For
                            End If
                        End If
                    End If
                Next
            End If

            If Not result Then
                ' Check all supervisor accounts
                For i = 0 To securityList.GetNumberOfUserNames - 1
                    If securityList.GetUserNameFromIndex(i) = myAspect.userNameTextBox.Text Then
                        If (securityList.GetAttributesFromIndex(i) And RECORDINGS_SUPERVISOR_MASK) > 0 Then
                            Dim myPassword As String = securityList.GetRecordingsPasswordFromIndex(i)

                            If myPassword = "" Then myPassword = securityList.GetPasswordFromIndex(i)

                            If myAspect.passwordTextBox.Text = myPassword Then
                                If Not securityList.GetExpiredFromIndex(i) Then
                                    f_data.Text = securityList.GetRecordingData(i)
                                    result = True
                                End If

                                Exit For
                            End If
                        End If
                    End If
                Next
            End If

            If result Then
                Dim mySql As String = "update " & USER_SECURITY_TABLE_NAME & " set lastLogonTime = GetDate(), logons = logons + 1 where name = " & WrapInSingleQuotes(SingleQuoteCheck(myAspect.userNameTextBox.Text))
                Dim myConnectionString = GetDatabaseConfigStringForSecurity()

                ExecuteNonQuery(myConnectionString, mySql)
            End If
        Else
            If loginButton.Text = "Recordings" Then
                result = True
            Else
                If ((myAspect.userNameTextBox.Text = USERNAME) Or (USERNAME.Length = 0)) And ((myAspect.passwordTextBox.Text = GetAdminPassword()) Or (GetAdminPassword.Length = 0)) Then
                    ' Show the master admin table if legacy version
                    myAspect.loginPanel.Visible = False

                    If usingNewInterface Then
                        myAspect.reportSelectionPanel.Visible = True
                    Else
                        adminChoiceTable.Visible = True
                    End If
                Else
                    ' Check all admin accounts
                    For i = 0 To myAdministrators.Count - 1
                        If myAspect.userNameTextBox.Text = myAdministrators(i).userName Then
                            If myAspect.passwordTextBox.Text = myAdministrators(i).password Then
                                result = True
                                Exit For
                            End If
                        End If
                    Next

                    If Not result Then
                        ' Check all supervisor accounts
                        For i = 0 To mySupervisors.Count - 1
                            If myAspect.userNameTextBox.Text = mySupervisors(i).userName Then
                                If myAspect.passwordTextBox.Text = mySupervisors(i).password Then
                                    result = True
                                    f_data.Text = mySupervisors(i).data
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If

        Return result
    End Function

    Sub DoLoginOK(ByVal p_gotoCalls As Boolean)
        LoadTimeDropDowns()
        LoadCallTypes()

        'If Security Then
        'Else
        'If ((f_loginExtension.Text = USERNAME) Or (USERNAME.Length = 0)) And ((f_loginPassword.Text = GetAdminPassword()) Or (GetAdminPassword.Length = 0)) Then gotoCalls = True
        'End If

        '   f_loginExtension.Text = ""
        myAspect.passwordTextBox.Text = ""

        If p_gotoCalls Then
            With myAspect
                .loginPanel.Visible = False
                .reportSelectionPanel.Visible = True
                .summaryPanel.Visible = False
                .callIdtextBox.Text = f_parm.Text
            End With
        End If
    End Sub

    Sub LoadTimeDropDowns()
        Dim myStartMinute As Integer = 0
        Dim myInterval As Integer = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.timeIntervalMinutes))
        Dim running As Boolean = True

        myAspect.startTimeDropDownList.Items.Clear()
        myAspect.endTimeDropDownList.Items.Clear()

        While running
            Dim myHour As Integer = myStartMinute \ 60
            Dim myMinute As Integer = myStartMinute Mod 60

            Dim x As String = ""

            If myHour < 10 Then x = "0"

            x = x & myHour & ":"

            If myMinute < 10 Then x &= "0"

            x = x & myMinute

            myAspect.startTimeDropDownList.Items.Add(x)

            myHour = (myStartMinute + myInterval - 1) \ 60
            myMinute = (myStartMinute + myInterval - 1) Mod 60
            x = ""

            If myHour < 10 Then x = "0"

            x = x & myHour & ":"

            If myMinute < 10 Then x &= "0"

            x = x & myMinute

            myAspect.endTimeDropDownList.Items.Add(x)

            myStartMinute += myInterval

            If myStartMinute >= (24 * 60) Then running = False
        End While

        myAspect.endTimeDropDownList.SelectedIndex = myAspect.endTimeDropDownList.Items.Count - 1
    End Sub

    Private Sub LoadCallTypes()
        Dim myList As New List(Of String)

        With myList
            .Add("Both Ways")
            .Add("Incoming")
            .Add("Outgoing")
        End With

        ' See if we can read anything from the CallOptionsTable
        Dim mySql As String = "select distinct CallDescription from CallOptionsTable"
        Dim myTable As New DataTable

        If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), mySql, myTable) Then
            For i = 0 To myTable.Rows.Count - 1
                With myTable.Rows(i)
                    If Not .Item(0) Is DBNull.Value Then
                        If .Item(0) <> "" Then myList.Add(.Item(0))
                    End If
                End With
            Next
        End If

        myList.Sort()

        With myAspect.callTypeList
            .Items.Clear()

            For i = 0 To myList.Count - 1
                .Items.Add(myList(i))
            Next
        End With
    End Sub

    Sub recordingsButtonPressed(ByVal Source As Object, ByVal e As EventArgs)
        adminChoiceTable.Visible = False

        With myAspect
            .reportSelectionPanel.Visible = True
            .summaryPanel.Visible = True
            .callIdtextBox.Text = f_parm.Text
        End With
    End Sub

    Sub closeButtonPressed(ByVal Source As Object, ByVal e As EventArgs)
        adminChoiceTable.Visible = False
        loginPanel.Visible = True
    End Sub

    Sub LogoutPressed(ByVal Source As Object, ByVal e As EventArgs)
        With myAspect
            .loginPanel.Visible = True
            .reportSelectionPanel.Visible = False
            nextButton.Visible = False
            prevButton.Visible = False
            myRepeater.Visible = False
            .summaryPanel.Visible = False
            recordsLabel.Text = ""
            .userNameTextBox.Text = ""
            .passwordTextBox.Text = ""
        End With
    End Sub

    Private Class CreateTimeComparer
        Implements IComparer

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim myX As String = x
            Dim myY As String = y

            Return -StrComp(myX.Substring(myX.Length - 15), myY.Substring(myY.Length - 15))
        End Function
    End Class

    Const DATE_AND_TIME_LENGTH As Integer = 15

    Sub WriteIt(ByRef p As String)
        Response.Write(p & Chr(10))
    End Sub

    Function GetNextFieldx(ByRef p As String) As String
        Dim result As String = p
        Dim pos As Integer = p.IndexOf("#")

        If pos >= 0 Then
            result = p.Substring(0, pos)
        Else
            If pos = 0 Then result = ""
        End If

        If pos >= 0 Then
            p = p.Substring(pos + 1)
        Else
            p = ""
        End If

        Return result
    End Function

    Function ParseDate(ByRef p As String) As String
        Dim x As String = p
        Dim index As Integer
        Dim result As String = ""
        Dim myDay, myMonth, myYear As Integer

        ' Must contain a slash at either index 1 or 2
        index = x.IndexOf("/")

        If index = 1 Or index = 2 Then
            myDay = CInt(x.Substring(0, index))
            x = x.Substring(index + 1)

            index = x.IndexOf("/")

            If index = 1 Or index = 2 Then
                myMonth = CInt(x.Substring(0, index))
                x = x.Substring(index + 1)

                If x.Length = 4 Then
                    myYear = CInt(x)

                    result = myYear

                    If myMonth < 10 Then result &= "0"

                    result &= myMonth

                    If myDay < 10 Then result &= "0"

                    result &= myDay
                End If
            End If
        End If

        Return result
    End Function

    Sub AddRow(ByRef p_dataSet As DataSet, ByVal p_tableName As String, ByRef p_fields As String())
        p_dataSet.Tables(p_tableName).Rows.Add(p_fields)
    End Sub

    Sub Submit(ByVal Source As Object, ByVal e As EventArgs)
        Dim myResultsOffset As Integer = 0

        HandleRequest(myResultsOffset)
    End Sub

    Sub NewInterfaceSubmit(ByVal Source As Object, ByVal e As EventArgs)
        Dim myResultsOffset As Integer = 0

        HandleRequest(myResultsOffset)
    End Sub

    Sub Clear(ByVal Source As Object, ByVal e As EventArgs)
        With myAspect
            f_startDate.Text = ""
            f_endDate.Text = ""
            .extensionTextBox.Text = ""
            .externalNumberTextBox.Text = ""
            .startDateCalendar.SelectedDate = New DateTime()
            .endDateCalendar.SelectedDate = New DateTime()
            .callIdtextBox.Text = ""
            .callTypeList.Text = "Both Ways"
            .startTimeDropDownList.SelectedIndex = 0
            .endTimeDropDownList.SelectedIndex = .endTimeDropDownList.Items.Count - 1

            If Not .timeSpanDropDownList Is Nothing Then .timeSpanDropDownList.SelectedIndex = 8

            If Not .ignoreTimeCheckbox Is Nothing Then .ignoreTimeCheckbox.Checked = True

            ' f_tag.Text = ""
        End With
    End Sub

    Sub AddToSelectString(ByRef p_selectString As String, ByRef p_auxData As String)
        ' Is there actually any condition to add ?
        If p_auxData.Length > 0 Then
            Dim myAuxData As String = "(" & p_auxData & ")"

            If p_selectString.Length > 0 Then
                p_selectString &= " AND " & myAuxData
            Else
                p_selectString = myAuxData
            End If
        End If
    End Sub

    Private Function GetIndexOfField(ByRef p_callRecordingFilenameFormat As String, ByRef p_fieldName As String) As Integer
        Dim result As Integer = -1
        Dim myFields() As String = p_callRecordingFilenameFormat.Split("#")

        For i = 0 To myFields.Count - 1
            If StrComp(myFields(i), p_fieldName, CompareMethod.Text) = 0 Then
                result = i
                Exit For
            End If
        Next

        Return result
    End Function

    Private Sub DigDirectory(ByRef p_path As String, ByRef p_filter As String, ByRef p_callIdFilter As String, ByRef p_list As List(Of String))
        Dim myDirectoryInfo As New IO.DirectoryInfo(p_path)

        For Each x As System.IO.FileSystemInfo In myDirectoryInfo.GetFileSystemInfos()
            If x.Attributes And IO.FileAttributes.Directory Then
                DigDirectory(x.FullName, p_filter, p_callIdFilter, p_list)
            Else
                Dim useFile As Boolean = True

                If p_filter.Length > 0 Then
                    If Not x.Name.ToLower.EndsWith(p_filter.Substring(1)) Then useFile = False
                End If

                If useFile Then
                    If p_callIdFilter.Length > 0 Then
                        If Not x.Name.Contains(p_callIdFilter) Then useFile = False
                    End If
                End If

                If useFile Then
                    p_list.Add(x.FullName)
                End If
            End If
        Next
    End Sub

    'Sub HandleRequest(ByVal resultOffset As Integer, Optional ByVal p_callId As Integer = 0)
    Sub HandleRequest(ByVal resultOffset As Integer, Optional ByRef p_parmDictionary As Dictionary(Of String, String) = Nothing)
        Dim myDataSet As New DataSet
        Dim maxRecordsToReturn As Integer = CInt(callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.maxRecordsToReturn)))
        Dim recordsMatched As Integer = 0
        Dim providedCallId As String = myAspect.callIdtextBox.Text.TrimStart("0")
        Dim p_callId As Integer = 0
        Dim useLegacyTable As Boolean = legacyCheckBox.Checked

        If USE_NEW_HANDLER Then
            Dim myCallId As Integer = 0
            Dim myStartDate As String = ParseDate(f_startDate.Text)
            Dim myEndDate As String = ParseDate(f_endDate.Text)
            Dim myExtension As String = myAspect.extensionTextBox.Text
            Dim myDestination As String = myAspect.externalNumberTextBox.Text
            Dim myStartTime As String = myAspect.startTimeDropDownList.Text
            Dim myEndTime As String = myAspect.endTimeDropDownList.Text

            If usingNewInterface Then
                If myAspect.ignoreTimeCheckbox.Checked Then
                    myStartTime = ""
                    myEndTime = ""
                End If
            End If

            If IsInteger(providedCallId) Then myCallId = CInt(providedCallId)

            If Not p_parmDictionary Is Nothing Then
                ' Handle any parameters passed in as part of the URL
                For Each myKey In p_parmDictionary.Keys
                    Select Case myKey.ToLower
                        Case "callid"
                            If IsIntegerByVal(p_parmDictionary.Item(myKey)) Then
                                p_callId = p_parmDictionary.Item(myKey)

                                If p_callId > 0 Then myCallId = p_callId
                            End If

                        Case "startdate"
                            myStartDate = p_parmDictionary.Item(myKey)

                        Case "enddate"
                            myEndDate = p_parmDictionary.Item(myKey)

                        Case "extension"
                            myExtension = p_parmDictionary.Item(myKey)

                        Case "externalnumber"
                            myDestination = p_parmDictionary.Item(myKey)

                        Case "starttime"
                            myStartTime = p_parmDictionary.Item(myKey)

                        Case "endtime"
                            myEndTime = p_parmDictionary.Item(myKey)
                    End Select
                Next
            End If

            'recordsMatched = NewHandleRequest(myDataSet, resultOffset, ParseDate(f_startDate.Text), ParseDate(f_endDate.Text), f_extension.Text, f_destination.Text, myCallId, callTypeList.SelectedValue, startTime.Text, endTime.Text, f_data.Text, f_loginExtension.Text)

            If USE_STORED_PROCEDURE Then
                recordsMatched = NewHandleRequestStoredProcedure(myDataSet, resultOffset, myStartDate, myEndDate, myExtension, myDestination, myCallId, myAspect.callTypeList.SelectedValue, myStartTime, myEndTime, f_data.Text, myAspect.userNameTextBox.Text, useLegacyTable)
            Else
                recordsMatched = NewHandleRequest(myDataSet, resultOffset, myStartDate, myEndDate, myExtension, myDestination, myCallId, myAspect.callTypeList.SelectedValue, myStartTime, myEndTime, f_data.Text, myAspect.userNameTextBox.Text, useLegacyTable)
            End If
        Else
            CallRecordingInterfaceInitialiseConfig()
            Options.LoadOptionsAsXML()

            ' Get all the files in the recordings directory
            Dim recordingsPath As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.recordingsPath))
            Dim secondaryrecordingsPath As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.secondaryRecordingsPath))
            Dim myDirectoryInfo As New IO.DirectoryInfo(recordingsPath)
            Dim i As Integer
            Dim myStartDate, myEndDate As String
            Dim recordsReturned As Integer = 0
            Dim filter As String = "*.wav"
            Dim myFileNameArray() As String
            Dim mySecondarySourceArray() As Boolean
            Dim useStructuredFileSystem As Boolean = False
            Dim callRecordingFilenameFormat As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.callRecordingFilenameFormat))
            Dim callDirectionIndex As Integer = GetIndexOfField(callRecordingFilenameFormat, "[calldirection]")
            Dim internalAddressIndex As Integer = GetIndexOfField(callRecordingFilenameFormat, "[internaladdress]")
            Dim remoteAddressIndex As Integer = GetIndexOfField(callRecordingFilenameFormat, "[remoteaddress]")
            Dim dateIndex As Integer = GetIndexOfField(callRecordingFilenameFormat, "[date]")
            Dim timeIndex As Integer = GetIndexOfField(callRecordingFilenameFormat, "[time]")
            Dim callIdIndex As Integer = GetIndexOfField(callRecordingFilenameFormat, "[callid]")
            Dim filterAtDatabase As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.filterAtDatabase))
            Dim searchSubDirectories As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.searchSubDirectories))
            Dim anySecondaryFiles As Boolean = False
            Dim myData As String = f_data.Text
            Dim myDataList As New List(Of String)
            Dim xrefDestination As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.xrefDestination))
            Dim useDatabase As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.useDatabase))
            Dim realDestinations(), transferTable() As String
            Dim fileLengths() As Integer = Nothing
            Dim statusArray() As Integer
            Dim xrefTransferredCalls As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.xrefTransferredCalls))
            Dim callIdAlias As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.callIdAlias))
            Dim providedCallIdList As New List(Of Integer)
            Dim archiveEnabled As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.archiveEnabled))
            Dim myTable As New DataTable
            Dim callRecordingsTableName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.callRecordingTable))

            If filterAtDatabase Then
                Dim mySQLStatement As New SQLStatementClass

                mySQLStatement.SetPrimaryTable(callRecordingsTableName)

                ' Get the total number of available records
                mySQLStatement.AddSelectString("COUNT(*)", "")

                If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), mySQLStatement.GetSQLStatement, myTable) Then
                    If myTable.Columns.Count > 0 Then
                        If myTable.Rows.Count > 0 Then
                            If Not myTable.Rows(0).Item(0) Is DBNull.Value Then
                                Dim startRow As Integer = resultOffset + 1
                                Dim endRow As Integer = startRow + maxRecordsToReturn - 1

                                recordsMatched = myTable.Rows(0).Item(0)

                                Dim x As String = "SELECT * from (SELECT *, ROW_NUMBER() OVER (ORDER BY timeStamp DESC) AS [rowNum] FROM " & callRecordingsTableName & ") t where rowNum between " & startRow & " AND " & endRow

                                myTable.Rows.Clear()
                                myTable.Columns.Clear()

                                If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), x, myTable) Then
                                    For i = 0 To myTable.Columns.Count - 1
                                        ' AddRow(myDataSet, False, False, False, 
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Parse any data into the data list
                ParseSecurityData(myData, myDataList)

                If callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.useStructuredFileSystem)).ToLower = "true" Then useStructuredFileSystem = True

                If p_callId > 0 Then providedCallId = p_callId

                ' Parse start and end dates if present
                myStartDate = ParseDate(f_startDate.Text)
                myEndDate = ParseDate(f_endDate.Text)

                If useDatabase Then
                    Dim callRecordingTable As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.callRecordingTable))
                    Dim cdrTableName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.cdrTableName))
                    Dim timeStampName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.timeStampName))
                    Dim useTempCDRTable As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.useTempCDRTable))
                    Dim mySQLStatement As New SQLStatementClass
                    Dim originalCDRTableName As String = cdrTableName

                    If archiveEnabled Then mySQLStatement.AddSelectString("A.status", "")

                    mySQLStatement.AddSelectString("A.fileLength", "")
                    mySQLStatement.AddSelectString("A.externalNumber", "")
                    mySQLStatement.AddSelectString("A.callId", "")
                    mySQLStatement.AddSelectString("A.timeStamp", "")

                    ' Only use temp CDR table if startdate is defined and it is for 2013 onwards ...
                    If useTempCDRTable Then
                        useTempCDRTable = False

                        If myStartDate.Length = 8 Then
                            Dim myYear As String = myStartDate.Substring(0, 4)

                            If IsInteger(myYear) Then
                                If myYear >= 2013 Then
                                    useTempCDRTable = True
                                    cdrTableName = "#tempCDR"
                                End If
                            End If
                        End If
                    End If

                    mySQLStatement.SetPrimaryTable(callRecordingTable)

                    If providedCallId.Length > 0 Then
                        ' If we have the callid, don't need any other criteria
                        mySQLStatement.AddCondition("A.callId=(SELECT dbo.GetCallRecordingCallID(" & providedCallId & "))")

                        If xrefDestination Then
                            mySQLStatement.AddSelectString("A.filename", "")
                            mySQLStatement.AddSelectString("A.InternalNumber", "")
                            mySQLStatement.AddSelectString("A.CallDirection", "")

                            mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, cdrTableName & " AS B", "CallId", "CallId")

                            If xrefTransferredCalls Then
                                mySQLStatement.AddSelectString("B.DestinationNumber", "")
                                mySQLStatement.AddSelectString("C.DestinationNumber", "XferDest")
                                mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, cdrTableName & " AS C", "B.TransferredToCallId", "CallId")
                            Else
                                mySQLStatement.AddSelectString("A.DestinationNumber", "")
                            End If
                        End If
                    Else
                        Dim countSqlString As String = ""

                        If filterAtDatabase Then
                            Dim startRow As Integer = resultOffset + 1
                            Dim endRow As Integer = resultOffset + maxRecordsToReturn

                            ' Count all the rows first
                            '  selectString = "SELECT filename"

                            ' If xrefDestination Then selectString &= ", InternalNumber, CallDirection, DestinationNumber"

                            ' selectString &= " FROM (SELECT ROW_NUMBER() OVER (ORDER BY " & timeStampName & " DESC) AS rowNum, "

                            '  AddToSelectString(rowAuxString, "rowNum >= " & startRow)
                            '  AddToSelectString(rowAuxString, "rowNum <= " & endRow)

                            '   If xrefDestination Then
                            'selectString &= "A.*, B.DestinationNumber FROM " & callRecordingTable & " AS A LEFT JOIN " & cdrTableName & " AS B ON A.CallId = B.CallId"
                            ' Else
                            '   selectString &= "* FROM " & callRecordingTable
                            ' End If
                        Else
                            ' Not filtering at database, are we getting xref info ?
                            If xrefDestination Then
                                mySQLStatement.AddSelectString("A.filename", "")
                                mySQLStatement.AddSelectString("A.InternalNumber", "")
                                mySQLStatement.AddSelectString("A.CallDirection", "")

                                mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, cdrTableName & " AS B", "CallId", "CallId")

                                If xrefTransferredCalls Then
                                    mySQLStatement.AddSelectString("B.DestinationNumber", "")
                                    mySQLStatement.AddSelectString("C.DestinationNumber", "XferDest")
                                    mySQLStatement.AddJoin(SQLStatementClass.JoinType.LEFT_JOIN, cdrTableName & " AS C", "B.TransferredToCallId", "CallId")
                                Else
                                    mySQLStatement.AddSelectString("A.DestinationNumber", "")
                                End If
                            Else
                                mySQLStatement.AddSelectString("A.filename", "")
                            End If
                        End If

                        AddDateCondition(myStartDate, myEndDate, mySQLStatement, timeStampName)

                        If myAspect.extensionTextBox.Text.Length > 0 Then
                            Dim x As String = myAspect.extensionTextBox.Text
                            Dim myOperator As String = "="

                            If x.Contains("*") Then
                                myOperator = "LIKE"
                                x = x.Replace("*", "%")
                            End If

                            If xrefDestination Then
                                If xrefTransferredCalls Then
                                    mySQLStatement.AddCondition("(InternalNumber " & myOperator & " '" & x & "') OR (B.DestinationNumber " & myOperator & " '" & x & "') OR (C.DestinationNumber " & myOperator & " '" & x & "')")
                                Else
                                    mySQLStatement.AddCondition("(InternalNumber " & myOperator & " '" & x & "') OR (DestinationNumber " & myOperator & " '" & x & "')")
                                End If
                            Else
                                mySQLStatement.AddCondition("InternalNumber " & myOperator & " '" & x & "'")
                            End If
                        End If

                        If myAspect.externalNumberTextBox.Text.Length > 0 Then
                            Dim x As String = myAspect.externalNumberTextBox.Text

                            If x.Contains("*") Then
                                mySQLStatement.AddCondition("ExternalNumber LIKE " & WrapInSingleQuotes(x.Replace("*", "%")))
                            Else
                                mySQLStatement.AddCondition("ExternalNumber = " & WrapInSingleQuotes(x))
                            End If
                        End If
                    End If

                    If filterAtDatabase Then
                        Dim countSqlString = "SELECT COUNT(*) AS recordsMatched FROM " & callRecordingTable

                        If xrefDestination Then
                            countSqlString &= " AS A LEFT JOIN " & cdrTableName & " AS B ON A.Callid = B.CallId"
                        End If

                        If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), countSqlString, myTable) Then
                            If myTable.Rows.Count > 0 Then
                                recordsMatched = myTable.Rows(0).Item("recordsMatched")
                                myTable.Rows.Clear()
                                myTable.Columns.Clear()
                            End If
                        End If
                    End If

                    '    selectString &= " ORDER BY " & timeStampName & " DESC"
                    mySQLStatement.AddOrderByString(timeStampName & " DESC")
                    Dim myCompositeSQLCommand As New CompositeSQLClass

                    If useTempCDRTable Then
                        Dim x As New SQLStatementClass

                        x.SetCreateTable(cdrTableName)
                        x.AddDeclaration("CallId", "INT")
                        x.AddDeclaration("DestinationNumber", "VARCHAR(80)")
                        x.AddDeclaration("TransferredToCallId", "INT")
                        myCompositeSQLCommand.AddCommand(x.GetSQLStatement)

                        x.Clear()
                        x.SetInsertIntoTable(cdrTableName)
                        x.AddSelectString("CallId", "")
                        x.AddSelectString("DestinationNumber", "")
                        x.AddSelectString("TransferredToCallId", "")
                        x.SetPrimaryTable(originalCDRTableName)
                        AddDateCondition(myStartDate, myEndDate, x, "StartTime")
                        myCompositeSQLCommand.AddCommand(x.GetSQLStatement)

                        myCompositeSQLCommand.AddCommand(mySQLStatement.GetSQLStatement)

                        myCompositeSQLCommand.DropTable("#tempCDR")
                    Else
                        myCompositeSQLCommand.AddCommand(mySQLStatement.GetSQLStatement)
                    End If

                    If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), myCompositeSQLCommand.GenerateCompositeCommand, myTable) Then
                        Dim numberOfRows As Integer = myTable.Rows.Count

                        If numberOfRows > 0 Then
                            ReDim myFileNameArray(numberOfRows - 1)
                            ReDim fileLengths(numberOfRows - 1)

                            For i = 0 To numberOfRows - 1
                                myFileNameArray(i) = myTable.Rows(i).Item("FileName").ToString

                                fileLengths(i) = -1
                                If Not myTable.Rows(i).Item("fileLength") Is DBNull.Value Then fileLengths(i) = myTable.Rows(i).Item("fileLength").ToString
                            Next

                            If archiveEnabled Then
                                ReDim statusArray(numberOfRows - 1)

                                For i = 0 To numberOfRows - 1
                                    statusArray(i) = 0

                                    With myTable.Rows(i)
                                        If .Item("status") IsNot DBNull.Value Then statusArray(i) = .Item("status")
                                    End With
                                Next
                            End If

                            If xrefDestination Then
                                ReDim realDestinations(numberOfRows - 1)

                                For i = 0 To numberOfRows - 1
                                    realDestinations(i) = ""

                                    With myTable.Rows(i)
                                        If .Item("InternalNumber") IsNot DBNull.Value Then
                                            If .Item("DestinationNumber") IsNot DBNull.Value Then
                                                If .Item("CallDirection") IsNot DBNull.Value Then
                                                    If .Item("CallDirection") = "In" Then
                                                        If Not .Item("InternalNumber") = .Item("DestinationNumber") Then
                                                            realDestinations(i) = .Item("DestinationNumber")
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End With
                                Next
                            End If

                            If xrefTransferredCalls Then
                                ReDim transferTable(numberOfRows - 1)

                                For i = 0 To numberOfRows - 1
                                    transferTable(i) = ""

                                    With myTable.Rows(i)
                                        If .Item("XferDest") IsNot DBNull.Value Then transferTable(i) = .Item("XferDest")
                                    End With
                                Next
                            End If
                        End If
                    End If
                Else
                    ' Not using database. If searching by callid, filter at file source
                    If providedCallId.Length > 0 Then
                        filter = "*" & GenerateCallIdString(providedCallId) & "*.wav"

                        If IsInteger(providedCallId) Then providedCallIdList.Add(CInt(providedCallId))
                    End If

                    If searchSubDirectories Then
                        Dim myFileList As New List(Of String)
                        Dim myFilter As String = "*.wav"

                        DigDirectory(recordingsPath, myFilter, providedCallId, myFileList)

                        If myFileList.Count > 0 Then
                            ReDim myFileNameArray(myFileList.Count - 1)

                            For i = 0 To myFileList.Count - 1
                                myFileNameArray(i) = myFileList(i)
                            Next
                        End If
                    Else
                        Try
                            myFileNameArray = IO.Directory.GetFiles(recordingsPath, filter)
                        Catch e As Exception
                            Logutil.LogError("GetFiles(" & recordingsPath & ", " & filter & ") failed", e.Message)
                        End Try

                        ' Were we trying to match on Call ID ?
                        If providedCallId.Length > 0 Then
                            ' Is Call ID aliasing enabled ?
                            If callIdAlias Then
                                Dim gotSomething As Boolean = False

                                If myFileNameArray IsNot Nothing Then
                                    If myFileNameArray.Length > 0 Then gotSomething = True
                                End If

                                If Not gotSomething Then
                                    ' See if we were transferred and can find some other Call IDs to try
                                    Dim cdrTableName As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.cdrTableName))

                                    myTable.Rows.Clear()
                                    myTable.Columns.Clear()

                                    If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), "SELECT TransferredCallId1 AS CallId1, TransferredCallId2 AS CallId2 FROM " & cdrTableName & " WHERE Callid = " & providedCallId, myTable) Then
                                        If myTable.Rows.Count = 1 Then
                                            Dim callId1 As Integer = 0
                                            Dim callId2 As Integer = 0

                                            If myTable.Rows(0).Item(0) IsNot DBNull.Value Then callId1 = myTable.Rows(0).Item(0)
                                            If myTable.Rows(0).Item(1) IsNot DBNull.Value Then callId2 = myTable.Rows(0).Item(1)

                                            providedCallIdList.Add(callId1)
                                            providedCallIdList.Add(callId2)

                                            Try
                                                myFileNameArray = IO.Directory.GetFiles(recordingsPath, "*" & GenerateCallIdString(callId1) & "*.wav")
                                            Catch e As Exception
                                                Logutil.LogError("GetFiles(" & recordingsPath & ", " & filter & ") failed", e.Message)
                                            End Try

                                            If myFileNameArray IsNot Nothing Then
                                                If myFileNameArray.Length > 0 Then gotSomething = True
                                            End If

                                            If Not gotSomething Then
                                                Try
                                                    myFileNameArray = IO.Directory.GetFiles(recordingsPath, "*" & GenerateCallIdString(callId2) & "*.wav")
                                                Catch e As Exception
                                                    Logutil.LogError("GetFiles(" & recordingsPath & ", " & filter & ") failed", e.Message)
                                                End Try
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                    ' The secondary path can only be flat, no subdirectories
                    If secondaryrecordingsPath.Length > 0 Then
                        Dim mySecondaryFileNameArray() As String = IO.Directory.GetFiles(secondaryrecordingsPath, filter)

                        If mySecondaryFileNameArray.Length > 0 Then
                            Dim myFileNameArrayLength As Integer = 0

                            If myFileNameArray IsNot Nothing Then myFileNameArrayLength = myFileNameArray.Length

                            ReDim Preserve myFileNameArray(myFileNameArrayLength + mySecondaryFileNameArray.Length - 1)
                            ReDim mySecondarySourceArray(myFileNameArrayLength + mySecondaryFileNameArray.Length - 1)

                            ' Append the file names from the secondary path to myFileNameArray
                            For i = 0 To mySecondaryFileNameArray.Length - 1
                                myFileNameArray(myFileNameArrayLength + i) = mySecondaryFileNameArray(i)
                            Next

                            ' Populate the flag array to show the source of each file
                            For i = 0 To myFileNameArrayLength - 1
                                mySecondarySourceArray(i) = False
                            Next

                            For i = 0 To mySecondaryFileNameArray.Length - 1
                                mySecondarySourceArray(myFileNameArrayLength + i) = True
                            Next

                            anySecondaryFiles = True
                        End If

                        mySecondaryFileNameArray = Nothing
                    End If
                End If ' If useDatabase

                CreateTable(myDataSet, "Records", myFieldNames)

                If myFileNameArray Is Nothing Then ReDim myFileNameArray(-1)

                If useDatabase Then
                    ' Create the filename for the URL
                    For i = 0 To myFileNameArray.Length - 1
                        myFileNameArray(i) = MapFilename(myFileNameArray(i), recordingsPath)
                    Next
                Else
                    Dim myKeyArray(myFileNameArray.Length - 1) As String
                    Dim logErrorOnFilenameFormat As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.logFilenameFormatErrors))

                    ' Sort the files into reverse date and time order (newest first)
                    For i = 0 To myFileNameArray.Length - 1
                        Dim hashCount As Integer = 0
                        Dim j As Integer
                        Dim lastIndex As Integer = myFileNameArray(i).LastIndexOf("\")
                        Dim pathWithSlash As String = recordingsPath.ToLower & "\"
                        Dim myIndex As Integer = 0

                        ' Either way, extract the raw filename with no path, and no ".wav" extension
                        If myFileNameArray(i).ToLower.StartsWith(pathWithSlash) Then
                            myFileNameArray(i) = myFileNameArray(i).Substring(pathWithSlash.Length, myFileNameArray(i).Length - (pathWithSlash.Length + ".wav".Length))
                        Else
                            If secondaryrecordingsPath.Length > 0 Then
                                pathWithSlash = secondaryrecordingsPath.ToLower & "\"

                                If myFileNameArray(i).ToLower.StartsWith(pathWithSlash) Then
                                    myFileNameArray(i) = myFileNameArray(i).Substring(pathWithSlash.Length, myFileNameArray(i).Length - (pathWithSlash.Length + ".wav".Length))
                                Else
                                    myFileNameArray(i) = myFileNameArray(i).Substring(lastIndex + 1, myFileNameArray(i).Length - (lastIndex + 1 + ".wav".Length))
                                End If
                            Else
                                myFileNameArray(i) = myFileNameArray(i).Substring(lastIndex + 1, myFileNameArray(i).Length - (lastIndex + 1 + ".wav".Length))
                            End If
                        End If

                        myKeyArray(i) = ""

                        ' There could still be some subdirectory stuff here.
                        If myFileNameArray(i).Contains("\") Then myIndex = myFileNameArray(i).LastIndexOf("\") + 1

                        ' Check filename formatFileNameArray(i).Substring(myindex)
                        For j = myIndex To myFileNameArray(i).Length - 1
                            If myFileNameArray(i)(j) = "#" Then
                                hashCount += 1

                                If hashCount = dateIndex Then
                                    If j + 1 >= myFileNameArray(i).Length Then
                                        If logErrorOnFilenameFormat Then LogError("CallRecordingGUI::HandleRequest()", "Illegal index for myFileNameArray(" & i & ").Substring(" & j + 1 & ", " & DATE_AND_TIME_LENGTH & ") value of " & myFileNameArray(i))
                                    Else
                                        If j + DATE_AND_TIME_LENGTH >= myFileNameArray(i).Length Then
                                            If logErrorOnFilenameFormat Then LogError("CallRecordingGUI::HandleRequest()", "Illegal index for myFileNameArray(" & i & ").Substring(" & j + 1 & ", " & DATE_AND_TIME_LENGTH & ") value of " & myFileNameArray(i))
                                        Else
                                            myKeyArray(i) = myFileNameArray(i).Substring(j + 1, DATE_AND_TIME_LENGTH)
                                        End If
                                    End If

                                    Exit For
                                End If
                            End If
                        Next
                    Next

                    ' Sort by date and time, most recent first
                    If anySecondaryFiles Then
                        Dim myKeyArrayCopy(myKeyArray.Length - 1) As String

                        Array.Copy(myKeyArray, myKeyArrayCopy, myKeyArray.Length)

                        Array.Sort(myKeyArray, myFileNameArray)
                        Array.Reverse(myFileNameArray)
                        Array.Sort(myKeyArrayCopy, mySecondarySourceArray)
                        Array.Reverse(mySecondarySourceArray)
                    Else
                        Array.Sort(myKeyArray, myFileNameArray)
                        Array.Reverse(myFileNameArray)
                    End If
                End If ' If useDatabase ...
            End If ' If filterAtDatabase ..

            For i = 0 To myFileNameArray.Length - 1
                Dim myIndex As Integer = 0
                Dim displayThisFile As Boolean = True
                Dim direction As String = ""
                Dim internalNumber As String = ""
                Dim externalNumber As String = ""
                Dim dateStamp As String = ""
                Dim timeStamp As String = ""
                Dim callId As String = ""
                Dim displayedInternalNumber As String = ""

                If useDatabase Then
                    ' Take the field values directly from the data table
                    Dim dateAndTime As New DateAndTimeClass

                    direction = SafeDBRead(myTable, i, "callDirection")
                    internalNumber = SafeDBRead(myTable, i, "internalNumber")
                    externalNumber = SafeDBRead(myTable, i, "externalNumber")
                    callId = SafeDBRead(myTable, i, "callId")

                    If Not myTable.Rows(i).Item("timeStamp") Is DBNull.Value Then
                        dateAndTime.SetFromDBReadString(myTable.Rows(i).Item("timeStamp"))
                        dateStamp = dateAndTime.AsCallRecordingFilenameDate
                        timeStamp = dateAndTime.AsCallRecordingFilenameTime
                    End If
                Else
                    If myFileNameArray(i).Contains("\") Then myIndex = myFileNameArray(i).LastIndexOf("\") + 1

                    Dim myFilenameFields() As String = myFileNameArray(i).Substring(myIndex).Split("#")

                    ' Extract the fields
                    If callDirectionIndex >= 0 And callDirectionIndex < myFilenameFields.Count Then direction = myFilenameFields(callDirectionIndex)
                    If internalAddressIndex >= 0 And internalAddressIndex < myFilenameFields.Count Then internalNumber = myFilenameFields(internalAddressIndex)
                    If remoteAddressIndex >= 0 And remoteAddressIndex < myFilenameFields.Count Then externalNumber = myFilenameFields(remoteAddressIndex)
                    If dateIndex >= 0 And dateIndex < myFilenameFields.Count Then dateStamp = myFilenameFields(dateIndex)
                    If timeIndex >= 0 And timeIndex < myFilenameFields.Count Then timeStamp = myFilenameFields(timeIndex)
                    If callIdIndex >= 0 And callIdIndex < myFilenameFields.Count Then callId = myFilenameFields(callIdIndex).TrimStart("0")
                End If

                ' Interpret the values of each field from the filename
                '   If IsInteger(externalNumber) Then
                'If externalNumber.Length > 2 Then
                'If externalNumber.StartsWith("90") Then
                'externalNumber = externalNumber.Substring(1)
                ' End If

                'If Not externalNumber.StartsWith("0") Then
                'externalNumber = "0" & externalNumber
                'End If
                'End If
                ' End If

                If externalNumber.ToLower.StartsWith("withheld") Then externalNumber = "WITHHELD"

                '    If internalNumber.ToLower.StartsWith("sip") Then
                'Dim x As Integer = internalNumber.IndexOf("@")

                '    If x >= 7 Then
                'internalNumber = internalNumber.Substring(x - 4, 4)
                '   End If
                '   End If

                '      If internalNumber.Length = 14 Then
                'If internalNumber.StartsWith("9000117223") Or internalNumber.StartsWith("9000292143") Then
                '  internalNumber = internalNumber.Substring(10)
                '    End If
                '  End If

                'If internalNumber = "2130" Or internalNumber = "2142" Then displayThisFile = False

                ' For FP Mailing
                If False Then
                    If internalNumber.Length > 0 Then
                        If externalNumber.Length > 0 Then
                            Dim swapIt As Boolean = False

                            If externalNumber.Length = 6 Then
                                If externalNumber.StartsWith("60") Then
                                    swapIt = True
                                End If
                            End If

                            If swapIt Then
                                Dim x As String = internalNumber

                                If externalNumber.Length >= 4 Then internalNumber = externalNumber.Substring(externalNumber.Length - 4)
                                externalNumber = x

                                If externalNumber.Length > 6 Then
                                    If Not externalNumber.StartsWith("0") Then externalNumber = "0" & externalNumber
                                End If

                                direction = "In"
                            Else
                                If internalNumber.Length >= 4 Then internalNumber = internalNumber.Substring(internalNumber.Length - 4)
                            End If
                        End If
                    Else
                        If externalNumber.Length > 0 Then
                            Dim swapIt As Boolean = False

                            ' Call could be O/G with no CLI, or I/C with no CLI
                            If externalNumber.Length = 6 Then
                                If externalNumber.StartsWith("60") Then
                                    swapIt = True
                                End If
                            End If

                            If swapIt Then
                                Dim x As String = internalNumber

                                If externalNumber.Length >= 4 Then internalNumber = externalNumber.Substring(externalNumber.Length - 4)
                                externalNumber = x
                                direction = "In"
                            End If
                        End If
                    End If
                End If
                ' For CTS
                '            If externalNumber.Length > 4 Then
                ' If externalNumber.StartsWith("0+44") Then
                'externalNumber = "0" & externalNumber.Substring(4)
                'End If
                'End If

                direction = MapDirection(direction)

                displayedInternalNumber = internalNumber

                If xrefDestination Then
                    If realDestinations(i).Length > 0 Then displayedInternalNumber &= " >> " & realDestinations(i)
                End If

                If xrefTransferredCalls Then
                    If transferTable(i).Length > 0 Then displayedInternalNumber &= " >> " & transferTable(i)
                End If

                If Not useDatabase Then
                    ' Start the comparisons
                    If providedCallId.Length > 0 Then
                        ' CallID has been provided, use this as the only comparison as it is absolute
                        displayThisFile = False

                        For j = 0 To providedCallIdList.Count - 1
                            If providedCallIdList(j) = callId Then displayThisFile = True
                        Next
                    End If

                    If myStartDate.Length > 0 Then
                        If StrComp(dateStamp, myStartDate) = -1 Then displayThisFile = False
                    End If

                    If displayThisFile Then
                        If myEndDate.Length > 0 Then
                            If StrComp(dateStamp, myEndDate) = 1 Then displayThisFile = False
                        End If

                        If displayThisFile Then
                            If myAspect.extensionTextBox.Text.Length > 0 Then
                                If Not myAspect.extensionTextBox.Text = internalNumber Then
                                    displayThisFile = False

                                    If xrefDestination Then
                                        If callId.Length > 0 Then
                                            Dim cdrTable As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.cdrTableName))
                                            Dim command As String = "SELECT DestinationNumber FROM " & cdrTable & " WHERE CallId = " & callId

                                            myTable.Rows.Clear()
                                            myTable.Columns.Clear()

                                            If FillTableFromCommand(CreateConnectionString(callRecordingInterfaceConfigDictionary), command, myTable) Then
                                                If myTable.Rows.Count > 0 Then
                                                    If myTable.Rows(0).Item(0).ToString = myAspect.extensionTextBox.Text Then
                                                        displayThisFile = True
                                                        displayedInternalNumber = myAspect.extensionTextBox.Text
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            If displayThisFile Then
                                If myAspect.externalNumberTextBox.Text.Length > 0 Then
                                    Dim x As String = myAspect.externalNumberTextBox.Text

                                    If x.Contains("*") Then
                                        If x.StartsWith("*") Then
                                            If x.EndsWith("*") Then
                                                ' x = "*abc*"
                                                If Not externalNumber.Contains(x.Trim("*")) Then displayThisFile = False
                                            Else
                                                ' x = "*abc"
                                                If Not externalNumber.EndsWith(x.Trim("*")) Then displayThisFile = False
                                            End If
                                        Else
                                            If x.EndsWith("*") Then
                                                ' x = "abc*"
                                                If Not externalNumber.StartsWith(x.Trim("*")) Then displayThisFile = False
                                            Else
                                                ' x = "a*bc"
                                                If Not (externalNumber.StartsWith(x.Substring(0, x.IndexOf("*"))) And externalNumber.EndsWith(x.Substring(x.IndexOf("*") + 1))) Then displayThisFile = False
                                            End If
                                        End If
                                    Else
                                        If Not externalNumber = x Then displayThisFile = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If displayThisFile Then
                    ' Any security to check ?
                    If myDataList.Count > 0 Then
                        Dim numberToMatchOn As String = internalNumber

                        displayThisFile = False

                        If xrefDestination Then
                            If realDestinations(i).Length > 0 Then numberToMatchOn = realDestinations(i)

                            If xrefTransferredCalls Then
                                If transferTable(i).Length > 0 Then numberToMatchOn = transferTable(i)
                            End If
                        End If

                        ' If the extension number is in the list then allow the recording
                        For j = 0 To myDataList.Count - 1
                            If myDataList(j) = numberToMatchOn Then
                                displayThisFile = True
                                Exit For
                            End If
                        Next
                    End If
                End If

                If displayThisFile Then
                    If Not filterAtDatabase Then recordsMatched += 1

                    If recordsMatched > resultOffset Then
                        If recordsReturned < maxRecordsToReturn Then
                            Dim myFileLength As Integer = -1
                            Dim myUseSecondaryFile As Boolean = False
                            Dim thisFileIsArchived As Boolean = False

                            If fileLengths IsNot Nothing Then myFileLength = fileLengths(i)

                            If anySecondaryFiles Then myUseSecondaryFile = mySecondarySourceArray(i)

                            If archiveEnabled Then thisFileIsArchived = statusArray(i)

                            AddRow(myDataSet, thisFileIsArchived, myUseSecondaryFile, useStructuredFileSystem, dateStamp, timeStamp, myIndex, myFileNameArray(i), myFileLength, displayedInternalNumber, externalNumber, direction, callId)
                            recordsReturned += 1
                        End If
                    End If
                End If
            Next
        End If

        myRepeater.DataSource = myDataSet
        myRepeater.DataBind()

        recordsLabel.Font.Bold = True

        ' Check if we can display all the records in one go
        If recordsMatched <= maxRecordsToReturn Then
            ' Yes we can. No need for Nav buttons
            Dim myString As String = recordsMatched & " Recording"

            If Not recordsMatched = 1 Then myString &= "s"

            recordsLabel.Text = myString & " Found"
            nextButton.Visible = False
            prevButton.Visible = False
        Else
            ' No we cannot. Setup Nav buttons
            Dim max As Integer = maxRecordsToReturn + resultOffset

            If recordsMatched <= max Then
                max = recordsMatched
                nextButton.Visible = False
            Else
                nextButton.Text = "Next " & maxRecordsToReturn
                nextButton.Visible = True
            End If

            Dim myString As String = "Displaying " & (resultOffset + 1) & " .. " & max & " of " & recordsMatched & " Recordings Found"

            recordsLabel.Text = myString

            If resultOffset > 0 Then
                prevButton.Text = "Previous " & maxRecordsToReturn
                prevButton.Visible = True
            Else
                prevButton.Visible = False
            End If
        End If

        f_offset.Text = resultOffset

        If recordsMatched = 0 Then
            'resultsTable.Visible = False
            myAspect.summaryPanel.Visible = True
            myRepeater.Visible = False
        Else
            'resultsTable.Visible = True
            myAspect.summaryPanel.Visible = True
            myRepeater.Visible = True
        End If
    End Sub

    Private Sub AddRow(ByRef p_dataSet As DataSet, ByVal p_thisFileIsArchived As Boolean, ByVal p_useSecondaryFile As Boolean, ByVal p_useStructuredFilesystem As Boolean, ByRef p_dateStamp As String, ByRef p_timeStamp As String, ByVal p_myIndex As Integer, ByRef p_filename As String, ByVal p_fileLength As Integer, ByRef p_displayedInternalNumber As String, ByRef p_externalNumber As String, ByRef p_direction As String, ByVal p_callid As Integer)
        Dim extra As String = ""
        Dim myPath As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.virtualPath))
        Dim thisFileArchived As Boolean = False
        Dim myDateAndTime As New DateAndTimeClass
        Dim myFields(myFieldNames.Length - 1) As String
        Dim duration As String = "--"
        Dim recordingRate As Integer = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.recordingRateBytesPerSecond))

        If p_thisFileIsArchived Then
            Dim archiveVirtualPath As String = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.archiveVirtualPath))

            If archiveVirtualPath.Length > 0 Then myPath = archiveVirtualPath

            thisFileArchived = True
        End If

        If p_useSecondaryFile Then myPath = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.secondaryHTMLRecordingsPath))

        If p_useStructuredFilesystem Or thisFileArchived Then
            extra = p_dateStamp.Substring(0, 4) & "/" & p_dateStamp.Substring(4, 2) & "/" & p_dateStamp.Substring(6, 2) & "/"
        Else
            If p_myIndex > 0 Then
                extra = BackSlashToForwardSlash(p_filename.Substring(0, p_myIndex))
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
        myFields(4) = p_direction
        myFields(5) = duration
        myFields(6) = p_callid
        Dim x As String = myPath & extra & PlusToSpace(Server.UrlEncode(p_filename.Substring(p_myIndex))) & ".wav"
        'myFields(7) = "<a href=" & WrapInQuotes(x) & " onclick=" & WrapInQuotes("window.open ('http://localhost/SMSHandler/SMSHandler.aspx?client=" & f_loginExtension.Text & "-" & x & "&billingStatus=2')") & ">Listen</a>"

        If FOR_ALDERMORE Then
            If FOR_NEW_ALDERMORE Then
                myFields(7) = "<a href=" & WrapInQuotes("https://STPV-SWYX-REA01/SecureRecordingAccess/SecureRecordingAccess.aspx?userName=" & myAspect.userNameTextBox.Text & "&requestedRecording=" & x) & ">Access Recording</a>"
            Else
                'myFields(7) = "<a href=" & WrapInQuotes("http://z2swma01/SecureRecordingAccess/SecureRecordingAccess.aspx?userName=" & f_loginExtension.Text & "&requestedRecording=" & x) & ">Access Recording</a>"
                myFields(7) = "<a href=" & WrapInQuotes("http://localhost/SecureRecordingAccess/SecureRecordingAccess.aspx?userName=" & myAspect.userNameTextBox.Text & "&requestedRecording=" & x) & ">Access Recording</a>"
            End If
        Else
            myFields(7) = "<a href=" & WrapInQuotes(x) & ">Listen</a>"
        End If

        AddRow(p_dataSet, "Records", myFields)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Target.SetTarget(TargetType.RECORDING_WEB_INTERFACE)
        CallRecordingInterfaceInitialiseConfig()

        Dim drillDownBypassesLogin As Boolean = callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.drillDownBypassesLogin))

        usingNewInterface = callRecordingInterfaceConfigDictionary.GetBooleanItem("useNewInterface")

        With myAspect
            If usingNewInterface Then
                .loginPanel = loginPanelNewInterface
                .userNameTextBox = userNameNewInterface
                .passwordTextBox = passwordNewInterface
                .styleSheet = "ReachAllCallRecordingStyleSheetNewInterface.css"
                .reportSelectionPanel = newInterfaceSearchPanel
                .callTypeList = directionDropDownListNewInterface
                .summaryPanel = summaryPanelNewInterface
                .extensionTextBox = extensionTextBoxNewInterface
                .externalNumberTextBox = externalNumberTextBoxNewInterface
                .callIdtextBox = callIdTextBoxNewInterface
                .startDateCalendar = calendar3
                .endDateCalendar = calendar4
                .startTimeDropDownList = startTimeNewInterfaceDropDownList
                .endTimeDropDownList = endTimeNewInterfaceDropDownList
                .ignoreTimeCheckbox = ignoreTimeNewInterfaceCheckBox
                .startTimeLabel = startTimeNewInterfaceLabel
                .endTimeLabel = endTimeNewInterfaceLabel
                .changePasswordButton = changePasswordButtonNewInterface
                .changePasswordResponseLabel = changePasswordResponseNewInterface
                .newPasswordTextBox = newPasswordTextBoxNewInterface
                .newPasswordAgainTextBox = newPasswordAgainTextBoxNewInterface
                .timeSpanDropDownList = timeSpanNewInterfaceDropDownList
            Else
                .loginPanel = loginPanel
                .userNameTextBox = f_userName
                .passwordTextBox = f_password
                .styleSheet = "ReachAllCallRecordingStyleSheet.css"
                .reportSelectionPanel = searchPanel
                .callTypeList = callTypeList
                .summaryPanel = summaryPanelNewInterface
                .extensionTextBox = f_extension
                .externalNumberTextBox = f_destination
                .callIdtextBox = f_callId
                .startDateCalendar = calendar1
                .endDateCalendar = calendar2
                .startTimeDropDownList = startTime
                .endTimeDropDownList = endTime
                .changePasswordButton = changePasswordButton
                .changePasswordResponseLabel = changePasswordResponseLegacy
                .newPasswordTextBox = newPasswordTextBox
                .newPasswordAgainTextBox = newPasswordAgainTextBox
            End If

            styleLiteral.Text = CreateStyleSheetLink(.styleSheet)
        End With

        If UsingNewSecurity(True) Then
            myAspect.changePasswordButton.Visible = True
        End If

        If Request.QueryString.Count > 0 Then
            loginPanel.Visible = False
            loginPanelNewInterface.Visible = False
            searchPanel.Visible = False
            newInterfaceSearchPanel.Visible = False

            If usingNewInterface Then summaryPanelNewInterface.Style.Add("margin-top", "100px")

            If drillDownBypassesLogin Then
                Dim parmDictionary As New Dictionary(Of String, String)

                If ALLOW_REQUEST_PARMS Then
                    For Each myName In Request.QueryString
                        If myName Is Nothing Then
                            parmDictionary.Add("callId", Request.QueryString(0))
                        Else
                            parmDictionary.Add(myName, Request(myName))
                        End If
                    Next
                Else
                    'HandleRequest(0, Request.QueryString(0))
                    parmDictionary.Add("callId", Request.QueryString(0))
                End If

                loginPanel.Visible = False
                HandleRequest(0, parmDictionary)
            Else
                f_parm.Text = Request.QueryString(0)
            End If
        Else
            f_parm.Text = ""

            If Not Page.IsPostBack Then
                ' This is the first page load, set everything up
                loginPanel.Visible = False
                loginPanelNewInterface.Visible = False
                searchPanel.Visible = False
                newInterfaceSearchPanel.Visible = False
                legacyCheckBox.Visible = USE_LEGACY_DATA

                With myAspect
                    .loginPanel.Visible = True
                    .summaryPanel.Visible = False
                End With

                If usingNewInterface Then
                    For i = 0 To timeSpans.Length - 1
                        myAspect.timeSpanDropDownList.Items.Add(timeSpans(i))
                    Next

                    myAspect.timeSpanDropDownList.SelectedIndex = 8
                End If
            End If
        End If
    End Sub

    Sub ChangePasswordPressed(ByVal Source As Object, ByVal e As EventArgs)
        If myAspect.userNameTextBox.Text = "Admin" Then
            myAspect.changePasswordResponseLabel.Text = "Cannot change Admin password"
        Else
            myAspect.changePasswordResponseLabel.Text = ""
            CallRecordingInterfaceInitialiseConfig()
            LoadAdminData(GetAdminDataFilename())

            If CheckLogin() Then
                If usingNewInterface Then myAspect.changePasswordButton.Visible = False

                ShowChangePasswordRows(True)
            End If
        End If
    End Sub

    Sub ShowChangePasswordRows(ByVal p As Boolean)
        If usingNewInterface Then
            changePasswordPanelNewInterface.Visible = p
        Else
            cpRow_1.Visible = p
            cpRow_2.Visible = p
            cpRow_3.Visible = p
        End If
    End Sub

    Sub ConfirmChangePasswordPressed(ByVal Source As Object, ByVal e As EventArgs)
        CallRecordingInterfaceInitialiseConfig()
        LoadAdminData(GetAdminDataFilename())

        Dim myIndex As Integer = securityList.GetIndexOfUser(myAspect.userNameTextBox.Text)
        Dim allOk As Boolean = False
        Dim myResponse As String = ""

        If myIndex >= 0 Then
            Dim currentPassword As String = securityList.GetRecordingsPasswordFromIndex(myIndex)

            If currentPassword = "" Then currentPassword = securityList.GetPasswordFromIndex(myIndex)

            If IsNewPasswordOK(currentPassword, myAspect.newPasswordTextBox.Text, myAspect.newPasswordAgainTextBox.Text, myResponse) Then
                If NewSetCallRecordingPassword(myAspect.userNameTextBox.Text, myAspect.newPasswordTextBox.Text) Then
                    NewMasterSaveSecuritySettings()
                    myResponse = "Password updated successfully - please login with new password"
                    allOk = True
                Else
                    myResponse = "Password could not be updated"
                End If
            End If
        Else
            myResponse = "Error in locating user"
        End If

        ShowChangePasswordRows(False)
        myAspect.changePasswordResponseLabel.Text = myResponse

        If allOk Then myAspect.changePasswordButton.Visible = False
    End Sub

    Protected Sub startDateCalendar_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles calendar1.SelectionChanged, calendar3.SelectionChanged
        f_startDate.Text = myAspect.startDateCalendar.SelectedDate
    End Sub

    Protected Sub endDateCalendar_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles calendar2.SelectionChanged, calendar4.SelectionChanged
        f_endDate.Text = myAspect.endDateCalendar.SelectedDate
    End Sub

    Sub PrevClicked(ByVal Source As Object, ByVal e As EventArgs)
        Dim myResultsOffset As Integer = f_offset.Text - CInt(callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.maxRecordsToReturn)))

        HandleRequest(myResultsOffset)
    End Sub

    Sub NextClicked(ByVal Source As Object, ByVal e As EventArgs)
        Dim myResultsOffset As Integer = f_offset.Text + CInt(callRecordingInterfaceConfigDictionary.GetItem([Enum].GetName(GetType(CallRecordingInterfaceConfigItems), CallRecordingInterfaceConfigItems.maxRecordsToReturn)))

        HandleRequest(myResultsOffset)
    End Sub

    Sub AdminButtonPressed(ByVal Source As Object, ByVal e As EventArgs)
        adminTable.Visible = True
        adminChoiceTable.Visible = False
        myAspect.callIdtextBox.Text = f_parm.Text

        myAspect.userNameTextBox.Text = ""
        myAspect.passwordTextBox.Text = ""
        LoadAdminData(GetAdminDataFilename())
        administratorsLit0.Text = GenerateAdministratorsUsernameTable()
        administratorsLit1.Text = GenerateAdministratorsPasswordnameTable()
        supervisorsLit0.Text = GenerateSupervisorsUsernameTable()
        supervisorsLit1.Text = GenerateSupervisorsPasswordnameTable()
        supervisorsLit2.Text = GenerateSupervisorsExtensionTable()
    End Sub

    Sub AdminOKButtonPressed(ByVal Source As Object, ByVal e As EventArgs)
        adminTable.Visible = False
        adminChoiceTable.Visible = True
    End Sub

    Sub AddAdministratorButtonPressed(ByVal Source As Object, ByVal e As EventArgs)
        If AdminAddHandler(administratorsUsernameText.Text, administratorsPasswordText.Text, "", myAdministrators) Then
            administratorsLit0.Text = GenerateAdministratorsUsernameTable()
            administratorsLit1.Text = GenerateAdministratorsPasswordnameTable()

            administratorsUsernameText.Text = ""
            administratorsPasswordText.Text = ""
        End If
    End Sub

    Sub AddSupervisorButtonPressed(ByVal Source As Object, ByVal e As EventArgs)
        If AdminAddHandler(supervisorsUsernameText.Text, supervisorsPasswordText.Text, supervisorsDataText.Text, mySupervisors) Then
            supervisorsLit0.Text = GenerateSupervisorsUsernameTable()
            supervisorsLit1.Text = GenerateSupervisorsPasswordnameTable()
            supervisorsLit2.Text = GenerateSupervisorsExtensionTable()

            supervisorsUsernameText.Text = ""
            supervisorsPasswordText.Text = ""
            supervisorsDataText.Text = ""
        End If
    End Sub

    Function AdminAddHandler(ByRef p_userName As String, ByRef p_password As String, ByRef p_data As String, ByRef p_list As List(Of LoginDetailsClass)) As Boolean
        Dim saveData As Boolean = False

        ' Username present ?
        If p_userName.Length > 0 Then
            LoadAdminData(GetAdminDataFilename())

            ' Is user already in list ?
            Dim myIndex As Integer = -1

            For i = 0 To p_list.Count - 1
                If p_list(i).userName = p_userName Then
                    myIndex = i
                    Exit For
                End If
            Next

            If myIndex = -1 Then
                ' User is not already present. Add new user if password is present
                If p_password.Length > 0 Then
                    Dim x As New LoginDetailsClass

                    x.userName = p_userName
                    x.password = p_password
                    x.data = p_data
                    p_list.Add(x)
                    saveData = True
                End If
            Else
                ' User is already present, is password present ?
                If p_password.Length > 0 Then
                    ' Yes. Only write if password or data has changed
                    If p_list(myIndex).password <> p_password Or p_list(myIndex).data <> p_data Then
                        ' Change password and data
                        p_list(myIndex).password = p_password
                        p_list(myIndex).data = p_data
                        saveData = True
                    End If
                Else
                    ' No password. Delete this user
                    p_list.RemoveAt(myIndex)
                    saveData = True
                End If
            End If

            If saveData Then
                SaveAdminData(GetAdminDataFilename())
                administratorsUsernameText.Text = ""
                administratorsPasswordText.Text = ""
                supervisorsUsernameText.Text = ""
                supervisorsPasswordText.Text = ""
            End If
        End If

        Return saveData
    End Function

    Sub AddLine(ByRef p_source As String, ByRef p_text As String)
        p_source &= p_text & vbCrLf
    End Sub

    Function GenerateAdministratorsUsernameTable() As String
        Return GenerateLiteral(myAdministrators, 0)
    End Function

    Function GenerateAdministratorsPasswordnameTable() As String
        Return GenerateLiteral(myAdministrators, 1)
    End Function

    Function GenerateSupervisorsUsernameTable() As String
        Return GenerateLiteral(mySupervisors, 0)
    End Function

    Function GenerateSupervisorsPasswordnameTable() As String
        Return GenerateLiteral(mySupervisors, 1)
    End Function

    Function GenerateSupervisorsExtensionTable() As String
        Return GenerateLiteral(mySupervisors, 2)
    End Function

    Function GenerateLiteral(ByRef p As List(Of LoginDetailsClass), ByVal p_columnIndex As Integer) As String
        Dim x As String = ""
        Dim headings() As String = {"Usernames", "Passwords", "Extensions"}

        AddLine(x, "<table border=1 bordercolor=" & WrapInQuotes("#ffffff") & ">")
        AddLine(x, "<tr>")
        AddLine(x, "<th width=140>" & headings(p_columnIndex) & "</th>")
        AddLine(x, "</tr>")

        For i = 0 To p.Count - 1
            Dim myData As String = p(i).userName

            If p_columnIndex = 1 Then myData = p(i).password

            If p_columnIndex = 2 Then
                myData = p(i).data

                If myData = "" Then myData = "---"
            End If

            AddLine(x, "<tr align=center>")
            AddLine(x, "<td>" & myData & "</td>")
            AddLine(x, "</tr>")
        Next

        If p.Count = 0 Then
            AddLine(x, "<tr align=center><td>---</td></tr>")
        End If

        AddLine(x, "</table>")

        Return x
    End Function

    Private Function GenerateCallIdString(ByVal p As Integer) As String
        Dim result As String = p.ToString

        While result.Length < CALLID_STRING_LENGTH
            result = "0" & result
        End While

        Return result
    End Function

    Sub timeSpanIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        With myAspect
            Select Case .timeSpanDropDownList.SelectedIndex
                Case 0
                    ' This hour
                    .ignoreTimeCheckbox.Checked = False
                    .startTimeDropDownList.SelectedIndex = Now.Hour * 6
                    .endTimeDropDownList.SelectedIndex = .startTimeDropDownList.SelectedIndex + 5
                    .startDateCalendar.SelectedDate = Today
                    .endDateCalendar.SelectedDate = .startDateCalendar.SelectedDate

                Case 1
                    ' Previous hour
                    .ignoreTimeCheckbox.Checked = False
                    .endTimeDropDownList.SelectedIndex = (((Now.Hour - 1) Mod 24) * 6) + 5
                    .startTimeDropDownList.SelectedIndex = .endTimeDropDownList.SelectedIndex - 5
                    .startDateCalendar.SelectedDate = Today
                    .endDateCalendar.SelectedDate = .startDateCalendar.SelectedDate

                Case 2
                    ' Today
                    .ignoreTimeCheckbox.Checked = True
                    .startDateCalendar.SelectedDate = Today
                    .endDateCalendar.SelectedDate = .startDateCalendar.SelectedDate

                Case 3
                    ' Yesterday
                    .ignoreTimeCheckbox.Checked = True
                    .startDateCalendar.SelectedDate = Today.AddDays(-1)
                    .endDateCalendar.SelectedDate = .startDateCalendar.SelectedDate

                Case 4
                    ' Last 7 days
                    .ignoreTimeCheckbox.Checked = True
                    .startDateCalendar.SelectedDate = Today.AddDays(-7)
                    .endDateCalendar.SelectedDate = .startDateCalendar.SelectedDate.AddDays(6)

                Case 5
                    ' Last week (Mon - Sun)
                    .ignoreTimeCheckbox.Checked = True
                    .startDateCalendar.SelectedDate = Today.AddDays(-(6 + CInt(Today.DayOfWeek)))
                    .endDateCalendar.SelectedDate = .startDateCalendar.SelectedDate.AddDays(6)

                Case 6
                    ' Last 30 days
                    .ignoreTimeCheckbox.Checked = True
                    .startDateCalendar.SelectedDate = Today.AddDays(-30)
                    .endDateCalendar.SelectedDate = .startDateCalendar.SelectedDate.AddDays(29)

                Case 7
                    ' Last month
                    .ignoreTimeCheckbox.Checked = True
                    .startDateCalendar.SelectedDate = Today.AddDays(-Today.Day)
                    .endDateCalendar.SelectedDate = .endDateCalendar.SelectedDate.AddDays(1 - Date.DaysInMonth(.endDateCalendar.SelectedDate.Year, .endDateCalendar.SelectedDate.Month))

                Case 8
                    .ignoreTimeCheckbox.Checked = True
                    .startTimeDropDownList.SelectedIndex = Now.Hour
                    .endTimeDropDownList.SelectedIndex = (.startTimeDropDownList.SelectedIndex + 1) Mod 24
                    .startDateCalendar.SelectedDate = New DateTime()
                    .endDateCalendar.SelectedDate = New DateTime()

                    .startTimeDropDownList.SelectedIndex = 0
                    .endTimeDropDownList.SelectedIndex = .endTimeDropDownList.Items.Count - 1

            End Select
        End With

        IgnoreTimeCheckBoxChanged_Handler()
    End Sub

    Sub IgnoreTimeCheckBoxChanged_Handler()
        Dim flag As Boolean = True

        With myAspect
            If .ignoreTimeCheckbox.Checked Then flag = False

            .startTimeDropDownList.Enabled = flag
            .endTimeDropDownList.Enabled = flag
            .startTimeLabel.Enabled = flag
            .endTimeLabel.Enabled = flag
        End With
    End Sub
End Class