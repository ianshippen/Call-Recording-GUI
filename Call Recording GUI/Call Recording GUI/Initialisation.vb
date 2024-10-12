Imports System.Xml

Module Initialisation
    Enum ConfigItems
        recordingsPath
        virtualPath
        archiveRecordingsPath
        archiveVirtualPath
        archiveEnabled
        secondaryRecordingsPath
        secondaryHTMLRecordingsPath
        maxRecordsToReturn
        useDatabase
        useStructuredFileSystem
        callRecordingTable
        timeStampName
        filterAtDatabase
        cdrTableName
        dataSource
        databaseName
        databaseUserId
        databasePassword
        searchSubDirectories
        security
        xrefDestination
        xrefTransferredCalls
        useNewDateSearch
        callIdAlias
        drillDownBypassesLogin
        useTempCDRTable
        callRecordingFilenameFormat
        logFilenameFormatErrors
        adminPassword
    End Enum

    Const DEFAULT_RECORDINGS_PATH As String = ""
    Const DEFAULT_VIRTUAL_PATH As String = ""
    Const DEFAULT_MAX_RECORDS_TO_RETURN As Integer = 100
    Const DEFAULT_USE_STRUCTURED_FILE_SYSTEM As Boolean = False
    Const DEFAULT_CALL_RECORDING_TABLE As String = "CallRecordingsTable"
    Const DEFAULT_TIMESTAMP_NAME As String = "timeStamp"
    Const DEFAULT_FILTER_AT_DATABASE As String = "False"
    Const DEFAULT_CDR_TABLE_NAME As String = "IpPbxCDR"
    Const DEFAULT_DATA_SOURCE As String = "."
    Const DEFAULT_DATABASE_NAME As String = ""
    Const DEFAULT_USER_ID As String = ""
    Const DEFAULT_PASSWORD As String = ""
    Const DEFAULT_SEARCH_SUB_DIRECTORIES As Boolean = False
    Const DEFAULT_SECURITY As Boolean = False
    Const DEFAULT_XREF_DESTINATION As Boolean = False
    Const DEFAULT_CALL_RECORDING_FILENAME_FORMAT As String = "[calldirection]#[internaladdress]#[remotename]#[remoteaddress]#[date]#[time]"
    Const CONFIG_FILENAME As String = "Config.xml"

    Public configDictionary As New DictionaryClass("configDictionary")

    Public Sub Initialise()
        InitLogutilConfig()
        InitConfig()
        LoadConfig()
    End Sub

    Private Sub InitConfig()
        configDictionary.Clear()

        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.recordingsPath), DEFAULT_RECORDINGS_PATH)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.virtualPath), DEFAULT_VIRTUAL_PATH)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.archiveRecordingsPath), "")
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.archiveVirtualPath), "")
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.archiveEnabled), False)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.secondaryRecordingsPath), DEFAULT_RECORDINGS_PATH)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.secondaryHTMLRecordingsPath), DEFAULT_VIRTUAL_PATH)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.maxRecordsToReturn), DEFAULT_MAX_RECORDS_TO_RETURN)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.useDatabase), False)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.useStructuredFileSystem), DEFAULT_USE_STRUCTURED_FILE_SYSTEM)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.callRecordingTable), DEFAULT_CALL_RECORDING_TABLE)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.timeStampName), DEFAULT_TIMESTAMP_NAME)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.filterAtDatabase), DEFAULT_FILTER_AT_DATABASE)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.cdrTableName), DEFAULT_CDR_TABLE_NAME)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.dataSource), DEFAULT_DATA_SOURCE)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.databaseName), DEFAULT_DATABASE_NAME)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.databaseUserId), DEFAULT_USER_ID)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.databasePassword), DEFAULT_PASSWORD)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.searchSubDirectories), DEFAULT_SEARCH_SUB_DIRECTORIES)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.security), DEFAULT_SECURITY)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.xrefDestination), DEFAULT_XREF_DESTINATION)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.xrefTransferredCalls), False)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.useNewDateSearch), False)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.callIdAlias), False)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.drillDownBypassesLogin), False)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.useTempCDRTable), False)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.logFilenameFormatErrors), False)
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.adminPassword), Encrypt(DEFAULT_ADMIN_PASSWORD))
        configDictionary.Add([Enum].GetName(GetType(ConfigItems), ConfigItems.callRecordingFilenameFormat), DEFAULT_CALL_RECORDING_FILENAME_FORMAT)
    End Sub

    Public Sub LoadConfig()
        Dim myFilename As String = GetApplicationPath() & CONFIG_FILENAME

        ' Check that the file exists before trying to load
        Dim myFileInfo As New IO.FileInfo(myFilename)

        If myFileInfo.Exists Then
            Dim myDoc As New XmlDocument
            Dim myRecord As XmlNode = Nothing

            myDoc.Load(myFilename)

            ' Loop over each parameter
            For Each myRecord In myDoc("Config")
                Dim recordName As String = myRecord.Name
                Dim recordData As String = ""

                If myRecord.HasChildNodes Then recordData = myRecord.FirstChild.Value

                ' Mappings for backwards compatiblity
                If recordName = "userId" Then recordName = "databaseUserid"
                If recordName = "password" Then recordName = "databasePassword"
                If recordName = "htmlRecordingsPath" Then recordName = "virtualPath"

                If configDictionary.ContainsKey(recordName) Then
                    configDictionary.SetItem(recordName, recordData)
                Else
                    Logutil.LogError("Unknown key in configDictionary", recordName)
                End If
            Next
        Else
            Logutil.LogError("Initialisation::LoadConfig()", "Cannot find config file: " & myFilename)
        End If
    End Sub

    Public Function GetAdminPassword() As String
        Return Decrypt(configDictionary.GetItem([Enum].GetName(GetType(ConfigItems), ConfigItems.adminPassword)))
    End Function
End Module