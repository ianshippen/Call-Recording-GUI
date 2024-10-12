Module LocalUtilities
    Public Function GetDatabaseConfigStringForSecurity() As String
        Return CreateConnectionString(callRecordingInterfaceConfigDictionary)
    End Function
End Module
