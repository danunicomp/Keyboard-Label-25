Module BoxGlobals
    Public sStationID As String

    Public Sub ReadConfig()
        sStationID = GetSetting("Unicomp Box Labels", "System", "StationID")
    End Sub

    Public Sub WriteConfig()
        SaveSetting("Unicomp Box Labels", "System", "StationID", sStationID)
    End Sub

End Module
