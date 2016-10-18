Module ModGlobals

    Public drBestGuess As DataRow
    Public drNewMaster As DataRow

    Public Structure WorkOrder
        Dim WONumber As String
        Dim Quantity As Integer
        Dim AMSPartNumber As String
        Dim StartDate As Date
        Dim RevLevel As String
        Dim Description As String
        Dim WODatarow As DataRow
    End Structure

    Public Structure Barcode
        Dim UniPartNumber As String
        Dim OEMPartNumber As String
        Dim SerialNumber As String
        Dim WorkOrder As String
        Dim QTY As Integer
        Dim SerialStart As String
        Dim SerialEnd As String
        Dim DateLastPrinted As Object
        Dim BarcodeString As String
        Dim KBProfile As String
        Dim BXProfile As String
        Dim KeyboardLabelLayout As String
        Dim BoxLabelLayout As String
        'Dim OEMCustomer As String
        Dim Sample As Boolean
        Dim Family As String
        Dim PNDigits As Int16
        Dim SNType As String
        Dim DECSerialStart As Long
        Dim DECSerialEnd As Long
        Dim MasterDataRow As DataRow
        Dim ReasonCannotPrint As String

    End Structure

    Public curWorkOrder As WorkOrder
    Public curBarCode As Barcode

    Public PNDigits As Int16

    'DEBUG STUFF
    Public NoPrint As Boolean
    Public OverrideExclusions As Boolean
    Public NoLogging As Boolean
    Public KeepDump As Boolean

    Public Sub ReadConfig()

        Dim sNoLogging As String
        Dim sOverrideExclusions As String
        Dim sNoPrint As String
        Dim sKeepDump As String

        sNoLogging = GetSetting("Unicomp Barcode System", "Debug", "No Logging")
        sOverrideExclusions = GetSetting("Unicomp Barcode System", "Debug", "Override Exclusions")
        sNoPrint = GetSetting("Unicomp Barcode System", "Debug", "No Printing")
        sKeepDump = GetSetting("Unicomp Barcode System", "Debug", "Keep Dump")

        If sNoLogging = "YES" Then NoLogging = True Else NoLogging = False
        If sOverrideExclusions = "YES" Then OverrideExclusions = True Else OverrideExclusions = False
        If sNoPrint = "YES" Then NoPrint = True Else NoPrint = False
        If sKeepDump = "YES" Then KeepDump = True Else KeepDump = False
    End Sub

    Public Sub WriteConfig()

        Dim sNoLogging As String
        Dim sOverrideExclusions As String
        Dim sNoPrint As String
        Dim sKeepDump As String

        If NoPrint Then sNoPrint = "YES" Else sNoPrint = "NO"
        If NoLogging Then sNoLogging = "YES" Else sNoLogging = "NO"
        If OverrideExclusions Then sOverrideExclusions = "YES" Else sOverrideExclusions = "NO"
        If KeepDump Then sKeepDump = "YES" Else sKeepDump = "NO"

        SaveSetting("Unicomp Barcode System", "Debug", "No Printing", sNoPrint)
        SaveSetting("Unicomp Barcode System", "Debug", "Override Exclusions", sOverrideExclusions)
        SaveSetting("Unicomp Barcode System", "Debug", "No Logging", sNoLogging)
        SaveSetting("Unicomp Barcode System", "Debug", "Keep Dump", sKeepDump)
    End Sub

End Module
