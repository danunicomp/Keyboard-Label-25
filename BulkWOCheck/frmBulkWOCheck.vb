Imports System.Data
Imports System.Data.OleDb
Public Class BulkWOCheck
    Inherits System.Windows.Forms.Form


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents lstOutput As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(BulkWOCheck))
        Me.btnStart = New System.Windows.Forms.Button
        Me.lstOutput = New System.Windows.Forms.ListBox
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(288, 296)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Check WO's"
        '
        'lstOutput
        '
        Me.lstOutput.Location = New System.Drawing.Point(16, 16)
        Me.lstOutput.Name = "lstOutput"
        Me.lstOutput.Size = New System.Drawing.Size(608, 264)
        Me.lstOutput.TabIndex = 1
        '
        'BulkWOCheck
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(640, 350)
        Me.Controls.Add(Me.lstOutput)
        Me.Controls.Add(Me.btnStart)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "BulkWOCheck"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private odcAlliance As System.Data.OleDb.OleDbConnection
    Private odaWorkOrders As System.Data.OleDb.OleDbDataAdapter
    Private cmdAlliance As System.Data.OleDb.OleDbCommandBuilder
    Private dsWSEWorkOrders As DataSet
    Private sqlStr As String

    Private odcBarcodeSystem As System.Data.OleDb.OleDbConnection
    Private odaBarcodeSystem As System.Data.OleDb.OleDbDataAdapter
    Private cmdBarcodeSystem As System.Data.OleDb.OleDbCommandBuilder
    Private dsExclusions As DataSet

    Private odcProfileMaster As System.Data.OleDb.OleDbConnection
    Private odaProfileMaster As System.Data.OleDb.OleDbDataAdapter
    Private cmdProfileMaster As System.Data.OleDb.OleDbCommandBuilder
    Private dsProfileMaster As DataSet
    'Private sqlStr As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub OpenAllDatabases()

        'ALLIANCE FOR WORK ORDER LIST WITH DESCRIPTION
        'Dim i As Integer
        Dim connStr As String

        connStr = "User ID=sa;Data Source=""HAL\AllianceMFG"";Tag with column collation when possible=False;Initial Catalog=KBD2002;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False;Packet Size=4096"
        odcAlliance = New OleDbConnection(connStr)
        odcAlliance.Open()
        dsWSEWorkOrders = New DataSet

        'sqlStr = "SELECT * FROM WOHeader"
        sqlStr = "SELECT WOHeader.WONumber, WOHeader.PartNumber, PartMaster.DescText FROM PartMaster INNER JOIN WOHeader ON PartMaster.PartNumber = WOHeader.PartNumber WHERE (((PartMaster.DescText) Like '%WSE%'))"
        odaWorkOrders = New OleDbDataAdapter(sqlStr, odcAlliance)
        'odaWorkOrders.SelectCommand.CommandText = sqlStr

        cmdAlliance = New OleDbCommandBuilder(odaWorkOrders)
        dsWSEWorkOrders.Clear()
        odaWorkOrders.Fill(dsWSEWorkOrders, "WOHeader")




        'sqlStr = "SELECT * FROM ProfileMaster WHERE"
    End Sub ' Open All Databases

    Private Function LookUpProfile(ByVal sPN As String) As Boolean
        Dim connStr As String
        'connStr = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Database Password=;Data Source=""\\Hal\unicomp\Share\KB_BARCODES\Profiles.mdb"";Password=;Jet OLEDB:Engine Type=5;Jet OLEDB:Global Bulk Transactions=1;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLEDB:SFP=False;Extended Properties=;Mode=Read;Jet OLEDB:New Database Password=;Jet OLEDB:Create System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;User ID=Admin;Jet OLEDB:Encrypt Database=False"
        'connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=Profiles;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        odcProfileMaster = New OleDbConnection(connStr)
        odcProfileMaster.Open()
        dsProfileMaster = New DataSet
        sPN = UCase(sPN)
        If Microsoft.VisualBasic.Left(sPN, 2) = "00" Then
            sPN = Mid(sPN, 3)
        Else
            If sPN = "042H1292U" Then sPN = "42H1292U"
        End If

        sqlStr = "SELECT * FROM ProfileMaster WHERE PartNumber = '" & sPN & "'"
        odaProfileMaster = New OleDbDataAdapter(sqlStr, odcProfileMaster)
        cmdProfileMaster = New OleDbCommandBuilder(odaProfileMaster)
        dsProfileMaster.Clear()
        odaProfileMaster.Fill(dsProfileMaster, "ProfileMaster")

        If dsProfileMaster.Tables("ProfileMaster").Rows.Count = 1 Then
            LookUpProfile = True
        Else
            LookUpProfile = False
        End If
    End Function

    Private Sub CloseAllDatabases()
        odcAlliance.Close()
        odcProfileMaster.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim r As DataRow
        Call OpenAllDatabases()

        lstOutput.Items.Clear()
        For Each r In dsWSEWorkOrders.Tables("WOHeader").Rows
            If LookUpProfile(r.Item("PartNumber")) = False Then
                'Debug.Write(r.Item("WONumber") & vbTab & r.Item("PartNumber") & vbCrLf)
                lstOutput.Items.Add(r.Item("WONumber") & vbTab & r.Item("PartNumber"))
            End If
            '& vbTab & LookUpProfile(r.Item("PartNumber")) & vbCrLf)
        Next

        Call CloseAllDatabases()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim connStr As String
        connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        odcBarcodeSystem = New OleDbConnection(connStr)
        odcBarcodeSystem.Open()
        dsExclusions = New DataSet

        sqlStr = "SELECT * FROM KBTemplatesToExclude"
        odaBarcodeSystem = New OleDbDataAdapter(sqlStr, odcBarcodeSystem)
        cmdBarcodeSystem = New OleDbCommandBuilder(odaBarcodeSystem)
        dsExclusions.Clear()
        odaBarcodeSystem.Fill(dsExclusions, "KBTemplatesToExclude")

        Dim r As DataRow
        'For Each r In dsExclusions.Tables("KBTemplatesToExclude").Rows
        '    Debug.Write(r.Item("KBLayoutID") & vbCrLf)
        'Next
        Call FillDatabase()

    End Sub

    Private Sub FillDatabase()
        Try


            'odaLog.Fill(myDs, "Logfile")
            'currentRow = 0
            'myDs.Clear()


            Dim newRow As DataRow
            Dim x As Integer

            'For x = 100 To 110
            newRow = dsExclusions.Tables("KBTemplatesToExclude").NewRow()
            newRow("KBLayoutID") = "DAN"

            dsExclusions.Tables("KBTemplatesToExclude").Rows.Add(newRow)
            'sqlInsert = "INSERT INTO LogFile;"
            'odaLog.InsertCommand.CommandText = sqlInsert
            Debug.Write(x & vbCrLf)
            odaBarcodeSystem.Update(dsExclusions, "KBTemplatesToExclude")

            'Next x
        Catch ex As Exception
            MsgBox(ex.ToString)
            Trace.WriteLine(ex.ToString)
        End Try
    End Sub
End Class
