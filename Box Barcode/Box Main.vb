
'History
'2.27 - removed logging for duplicate scans

'2.26 - Initial release to line


Imports System.Data
Imports System.Data.OleDb
Public Class frmBoxMain
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
    Friend WithEvents txtKBScan As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblBoxStationID As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtWorkOrder As System.Windows.Forms.TextBox
    Friend WithEvents txtQTYLeft As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtQTYReq As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents mnuSetup As System.Windows.Forms.MenuItem
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBoxMain))
        Me.txtKBScan = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblBoxStationID = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtWorkOrder = New System.Windows.Forms.TextBox
        Me.txtQTYLeft = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtQTYReq = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.mnuMain = New System.Windows.Forms.MainMenu
        Me.mnuSetup = New System.Windows.Forms.MenuItem
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtKBScan
        '
        Me.txtKBScan.Location = New System.Drawing.Point(16, 48)
        Me.txtKBScan.Name = "txtKBScan"
        Me.txtKBScan.Size = New System.Drawing.Size(224, 20)
        Me.txtKBScan.TabIndex = 0
        Me.txtKBScan.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(64, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 22)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Station:"
        '
        'lblBoxStationID
        '
        Me.lblBoxStationID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoxStationID.Location = New System.Drawing.Point(136, 16)
        Me.lblBoxStationID.Name = "lblBoxStationID"
        Me.lblBoxStationID.Size = New System.Drawing.Size(40, 23)
        Me.lblBoxStationID.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 90)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 18)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Work Order:"
        '
        'txtWorkOrder
        '
        Me.txtWorkOrder.BackColor = System.Drawing.SystemColors.Control
        Me.txtWorkOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkOrder.Location = New System.Drawing.Point(88, 88)
        Me.txtWorkOrder.Name = "txtWorkOrder"
        Me.txtWorkOrder.Size = New System.Drawing.Size(72, 22)
        Me.txtWorkOrder.TabIndex = 4
        Me.txtWorkOrder.TabStop = False
        Me.txtWorkOrder.Text = ""
        '
        'txtQTYLeft
        '
        Me.txtQTYLeft.BackColor = System.Drawing.SystemColors.Control
        Me.txtQTYLeft.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQTYLeft.Location = New System.Drawing.Point(248, 120)
        Me.txtQTYLeft.Name = "txtQTYLeft"
        Me.txtQTYLeft.Size = New System.Drawing.Size(56, 22)
        Me.txtQTYLeft.TabIndex = 6
        Me.txtQTYLeft.TabStop = False
        Me.txtQTYLeft.Text = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(176, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 18)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "QTY Left:"
        '
        'txtQTYReq
        '
        Me.txtQTYReq.BackColor = System.Drawing.SystemColors.Control
        Me.txtQTYReq.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQTYReq.Location = New System.Drawing.Point(112, 120)
        Me.txtQTYReq.Name = "txtQTYReq"
        Me.txtQTYReq.Size = New System.Drawing.Size(48, 22)
        Me.txtQTYReq.TabIndex = 8
        Me.txtQTYReq.TabStop = False
        Me.txtQTYReq.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(8, 122)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 18)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "QTY Required:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.lblBoxStationID)
        Me.GroupBox1.Controls.Add(Me.txtKBScan)
        Me.GroupBox1.Location = New System.Drawing.Point(31, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(256, 80)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSetup})
        '
        'mnuSetup
        '
        Me.mnuSetup.Index = 0
        Me.mnuSetup.Text = "Setup"
        '
        'frmBoxMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(318, 156)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtQTYReq)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtQTYLeft)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtWorkOrder)
        Me.Controls.Add(Me.Label2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.mnuMain
        Me.Name = "frmBoxMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "BOX Barcode"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    '42H1292U-1234567
    'UNI0496--1236545

    '42H1292U-0132818
    Public Structure Barcode
        Dim PartNumber As String
        Dim OEMPartNumber As String
        Dim SerialNumber As String
        Dim BoxProfileFileName As String
        Dim MainBarCode As String
        Dim MasterProfileDatarow As DataRow
        Dim WorkOrder As String
        Dim QTYReq As Integer
        Dim QTYLeft As Integer
    End Structure

    Dim curBarcode As Barcode
    Dim TITLE As String
    Dim AllowPrinting As Boolean

    Private Sub frmBoxMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TITLE = "Unicomp Box v" & Application.ProductVersion
        AllowPrinting = True
        Me.Text = TITLE
        Call ReadConfig()
        If sStationID = "" Then
            Dim frmCNF As New frmBoxSetup

            frmCNF.ShowDialog()
            frmCNF = Nothing 'Better the free variable 
        End If
        Call ResetAll()

    End Sub
    Protected Overrides Function ProcessDialogKey(ByVal keydata As System.Windows.Forms.Keys) As Boolean
        Dim key As System.Windows.Forms.Keys = keydata
        If key = Keys.Enter Then
            If txtKBScan.Focused Then
                Call ProcessInput()
            End If
            'If txtCustomPN.Focused Then
            '    Call ProcessCustom()
            'End If
        End If

        If key = Keys.Tab Then
            If txtKBScan.Focused Then
                Call ProcessInput()
            End If
        End If
        Return MyBase.ProcessDialogKey(keydata)
    End Function



    Private Sub ProcessInput()
        Dim sPartNumber As String
        Dim sSerial As String
        Dim sFullBarcode As String

        sFullBarcode = UCase(txtKBScan.Text)
        If Len(sFullBarcode) <> 16 Then
            MessageBox.Show("Invalid Barcode Length: " & Len(sFullBarcode))
            Call ResetAll()
            Exit Sub
        End If

        sPartNumber = Microsoft.VisualBasic.Left(sFullBarcode, 9)
        sSerial = Microsoft.VisualBasic.Right(sFullBarcode, 7)

        If Microsoft.VisualBasic.Right(sPartNumber, 2) = "--" Then
            sPartNumber = Microsoft.VisualBasic.Left(sPartNumber, 7)
        ElseIf Microsoft.VisualBasic.Right(sPartNumber, 1) = "-" Then
            sPartNumber = Microsoft.VisualBasic.Left(sPartNumber, 8)
        End If

        curBarcode.PartNumber = sPartNumber
        curBarcode.SerialNumber = sSerial

        If curBarcode.SerialNumber = "XXXXXXX" Then
            MessageBox.Show("Invalid Serial Number: XXXXXXX" & vbCrLf & "Its the sample KB Label", "Scanned Sample", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ResetAll()
            Return
        End If

        curBarcode.MasterProfileDatarow = LookUpInMasterProfileQuery(curBarcode.PartNumber)
        If Not IsNothing(curBarcode.MasterProfileDatarow) Then
            Call ParseMasterDatarow(curBarcode.MasterProfileDatarow)
            Call GetWorkOrder(curBarcode.PartNumber, curBarcode.SerialNumber)
            If curBarcode.WorkOrder = "0" Then
                If MessageBox.Show("There is no Work Order assocciated with this serial." & vbCrLf & "Print Anyway?", "No record of WO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then

                    Call ResetAll()
                    Return
                End If
            End If
            Call UpdateWorkOrderLOG()
            Call LookUpInWorkOrderTable()
            txtWorkOrder.Text = curBarcode.WorkOrder
            txtQTYReq.Text = curBarcode.QTYReq
            txtQTYLeft.Text = curBarcode.QTYLeft
            If AlreadyScanned() Then
                If MessageBox.Show("Serial already printed. Reprint?", "Rescan", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Call CreateDump()
                    Call PrintLabelObject()
                    'Call AddToLogfile()
                    curBarcode = Nothing
                Else
                    Call ResetAll()
                    Exit Sub
                End If
            Else
                Call CreateDump()
                Call PrintLabelObject()
                Call AddToLogfile()
            End If


        Else
            MessageBox.Show("Part Number: " & curBarcode.PartNumber & vbCrLf & "Not in system", "P/N Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

        Call ResetAll()

    End Sub

    Private Function LookUpInMasterProfileQuery(ByVal sUnicompPartNumber As String) As DataRow

        Try
            'returns datarow from master query
            Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
            Dim odaProfiles As System.Data.OleDb.OleDbDataAdapter
            Dim cmdProfiles As System.Data.OleDb.OleDbCommandBuilder
            Dim dsProfile As DataSet
            Dim sqlStr As String
            Dim drProfile As DataRow


            Dim connStr As String
            Dim sWorkOrder As String

            'Dim sSerial As String
            'Dim sFamily As String
            'Dim sSerialType As String

            connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
            odcUnicompMain = New OleDbConnection(connStr)
            odcUnicompMain.Open()
            dsProfile = New DataSet

            sqlStr = "SELECT * FROM qProfile WHERE PartNumber = '" & sUnicompPartNumber & "'"
            odaProfiles = New OleDbDataAdapter(sqlStr, odcUnicompMain)

            cmdProfiles = New OleDbCommandBuilder(odaProfiles)
            dsProfile.Clear()
            odaProfiles.Fill(dsProfile, "qProfile")
            If dsProfile.Tables("qProfile").Rows.Count = 1 Then
                LookUpInMasterProfileQuery = dsProfile.Tables("qProfile").Rows(0)
            Else
                LookUpInMasterProfileQuery = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show("Problem connecting to master Profile database." & vbCrLf & "Check network connection" & vbCrLf & ex.ToString, "Network Problem", MessageBoxButtons.OK, MessageBoxIcon.Error)
            LookUpInMasterProfileQuery = Nothing

        End Try
    End Function    'end LookUpInMasterProfileQuery

    Private Sub ParseMasterDatarow(ByVal drMaster As DataRow)
        curBarcode.OEMPartNumber = RTrim(IIf(IsDBNull(drMaster.Item("OEMPartNumber")), curBarcode.PartNumber, drMaster.Item("OEMPartNumber")))
        curBarcode.BoxProfileFileName = RTrim(IIf(IsDBNull(drMaster.Item("BoxTemplateFilename")), "NOTFOUND", drMaster.Item("BoxTemplateFilename")))
        If curBarcode.OEMPartNumber = "0" Then curBarcode.OEMPartNumber = curBarcode.PartNumber
    End Sub


    Private Sub mnuSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frmCNF As New frmBoxSetup

        frmCNF.ShowDialog()
        frmCNF = Nothing 'Better the free variable 
        Call ReadConfig()
        Call ResetAll()
    End Sub

    Private Sub ResetAll()
        lblBoxStationID.Text = sStationID
        txtKBScan.Clear()
        curBarcode = Nothing
    End Sub


#Region "PRinting Routines"

    Private Sub CreateDump()
        Dim sSpace As String
        Select Case Len(curBarcode.PartNumber)
            Case 7
                sSpace = "--"
            Case 8
                sSpace = "-"
            Case 9
                sSpace = ""
            Case Else
                sSpace = ""
        End Select

        curBarcode.MainBarCode = curBarcode.PartNumber & sSpace & curBarcode.SerialNumber

        FileOpen(1, "C:\Unicomp\Templates\BOXDUMP.csv", OpenMode.Output)
        PrintLine(1, "LabelNumber,UniPN,OEMPN,Serial,MainBarcode,Date")
        PrintLine(1, "1," & _
        curBarcode.PartNumber & "," & _
        curBarcode.OEMPartNumber & "," & _
        curBarcode.SerialNumber & "," & _
        curBarcode.MainBarCode & "," & _
        Date.Today)

        FileClose(1)

    End Sub

    Private Sub PrintLabelObject()

        If Not AllowPrinting Then Exit Sub

        Dim LW As Object = CreateObject("Lworks3.LabelEngine")
        Try
            LW.FileName = "C:\Unicomp\Templates\" & curBarcode.BoxProfileFileName & ".LW3"
            LW.Copies = 1
            LW.StartLabel = 1
            LW.TotalLabels = 1
            LW.UpdateSerials = False

            LW.PrintLabels()

            LW = Nothing
        Catch ex As System.Runtime.InteropServices.COMException
            MessageBox.Show("Error: " & curBarcode.BoxProfileFileName & ".LW3" & " not found" & vbCrLf & "in C:\Unicomp\Templates\", "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            'Debug.Write(ex.ToString)
        End Try
    End Sub
#End Region

    Private Function AlreadyScanned() As Boolean

        Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
        Dim odaLOGBox As System.Data.OleDb.OleDbDataAdapter
        Dim cmdLOGBox As System.Data.OleDb.OleDbCommandBuilder
        Dim dsLOGBox As DataSet
        Dim sqlStr As String
        Dim drLOGBox As DataRow
        Dim connStr As String

        Try
            connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
            odcUnicompMain = New OleDbConnection(connStr)
            odcUnicompMain.Open()
            dsLOGBox = New DataSet

            sqlStr = "SELECT * FROM LOG_Box WHERE PartNumber = '" & curBarcode.PartNumber & "' and SerialNumber ='" & curBarcode.SerialNumber & "'"
            'sqlStr = "SELECT * FROM qProfile WHERE PartNumber = '" & sUnicompPartNumber & "'"
            odaLOGBox = New OleDbDataAdapter(sqlStr, odcUnicompMain)

            cmdLOGBox = New OleDbCommandBuilder(odaLOGBox)
            dsLOGBox.Clear()
            odaLOGBox.Fill(dsLOGBox, "LOG_Box")
            If dsLOGBox.Tables("LOG_Box").Rows.Count > 0 Then
                AlreadyScanned = True
            Else
                AlreadyScanned = False
            End If
            odcUnicompMain.Close()
        Catch ex As Exception
            MessageBox.Show("Problem with Log File" & vbCrLf & ex.ToString, "Log File Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Function

    Private Sub GetWorkOrder(ByVal sPN As String, ByVal sSN As String)
        Try
            'LOOk in LOG_KEYBOARD Table for Workorder matching PN and SN
            Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
            Dim odaLOGKB As System.Data.OleDb.OleDbDataAdapter
            Dim cmdLOGKB As System.Data.OleDb.OleDbCommandBuilder
            Dim dsLOGKB As DataSet
            Dim sqlStr As String
            Dim connStr As String

            connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
            odcUnicompMain = New OleDbConnection(connStr)
            odcUnicompMain.Open()
            dsLOGKB = New DataSet

            sqlStr = "SELECT * FROM LOG_Keyboard WHERE PartNumber = '" & sPN & "' and SerialNumber ='" & sSN & "'"
            odaLOGKB = New OleDbDataAdapter(sqlStr, odcUnicompMain)

            cmdLOGKB = New OleDbCommandBuilder(odaLOGKB)
            dsLOGKB.Clear()
            odaLOGKB.Fill(dsLOGKB, "LOG_Keyboard")
            If dsLOGKB.Tables("LOG_Keyboard").Rows.Count > 0 Then
                curBarcode.WorkOrder = RTrim(dsLOGKB.Tables("LOG_Keyboard").Rows(0).Item("WorkOrder"))

            Else
                curBarcode.WorkOrder = "0"
            End If
            odcUnicompMain.Close()
        Catch ex As Exception
            MessageBox.Show("Problem finding workorder in keyboard log" & vbCrLf & ex.ToString, "Keyboard Log error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub LookUpInWorkOrderTable()
        Try
            Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
            Dim odaWO As System.Data.OleDb.OleDbDataAdapter
            Dim cmdWO As System.Data.OleDb.OleDbCommandBuilder
            Dim dsWO As DataSet
            Dim sqlStr As String
            Dim connStr As String

            connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
            odcUnicompMain = New OleDbConnection(connStr)
            odcUnicompMain.Open()
            dsWO = New DataSet

            sqlStr = "SELECT * FROM LOG_WorkOrders WHERE WorkOrder = '" & curBarcode.WorkOrder & "'"
            odaWO = New OleDbDataAdapter(sqlStr, odcUnicompMain)

            cmdWO = New OleDbCommandBuilder(odaWO)
            dsWO.Clear()
            odaWO.Fill(dsWO, "LOG_WorkOrders")
            If dsWO.Tables("LOG_WorkOrders").Rows.Count > 0 Then
                'curBarcode.WorkOrder = RTrim(dsWO.Tables("LOG_WorkOrders").Rows(0).Item("WorkOrder"))
                curBarcode.QTYReq = dsWO.Tables("LOG_WorkOrders").Rows(0).Item("QTYRequired")
                curBarcode.QTYLeft = curBarcode.QTYReq - dsWO.Tables("LOG_WorkOrders").Rows(0).Item("QTYBoxed")
            Else
                curBarcode.QTYReq = 0
                curBarcode.QTYLeft = 0
                'curBarcode.WorkOrder = "0"
            End If
        odcUnicompMain.Close()
        Catch ex As Exception
            MessageBox.Show("Problem looking up workordr in Work Order table" & vbCrLf & ex.ToString, "Work Order database", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub UpdateWorkOrderLOG()
        Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
        Dim odaWO As System.Data.OleDb.OleDbDataAdapter
        Dim cmdWO As System.Data.OleDb.OleDbCommandBuilder
        Dim dsWO As DataSet
        Dim sqlStr As String
        Dim connStr As String
        Dim drDataRow As DataRow

        Dim iQTYBoxed As Integer

        Try
            connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
            odcUnicompMain = New OleDbConnection(connStr)
            odcUnicompMain.Open()
            dsWO = New DataSet

            sqlStr = "SELECT * FROM LOG_WorkOrders WHERE WorkOrder = '" & curBarcode.WorkOrder & "'"
            odaWO = New OleDbDataAdapter(sqlStr, odcUnicompMain)

            cmdWO = New OleDbCommandBuilder(odaWO)
            dsWO.Clear()
            odaWO.Fill(dsWO, "LOG_WorkOrders")

            'drDataRow = dsWO.Tables("LOG_Box").Rows(0)
            If dsWO.Tables("LOG_WorkOrders").Rows.Count = 1 Then
                iQTYBoxed = dsWO.Tables("LOG_WorkOrders").Rows(0).Item("QTYBoxed")
                iQTYBoxed = iQTYBoxed + 1
                dsWO.Tables("LOG_WorkOrders").Rows(0).Item("QTYBoxed") = iQTYBoxed
                odaWO.Update(dsWO, "LOG_WorkOrders")
            End If

            odcUnicompMain.Close()
        Catch ex As Exception
            MessageBox.Show("Problem updating workorder log file" & vbCrLf & ex.ToString, "Work Order logfile", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub AddToLogfile()

        Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
        Dim odaLOGBox As System.Data.OleDb.OleDbDataAdapter
        Dim cmdLOGBox As System.Data.OleDb.OleDbCommandBuilder
        Dim dsLOGBox As DataSet
        Dim sqlStr As String
        Dim drLOGBox As DataRow
        Dim connStr As String

        Dim drNewRow As System.Data.DataRow
        Try
            connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"

            odcUnicompMain = New OleDbConnection(connStr)
            odcUnicompMain.Open()
            dsLOGBox = New DataSet
            sqlStr = "SELECT * FROM LOG_Box"
            odaLOGBox = New OleDbDataAdapter(sqlStr, odcUnicompMain)

            cmdLOGBox = New OleDbCommandBuilder(odaLOGBox)
            dsLOGBox.Clear()
            odaLOGBox.Fill(dsLOGBox, "LOG_Box")

            drNewRow = dsLOGBox.Tables("LOG_Box").NewRow

            drNewRow.Item("Station") = sStationID
            drNewRow.Item("PartNumber") = curBarcode.PartNumber
            drNewRow.Item("SerialNumber") = curBarcode.SerialNumber
            drNewRow.Item("DateBoxed") = Now
            drNewRow.Item("WorkOrder") = curBarcode.WorkOrder

            dsLOGBox.Tables("LOG_Box").Rows.Add(drNewRow)
            odaLOGBox.Update(dsLOGBox, "LOG_Box")
            odcUnicompMain.Close()
        Catch ex As Exception
            MessageBox.Show("Problem writting log file." & vbCrLf & ex.ToString, "Writting Log", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub txtWorkOrder_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWorkOrder.KeyPress
        e.Handled = True
    End Sub

    Private Sub txtQTYLeft_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQTYLeft.KeyPress
        e.Handled = True
    End Sub

    Private Sub txtQTYReq_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQTYReq.KeyPress
        e.Handled = True
    End Sub

    Private Sub mnuSetup_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetup.Click
        Dim frmCNF As New frmBoxSetup

        frmCNF.ShowDialog()
        frmCNF = Nothing 'Better the free variable 
        Call ReadConfig()
        Call ResetAll()
    End Sub

    Private Sub txtKBScan_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKBScan.TextChanged

    End Sub
End Class
