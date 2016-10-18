Imports System
Imports System.IO
Imports System.Text

Imports System.Data
Imports System.Data.OleDb

'TO DO
' (3/22/05) Add a table to cross-reference part numbers. convert AMSP/N to a Unicomp PN


'History
'v.2.51 (3/25/05)
'-  Send Work Order, Rev Level and PNSN (no hyphen) to template dump data
'-  Add debug window for supressing print, etc

'v.2.50  (3/17/05)
'-  Move to SQLServer database.
'-  Change serial Numbers to numbers
'- Moved log file to SQL Server
'- added work order logfile
'
'
'v.2.29
'-  corrected problem with 42H1292U.  curworkorder.partnumber is actually checked
'       in add new, instead of constructing AMS p/n.  Effect: Part Number will ALWAYS
'       be found.  Bad thing, can't print stuff without a workorder. Fix Later.
'v2.28
'- added override exclusion variable
'- removed select for each format type
'
'v2.27 
'- Release to Floor

Public Class frmKBMain
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
    Friend WithEvents txtWorkOrder As System.Windows.Forms.TextBox
    Friend WithEvents lblWorkOrder As System.Windows.Forms.Label
    Friend WithEvents txtQTY As System.Windows.Forms.TextBox
    Friend WithEvents lblWOPN As System.Windows.Forms.Label
    Friend WithEvents lblQTY As System.Windows.Forms.Label
    Friend WithEvents lblStartlDate As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblRev As System.Windows.Forms.Label
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents lblStartingSerial As System.Windows.Forms.Label
    Friend WithEvents txtStartingSerial As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtWOPN As System.Windows.Forms.TextBox
    Friend WithEvents txtRevLevel As System.Windows.Forms.TextBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grpCustom As System.Windows.Forms.GroupBox
    Friend WithEvents grpNormal As System.Windows.Forms.GroupBox
    Friend WithEvents rdoModeNormal As System.Windows.Forms.RadioButton
    Friend WithEvents rdoModeCustom As System.Windows.Forms.RadioButton
    Friend WithEvents btnCustomGO As System.Windows.Forms.Button
    Friend WithEvents txtCustomStartSN As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomPN As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomQTY As System.Windows.Forms.TextBox
    Friend WithEvents grpMode As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtOEMPN As System.Windows.Forms.TextBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtCustomOEMPN As System.Windows.Forms.TextBox
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents btnDebug As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmKBMain))
        Me.txtWorkOrder = New System.Windows.Forms.TextBox
        Me.lblWorkOrder = New System.Windows.Forms.Label
        Me.txtQTY = New System.Windows.Forms.TextBox
        Me.lblWOPN = New System.Windows.Forms.Label
        Me.lblQTY = New System.Windows.Forms.Label
        Me.lblStartlDate = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.lblRev = New System.Windows.Forms.Label
        Me.lblDescription = New System.Windows.Forms.Label
        Me.lblStartDate = New System.Windows.Forms.Label
        Me.lblMessage = New System.Windows.Forms.Label
        Me.lblStartingSerial = New System.Windows.Forms.Label
        Me.txtStartingSerial = New System.Windows.Forms.TextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtWOPN = New System.Windows.Forms.TextBox
        Me.txtRevLevel = New System.Windows.Forms.TextBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.rdoModeNormal = New System.Windows.Forms.RadioButton
        Me.rdoModeCustom = New System.Windows.Forms.RadioButton
        Me.grpCustom = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtCustomOEMPN = New System.Windows.Forms.TextBox
        Me.btnCustomGO = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtCustomStartSN = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtCustomPN = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtCustomQTY = New System.Windows.Forms.TextBox
        Me.grpNormal = New System.Windows.Forms.GroupBox
        Me.txtOEMPN = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.grpMode = New System.Windows.Forms.GroupBox
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnReset = New System.Windows.Forms.Button
        Me.btnAddNew = New System.Windows.Forms.Button
        Me.btnDebug = New System.Windows.Forms.Button
        Me.grpCustom.SuspendLayout()
        Me.grpNormal.SuspendLayout()
        Me.grpMode.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtWorkOrder
        '
        Me.txtWorkOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkOrder.Location = New System.Drawing.Point(144, 24)
        Me.txtWorkOrder.Name = "txtWorkOrder"
        Me.txtWorkOrder.Size = New System.Drawing.Size(96, 29)
        Me.txtWorkOrder.TabIndex = 0
        Me.txtWorkOrder.Text = ""
        '
        'lblWorkOrder
        '
        Me.lblWorkOrder.AutoSize = True
        Me.lblWorkOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWorkOrder.Location = New System.Drawing.Point(16, 24)
        Me.lblWorkOrder.Name = "lblWorkOrder"
        Me.lblWorkOrder.Size = New System.Drawing.Size(125, 27)
        Me.lblWorkOrder.TabIndex = 1
        Me.lblWorkOrder.Text = "Work Order:"
        '
        'txtQTY
        '
        Me.txtQTY.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQTY.Location = New System.Drawing.Point(104, 160)
        Me.txtQTY.Name = "txtQTY"
        Me.txtQTY.Size = New System.Drawing.Size(80, 26)
        Me.txtQTY.TabIndex = 4
        Me.txtQTY.Text = ""
        '
        'lblWOPN
        '
        Me.lblWOPN.AutoSize = True
        Me.lblWOPN.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWOPN.Location = New System.Drawing.Point(16, 104)
        Me.lblWOPN.Name = "lblWOPN"
        Me.lblWOPN.Size = New System.Drawing.Size(84, 18)
        Me.lblWOPN.TabIndex = 6
        Me.lblWOPN.Text = "Part Number:"
        '
        'lblQTY
        '
        Me.lblQTY.AutoSize = True
        Me.lblQTY.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQTY.Location = New System.Drawing.Point(64, 168)
        Me.lblQTY.Name = "lblQTY"
        Me.lblQTY.Size = New System.Drawing.Size(36, 18)
        Me.lblQTY.TabIndex = 7
        Me.lblQTY.Text = "QTY:"
        '
        'lblStartlDate
        '
        Me.lblStartlDate.AutoSize = True
        Me.lblStartlDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartlDate.Location = New System.Drawing.Point(56, 64)
        Me.lblStartlDate.Name = "lblStartlDate"
        Me.lblStartlDate.Size = New System.Drawing.Size(82, 16)
        Me.lblStartlDate.TabIndex = 8
        Me.lblStartlDate.Text = "Released Date:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Location = New System.Drawing.Point(0, 216)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(552, 8)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        '
        'lblRev
        '
        Me.lblRev.AutoSize = True
        Me.lblRev.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRev.Location = New System.Drawing.Point(228, 88)
        Me.lblRev.Name = "lblRev"
        Me.lblRev.Size = New System.Drawing.Size(32, 18)
        Me.lblRev.TabIndex = 13
        Me.lblRev.Text = "Rev:"
        '
        'lblDescription
        '
        Me.lblDescription.AutoSize = True
        Me.lblDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescription.Location = New System.Drawing.Point(24, 136)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(76, 18)
        Me.lblDescription.TabIndex = 15
        Me.lblDescription.Text = "Description:"
        '
        'lblStartDate
        '
        Me.lblStartDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartDate.Location = New System.Drawing.Point(144, 64)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(96, 17)
        Me.lblStartDate.TabIndex = 17
        '
        'lblMessage
        '
        Me.lblMessage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessage.Location = New System.Drawing.Point(256, 24)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(224, 56)
        Me.lblMessage.TabIndex = 21
        '
        'lblStartingSerial
        '
        Me.lblStartingSerial.AutoSize = True
        Me.lblStartingSerial.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartingSerial.Location = New System.Drawing.Point(216, 168)
        Me.lblStartingSerial.Name = "lblStartingSerial"
        Me.lblStartingSerial.Size = New System.Drawing.Size(127, 18)
        Me.lblStartingSerial.TabIndex = 25
        Me.lblStartingSerial.Text = "STARTING SERIAL:"
        '
        'txtStartingSerial
        '
        Me.txtStartingSerial.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStartingSerial.Location = New System.Drawing.Point(344, 160)
        Me.txtStartingSerial.Name = "txtStartingSerial"
        Me.txtStartingSerial.Size = New System.Drawing.Size(136, 26)
        Me.txtStartingSerial.TabIndex = 24
        Me.txtStartingSerial.Text = ""
        '
        'GroupBox3
        '
        Me.GroupBox3.Location = New System.Drawing.Point(0, 432)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(552, 8)
        Me.GroupBox3.TabIndex = 26
        Me.GroupBox3.TabStop = False
        '
        'txtWOPN
        '
        Me.txtWOPN.BackColor = System.Drawing.SystemColors.Control
        Me.txtWOPN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWOPN.Location = New System.Drawing.Point(104, 104)
        Me.txtWOPN.Name = "txtWOPN"
        Me.txtWOPN.Size = New System.Drawing.Size(104, 26)
        Me.txtWOPN.TabIndex = 27
        Me.txtWOPN.TabStop = False
        Me.txtWOPN.Text = ""
        '
        'txtRevLevel
        '
        Me.txtRevLevel.BackColor = System.Drawing.SystemColors.Control
        Me.txtRevLevel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRevLevel.Location = New System.Drawing.Point(216, 112)
        Me.txtRevLevel.Name = "txtRevLevel"
        Me.txtRevLevel.Size = New System.Drawing.Size(56, 20)
        Me.txtRevLevel.TabIndex = 28
        Me.txtRevLevel.TabStop = False
        Me.txtRevLevel.Text = ""
        '
        'txtDesc
        '
        Me.txtDesc.BackColor = System.Drawing.SystemColors.Control
        Me.txtDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesc.Location = New System.Drawing.Point(104, 136)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(376, 20)
        Me.txtDesc.TabIndex = 29
        Me.txtDesc.TabStop = False
        Me.txtDesc.Text = ""
        '
        'rdoModeNormal
        '
        Me.rdoModeNormal.Location = New System.Drawing.Point(24, 24)
        Me.rdoModeNormal.Name = "rdoModeNormal"
        Me.rdoModeNormal.Size = New System.Drawing.Size(72, 24)
        Me.rdoModeNormal.TabIndex = 50
        Me.rdoModeNormal.Text = "Normal"
        '
        'rdoModeCustom
        '
        Me.rdoModeCustom.Location = New System.Drawing.Point(24, 56)
        Me.rdoModeCustom.Name = "rdoModeCustom"
        Me.rdoModeCustom.Size = New System.Drawing.Size(72, 24)
        Me.rdoModeCustom.TabIndex = 51
        Me.rdoModeCustom.Text = "Custom"
        '
        'grpCustom
        '
        Me.grpCustom.Controls.Add(Me.Label6)
        Me.grpCustom.Controls.Add(Me.txtCustomOEMPN)
        Me.grpCustom.Controls.Add(Me.btnCustomGO)
        Me.grpCustom.Controls.Add(Me.Label3)
        Me.grpCustom.Controls.Add(Me.txtCustomStartSN)
        Me.grpCustom.Controls.Add(Me.Label2)
        Me.grpCustom.Controls.Add(Me.txtCustomPN)
        Me.grpCustom.Controls.Add(Me.Label1)
        Me.grpCustom.Controls.Add(Me.txtCustomQTY)
        Me.grpCustom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpCustom.Location = New System.Drawing.Point(160, 232)
        Me.grpCustom.Name = "grpCustom"
        Me.grpCustom.Size = New System.Drawing.Size(352, 176)
        Me.grpCustom.TabIndex = 39
        Me.grpCustom.TabStop = False
        Me.grpCustom.Text = "Custom Labels"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 72)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(65, 18)
        Me.Label6.TabIndex = 47
        Me.Label6.Text = "OEM P/N:"
        '
        'txtCustomOEMPN
        '
        Me.txtCustomOEMPN.Enabled = False
        Me.txtCustomOEMPN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomOEMPN.Location = New System.Drawing.Point(88, 64)
        Me.txtCustomOEMPN.Name = "txtCustomOEMPN"
        Me.txtCustomOEMPN.Size = New System.Drawing.Size(96, 26)
        Me.txtCustomOEMPN.TabIndex = 46
        Me.txtCustomOEMPN.TabStop = False
        Me.txtCustomOEMPN.Text = ""
        '
        'btnCustomGO
        '
        Me.btnCustomGO.Location = New System.Drawing.Point(264, 144)
        Me.btnCustomGO.Name = "btnCustomGO"
        Me.btnCustomGO.TabIndex = 45
        Me.btnCustomGO.Text = "GO"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(192, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(56, 31)
        Me.Label3.TabIndex = 44
        Me.Label3.Text = "Starting S/N:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCustomStartSN
        '
        Me.txtCustomStartSN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomStartSN.Location = New System.Drawing.Point(256, 24)
        Me.txtCustomStartSN.Name = "txtCustomStartSN"
        Me.txtCustomStartSN.Size = New System.Drawing.Size(80, 26)
        Me.txtCustomStartSN.TabIndex = 43
        Me.txtCustomStartSN.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(32, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(31, 18)
        Me.Label2.TabIndex = 42
        Me.Label2.Text = "P/N:"
        '
        'txtCustomPN
        '
        Me.txtCustomPN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomPN.Location = New System.Drawing.Point(88, 24)
        Me.txtCustomPN.Name = "txtCustomPN"
        Me.txtCustomPN.Size = New System.Drawing.Size(96, 26)
        Me.txtCustomPN.TabIndex = 41
        Me.txtCustomPN.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(216, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 18)
        Me.Label1.TabIndex = 40
        Me.Label1.Text = "QTY:"
        '
        'txtCustomQTY
        '
        Me.txtCustomQTY.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomQTY.Location = New System.Drawing.Point(256, 64)
        Me.txtCustomQTY.Name = "txtCustomQTY"
        Me.txtCustomQTY.Size = New System.Drawing.Size(48, 26)
        Me.txtCustomQTY.TabIndex = 44
        Me.txtCustomQTY.Text = ""
        '
        'grpNormal
        '
        Me.grpNormal.Controls.Add(Me.txtOEMPN)
        Me.grpNormal.Controls.Add(Me.Label5)
        Me.grpNormal.Controls.Add(Me.txtDesc)
        Me.grpNormal.Controls.Add(Me.lblDescription)
        Me.grpNormal.Controls.Add(Me.lblQTY)
        Me.grpNormal.Controls.Add(Me.txtQTY)
        Me.grpNormal.Controls.Add(Me.lblStartingSerial)
        Me.grpNormal.Controls.Add(Me.txtStartingSerial)
        Me.grpNormal.Controls.Add(Me.lblRev)
        Me.grpNormal.Controls.Add(Me.txtWOPN)
        Me.grpNormal.Controls.Add(Me.txtRevLevel)
        Me.grpNormal.Controls.Add(Me.lblWOPN)
        Me.grpNormal.Controls.Add(Me.txtWorkOrder)
        Me.grpNormal.Controls.Add(Me.lblWorkOrder)
        Me.grpNormal.Controls.Add(Me.lblStartlDate)
        Me.grpNormal.Controls.Add(Me.lblStartDate)
        Me.grpNormal.Controls.Add(Me.lblMessage)
        Me.grpNormal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpNormal.Location = New System.Drawing.Point(16, 16)
        Me.grpNormal.Name = "grpNormal"
        Me.grpNormal.Size = New System.Drawing.Size(496, 200)
        Me.grpNormal.TabIndex = 40
        Me.grpNormal.TabStop = False
        Me.grpNormal.Text = "Label By Work Order"
        '
        'txtOEMPN
        '
        Me.txtOEMPN.BackColor = System.Drawing.SystemColors.Control
        Me.txtOEMPN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOEMPN.Location = New System.Drawing.Point(360, 104)
        Me.txtOEMPN.Name = "txtOEMPN"
        Me.txtOEMPN.Size = New System.Drawing.Size(120, 26)
        Me.txtOEMPN.TabIndex = 31
        Me.txtOEMPN.TabStop = False
        Me.txtOEMPN.Text = ""
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(288, 108)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(69, 18)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "OEM P/N::"
        '
        'grpMode
        '
        Me.grpMode.Controls.Add(Me.rdoModeNormal)
        Me.grpMode.Controls.Add(Me.rdoModeCustom)
        Me.grpMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpMode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpMode.Location = New System.Drawing.Point(16, 232)
        Me.grpMode.Name = "grpMode"
        Me.grpMode.Size = New System.Drawing.Size(104, 104)
        Me.grpMode.TabIndex = 41
        Me.grpMode.TabStop = False
        Me.grpMode.Text = "Mode"
        '
        'btnPrint
        '
        Me.btnPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(416, 448)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(104, 48)
        Me.btnPrint.TabIndex = 43
        Me.btnPrint.Text = "Print"
        '
        'btnReset
        '
        Me.btnReset.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReset.Location = New System.Drawing.Point(304, 464)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(88, 32)
        Me.btnReset.TabIndex = 44
        Me.btnReset.Text = "Reset"
        '
        'btnAddNew
        '
        Me.btnAddNew.Location = New System.Drawing.Point(16, 480)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.TabIndex = 45
        Me.btnAddNew.Text = "Add New"
        '
        'btnDebug
        '
        Me.btnDebug.Location = New System.Drawing.Point(16, 448)
        Me.btnDebug.Name = "btnDebug"
        Me.btnDebug.TabIndex = 46
        Me.btnDebug.Text = "Debug"
        '
        'frmKBMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(544, 508)
        Me.Controls.Add(Me.btnDebug)
        Me.Controls.Add(Me.btnAddNew)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.grpMode)
        Me.Controls.Add(Me.grpNormal)
        Me.Controls.Add(Me.grpCustom)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmKBMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Keyboard Barcode Labeling System"
        Me.grpCustom.ResumeLayout(False)
        Me.grpNormal.ResumeLayout(False)
        Me.grpMode.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Abort As Boolean
    Public Custom As Boolean

    Dim BarcodeFieldType(9) As String
    Dim BarcodeFieldLength(9) As Integer

    Dim sExcludedReason As String

#Region "Key Press Supression"

    Private Sub _KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWOPN.KeyPress
        e.Handled = True
    End Sub
    Private Sub txtRevLevel_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRevLevel.KeyPress
        e.Handled = True
    End Sub
    Private Sub txtDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDesc.KeyPress
        e.Handled = True
    End Sub

    Private Sub txtStartingSerial_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStartingSerial.KeyPress
        e.Handled = True
    End Sub

    Private Sub txtQTY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQTY.KeyPress
        e.Handled = True
    End Sub
    Private Sub txtCustomLayout_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        e.Handled = True
    End Sub
    Private Sub txtOEMPN_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOEMPN.KeyPress
        e.Handled = True
    End Sub
    Private Sub txtCustomOEMPN_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomOEMPN.KeyDown
        e.Handled = True
    End Sub
#End Region

    Private Sub frmKBMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        rdoModeNormal.Checked = True
        ShowMessage("Ready", Color.Green)

        '------------------------VARIABLES
        Me.Text = "Unicomp Barcode System v" & Application.ProductVersion()

        Call ResetAll()

        'AllowPrint = True
        'OverrideExclusions = False
        'Debug.Write(Application.ProductVersion())
    End Sub

    Protected Overrides Function ProcessDialogKey(ByVal keydata As System.Windows.Forms.Keys) As Boolean
        Dim key As System.Windows.Forms.Keys = keydata
        If key = Keys.Enter Then
            If txtWorkOrder.Focused Then
                Call MainProcess()
            End If
            'If txtCustomPN.Focused Then
            '    Call ProcessCustom()
            'End If
        End If

        Return MyBase.ProcessDialogKey(keydata)
    End Function
    Private Sub OpenAllDatabases()


    End Sub ' Open All Databases

    Private Sub MainProcess()

        Try
            Dim sWON As String
            sWON = UCase(RTrim(LTrim(txtWorkOrder.Text)))

            curWorkOrder.WODatarow = LookUpWorkOrder(sWON)
            If IsNothing(curWorkOrder.WODatarow) Then
                MessageBox.Show("Could not find Work Order: " & sWON)
                ResetAll()
                Exit Sub
            End If
            Call ParseWODatarow(curWorkOrder.WODatarow)
            curBarCode.UniPartNumber = UnicompPartNumber(curWorkOrder.AMSPartNumber)
            curBarCode.WorkOrder = curWorkOrder.WONumber
            curBarCode.QTY = curWorkOrder.Quantity
            curBarCode.MasterDataRow = LookUpInMasterProfileQuery(curBarCode.UniPartNumber)
            If IsNothing(curBarCode.MasterDataRow) Then
                MsgBox("NOT PRINTED BEFORE. ADD TO MASTER PROFILE")
                Call AddToMaster()
                Exit Sub
            Else
                ParseMasterDatarow(curBarCode.MasterDataRow)
                Call UpdateDisplayFromWO()
                If InFamilyExclusions() Then
                    MessageBox.Show("Family cannot be processed." & vbCrLf & "Part Number: " & curBarCode.UniPartNumber & vbCrLf & "Family: " & curBarCode.Family & vbCrLf & "Reason: " & curBarCode.ReasonCannotPrint)
                    Call ResetAll()
                    Exit Sub
                End If
                If curBarCode.OEMPartNumber = "0" Then
                    MessageBox.Show("ERROR.  P/N: " & curBarCode.UniPartNumber & " Has '0' For OEM Part Number in database", "Database problem", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Call ResetAll()
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("General Failure, better get Dan" & vbCrLf & ex.ToString, "Get Dan", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        'Call CreateTemplateDumpDB()



    End Sub

    Private Function LookUpWorkOrder(ByVal sWONumber As String) As DataRow
        Dim odcAlliance As System.Data.OleDb.OleDbConnection
        Dim odaWorkOrders As System.Data.OleDb.OleDbDataAdapter
        Dim cmdAlliance As System.Data.OleDb.OleDbCommandBuilder
        Dim dsWorkOrder As DataSet
        Dim sqlStr As String

        'ALLIANCE FOR WORK ORDER LIST WITH DESCRIPTION
        Dim connStr As String
        Dim sWorkOrder As String
        Try
            sWorkOrder = LTrim(RTrim(txtWorkOrder.Text))

            connStr = "User ID=sa;Data Source=""HAL\AllianceMFG"";Tag with column collation when possible=False;Initial Catalog=KBD2002;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False;Packet Size=4096"
            odcAlliance = New OleDbConnection(connStr)
            odcAlliance.Open()
            dsWorkOrder = New DataSet

            sqlStr = "SELECT WOHeader.WONumber, WOHeader.QuantityRequired, WOHeader.PartNumber, WOHeader.StartDate, PartMaster.DescText , PartMaster.Revision FROM PartMaster INNER JOIN WOHeader ON PartMaster.PartNumber = WOHeader.PartNumber WHERE (((WOHeader.WONumber) = '" & sWorkOrder & "'))"
            odaWorkOrders = New OleDbDataAdapter(sqlStr, odcAlliance)

            cmdAlliance = New OleDbCommandBuilder(odaWorkOrders)
            dsWorkOrder.Clear()
            odaWorkOrders.Fill(dsWorkOrder, "WOHeader")
            If dsWorkOrder.Tables("WOHeader").Rows.Count = 1 Then
                LookUpWorkOrder = dsWorkOrder.Tables("WOHeader").Rows(0)
            Else
                LookUpWorkOrder = Nothing
            End If
            odcAlliance.Close()
        Catch ex As Exception
            MessageBox.Show("Error connecting to database server." & vbCrLf & "Check network connection", "Check Network", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub ParseWODatarow(ByVal drWorkOrder As DataRow)
        With curWorkOrder
            .WONumber = IIf(IsDBNull(drWorkOrder.Item("WONumber")), "", drWorkOrder.Item("WONumber"))
            .AMSPartNumber = IIf(IsDBNull(drWorkOrder.Item("PartNumber")), "0", drWorkOrder.Item("PartNumber"))
            .Quantity = IIf(IsDBNull(drWorkOrder.Item("QuantityRequired")), "0", drWorkOrder.Item("QuantityRequired"))
            .StartDate = IIf(IsDBNull(drWorkOrder.Item("StartDate")), "#01/01/1980#", drWorkOrder.Item("StartDate"))
            .Description = IIf(IsDBNull(drWorkOrder.Item("DescText")), "", drWorkOrder.Item("DescText"))
            .RevLevel = IIf(IsDBNull(drWorkOrder.Item("Revision")), "", drWorkOrder.Item("Revision"))
        End With
    End Sub

    Private Sub UpdateDisplayFromWO()
        ''UPDATE DISPLAY
        txtWOPN.Text = curBarCode.UniPartNumber
        txtOEMPN.Text = curBarCode.OEMPartNumber
        txtQTY.Text = curWorkOrder.Quantity
        lblStartDate.Text = curWorkOrder.StartDate
        txtRevLevel.Text = curWorkOrder.RevLevel
        txtDesc.Text = curWorkOrder.Description

        txtStartingSerial.Text = LTrim(Str(curBarCode.DECSerialStart)).PadLeft(7, "0")
    End Sub

    Private Function UnicompPartNumber(ByVal sAMSPN As String) As String
        'CONSTRUCT UNICOMP PN FROM AMS PN
        If Microsoft.VisualBasic.Left(sAMSPN, 2) = "00" And Len(sAMSPN) = 9 Then
            UnicompPartNumber = UCase(Mid(sAMSPN, 3))
            PNDigits = 7
        Else
            ' PART NUMBER EXCEPTIONS
            Select Case sAMSPN
                Case "042H1292U"
                    curBarCode.PNDigits = 8
                    UnicompPartNumber = UCase(Mid(curWorkOrder.AMSPartNumber, 2))
                Case "0098U0181ZZ"
                    curBarCode.PNDigits = 9
                    UnicompPartNumber = UCase(Mid(curWorkOrder.AMSPartNumber, 3))
                Case Else
                    MsgBox("CANNOT PROCESS P/N: " & curWorkOrder.AMSPartNumber & vbCrLf & _
                               "Illegal part number format" & vbCrLf & _
                               "TOO MANY DIGITS OR NO LEADING 00. Cannot Serialize P/N")
                    Abort = True
            End Select
        End If
    End Function

    Private Function LookUpInMasterProfileQuery(ByVal sUnicompPartNumber As String) As DataRow
        'returns datarow from master query
        Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
        Dim odaProfiles As System.Data.OleDb.OleDbDataAdapter
        Dim cmdProfiles As System.Data.OleDb.OleDbCommandBuilder
        Dim dsProfile As DataSet
        Dim sqlStr As String
        Dim drProfile As DataRow

        'ALLIANCE FOR WORK ORDER LIST WITH DESCRIPTION
        Dim connStr As String
        Dim sWorkOrder As String

        Dim sSerial As String
        Dim sFamily As String
        Dim sSerialType As String

        connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        odcUnicompMain = New OleDbConnection(connStr)
        odcUnicompMain.Open()
        dsProfile = New DataSet

        sqlStr = "SELECT * FROM qProfile WHERE PartNumber = '" & curBarCode.UniPartNumber & "'"
        odaProfiles = New OleDbDataAdapter(sqlStr, odcUnicompMain)

        cmdProfiles = New OleDbCommandBuilder(odaProfiles)
        dsProfile.Clear()
        Try
            odaProfiles.Fill(dsProfile, "qProfile")

        If dsProfile.Tables("qProfile").Rows.Count = 1 Then
            LookUpInMasterProfileQuery = dsProfile.Tables("qProfile").Rows(0)
        Else
            LookUpInMasterProfileQuery = Nothing

            End If
        Catch ex As Exception
            MessageBox.Show("Unable to connect to master database." & vbCrLf & "Check network connection", "Check Network", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function    'end LookUpInMasterProfileQuery

    Private Sub ParseMasterDatarow(ByVal drMaster As DataRow)
        Try
            curBarCode.OEMPartNumber = RTrim(IIf(IsDBNull(drMaster.Item("OEMPartNumber")), curBarCode.UniPartNumber, drMaster.Item("OEMPartNumber")))
            curBarCode.DateLastPrinted = IIf(IsDBNull(drMaster.Item("DateLastPrinted")), vbNull, drMaster.Item("DateLastPrinted"))
            curBarCode.KBProfile = RTrim(IIf(IsDBNull(drMaster.Item("KeyboardLabelTemplate")), "0", drMaster.Item("KeyboardLabelTemplate")))
            curBarCode.BXProfile = RTrim(IIf(IsDBNull(drMaster.Item("BoxLabelTemplate")), "0", drMaster.Item("BoxLabelTemplate")))
            curBarCode.KeyboardLabelLayout = RTrim(IIf(IsDBNull(drMaster.Item("KBTemplateFilename")), "NOTFOUND", drMaster.Item("KBTemplateFilename")))
            curBarCode.BoxLabelLayout = RTrim(IIf(IsDBNull(drMaster.Item("BoxTemplateFilename")), "NOTFOUND", drMaster.Item("BoxTemplateFilename")))
            curBarCode.Family = IIf(IsDBNull(drMaster.Item("FamilyCode")), "GM", drMaster.Item("FamilyCode"))
            curBarCode.DECSerialStart = StartingSerial(drMaster)
        Catch ex As Exception
            MessageBox.Show("Error parsing Profile data:" & vbCrLf & ex.ToString, "Error Parsing", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Function StartingSerial(ByVal drMaster As DataRow) As Long
        Dim sSerial As String
        Try
            sSerial = RTrim(IIf(IsDBNull(drMaster.Item("NextSerialNumber")), "1", drMaster.Item("NextSerialNumber")))
            StartingSerial = Val(sSerial)
        Catch ex As Exception
            MessageBox.Show("Error getting starting serial number:" & vbCrLf & ex.ToString, "Error with Starting Serial", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub CreateTemplateDumpDB(ByVal sMode As String)
        Dim odcTemplateDump As System.Data.OleDb.OleDbConnection
        Dim odaTemplateDump As System.Data.OleDb.OleDbDataAdapter
        Dim cmdTemplateDump As System.Data.OleDb.OleDbCommandBuilder
        Dim dsTemplateDump As DataSet
        Dim sqlStr As String
        Dim connStr As String

        Dim r As DataRow

        connStr = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Database Password=;Data Source=""C:\Unicomp\Templates\TemplateData.mdb"";Password=;Jet OLEDB:Engine Type=5;Jet OLEDB:Global Bulk Transactions=1;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLEDB:SFP=False;Extended Properties=;Mode=Share Deny None;Jet OLEDB:New Database Password=;Jet OLEDB:Create System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;User ID=Admin;Jet OLEDB:Encrypt Database=False"
        odcTemplateDump = New OleDbConnection(connStr)
        odcTemplateDump.Open()
        dsTemplateDump = New DataSet
        sqlStr = "SELECT * FROM TemplateData"
        odaTemplateDump = New OleDbDataAdapter(sqlStr, odcTemplateDump)

        cmdTemplateDump = New OleDbCommandBuilder(odaTemplateDump)
        dsTemplateDump.Clear()
        Try
            odaTemplateDump.Fill(dsTemplateDump, "TemplateData")
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Exit Sub
        End Try
        'CLEAR DUMP DATABASE
        Try
            For Each r In dsTemplateDump.Tables("TemplateData").Rows
                r.Delete()
            Next
            odaTemplateDump.Update(dsTemplateDump, "TemplateData")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Dim s As String

        Try
            Select Case Len(curBarCode.UniPartNumber)
                Case 7 : s = "--"
                Case 8 : s = "-"
                Case 9 : s = ""
                Case Else
                    'MsgBox("ERROR WITH PART NMBER LENGTH")
                    'Exit Sub
            End Select

            Select Case sMode
                Case "Sample"
                    r = dsTemplateDump.Tables("TemplateData").NewRow
                    r.Item("UniPN") = curBarCode.UniPartNumber
                    r.Item("OEMPN") = curBarCode.OEMPartNumber
                    r.Item("Serial") = "XXXXXXX"
                    r.Item("MainBarcode") = curBarCode.UniPartNumber & s & "XXXXXXX"
                    r.Item("Sample") = "SAMPLE DO NOT USE"
                    r.Item("BarcodeWithNoHyphen") = curBarCode.UniPartNumber & "XXXXXXX"
                    r.Item("WorkOrder") = curBarCode.WorkOrder
                    r.Item("RevLevel") = curWorkOrder.RevLevel
                    r.Item("BuildDate") = Now.Date
                    dsTemplateDump.Tables("TemplateData").Rows.Add(r)
                    odaTemplateDump.Update(dsTemplateDump, "TemplateData")
                Case "Full"
                    Dim x As Integer
                    For x = 0 To curBarCode.QTY - 1
                        r = dsTemplateDump.Tables("TemplateData").NewRow
                        r.Item("UniPN") = curBarCode.UniPartNumber
                        r.Item("OEMPN") = curBarCode.OEMPartNumber
                        r.Item("Serial") = LTrim(curBarCode.DECSerialStart + x).PadLeft(7, "0")
                        r.Item("MainBarcode") = curBarCode.UniPartNumber & s & LTrim(curBarCode.DECSerialStart + x).PadLeft(7, "0")
                        r.Item("Sample") = ""
                        r.Item("BuildDate") = Now.Date
                        r.Item("BarcodeWithNoHyphen") = curBarCode.UniPartNumber & LTrim(curBarCode.DECSerialStart + x).PadLeft(7, "0")
                        r.Item("WorkOrder") = curBarCode.WorkOrder
                        r.Item("RevLevel") = curWorkOrder.RevLevel
                        dsTemplateDump.Tables("TemplateData").Rows.Add(r)
                    Next
                    Try
                        odaTemplateDump.Update(dsTemplateDump, "TemplateData")
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try
                Case "Clear"
                    '
                Case Else
                    MsgBox("ERROR IN CREATE TEMPLATE DATABASE FUNCTION CALL")
            End Select

            odcTemplateDump.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


    Private Sub ShowMessage(ByVal sMessage As String, ByVal colColor As Color)
        lblMessage.Text = sMessage
        lblMessage.ForeColor = colColor
    End Sub


    Private Sub frmKBMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        txtWorkOrder.Focus()
    End Sub



    Private Sub butManualMaster_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call AddToMaster()
    End Sub



#Region "Old Routines"

    'Private Function LookupInGuess() As Boolean
    '    'check for full part number first
    '    Dim drGuess As DataRow
    '    DsGuess1.Clear()
    '    selGuess.CommandText = "SELECT * FROM qryGuess WHERE Suffix = '" & curBarCode.UniPartNumber & "'"
    '    odaGuess.Fill(DsGuess1)

    '    If DsGuess1.Tables("qryGuess").Rows.Count = 1 Then
    '        drBestGuess = DsGuess1.Tables("qryGuess").Rows.Item(0)
    '        LookupInGuess = True
    '    Else    'not found whole number, look for Suffix
    '        DsGuess1.Clear()
    '        selGuess.CommandText = "SELECT * FROM qryGuess WHERE Suffix = '" & Microsoft.VisualBasic.Left(curBarCode.UniPartNumber, 3) & "'"
    '        odaGuess.Fill(DsGuess1)
    '        If DsGuess1.Tables("qryGuess").Rows.Count = 1 Then
    '            'found prefix
    '            drBestGuess = DsGuess1.Tables("qryGuess").Rows.Item(0)

    '            LookupInGuess = True
    '        Else
    '            'NOT FOUND ANYWHERE
    '            MsgBox("CANNOT GUESS P/N, HAVE TO ENTER MANUALLY")
    '            Abort = True
    '            LookupInGuess = False
    '            'Call AddToMaster()
    '            Exit Function
    '        End If
    '    End If
    '    'curBarCode.SerialStart = IIf(IsDBNull(drGuess.Item("SerialNext")), "", drGuess.Item("SerialNext"))
    '    'curBarCode.SerialStart = "1"
    '    'If drGuess.Item("HasOEMPartNumber") = True Then
    '    '    curBarCode.OEMPartNumber = InputBox("ENTER OEM P/N FOR" & vbCrLf & curBarCode.UniPartNumber)
    '    'Else
    '    '    curBarCode.OEMPartNumber = "NONE"
    '    'End If

    '    'curBarCode.QTY = curWorkOrder.Quantity
    '    ''curBarCode.DateLastPrinted = IIf(IsDBNull(drGuess.Item("DateLastPrinted")), vbNull, drGuess.Item("DateLastPrinted"))
    '    'curBarCode.FormatType = IIf(IsDBNull(drGuess.Item("FormatType")), "0", drGuess.Item("FormatType"))
    '    'curBarCode.KeyboardLabelLayout = IIf(IsDBNull(drGuess.Item("KB_RawLayout")), "NOTFOUND", drGuess.Item("KB_RawLayout"))
    '    'curBarCode.BoxLabelLayout = IIf(IsDBNull(drGuess.Item("BOX_RawLayout")), "NOTFOUND", drGuess.Item("BOX_RawLayout"))
    '    'curBarCode.OEMCustomer = IIf(IsDBNull(drGuess.Item("CustomerDesc")), "", drGuess.Item("CustomerDesc"))


    'End Function

    'Private Sub ProcessCustom()
    '    txtCustomLayout.Clear()
    '    DsGuess1.Clear()
    '    selGuess.CommandText = "SELECT * FROM qryGuess WHERE Suffix = '" & Microsoft.VisualBasic.Left(txtCustomPN.Text, 3) & "'"
    '    odaGuess.Fill(DsGuess1)
    '    If DsGuess1.Tables("qryGuess").Rows.Count = 1 Then
    '        txtCustomLayout.Text = DsGuess1.Tables("qryGuess").Rows(0).Item("Description")
    '    Else
    '        DsGuess1.Clear()
    '        selGuess.CommandText = "SELECT * FROM qryGuess WHERE Suffix = '" & txtCustomPN.Text & "'"
    '        odaGuess.Fill(DsGuess1)
    '        If DsGuess1.Tables("qryGuess").Rows.Count = 1 Then
    '            txtCustomLayout.Text = DsGuess1.Tables("qryGuess").Rows(0).Item("Description")
    '        End If
    '    End If

    'End Sub

#End Region



#Region "Printing: Summary"


    Private Sub CreateSummaryDump()
        Dim iLBLNumber As Integer
        Dim sDump As String
        Dim iCurrentSerial As Integer

        'sDump = "LabelNumber,OEMCust,UniPN,OEMPN,Serial,Batch,Date"



        sDump = Nothing
        'lstOutput.Items.Clear()
        Try
            FileOpen(1, "C:\Unicomp\Templates\SUMMARY.csv", OpenMode.Output)

            PrintLine(1, "UniPN,OEMPN,SerialStart,SerialEnd,QTY,WorkOrder,Date")
            'lstOutput.Items.Add("OEMCust,UniPN,OEMPN,SerialStart,SerialEnd,QTY,WorkOrder,PrintReprint,Date")
            sDump = sDump & _
            curBarCode.UniPartNumber & "," & _
            curBarCode.OEMPartNumber & "," & _
            LTrim(Str(curBarCode.DECSerialStart)).PadLeft(7, "0") & "," & _
            LTrim(Str(curBarCode.DECSerialEnd)).PadLeft(7, "0") & "," & _
            curBarCode.QTY & "," & _
            curBarCode.WorkOrder & "," & _
            Now()

            PrintLine(1, sDump)
            'lstOutput.Items.Add(sDump)
        Catch ex As Exception
            MessageBox.Show("Problem with Summary dump file:" & vbCrLf & ex.ToString, "Summary Dump", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

        sDump = Nothing

        FileClose(1)
    End Sub
    Private Sub PrintLabelSummary()

        Dim LW As Object = CreateObject("Lworks3.LabelEngine")

        Try
            LW.FileName = "C:\Unicomp\Templates\SUMMARY.lw3"
            LW.Copies = 1
            LW.StartLabel = 1
            LW.TotalLabels = 1
            LW.UpdateSerials = False

            If Not NoPrint Then LW.PrintLabels()

            LW = Nothing
        Catch ex As System.Runtime.InteropServices.COMException
            MessageBox.Show("Error: " & "C:\Unicomp\Templates\SUMMARY.lw3" & " not found" & vbCrLf & "in C:\Unicomp\Templates\", "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub PrintLabels()

        Dim LW As Object
        Try
            LW = CreateObject("Lworks3.LabelEngine")

            'Open the label file we want to print

            LW.FileName = "C:\Unicomp\Templates\" & curBarCode.KeyboardLabelLayout & ".LW3"

            'Set up the label print job.

            LW.Copies = 1
            LW.StartLabel = 1
            LW.TotalLabels = curBarCode.QTY
            LW.UpdateSerials = False

            'Run the print job

            If Not NoPrint Then LW.PrintLabels()

            'Close down LabelWorks

            LW = Nothing
        Catch ex As System.Runtime.InteropServices.COMException
            MessageBox.Show("Error: " & curBarCode.KeyboardLabelLayout & ".LW3" & " not found" & vbCrLf & "in C:\Unicomp\Templates\", "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    



    Private Function PrintSample() As MsgBoxResult
        Dim LW As Object
        Try
            LW = CreateObject("Lworks3.LabelEngine")

            'Open the label file we want to print

            LW.FileName = "C:\Unicomp\Templates\" & curBarCode.KeyboardLabelLayout & ".LW3"

            'Set up the label print job.

            LW.Copies = 1
            LW.StartLabel = 1
            LW.TotalLabels = 1
            LW.UpdateSerials = False

            'Run the print job

            If Not NoPrint Then LW.PrintLabels()

            'Close down LabelWorks

            LW = Nothing
            PrintSample = MessageBox.Show("IS SAMPLE CORRECT?", "Sample Verify", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        Catch ex As System.Runtime.InteropServices.COMException
            MessageBox.Show("Error: " & curBarCode.KeyboardLabelLayout & ".LW3" & " not found" & vbCrLf & "in C:\Unicomp\Templates\", "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try


    End Function

#End Region

    Private Sub btnTestSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call CreateSummaryDump()
        Call PrintLabelSummary()
    End Sub


    Private Sub WriteToLog()
        Dim odcLogFile As System.Data.OleDb.OleDbConnection
        Dim odaKeyboardLog As System.Data.OleDb.OleDbDataAdapter
        Dim cmdKeyboardLog As System.Data.OleDb.OleDbCommandBuilder
        Dim dsKeyboardLog As DataSet
        Dim sqlStr As String
        Dim connStr As String
        Dim drNewRow As DataRow
        Dim x As Integer

        connStr = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Database Password=;Data Source=""C:\Unicomp\SerialLog.mdb"";Password=;Jet OLEDB:Engine Type=5;Jet OLEDB:Global Bulk Transactions=1;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLEDB:SFP=False;Extended Properties=;Mode=Share Deny None;Jet OLEDB:New Database Password=;Jet OLEDB:Create System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;User ID=Admin;Jet OLEDB:Encrypt Database=False"
        odcLogFile = New OleDbConnection(connStr)
        odcLogFile.Open()
        dsKeyboardLog = New DataSet
        sqlStr = "SELECT * FROM LogKeyboard"
        odaKeyboardLog = New OleDbDataAdapter(sqlStr, odcLogFile)

        cmdKeyboardLog = New OleDbCommandBuilder(odaKeyboardLog)
        dsKeyboardLog.Clear()
        odaKeyboardLog.Fill(dsKeyboardLog, "LogKeyboard")

        For x = 0 To curBarCode.QTY - 1
            drNewRow = dsKeyboardLog.Tables("LogKeyboard").NewRow

            drNewRow.Item("PartNumber") = curBarCode.UniPartNumber
            drNewRow.Item("DatePrinted") = Now
            drNewRow.Item("WorkOrder") = curBarCode.WorkOrder
            'drNewRow.Item("Quantity") = curBarCode.QTY
            drNewRow.Item("SerialNumber") = LTrim(Str(curBarCode.DECSerialStart + x)).PadLeft(7, "0")
            'drNewRow.Item("SerialNumberEnd") = LTrim(Str(curBarCode.DECSerialEnd)).PadLeft(7, "0")

            drNewRow.Item("Station") = "1"
            'drNewRow.Item("PrintReprint") = "Print"
            dsKeyboardLog.Tables("LogKeyboard").Rows.Add(drNewRow)
        Next
        odaKeyboardLog.Update(dsKeyboardLog, "LogKeyboard")



        'Dim drNewRow As DataRow
        'Dim drUpdateProfile As DataRow

        'DsLog1.Clear()
        'selLog.CommandText = "SELECT * FROM  LogKeyboard"
        'Try
        '    odaLog.Fill(DsLog1)
        'Catch ex As OleDb.OleDbException
        '    MessageBox.Show("Problem with C:\unicomp\Barcode Logs.mdb" & vbCrLf & "Be sure it exists", "Log File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '    Exit Sub
        'End Try

        'drNewRow = DsLog1.Tables("LogKeyboard").NewRow

        'drNewRow.Item("JobID") = Now.Ticks
        'drNewRow.Item("PartNumber") = curBarCode.UniPartNumber

        'drNewRow.Item("Date") = Now
        'drNewRow.Item("WorkOrder") = curBarCode.WorkOrder
        'drNewRow.Item("Quantity") = curBarCode.QTY
        'drNewRow.Item("SerialNumberStart") = curBarCode.SerialStart
        'drNewRow.Item("SerialNumberEnd") = curBarCode.SerialEnd

        'drNewRow.Item("Station") = "0"
        'drNewRow.Item("PrintReprint") = "Print"

        'DsLog1.Tables("LogKeyboard").Rows.Add(drNewRow)
        'odaLog.Update(DsLog1)

        'Try
        '    DsProfileMaster1.Clear()
        '    selProfileMaster.CommandText = "SELECT * FROM ProfileMaster WHERE PartNumber = '" & curBarCode.UniPartNumber & "'"

        '    odaProfileMaster.Fill(DsProfileMaster1)

        '    If DsProfileMaster1.Tables("ProfileMaster").Rows.Count = 1 Then

        '        drUpdateProfile = DsProfileMaster1.Tables("ProfileMaster").Rows(0)
        '        'MsgBox(drUpdateProfile.Item("DateLastPrinted"))
        '        drUpdateProfile.Item("DateLastPrinted") = Now
        '        'MsgBox(drUpdateProfile.Item("DateLastPrinted"))
        '        'dsProfileMaster1.Tables("ProfileMaster").Rows(0).Item("DateLastPrinted") = Now
        '        odaProfileMaster.Update(DsProfileMaster1)
        '    Else
        '        MessageBox.Show("P/N: " & curBarCode.UniPartNumber & " Not in Master Profile database", "Not in Master Profile", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '    End If
        'Catch ex As OleDb.OleDbException
        '    MessageBox.Show("Error with Profile Master." & vbCrLf & "Check Network", "Profile Master Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        'Catch ex As Exception
        '    MsgBox(ex.ToString)
        'End Try
    End Sub


    Private Sub WriteToKBLogSQL()
        Dim odcLogFile As System.Data.OleDb.OleDbConnection
        Dim odaKeyboardLog As System.Data.OleDb.OleDbDataAdapter
        Dim cmdKeyboardLog As System.Data.OleDb.OleDbCommandBuilder
        Dim dsKeyboardLog As DataSet
        Dim sqlStr As String
        Dim connStr As String
        Dim drNewRow As DataRow
        Dim x As Integer

        connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        odcLogFile = New OleDbConnection(connStr)
        odcLogFile.Open()
        dsKeyboardLog = New DataSet
        sqlStr = "SELECT * FROM LOG_Keyboard"
        odaKeyboardLog = New OleDbDataAdapter(sqlStr, odcLogFile)

        cmdKeyboardLog = New OleDbCommandBuilder(odaKeyboardLog)
        dsKeyboardLog.Clear()
        odaKeyboardLog.Fill(dsKeyboardLog, "LOG_Keyboard")

        For x = 0 To curBarCode.QTY - 1
            drNewRow = dsKeyboardLog.Tables("LOG_Keyboard").NewRow

            drNewRow.Item("PartNumber") = curBarCode.UniPartNumber
            drNewRow.Item("DatePrinted") = Now
            drNewRow.Item("WorkOrder") = curBarCode.WorkOrder
            'drNewRow.Item("Quantity") = curBarCode.QTY
            drNewRow.Item("SerialNumber") = LTrim(Str(curBarCode.DECSerialStart + x)).PadLeft(7, "0")
            'drNewRow.Item("SerialNumberEnd") = LTrim(Str(curBarCode.DECSerialEnd)).PadLeft(7, "0")

            '            drNewRow.Item("Station") = "1"
            'drNewRow.Item("PrintReprint") = "Print"
            dsKeyboardLog.Tables("LOG_Keyboard").Rows.Add(drNewRow)
        Next
        odaKeyboardLog.Update(dsKeyboardLog, "LOG_Keyboard")
        odcLogFile.Close()
    End Sub

    Private Sub WriteToWorkOrderLogSQL()
        Dim odcLogFile As System.Data.OleDb.OleDbConnection
        Dim odaWorkOrderLog As System.Data.OleDb.OleDbDataAdapter
        Dim cmdWorkOrderLog As System.Data.OleDb.OleDbCommandBuilder
        Dim dsWorkOrderLog As DataSet
        Dim sqlStr As String
        Dim connStr As String
        Dim drNewRow As DataRow
        Dim x As Integer

        connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        odcLogFile = New OleDbConnection(connStr)
        odcLogFile.Open()
        dsWorkOrderLog = New DataSet
        sqlStr = "SELECT * FROM LOG_WorkOrders"
        odaWorkOrderLog = New OleDbDataAdapter(sqlStr, odcLogFile)

        cmdWorkOrderLog = New OleDbCommandBuilder(odaWorkOrderLog)
        dsWorkOrderLog.Clear()
        odaWorkOrderLog.Fill(dsWorkOrderLog, "LOG_WorkOrders")
        '
        drNewRow = dsWorkOrderLog.Tables("LOG_WorkOrders").NewRow

        drNewRow.Item("WorkOrder") = curBarCode.WorkOrder
        drNewRow.Item("PartNumber") = curBarCode.UniPartNumber
        drNewRow.Item("SerialStart") = curBarCode.DECSerialStart
        drNewRow.Item("SerialEnd") = curBarCode.DECSerialEnd
        drNewRow.Item("QTYRequired") = curBarCode.QTY
        drNewRow.Item("QTYBoxed") = 0
        'drNewRow.Item("DateComplete") = Now

        dsWorkOrderLog.Tables("LOG_WorkOrders").Rows.Add(drNewRow)
        odaWorkOrderLog.Update(dsWorkOrderLog, "LOG_WorkOrders")
        odcLogFile.Close()
    End Sub

    Private Sub UpdateMaster()
        Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
        Dim odaFamily As System.Data.OleDb.OleDbDataAdapter
        Dim cmdFamily As System.Data.OleDb.OleDbCommandBuilder
        Dim dsFamily As DataSet
        Dim sqlStr As String
        Dim drFamily As DataRow

        'ALLIANCE FOR WORK ORDER LIST WITH DESCRIPTION
        Dim connStr As String
        Dim sWorkOrder As String

        Dim sSerial As String
        Dim sFamily As String
        Dim sSerialType As String
        Dim drUpdatedRow As DataRow

        connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        odcUnicompMain = New OleDbConnection(connStr)
        odcUnicompMain.Open()
        dsFamily = New DataSet

        sqlStr = "SELECT * FROM Family WHERE FamilyCode = '" & curBarCode.Family & "'"
        odaFamily = New OleDbDataAdapter(sqlStr, odcUnicompMain)

        cmdFamily = New OleDbCommandBuilder(odaFamily)
        dsFamily.Clear()
        odaFamily.Fill(dsFamily, "Family")
        If dsFamily.Tables("Family").Rows.Count = 1 Then
            dsFamily.Tables("Family").Rows(0).Item("NextSerialNumber") = LTrim(Str(curBarCode.DECSerialEnd + 1)).PadLeft(7, "0")
            dsFamily.Tables("Family").Rows(0).Item("DateLastUsed") = Now
            'dsProfile.Tables("qProfile").Rows
            odaFamily.Update(dsFamily, "Family")
        End If
    End Sub

    Private Sub btnTESTWriteLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call WriteToLog()
    End Sub
    Private Sub CleanUp()
        curBarCode = Nothing
        rdoModeNormal.Checked = True
        If Not KeepDump Then Call CreateTemplateDumpDB("Clear")
        Call ResetAll()
    End Sub

    Private Sub btnCustomGO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCustomGO.Click

        Try
            Custom = True
            curBarCode.UniPartNumber = UCase(txtCustomPN.Text)
            curBarCode.DECSerialStart = UCase(txtCustomStartSN.Text)
            curBarCode.QTY = Val(txtCustomQTY.Text)

            If curBarCode.QTY <= 0 Then
                MessageBox.Show("Quantity cannot be Less than 1!", "QTY Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Call ResetAll()
                Exit Sub
            End If

            'If Len(curBarCode.DECSerialStart) <> 7 Then
            '    MessageBox.Show("Serial Number must be 7 digits", "SN Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Call ResetAll()
            '    Exit Sub
            'End If
            curBarCode.WorkOrder = "CUSTOM"

            PNDigits = Len(curBarCode.UniPartNumber)
            curBarCode.MasterDataRow = LookUpInMasterProfileQuery(curBarCode.UniPartNumber)
            ParseMasterDatarow(curBarCode.MasterDataRow)
            If IsNothing(curBarCode.MasterDataRow) Then
                MessageBox.Show("Part Number: " & curBarCode.UniPartNumber & " Not in Database", "Not in database", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                'ResetAll()
            Else

                If Not InFamilyExclusions() Then
                    If curBarCode.OEMPartNumber = "0" Then
                        MessageBox.Show("P/N: " & curBarCode.UniPartNumber & " Should have an OEM P/N" & vbCrLf & "See Chuck", "No OEM P/N Listed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        ResetAll()
                        Exit Sub
                    Else
                        curBarCode.DECSerialStart = Val(txtCustomStartSN.Text)
                        curBarCode.DECSerialEnd = curBarCode.DECSerialStart + curBarCode.QTY - 1
                        txtCustomOEMPN.Text = curBarCode.OEMPartNumber
                        Call CreateTemplateDumpDB("Full")
                        Call PrintLabels()
                        '            'Call GenerateSampleDump()
                        '            'If PrintSample() = MsgBoxResult.Yes Then

                        Call CreateSummaryDump()
                        Call PrintLabelSummary()
                        '            'Else
                        '            ' Call ResetAll()
                        '            'End If
                    End If
                Else
                    MessageBox.Show("Family cannot be processed." & vbCrLf & "Part Number: " & curBarCode.UniPartNumber & vbCrLf & "Family: " & curBarCode.Family & vbCrLf & "Reason: " & curBarCode.ReasonCannotPrint)
                    Call ResetAll()
                End If

                'Else
                '    MsgBox("PART HAS NOT BEEN PRINTED BEFORE." & vbCrLf & "PLEASE ENTER INFO FOR FIRST RUN")
                '    Call AddToMaster()
            End If
            rdoModeNormal.Checked = True
            Call ResetAll()
        Catch ex As Exception
            MessageBox.Show("General Error in Print_Custom suboutine", "General Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Function InFamilyExclusions() As Boolean

        If OverrideExclusions Then
            InFamilyExclusions = False
            Exit Function
        End If


        'returns datarow from master query
        Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
        Dim odaFamilyExclusions As System.Data.OleDb.OleDbDataAdapter
        Dim cmdFamilyExclusions As System.Data.OleDb.OleDbCommandBuilder
        Dim dsFamilyExclusions As DataSet
        Dim sqlStr As String
        'Dim drProfile As DataRow

        Dim connStr As String

        connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        odcUnicompMain = New OleDbConnection(connStr)
        odcUnicompMain.Open()
        dsFamilyExclusions = New DataSet

        sqlStr = "SELECT * FROM FamilyExclusion WHERE FamilyToExclude = '" & curBarCode.Family & "'"
        odaFamilyExclusions = New OleDbDataAdapter(sqlStr, odcUnicompMain)

        cmdFamilyExclusions = New OleDbCommandBuilder(odaFamilyExclusions)
        dsFamilyExclusions.Clear()

        odaFamilyExclusions.Fill(dsFamilyExclusions, "FamilyExclusion")

        If dsFamilyExclusions.Tables("FamilyExclusion").Rows.Count = 1 Then
            InFamilyExclusions = True
            curBarCode.ReasonCannotPrint = RTrim(dsFamilyExclusions.Tables("FamilyExclusion").Rows(0).Item("Reason"))
        End If

        odcUnicompMain.Close()

    End Function

    Private Function InExlusions() As Boolean
        ''InExlusions = True

        'If OverrideExclusions Then
        '    InExlusions = False
        '    Exit Function
        'End If

        'DsExlusions1.Clear()
        'selExclusions.CommandText = "SELECT * FROM  KBProfilesToExclude WHERE KBLayoutID = '" & curBarCode.KBProfile & "'"
        'Try
        '    odaExlusions.Fill(DsExlusions1)
        'Catch ex As OleDb.OleDbException
        '    MessageBox.Show("Problem with exlusions.mdb" & vbCrLf & "Be sure it exists", "Log File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '    InExlusions = True
        '    Exit Function
        'End Try

        'If DsExlusions1.Tables("KBProfilesToExclude").Rows.Count = 1 Then
        '    InExlusions = True
        '    sExcludedReason = DsExlusions1.Tables("KBProfilesToExclude").Rows(0).Item("Reason")
        'Else
        '    InExlusions = False
        'End If

    End Function

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Custom = False
        If curBarCode.UniPartNumber = "" Then Exit Sub

        Call CreateTemplateDumpDB("Sample")
        If PrintSample() = MsgBoxResult.Yes Then
            Call CreateTemplateDumpDB("Full")
            Call PrintLabels()

            If MessageBox.Show("Did Labels Print OK?", "Verify Print", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                curBarCode.DECSerialEnd = curBarCode.DECSerialStart + curBarCode.QTY - 1
                If Not Custom Then
                    Call CreateSummaryDump()
                    Call PrintLabelSummary()
                    If Not NoLogging Then
                        Call WriteToKBLogSQL()
                        Call WriteToWorkOrderLogSQL()
                        Call UpdateMaster()
                    End If
                End If


                Call CleanUp()
                Call ResetAll()
            Else
                Call ResetAll()
            End If
        Else
            MsgBox("FAIL SAMPLE")
            Call ResetAll()
        End If

    End Sub


#Region "Interface Stuff"
    Private Sub AddToMaster()   'Add to master based completely on guess
        Dim frmNewPN As New AddToMaster


        frmNewPN.ShowDialog()
        Call ResetAll()
    End Sub

    Private Sub rdoModeNormal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoModeNormal.CheckedChanged
        Call ResetAll()
        rdoModeNormal.TabStop = False
        If rdoModeNormal.Checked = True Then
            grpNormal.Enabled = True
            grpCustom.Enabled = False

            btnPrint.Enabled = True
            txtWorkOrder.Focus()


        Else
            grpNormal.Enabled = False
            grpCustom.Enabled = True
            txtCustomPN.Focus()
            btnPrint.Enabled = False
        End If
    End Sub

    Private Sub rdoModeCustom_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoModeCustom.CheckedChanged
        Call ResetAll()
        rdoModeCustom.TabStop = False
        If rdoModeCustom.Checked = True Then
            grpNormal.Enabled = False
            grpCustom.Enabled = True
            txtCustomPN.Focus()
            btnPrint.Enabled = False
            txtCustomPN.Focus()
        Else
            grpNormal.Enabled = True
            grpCustom.Enabled = False
            btnPrint.Enabled = True
            txtWorkOrder.Focus()
        End If
    End Sub

    Private Sub btnTESTPrintSample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Call GenerateSampleDump()
        If PrintSample() = MsgBoxResult.Yes Then
            MsgBox("GOOD, OK TO PRINT")
        Else
            'MsgBox("FAIL SAMPLE")
            Call ResetAll()
        End If
    End Sub

    Private Sub ResetAll()

        txtWorkOrder.Clear()
        txtWOPN.Clear()
        txtQTY.Clear()
        txtRevLevel.Clear()
        txtStartingSerial.Clear()
        txtDesc.Clear()
        txtOEMPN.Clear()
        lblStartDate.Text = ""

        curBarCode = Nothing
        curWorkOrder = Nothing
        sExcludedReason = Nothing
        PNDigits = 0
        'BarcodeFieldType = Nothing
        'BarcodeFieldLength = Nothing
        Abort = False
        'MsgBox("RESET")
        '
        txtCustomPN.Clear()
        txtCustomQTY.Clear()
        txtCustomStartSN.Clear()

        txtCustomPN.Clear()
        txtCustomStartSN.Clear()
        txtCustomOEMPN.Clear()
        txtCustomQTY.Clear()

        txtWorkOrder.Focus()

        Call ReadConfig()
    End Sub

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        Dim frmNewPN As New AddToMaster


        frmNewPN.ShowDialog()
        Call ResetAll()
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Call CleanUp()
    End Sub

#End Region


    Private Sub txtWorkOrder_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWorkOrder.TextChanged
        txtWOPN.Clear()
        txtQTY.Clear()
        lblStartDate.Text = ""
        txtDesc.Clear()
        txtRevLevel.Clear()
        curBarCode = Nothing
        curWorkOrder = Nothing
        txtStartingSerial.Clear()
    End Sub


#Region "v2.26 Routines"
    'Private Sub GenerateSampleDump()
    '    Dim sSerial As String
    '    Dim sDataLine As String
    '    ' Select Case curBarCode.KBProfile
    '    ' Case "UNI"

    '    FileOpen(1, "C:\Unicomp\Templates\DUMPFILE.csv", OpenMode.Output)
    '    PrintLine(1, "LabelNumber,UniPN,OEMPN,Serial,MainBarcode,Sample,Date")
    '    sSerial = "XXXXXXX"
    '    sDataLine = "1," & curBarCode.UniPartNumber & "," & _
    '    curBarCode.OEMPartNumber & "," & _
    '    sSerial & "," & _
    '    curBarCode.UniPartNumber
    '    Select Case PNDigits
    '        Case 0
    '            MsgBox("FAILURE")
    '        Case 7
    '            sDataLine = sDataLine & "--"
    '        Case 8
    '            sDataLine = sDataLine & "-"

    '    End Select


    '    sDataLine = sDataLine & sSerial & "," & _
    '    "SAMPLE DO NOT USE" & "," & _
    '    Now.Date()
    '    PrintLine(1, sDataLine)
    '    FileClose(1)


    '    'End Select
    'End Sub

    'Private Sub GenerateDumpFile()
    '    'LabelNumber,UniPN,OEMPN,Serial,MainBarcode,Sample,Date

    '    Dim i As Integer
    '    Dim sDataLine As String
    '    Dim sSerial As String
    '    FileOpen(1, "C:\Unicomp\Templates\DUMPFILE.csv", OpenMode.Output)
    '    PrintLine(1, "LabelNumber,UniPN,OEMPN,Serial,MainBarcode,Sample,Date")
    '    'LabelNumber,UniPN,OEMPN,Serial,MainBarcode,Sample,Date
    '    sSerial = curBarCode.SerialStart
    '    For i = 1 To curBarCode.QTY
    '        'sSerial = AddOneToBase36(sSerial)
    '        sDataLine = Str(i) & "," & curBarCode.UniPartNumber & "," & _
    '        curBarCode.OEMPartNumber & "," & _
    '        sSerial & "," & _
    '        curBarCode.UniPartNumber
    '        Select Case PNDigits
    '            Case 0
    '                MsgBox("FAILURE")
    '            Case 7
    '                sDataLine = sDataLine & "--"
    '            Case 8
    '                sDataLine = sDataLine & "-"

    '        End Select


    '        sDataLine = sDataLine & sSerial & ",," & _
    '        Now.Date()
    '        sSerial = AddOneToBase36(sSerial)
    '        'Debug.Write(sDataLine & vbCrLf)
    '        PrintLine(1, sDataLine)
    '    Next
    '    curBarCode.SerialEnd = AddToBase36(sSerial, -1)
    '    Debug.Write("Range: " & curBarCode.SerialStart & " To " & curBarCode.SerialEnd & vbCrLf)
    '    FileClose(1)
    '    'If Not Custom Then Call WriteNextSerial(sSerial)

    '    ''PrintLine(1, "UniPN,OEMPN,SerialStart,SerialEnd,QTY,WorkOrder,Date")
    '    'Select Case curBarCode.KBProfile
    '    '    Case "UNI"
    '    '        FileOpen(1, "C:\Unicomp\Templates\DUMPFILE.csv", OpenMode.Output)
    '        PrintLine(1, "LabelNumber,UniPN,OEMPN,Serial,MainBarcode,Sample,Date")
    '        'LabelNumber,UniPN,OEMPN,Serial,MainBarcode,Sample,Date
    '        sSerial = curBarCode.SerialStart
    '        For i = 1 To curBarCode.QTY
    '            'sSerial = AddOneToBase36(sSerial)
    '            sDataLine = Str(i) & "," & curBarCode.UniPartNumber & "," & _
    '            curBarCode.OEMPartNumber & "," & _
    '            sSerial & "," & _
    '            curBarCode.UniPartNumber
    '            Select Case PNDigits
    '                Case 0
    '                    MsgBox("FAILURE")
    '                Case 7
    '                    sDataLine = sDataLine & "--"
    '                Case 8
    '                    sDataLine = sDataLine & "-"

    '            End Select


    '            sDataLine = sDataLine & sSerial & ",," & _
    '            Now.Date()
    '            sSerial = AddOneToBase36(sSerial)
    '            'Debug.Write(sDataLine & vbCrLf)
    '            PrintLine(1, sDataLine)
    '        Next
    '        curBarCode.SerialEnd = AddToBase36(sSerial, -1)
    '        Debug.Write("Range: " & curBarCode.SerialStart & " To " & curBarCode.SerialEnd & vbCrLf)
    '        FileClose(1)
    '        If Not Custom Then Call WriteNextSerial(sSerial)
    '    Case "ACP"
    '        FileOpen(1, "C:\Unicomp\Templates\DUMPFILE.csv", OpenMode.Output)
    '        PrintLine(1, "LabelNumber,UniPN,OEMPN,Serial,MainBarcode,Sample,Date")
    '        'LabelNumber,UniPN,OEMPN,Serial,MainBarcode,Sample,Date
    '        sSerial = curBarCode.SerialStart
    '        For i = 1 To curBarCode.QTY
    '            'sSerial = AddOneToBase36(sSerial)
    '            sDataLine = Str(i) & "," & curBarCode.UniPartNumber & "," & _
    '            curBarCode.OEMPartNumber & "," & _
    '            sSerial & "," & _
    '            curBarCode.UniPartNumber
    '            Select Case PNDigits
    '                Case 0
    '                    MsgBox("FAILURE")
    '                Case 7
    '                    sDataLine = sDataLine & "--"
    '                Case 8
    '                    sDataLine = sDataLine & "-"

    '            End Select


    '            sDataLine = sDataLine & sSerial & ",," & _
    '            Now.Date()
    '            sSerial = AddOneToBase36(sSerial)
    '            'Debug.Write(sDataLine & vbCrLf)
    '            PrintLine(1, sDataLine)
    '        Next
    '        curBarCode.SerialEnd = AddToBase36(sSerial, -1)
    '        Debug.Write("Range: " & curBarCode.SerialStart & " To " & curBarCode.SerialEnd & vbCrLf)
    '        FileClose(1)
    '        If Not Custom Then Call WriteNextSerial(sSerial)

    '    Case Else
    '        MessageBox.Show("ERROR IN DATABASE" & vbCrLf & "No select for " & curBarCode.KBProfile)
    'End Select
    'End Sub
    'Private Function LookupInMaster(ByVal sPN As String) As Boolean
    '    Dim odcUnicompMain As System.Data.OleDb.OleDbConnection
    '    Dim odaProfiles As System.Data.OleDb.OleDbDataAdapter
    '    Dim cmdProfiles As System.Data.OleDb.OleDbCommandBuilder
    '    Dim dsProfile As DataSet
    '    Dim sqlStr As String
    '    Dim drProfile As DataRow

    '    'ALLIANCE FOR WORK ORDER LIST WITH DESCRIPTION
    '    'Dim i As Integer
    '    Dim connStr As String
    '    Dim sWorkOrder As String

    '    Dim sSerial As String
    '    Dim sFamily As String
    '    Dim sSerialType As String

    '    'sWorkOrder = LTrim(RTrim(txtWorkOrder.Text))

    '    connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
    '    odcUnicompMain = New OleDbConnection(connStr)
    '    odcUnicompMain.Open()
    '    dsProfile = New DataSet

    '    'sqlStr = "SELECT * FROM WOHeader"
    '    sqlStr = "SELECT * FROM qProfile WHERE PartNumber = '" & curBarCode.UniPartNumber & "'"
    '    odaProfiles = New OleDbDataAdapter(sqlStr, odcUnicompMain)
    '    'odaWorkOrders.SelectCommand.CommandText = sqlStr

    '    cmdProfiles = New OleDbCommandBuilder(odaProfiles)
    '    dsProfile.Clear()
    '    odaProfiles.Fill(dsProfile, "qProfile")

    '    If dsProfile.Tables("qProfile").Rows.Count = 1 Then
    '        drProfile = dsProfile.Tables("qProfile").Rows(0)
    '        LookupInMaster = True
    '        'curBarCode.UniPartNumber = curWorkOrder.PartNumber
    '        'curBarCode.SerialStart = IIf(IsDBNull(drMaster.Item("NextSerialNumber")), "0", drMaster.Item("NextSerialNumber"))
    '        curBarCode.OEMPartNumber = RTrim(IIf(IsDBNull(drProfile.Item("OEMPartNumber")), curBarCode.UniPartNumber, drProfile.Item("OEMPartNumber")))

    '        curBarCode.DateLastPrinted = IIf(IsDBNull(drProfile.Item("DateLastPrinted")), vbNull, drProfile.Item("DateLastPrinted"))
    '        curBarCode.KBProfile = RTrim(IIf(IsDBNull(drProfile.Item("KeyboardLabelTemplate")), "0", drProfile.Item("KeyboardLabelTemplate")))
    '        curBarCode.BXProfile = RTrim(IIf(IsDBNull(drProfile.Item("BoxLabelTemplate")), "0", drProfile.Item("BoxLabelTemplate")))
    '        curBarCode.KeyboardLabelLayout = RTrim(IIf(IsDBNull(drProfile.Item("KBTemplateFilename")), "NOTFOUND", drProfile.Item("KBTemplateFilename")))
    '        curBarCode.BoxLabelLayout = RTrim(IIf(IsDBNull(drProfile.Item("BoxTemplateFilename")), "NOTFOUND", drProfile.Item("BoxTemplateFilename")))
    '        curBarCode.Family = IIf(IsDBNull(drProfile.Item("FamilyCode")), "GM", drProfile.Item("FamilyCode"))

    '        'DO SERIAL NUMBER
    '        sSerial = RTrim(IIf(IsDBNull(drProfile.Item("NextSerialNumber")), "1", drProfile.Item("NextSerialNumber")))
    '        sSerialType = RTrim(IIf(IsDBNull(drProfile.Item("SerialType")), "D7_OEM", drProfile.Item("SerialType")))
    '        Select Case sSerialType
    '            Case "D7_OEM"
    '                'MsgBox("DIGIT7")
    '                curBarCode.SNType = "D7_OEM"
    '                curBarCode.SerialStart = sSerial.PadLeft(7, "0")
    '            Case "A7_UNI"
    '                MsgBox("ALPHA")
    '        End Select


    '    Else
    '        drProfile = Nothing
    '        LookupInMaster = False
    '    End If
    '    odcUnicompMain.Close()

    'End Function

#End Region


    Private Sub btnDebug_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDebug.Click
        Dim frmNewPN As New frmDebug


        frmNewPN.ShowDialog()
        Call ResetAll()
    End Sub
End Class
