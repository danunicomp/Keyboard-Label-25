Imports System
Imports System.IO
Imports System.Text

Imports System.Data
Imports System.Data.OleDb
Public Class AddToMaster
    Inherits System.Windows.Forms.Form

    Private Structure MasterFields
        Dim UniPN As String
        Dim OEMPN As String
        Dim CustomerID As String
        Dim SerialStart As String
        Dim FormatType As String
    End Structure

    Private Structure Profile
        Dim PartNumber As String
        Dim Family As String
        Dim OEMPartNumber As String
        Dim Description As String
        Dim KBTemplate As String
        Dim BOXTemplate As String
        Dim DateAdded As Date
        Dim DateLastPrinted As Date
        Dim StartingSerial As String
    End Structure
    Dim newProfile As Profile
    Dim ABORT As Boolean

    Private curNewMaster As MasterFields

    Dim odcUnicompMain As System.Data.OleDb.OleDbConnection

    Dim odaFamily As System.Data.OleDb.OleDbDataAdapter
    Dim cmdFamily As System.Data.OleDb.OleDbCommandBuilder
    Dim dsFamily As DataSet

    Dim odaBoxTemplate As System.Data.OleDb.OleDbDataAdapter
    Dim cmdBoxTemplates As System.Data.OleDb.OleDbCommandBuilder
    Dim dsBoxTemplates As DataSet

    Dim odaKBTemplate As System.Data.OleDb.OleDbDataAdapter
    Dim cmdKBTemplates As System.Data.OleDb.OleDbCommandBuilder
    Dim dsKBTemplates As DataSet




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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtNewPN As System.Windows.Forms.TextBox
    Friend WithEvents txtNewOEMPN As System.Windows.Forms.TextBox
    Friend WithEvents txtNewStartSN As System.Windows.Forms.TextBox
    Friend WithEvents chkHasOEMPN As System.Windows.Forms.CheckBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSample As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtKBLabelFN As System.Windows.Forms.TextBox
    Friend WithEvents txtBoxLabelFN As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtBoxSample As System.Windows.Forms.TextBox
    Friend WithEvents cmbProductFamily As System.Windows.Forms.ComboBox
    Friend WithEvents txtAllianceDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cmbKBProfile As System.Windows.Forms.ComboBox
    Friend WithEvents cmbBoxProfile As System.Windows.Forms.ComboBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents odcAddProfile As System.Data.OleDb.OleDbConnection
    Friend WithEvents odaAddProfile As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents selAddProfile As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    'Friend WithEvents DsAddProfile1 As Unicomp_Barcode_System.dsAddProfile
    Friend WithEvents odcProductCodes As System.Data.OleDb.OleDbConnection
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents odaProductCodes As System.Data.OleDb.OleDbDataAdapter
    'Friend WithEvents DsPC1 As Unicomp_Barcode_System.dsPC
    Friend WithEvents odaKeyboardProfiles As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents selKBProfile As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand3 As System.Data.OleDb.OleDbCommand
    'Friend WithEvents DsKBProfiles1 As Unicomp_Barcode_System.dsKBProfiles
    Friend WithEvents odaBoxProfiles As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents selBoxProfile As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand4 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand4 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand4 As System.Data.OleDb.OleDbCommand
    'Friend WithEvents DsBoxProfiles1 As Unicomp_Barcode_System.dsBoxProfiles
    Friend WithEvents odaAMSPArtMaster As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbInsertCommand5 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand5 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand5 As System.Data.OleDb.OleDbCommand
    Friend WithEvents odcAlliancePartMaster As System.Data.OleDb.OleDbConnection
    'Friend WithEvents DsAMSPartMaster1 As Unicomp_Barcode_System.dsAMSPartMaster
    Friend WithEvents selAMSPArtMaster As System.Data.OleDb.OleDbCommand
    Friend WithEvents odaStructureID As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents selStructureID As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand6 As System.Data.OleDb.OleDbCommand
    'Friend WithEvents DsPNStructureID1 As Unicomp_Barcode_System.dsPNStructureID
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(AddToMaster))
        Me.txtNewPN = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtNewOEMPN = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtNewStartSN = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.chkHasOEMPN = New System.Windows.Forms.CheckBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtSample = New System.Windows.Forms.TextBox
        Me.txtKBLabelFN = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtBoxLabelFN = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtBoxSample = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cmbProductFamily = New System.Windows.Forms.ComboBox
        Me.txtAllianceDescription = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmbKBProfile = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cmbBoxProfile = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.odcAddProfile = New System.Data.OleDb.OleDbConnection
        Me.odaAddProfile = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.selAddProfile = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.odcProductCodes = New System.Data.OleDb.OleDbConnection
        Me.odaProductCodes = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand2 = New System.Data.OleDb.OleDbCommand
        Me.odaKeyboardProfiles = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand3 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand3 = New System.Data.OleDb.OleDbCommand
        Me.selKBProfile = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand3 = New System.Data.OleDb.OleDbCommand
        Me.odaBoxProfiles = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand4 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand4 = New System.Data.OleDb.OleDbCommand
        Me.selBoxProfile = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand4 = New System.Data.OleDb.OleDbCommand
        Me.odaAMSPArtMaster = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand5 = New System.Data.OleDb.OleDbCommand
        Me.odcAlliancePartMaster = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand5 = New System.Data.OleDb.OleDbCommand
        Me.selAMSPArtMaster = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand5 = New System.Data.OleDb.OleDbCommand
        Me.odaStructureID = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbInsertCommand6 = New System.Data.OleDb.OleDbCommand
        Me.selStructureID = New System.Data.OleDb.OleDbCommand
        Me.SuspendLayout()
        '
        'txtNewPN
        '
        Me.txtNewPN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewPN.Location = New System.Drawing.Point(176, 40)
        Me.txtNewPN.Name = "txtNewPN"
        Me.txtNewPN.Size = New System.Drawing.Size(112, 26)
        Me.txtNewPN.TabIndex = 0
        Me.txtNewPN.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(64, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 22)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Part Number:"
        '
        'txtNewOEMPN
        '
        Me.txtNewOEMPN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewOEMPN.Location = New System.Drawing.Point(176, 112)
        Me.txtNewOEMPN.Name = "txtNewOEMPN"
        Me.txtNewOEMPN.Size = New System.Drawing.Size(160, 26)
        Me.txtNewOEMPN.TabIndex = 2
        Me.txtNewOEMPN.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Enabled = False
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(72, 192)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(99, 22)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Starting S/N:"
        '
        'txtNewStartSN
        '
        Me.txtNewStartSN.Enabled = False
        Me.txtNewStartSN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewStartSN.Location = New System.Drawing.Point(176, 192)
        Me.txtNewStartSN.Name = "txtNewStartSN"
        Me.txtNewStartSN.TabIndex = 4
        Me.txtNewStartSN.Text = ""
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(0, 152)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(164, 22)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Product Family Code:"
        '
        'chkHasOEMPN
        '
        Me.chkHasOEMPN.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkHasOEMPN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkHasOEMPN.Location = New System.Drawing.Point(24, 112)
        Me.chkHasOEMPN.Name = "chkHasOEMPN"
        Me.chkHasOEMPN.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkHasOEMPN.Size = New System.Drawing.Size(144, 24)
        Me.chkHasOEMPN.TabIndex = 11
        Me.chkHasOEMPN.Text = "Has OEM P/N?"
        Me.chkHasOEMPN.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(104, 288)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(66, 22)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Sample:"
        '
        'txtSample
        '
        Me.txtSample.BackColor = System.Drawing.SystemColors.Control
        Me.txtSample.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSample.Location = New System.Drawing.Point(176, 288)
        Me.txtSample.Name = "txtSample"
        Me.txtSample.Size = New System.Drawing.Size(312, 26)
        Me.txtSample.TabIndex = 13
        Me.txtSample.Text = ""
        '
        'txtKBLabelFN
        '
        Me.txtKBLabelFN.BackColor = System.Drawing.SystemColors.Control
        Me.txtKBLabelFN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtKBLabelFN.Location = New System.Drawing.Point(176, 320)
        Me.txtKBLabelFN.Name = "txtKBLabelFN"
        Me.txtKBLabelFN.Size = New System.Drawing.Size(80, 20)
        Me.txtKBLabelFN.TabIndex = 15
        Me.txtKBLabelFN.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(64, 320)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(103, 16)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "KB Label Filename:"
        '
        'txtBoxLabelFN
        '
        Me.txtBoxLabelFN.BackColor = System.Drawing.SystemColors.Control
        Me.txtBoxLabelFN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBoxLabelFN.Location = New System.Drawing.Point(176, 424)
        Me.txtBoxLabelFN.Name = "txtBoxLabelFN"
        Me.txtBoxLabelFN.Size = New System.Drawing.Size(80, 20)
        Me.txtBoxLabelFN.TabIndex = 17
        Me.txtBoxLabelFN.Text = ""
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(64, 424)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(107, 16)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Box Label Filename:"
        '
        'txtBoxSample
        '
        Me.txtBoxSample.BackColor = System.Drawing.SystemColors.Control
        Me.txtBoxSample.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBoxSample.Location = New System.Drawing.Point(176, 392)
        Me.txtBoxSample.Name = "txtBoxSample"
        Me.txtBoxSample.Size = New System.Drawing.Size(312, 26)
        Me.txtBoxSample.TabIndex = 21
        Me.txtBoxSample.Text = ""
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(104, 392)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(66, 22)
        Me.Label8.TabIndex = 20
        Me.Label8.Text = "Sample:"
        '
        'cmbProductFamily
        '
        Me.cmbProductFamily.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbProductFamily.Location = New System.Drawing.Point(176, 152)
        Me.cmbProductFamily.Name = "cmbProductFamily"
        Me.cmbProductFamily.Size = New System.Drawing.Size(312, 28)
        Me.cmbProductFamily.TabIndex = 3
        '
        'txtAllianceDescription
        '
        Me.txtAllianceDescription.BackColor = System.Drawing.SystemColors.Control
        Me.txtAllianceDescription.Location = New System.Drawing.Point(176, 80)
        Me.txtAllianceDescription.Name = "txtAllianceDescription"
        Me.txtAllianceDescription.Size = New System.Drawing.Size(312, 20)
        Me.txtAllianceDescription.TabIndex = 23
        Me.txtAllianceDescription.Text = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(72, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(93, 22)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Description:"
        '
        'cmbKBProfile
        '
        Me.cmbKBProfile.Enabled = False
        Me.cmbKBProfile.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbKBProfile.Location = New System.Drawing.Point(176, 248)
        Me.cmbKBProfile.Name = "cmbKBProfile"
        Me.cmbKBProfile.Size = New System.Drawing.Size(312, 28)
        Me.cmbKBProfile.TabIndex = 5
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(40, 256)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(128, 22)
        Me.Label9.TabIndex = 25
        Me.Label9.Text = "KB Label Profile:"
        '
        'cmbBoxProfile
        '
        Me.cmbBoxProfile.Enabled = False
        Me.cmbBoxProfile.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbBoxProfile.Location = New System.Drawing.Point(176, 352)
        Me.cmbBoxProfile.Name = "cmbBoxProfile"
        Me.cmbBoxProfile.Size = New System.Drawing.Size(312, 28)
        Me.cmbBoxProfile.TabIndex = 6
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(32, 360)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(135, 22)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "Box Label Profile:"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(424, 456)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 40)
        Me.btnAdd.TabIndex = 29
        Me.btnAdd.Text = "ADD"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(320, 456)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 40)
        Me.btnCancel.TabIndex = 30
        Me.btnCancel.Text = "Cancel"
        '
        'odcAddProfile
        '
        Me.odcAddProfile.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Jet OLEDB:Database Password=;Data Source=""Z:\Profiles.mdb"";Passwor" & _
        "d=;Jet OLEDB:Engine Type=5;Jet OLEDB:Global Bulk Transactions=1;Provider=""Micros" & _
        "oft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLEDB:SFP=False;Extended Prope" & _
        "rties=;Mode=Share Deny None;Jet OLEDB:New Database Password=;Jet OLEDB:Create Sy" & _
        "stem Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compa" & _
        "ct Without Replica Repair=False;User ID=Admin;Jet OLEDB:Encrypt Database=False"
        '
        'odaAddProfile
        '
        Me.odaAddProfile.DeleteCommand = Me.OleDbDeleteCommand1
        Me.odaAddProfile.InsertCommand = Me.OleDbInsertCommand1
        Me.odaAddProfile.SelectCommand = Me.selAddProfile
        Me.odaAddProfile.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "ProfileMaster", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("BOXLabelProfile", "BOXLabelProfile"), New System.Data.Common.DataColumnMapping("DateAdded", "DateAdded"), New System.Data.Common.DataColumnMapping("DateLastPrinted", "DateLastPrinted"), New System.Data.Common.DataColumnMapping("Description", "Description"), New System.Data.Common.DataColumnMapping("KeyboardLabelProfile", "KeyboardLabelProfile"), New System.Data.Common.DataColumnMapping("OEMPartNumber", "OEMPartNumber"), New System.Data.Common.DataColumnMapping("PartNumber", "PartNumber"), New System.Data.Common.DataColumnMapping("ProductCode", "ProductCode")})})
        Me.odaAddProfile.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM ProfileMaster WHERE (PartNumber = ?) AND (BOXLabelProfile = ? OR ? IS" & _
        " NULL AND BOXLabelProfile IS NULL) AND (DateAdded = ? OR ? IS NULL AND DateAdded" & _
        " IS NULL) AND (DateLastPrinted = ? OR ? IS NULL AND DateLastPrinted IS NULL) AND" & _
        " (Description = ? OR ? IS NULL AND Description IS NULL) AND (KeyboardLabelProfil" & _
        "e = ? OR ? IS NULL AND KeyboardLabelProfile IS NULL) AND (OEMPartNumber = ? OR ?" & _
        " IS NULL AND OEMPartNumber IS NULL) AND (ProductCode = ? OR ? IS NULL AND Produc" & _
        "tCode IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.odcAddProfile
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BOXLabelProfile", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BOXLabelProfile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BOXLabelProfile1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BOXLabelProfile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DateAdded", System.Data.OleDb.OleDbType.Date, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateAdded", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DateAdded1", System.Data.OleDb.OleDbType.Date, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateAdded", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DateLastPrinted", System.Data.OleDb.OleDbType.Date, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateLastPrinted", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DateLastPrinted1", System.Data.OleDb.OleDbType.Date, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateLastPrinted", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Description", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Description", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Description1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Description", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KeyboardLabelProfile", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KeyboardLabelProfile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KeyboardLabelProfile1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KeyboardLabelProfile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OEMPartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OEMPartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OEMPartNumber1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OEMPartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO ProfileMaster(BOXLabelProfile, DateAdded, DateLastPrinted, Descriptio" & _
        "n, KeyboardLabelProfile, OEMPartNumber, PartNumber, ProductCode) VALUES (?, ?, ?" & _
        ", ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.odcAddProfile
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("BOXLabelProfile", System.Data.OleDb.OleDbType.VarWChar, 50, "BOXLabelProfile"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("DateAdded", System.Data.OleDb.OleDbType.Date, 0, "DateAdded"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("DateLastPrinted", System.Data.OleDb.OleDbType.Date, 0, "DateLastPrinted"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 50, "Description"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("KeyboardLabelProfile", System.Data.OleDb.OleDbType.VarWChar, 50, "KeyboardLabelProfile"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("OEMPartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, "OEMPartNumber"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, "PartNumber"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, "ProductCode"))
        '
        'selAddProfile
        '
        Me.selAddProfile.CommandText = "SELECT BOXLabelProfile, DateAdded, DateLastPrinted, Description, KeyboardLabelPro" & _
        "file, OEMPartNumber, PartNumber, ProductCode FROM ProfileMaster"
        Me.selAddProfile.Connection = Me.odcAddProfile
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE ProfileMaster SET BOXLabelProfile = ?, DateAdded = ?, DateLastPrinted = ?," & _
        " Description = ?, KeyboardLabelProfile = ?, OEMPartNumber = ?, PartNumber = ?, P" & _
        "roductCode = ? WHERE (PartNumber = ?) AND (BOXLabelProfile = ? OR ? IS NULL AND " & _
        "BOXLabelProfile IS NULL) AND (DateAdded = ? OR ? IS NULL AND DateAdded IS NULL) " & _
        "AND (DateLastPrinted = ? OR ? IS NULL AND DateLastPrinted IS NULL) AND (Descript" & _
        "ion = ? OR ? IS NULL AND Description IS NULL) AND (KeyboardLabelProfile = ? OR ?" & _
        " IS NULL AND KeyboardLabelProfile IS NULL) AND (OEMPartNumber = ? OR ? IS NULL A" & _
        "ND OEMPartNumber IS NULL) AND (ProductCode = ? OR ? IS NULL AND ProductCode IS N" & _
        "ULL)"
        Me.OleDbUpdateCommand1.Connection = Me.odcAddProfile
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("BOXLabelProfile", System.Data.OleDb.OleDbType.VarWChar, 50, "BOXLabelProfile"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("DateAdded", System.Data.OleDb.OleDbType.Date, 0, "DateAdded"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("DateLastPrinted", System.Data.OleDb.OleDbType.Date, 0, "DateLastPrinted"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 50, "Description"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("KeyboardLabelProfile", System.Data.OleDb.OleDbType.VarWChar, 50, "KeyboardLabelProfile"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("OEMPartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, "OEMPartNumber"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, "PartNumber"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, "ProductCode"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BOXLabelProfile", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BOXLabelProfile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BOXLabelProfile1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BOXLabelProfile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DateAdded", System.Data.OleDb.OleDbType.Date, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateAdded", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DateAdded1", System.Data.OleDb.OleDbType.Date, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateAdded", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DateLastPrinted", System.Data.OleDb.OleDbType.Date, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateLastPrinted", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DateLastPrinted1", System.Data.OleDb.OleDbType.Date, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateLastPrinted", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Description", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Description", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Description1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Description", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KeyboardLabelProfile", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KeyboardLabelProfile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KeyboardLabelProfile1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KeyboardLabelProfile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OEMPartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OEMPartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OEMPartNumber1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OEMPartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        '
        'odcProductCodes
        '
        Me.odcProductCodes.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Jet OLEDB:Database Password=;Data Source=""Z:\SerialNumbers.mdb"";Pa" & _
        "ssword=;Jet OLEDB:Engine Type=5;Jet OLEDB:Global Bulk Transactions=1;Provider=""M" & _
        "icrosoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLEDB:SFP=False;Extended " & _
        "Properties=;Mode=Share Deny None;Jet OLEDB:New Database Password=;Jet OLEDB:Crea" & _
        "te System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:" & _
        "Compact Without Replica Repair=False;User ID=Admin;Jet OLEDB:Encrypt Database=Fa" & _
        "lse"
        '
        'odaProductCodes
        '
        Me.odaProductCodes.DeleteCommand = Me.OleDbDeleteCommand2
        Me.odaProductCodes.InsertCommand = Me.OleDbInsertCommand2
        Me.odaProductCodes.SelectCommand = Me.OleDbSelectCommand1
        Me.odaProductCodes.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SerialNumbers", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ProductCode", "ProductCode"), New System.Data.Common.DataColumnMapping("PCDescription", "PCDescription")})})
        Me.odaProductCodes.UpdateCommand = Me.OleDbUpdateCommand2
        '
        'OleDbDeleteCommand2
        '
        Me.OleDbDeleteCommand2.CommandText = "DELETE FROM SerialNumbers WHERE (ProductCode = ?) AND (PCDescription = ? OR ? IS " & _
        "NULL AND PCDescription IS NULL)"
        Me.OleDbDeleteCommand2.Connection = Me.odcProductCodes
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PCDescription", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PCDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PCDescription1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PCDescription", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand2
        '
        Me.OleDbInsertCommand2.CommandText = "INSERT INTO SerialNumbers(ProductCode, PCDescription) VALUES (?, ?)"
        Me.OleDbInsertCommand2.Connection = Me.odcProductCodes
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, "ProductCode"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("PCDescription", System.Data.OleDb.OleDbType.VarWChar, 50, "PCDescription"))
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT ProductCode, PCDescription, NextSerialNumber FROM SerialNumbers"
        Me.OleDbSelectCommand1.Connection = Me.odcProductCodes
        '
        'OleDbUpdateCommand2
        '
        Me.OleDbUpdateCommand2.CommandText = "UPDATE SerialNumbers SET ProductCode = ?, PCDescription = ? WHERE (ProductCode = " & _
        "?) AND (PCDescription = ? OR ? IS NULL AND PCDescription IS NULL)"
        Me.OleDbUpdateCommand2.Connection = Me.odcProductCodes
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, "ProductCode"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("PCDescription", System.Data.OleDb.OleDbType.VarWChar, 50, "PCDescription"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PCDescription", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PCDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PCDescription1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PCDescription", System.Data.DataRowVersion.Original, Nothing))
        '
        'odaKeyboardProfiles
        '
        Me.odaKeyboardProfiles.DeleteCommand = Me.OleDbDeleteCommand3
        Me.odaKeyboardProfiles.InsertCommand = Me.OleDbInsertCommand3
        Me.odaKeyboardProfiles.SelectCommand = Me.selKBProfile
        Me.odaKeyboardProfiles.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "KBProfiles", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("KBPDescription", "KBPDescription"), New System.Data.Common.DataColumnMapping("KBPLabel1", "KBPLabel1"), New System.Data.Common.DataColumnMapping("KBPLabel2", "KBPLabel2"), New System.Data.Common.DataColumnMapping("KBPLabel3", "KBPLabel3"), New System.Data.Common.DataColumnMapping("KBPLabel4", "KBPLabel4"), New System.Data.Common.DataColumnMapping("KBPLabel5", "KBPLabel5"), New System.Data.Common.DataColumnMapping("KBPLabel6", "KBPLabel6"), New System.Data.Common.DataColumnMapping("KBPLabel7", "KBPLabel7"), New System.Data.Common.DataColumnMapping("KBProfileName", "KBProfileName"), New System.Data.Common.DataColumnMapping("KeyboardLabelFile", "KeyboardLabelFile")})})
        Me.odaKeyboardProfiles.UpdateCommand = Me.OleDbUpdateCommand3
        '
        'OleDbDeleteCommand3
        '
        Me.OleDbDeleteCommand3.CommandText = "DELETE FROM KBProfiles WHERE (KBProfileName = ?) AND (KBPDescription = ? OR ? IS " & _
        "NULL AND KBPDescription IS NULL) AND (KBPLabel1 = ? OR ? IS NULL AND KBPLabel1 I" & _
        "S NULL) AND (KBPLabel2 = ? OR ? IS NULL AND KBPLabel2 IS NULL) AND (KBPLabel3 = " & _
        "? OR ? IS NULL AND KBPLabel3 IS NULL) AND (KBPLabel4 = ? OR ? IS NULL AND KBPLab" & _
        "el4 IS NULL) AND (KBPLabel5 = ? OR ? IS NULL AND KBPLabel5 IS NULL) AND (KBPLabe" & _
        "l6 = ? OR ? IS NULL AND KBPLabel6 IS NULL) AND (KBPLabel7 = ? OR ? IS NULL AND K" & _
        "BPLabel7 IS NULL) AND (KeyboardLabelFile = ? OR ? IS NULL AND KeyboardLabelFile " & _
        "IS NULL)"
        Me.OleDbDeleteCommand3.Connection = Me.odcAddProfile
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBProfileName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBProfileName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPDescription", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPDescription1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel11", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel2", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel21", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel3", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel31", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel4", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel41", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel5", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel5", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel51", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel5", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel6", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel6", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel61", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel6", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel7", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel7", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel71", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel7", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KeyboardLabelFile", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KeyboardLabelFile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KeyboardLabelFile1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KeyboardLabelFile", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand3
        '
        Me.OleDbInsertCommand3.CommandText = "INSERT INTO KBProfiles(KBPDescription, KBPLabel1, KBPLabel2, KBPLabel3, KBPLabel4" & _
        ", KBPLabel5, KBPLabel6, KBPLabel7, KBProfileName, KeyboardLabelFile) VALUES (?, " & _
        "?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand3.Connection = Me.odcAddProfile
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPDescription", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPDescription"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel1", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel1"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel2", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel2"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel3", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel3"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel4", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel4"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel5", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel5"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel6", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel6"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel7", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel7"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBProfileName", System.Data.OleDb.OleDbType.VarWChar, 50, "KBProfileName"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KeyboardLabelFile", System.Data.OleDb.OleDbType.VarWChar, 50, "KeyboardLabelFile"))
        '
        'selKBProfile
        '
        Me.selKBProfile.CommandText = "SELECT KBPDescription, KBPLabel1, KBPLabel2, KBPLabel3, KBPLabel4, KBPLabel5, KBP" & _
        "Label6, KBPLabel7, KBProfileName, KeyboardLabelFile FROM KBProfiles"
        Me.selKBProfile.Connection = Me.odcAddProfile
        '
        'OleDbUpdateCommand3
        '
        Me.OleDbUpdateCommand3.CommandText = "UPDATE KBProfiles SET KBPDescription = ?, KBPLabel1 = ?, KBPLabel2 = ?, KBPLabel3" & _
        " = ?, KBPLabel4 = ?, KBPLabel5 = ?, KBPLabel6 = ?, KBPLabel7 = ?, KBProfileName " & _
        "= ?, KeyboardLabelFile = ? WHERE (KBProfileName = ?) AND (KBPDescription = ? OR " & _
        "? IS NULL AND KBPDescription IS NULL) AND (KBPLabel1 = ? OR ? IS NULL AND KBPLab" & _
        "el1 IS NULL) AND (KBPLabel2 = ? OR ? IS NULL AND KBPLabel2 IS NULL) AND (KBPLabe" & _
        "l3 = ? OR ? IS NULL AND KBPLabel3 IS NULL) AND (KBPLabel4 = ? OR ? IS NULL AND K" & _
        "BPLabel4 IS NULL) AND (KBPLabel5 = ? OR ? IS NULL AND KBPLabel5 IS NULL) AND (KB" & _
        "PLabel6 = ? OR ? IS NULL AND KBPLabel6 IS NULL) AND (KBPLabel7 = ? OR ? IS NULL " & _
        "AND KBPLabel7 IS NULL) AND (KeyboardLabelFile = ? OR ? IS NULL AND KeyboardLabel" & _
        "File IS NULL)"
        Me.OleDbUpdateCommand3.Connection = Me.odcAddProfile
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPDescription", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPDescription"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel1", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel1"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel2", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel2"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel3", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel3"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel4", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel4"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel5", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel5"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel6", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel6"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBPLabel7", System.Data.OleDb.OleDbType.VarWChar, 50, "KBPLabel7"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KBProfileName", System.Data.OleDb.OleDbType.VarWChar, 50, "KBProfileName"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("KeyboardLabelFile", System.Data.OleDb.OleDbType.VarWChar, 50, "KeyboardLabelFile"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBProfileName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBProfileName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPDescription", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPDescription1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel11", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel2", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel21", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel3", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel31", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel4", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel41", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel5", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel5", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel51", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel5", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel6", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel6", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel61", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel6", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel7", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel7", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KBPLabel71", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KBPLabel7", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KeyboardLabelFile", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KeyboardLabelFile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_KeyboardLabelFile1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "KeyboardLabelFile", System.Data.DataRowVersion.Original, Nothing))
        '
        'odaBoxProfiles
        '
        Me.odaBoxProfiles.DeleteCommand = Me.OleDbDeleteCommand4
        Me.odaBoxProfiles.InsertCommand = Me.OleDbInsertCommand4
        Me.odaBoxProfiles.SelectCommand = Me.selBoxProfile
        Me.odaBoxProfiles.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "BXProfiles", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("BoxDescription", "BoxDescription"), New System.Data.Common.DataColumnMapping("BoxLabel1", "BoxLabel1"), New System.Data.Common.DataColumnMapping("BoxLabel2", "BoxLabel2"), New System.Data.Common.DataColumnMapping("BoxLabel3", "BoxLabel3"), New System.Data.Common.DataColumnMapping("BoxLabel4", "BoxLabel4"), New System.Data.Common.DataColumnMapping("BoxLabel5", "BoxLabel5"), New System.Data.Common.DataColumnMapping("BoxLabel6", "BoxLabel6"), New System.Data.Common.DataColumnMapping("BoxLabel7", "BoxLabel7"), New System.Data.Common.DataColumnMapping("BoxLabelFile", "BoxLabelFile"), New System.Data.Common.DataColumnMapping("BXProfileName", "BXProfileName")})})
        Me.odaBoxProfiles.UpdateCommand = Me.OleDbUpdateCommand4
        '
        'OleDbDeleteCommand4
        '
        Me.OleDbDeleteCommand4.CommandText = "DELETE FROM BXProfiles WHERE (BXProfileName = ?) AND (BoxDescription = ? OR ? IS " & _
        "NULL AND BoxDescription IS NULL) AND (BoxLabel1 = ? OR ? IS NULL AND BoxLabel1 I" & _
        "S NULL) AND (BoxLabel2 = ? OR ? IS NULL AND BoxLabel2 IS NULL) AND (BoxLabel3 = " & _
        "? OR ? IS NULL AND BoxLabel3 IS NULL) AND (BoxLabel4 = ? OR ? IS NULL AND BoxLab" & _
        "el4 IS NULL) AND (BoxLabel5 = ? OR ? IS NULL AND BoxLabel5 IS NULL) AND (BoxLabe" & _
        "l6 = ? OR ? IS NULL AND BoxLabel6 IS NULL) AND (BoxLabel7 = ? OR ? IS NULL AND B" & _
        "oxLabel7 IS NULL) AND (BoxLabelFile = ? OR ? IS NULL AND BoxLabelFile IS NULL)"
        Me.OleDbDeleteCommand4.Connection = Me.odcAddProfile
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BXProfileName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BXProfileName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxDescription", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxDescription1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel11", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel2", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel21", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel3", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel31", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel4", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel41", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel5", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel5", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel51", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel5", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel6", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel6", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel61", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel6", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel7", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel7", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel71", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel7", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabelFile", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabelFile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabelFile1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabelFile", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand4
        '
        Me.OleDbInsertCommand4.CommandText = "INSERT INTO BXProfiles(BoxDescription, BoxLabel1, BoxLabel2, BoxLabel3, BoxLabel4" & _
        ", BoxLabel5, BoxLabel6, BoxLabel7, BoxLabelFile, BXProfileName) VALUES (?, ?, ?," & _
        " ?, ?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand4.Connection = Me.odcAddProfile
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxDescription", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxDescription"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel1", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel1"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel2", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel2"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel3", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel3"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel4", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel4"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel5", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel5"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel6", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel6"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel7", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel7"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabelFile", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabelFile"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BXProfileName", System.Data.OleDb.OleDbType.VarWChar, 50, "BXProfileName"))
        '
        'selBoxProfile
        '
        Me.selBoxProfile.CommandText = "SELECT BoxDescription, BoxLabel1, BoxLabel2, BoxLabel3, BoxLabel4, BoxLabel5, Box" & _
        "Label6, BoxLabel7, BoxLabelFile, BXProfileName FROM BXProfiles"
        Me.selBoxProfile.Connection = Me.odcAddProfile
        '
        'OleDbUpdateCommand4
        '
        Me.OleDbUpdateCommand4.CommandText = "UPDATE BXProfiles SET BoxDescription = ?, BoxLabel1 = ?, BoxLabel2 = ?, BoxLabel3" & _
        " = ?, BoxLabel4 = ?, BoxLabel5 = ?, BoxLabel6 = ?, BoxLabel7 = ?, BoxLabelFile =" & _
        " ?, BXProfileName = ? WHERE (BXProfileName = ?) AND (BoxDescription = ? OR ? IS " & _
        "NULL AND BoxDescription IS NULL) AND (BoxLabel1 = ? OR ? IS NULL AND BoxLabel1 I" & _
        "S NULL) AND (BoxLabel2 = ? OR ? IS NULL AND BoxLabel2 IS NULL) AND (BoxLabel3 = " & _
        "? OR ? IS NULL AND BoxLabel3 IS NULL) AND (BoxLabel4 = ? OR ? IS NULL AND BoxLab" & _
        "el4 IS NULL) AND (BoxLabel5 = ? OR ? IS NULL AND BoxLabel5 IS NULL) AND (BoxLabe" & _
        "l6 = ? OR ? IS NULL AND BoxLabel6 IS NULL) AND (BoxLabel7 = ? OR ? IS NULL AND B" & _
        "oxLabel7 IS NULL) AND (BoxLabelFile = ? OR ? IS NULL AND BoxLabelFile IS NULL)"
        Me.OleDbUpdateCommand4.Connection = Me.odcAddProfile
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxDescription", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxDescription"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel1", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel1"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel2", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel2"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel3", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel3"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel4", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel4"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel5", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel5"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel6", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel6"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabel7", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabel7"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BoxLabelFile", System.Data.OleDb.OleDbType.VarWChar, 50, "BoxLabelFile"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("BXProfileName", System.Data.OleDb.OleDbType.VarWChar, 50, "BXProfileName"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BXProfileName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BXProfileName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxDescription", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxDescription1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxDescription", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel11", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel2", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel21", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel3", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel31", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel4", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel41", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel5", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel5", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel51", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel5", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel6", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel6", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel61", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel6", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel7", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel7", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabel71", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabel7", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabelFile", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabelFile", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_BoxLabelFile1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "BoxLabelFile", System.Data.DataRowVersion.Original, Nothing))
        '
        'odaAMSPArtMaster
        '
        Me.odaAMSPArtMaster.DeleteCommand = Me.OleDbDeleteCommand5
        Me.odaAMSPArtMaster.InsertCommand = Me.OleDbInsertCommand5
        Me.odaAMSPArtMaster.SelectCommand = Me.selAMSPArtMaster
        Me.odaAMSPArtMaster.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "PartMaster", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("PartNumber", "PartNumber"), New System.Data.Common.DataColumnMapping("DescText", "DescText")})})
        Me.odaAMSPArtMaster.UpdateCommand = Me.OleDbUpdateCommand5
        '
        'OleDbDeleteCommand5
        '
        Me.OleDbDeleteCommand5.CommandText = "DELETE FROM PartMaster WHERE (PartNumber = ?) AND (DescText = ? OR ? IS NULL AND " & _
        "DescText IS NULL)"
        Me.OleDbDeleteCommand5.Connection = Me.odcAlliancePartMaster
        Me.OleDbDeleteCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DescText", System.Data.OleDb.OleDbType.VarWChar, 60, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DescText", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DescText1", System.Data.OleDb.OleDbType.VarWChar, 60, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DescText", System.Data.DataRowVersion.Original, Nothing))
        '
        'odcAlliancePartMaster
        '
        Me.odcAlliancePartMaster.ConnectionString = "User ID=sa;Data Source=""HAL\AllianceMFG"";Tag with column collation when possible=" & _
        "False;Initial Catalog=KBD2002;Use Procedure for Prepare=1;Auto Translate=True;Pe" & _
        "rsist Security Info=False;Provider=SQLOLEDB;Workstation ID=SHAFFER2;Use Encrypti" & _
        "on for Data=False;Packet Size=4096"
        '
        'OleDbInsertCommand5
        '
        Me.OleDbInsertCommand5.CommandText = "INSERT INTO PartMaster(PartNumber, DescText) VALUES (?, ?); SELECT PartNumber, De" & _
        "scText FROM PartMaster WHERE (PartNumber = ?)"
        Me.OleDbInsertCommand5.Connection = Me.odcAlliancePartMaster
        Me.OleDbInsertCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        Me.OleDbInsertCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("DescText", System.Data.OleDb.OleDbType.VarWChar, 60, "DescText"))
        Me.OleDbInsertCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Select_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        '
        'selAMSPArtMaster
        '
        Me.selAMSPArtMaster.CommandText = "SELECT PartNumber, DescText FROM PartMaster"
        Me.selAMSPArtMaster.Connection = Me.odcAlliancePartMaster
        '
        'OleDbUpdateCommand5
        '
        Me.OleDbUpdateCommand5.CommandText = "UPDATE PartMaster SET PartNumber = ?, DescText = ? WHERE (PartNumber = ?) AND (De" & _
        "scText = ? OR ? IS NULL AND DescText IS NULL); SELECT PartNumber, DescText FROM " & _
        "PartMaster WHERE (PartNumber = ?)"
        Me.OleDbUpdateCommand5.Connection = Me.odcAlliancePartMaster
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("DescText", System.Data.OleDb.OleDbType.VarWChar, 60, "DescText"))
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DescText", System.Data.OleDb.OleDbType.VarWChar, 60, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DescText", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DescText1", System.Data.OleDb.OleDbType.VarWChar, 60, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DescText", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Select_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        '
        'odaStructureID
        '
        Me.odaStructureID.InsertCommand = Me.OleDbInsertCommand6
        Me.odaStructureID.SelectCommand = Me.selStructureID
        Me.odaStructureID.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "qrySerialNumbers", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("PNStructureID", "PNStructureID"), New System.Data.Common.DataColumnMapping("ProductCode", "ProductCode")})})
        '
        'OleDbInsertCommand6
        '
        Me.OleDbInsertCommand6.CommandText = "INSERT INTO qrySerialNumbers(PNStructureID, ProductCode) VALUES (?, ?)"
        Me.OleDbInsertCommand6.Connection = Me.odcProductCodes
        Me.OleDbInsertCommand6.Parameters.Add(New System.Data.OleDb.OleDbParameter("PNStructureID", System.Data.OleDb.OleDbType.VarWChar, 50, "PNStructureID"))
        Me.OleDbInsertCommand6.Parameters.Add(New System.Data.OleDb.OleDbParameter("ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, "ProductCode"))
        '
        'selStructureID
        '
        Me.selStructureID.CommandText = "SELECT PNStructureID, ProductCode FROM qrySerialNumbers"
        Me.selStructureID.Connection = Me.odcProductCodes
        '
        'AddToMaster
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(522, 504)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.cmbBoxProfile)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.cmbKBProfile)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtAllianceDescription)
        Me.Controls.Add(Me.cmbProductFamily)
        Me.Controls.Add(Me.txtBoxSample)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtBoxLabelFN)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtKBLabelFN)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSample)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtNewStartSN)
        Me.Controls.Add(Me.txtNewOEMPN)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNewPN)
        Me.Controls.Add(Me.chkHasOEMPN)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AddToMaster"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add to Master List"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub AddToMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ABORT = False
        txtNewPN.Text = curBarCode.UniPartNumber
        txtAllianceDescription.Text = curWorkOrder.Description
        'Call PopulateCustomerChoices()
        txtNewOEMPN.Enabled = False
        txtNewOEMPN.Text = UCase(txtNewPN.Text)
        Call PopulateCombos()
        'Call PopulateFormatChoices()

    End Sub

    Private Sub chkHasOEMPN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHasOEMPN.CheckedChanged
        If chkHasOEMPN.Checked = True Then
            txtNewOEMPN.Enabled = True
            txtNewOEMPN.Clear()
            txtNewOEMPN.Focus()
        Else
            txtNewOEMPN.Enabled = False
            txtNewOEMPN.Text = UCase(txtNewPN.Text)
        End If
    End Sub

    'Private Sub cmbNewCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbNewCustomer.SelectedIndexChanged

    '    DsCustomers1.Clear()
    '    selCustomers.CommandText = "SELECT * FROM  Customers WHERE CustomerDesc = '" & cmbNewCustomer.SelectedItem & "'"
    '    odaCustomers.Fill(DsCustomers1)

    '    'drCustomer = DsCustomers1.Tables("Customers").Rows(0)

    '    curNewMaster.CustomerID = DsCustomers1.Tables("Customers").Rows(0).Item("CustomerID")

    'End Sub

    Private Sub cmbNewFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Dim drNewFormat As DataRow

        'DsNMFormats1.Clear()
        'selFormats.CommandText = "SELECT * FROM  tblFormats WHERE Description = '" & cmbNewFormat.SelectedItem & "'"
        'odaFormats.Fill(DsNMFormats1)

        ''drCustomer = DsCustomers1.Tables("Customers").Rows(0)
        'drNewFormat = DsNMFormats1.Tables("tblFormats").Rows(0)
        'curNewMaster.FormatType = drNewFormat.Item("FormatGroupID")

        'Call PopulateFormatFields(drNewFormat)
    End Sub
    Private Sub PopulateFormatFields(ByVal drFormat As DataRow)

        'txtKBLabelFN.Text = drFormat.Item("KB_RawLayout")
        'txtBoxLabelFN.Text = drFormat.Item("BOX_RawLayout")

        ''BUILD SAMPLE TXT BOX
        'Dim i, x As Int16
        'Dim sFieldID As String
        'Dim iFieldLength As Int16
        'Dim sBCSample As String

        'For i = 1 To 8
        '    sFieldID = drFormat.Item("DefaultBarcodeField_" & i)
        '    DsNMFields1.Clear()
        '    selFields.CommandText = "SELECT * FROM tblBarcodeFields WHERE FieldID = '" & sFieldID & "'"
        '    odaFields.Fill(DsNMFields1)
        '    iFieldLength = DsNMFields1.Tables("tblBarcodeFields").Rows(0).Item("Digits")
        '    For x = 1 To iFieldLength
        '        sBCSample = sBCSample & Str(i).TrimStart(" ")
        '    Next
        'Next
        'txtSample.Text = sBCSample
    End Sub

#Region "Populate Combo Boxes"
    Private Sub PopulateFormatChoices()
        '    Dim i As Int16
        '    Dim iCount As Integer
        '    DsNMFormats1.Clear()
        '    selFormats.CommandText = "SELECT * FROM  tblFormats"
        '    odaFormats.Fill(DsNMFormats1)
        '    iCount = DsNMFormats1.Tables("tblFormats").Rows.Count
        '    cmbNewFormat.Items.Clear()
        '    For i = 0 To iCount - 1
        '        cmbNewFormat.Items.Add(DsNMFormats1.Tables("tblFormats").Rows(i).Item("Description"))
        '    Next



    End Sub

    'Private Sub PopulateCustomerChoices()
    '    Dim i As Int16
    '    Dim iCount As Integer
    '    DsCustomers1.Clear()
    '    selCustomers.CommandText = "SELECT * FROM  Customers"
    '    odaCustomers.Fill(DsCustomers1)

    '    iCount = DsCustomers1.Tables("Customers").Rows.Count

    '    cmbNewCustomer.Items.Clear()
    '    For i = 0 To iCount - 1
    '        cmbNewCustomer.Items.Add(DsCustomers1.Tables("Customers").Rows(i).Item("CustomerDesc"))
    '    Next

    'End Sub

#End Region

#Region "Key Presses"
    Private Sub cmbNewFormat_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        e.Handled = True
    End Sub

    Private Sub cmbNewCustomer_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        e.Handled = True
    End Sub

    Private Sub txtSample_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSample.KeyPress
        e.Handled = True
    End Sub

    Private Sub txtBoxLabelFN_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBoxLabelFN.KeyPress
        e.Handled = True
    End Sub
    Private Sub txtKBLabelFN_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtKBLabelFN.KeyPress
        e.Handled = True
    End Sub
    Private Sub txtAllianceDescription_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAllianceDescription.KeyPress
        e.Handled = True
    End Sub

    Private Sub cmbBoxProfile_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbBoxProfile.KeyPress
        e.Handled = True
    End Sub
    Private Sub cmbKBProfile_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbKBProfile.KeyPress
        e.Handled = True
    End Sub
#End Region



    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Function VerifyFields() As Boolean
        VerifyFields = True

        'cmbBoxProfile.BackColor = Color.White
        'cmbKBProfile.BackColor = Color.White
        cmbProductFamily.BackColor = Color.White

        'DsAddProfile1.Clear()
        'selAddProfile.CommandText = "SELECT * FROM  ProfileMaster WHERE PartNumber = '" & txtNewPN.Text & "'"
        'odaAddProfile.Fill(DsAddProfile1)
        'If DsAddProfile1.Tables("ProfileMaster").Rows.Count > 0 Then
        '    MsgBox("PART NUMBER ALREADY EXISTS")
        '    VerifyFields = False
        'End If

        If chkHasOEMPN.Checked = False Then
            txtNewOEMPN.Text = UCase(txtNewPN.Text)
        End If

        If txtNewPN.Text = "" Then
            txtNewPN.BackColor = Color.Yellow
        End If

        If txtAllianceDescription.Text = "" Then txtAllianceDescription.Text = "Manually Added"

        'If cmbBoxProfile.Text = "" Then
        '    cmbBoxProfile.BackColor = Color.Yellow
        '    VerifyFields = False
        'End If

        'If cmbKBProfile.Text = "" Then
        '    cmbKBProfile.BackColor = Color.Yellow
        '    'MsgBox("No KB Profile")
        '    VerifyFields = False
        'End If

        If cmbProductFamily.Text = "" Then
            'MsgBox("No Family Code")
            cmbProductFamily.BackColor = Color.Yellow
            VerifyFields = False
        End If

        If Not VerifyFields Then
            MessageBox.Show("Errors on Add", "Errors", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Function
    Private Sub WriteToDatabase()

        Dim connStr As String
        Dim sqlStr As String
        Dim drNewRow As DataRow

        Dim dsProfile As DataSet
        Dim odaProfile As System.Data.OleDb.OleDbDataAdapter
        Dim cmdProfile As System.Data.OleDb.OleDbCommandBuilder

        Dim resUpdate As DialogResult

        connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        odcUnicompMain = New OleDbConnection(connStr)
        odcUnicompMain.Open()



        dsProfile = New DataSet
        sqlStr = "SELECT * FROM ProfileMaster WHERE PartNumber = '" & newProfile.PartNumber & "'"
        odaProfile = New OleDbDataAdapter(sqlStr, odcUnicompMain)

        cmdProfile = New OleDbCommandBuilder(odaProfile)
        dsProfile.Clear()
        odaProfile.Fill(dsProfile, "ProfileMaster")
        Try
            If dsProfile.Tables("ProfileMaster").Rows.Count = 1 Then
                resUpdate = MessageBox.Show("Profile for P/N: " & newProfile.PartNumber & vbCrLf & "Already Exists. Update?", "Update?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If resUpdate = DialogResult.Yes Then
                    drNewRow = dsProfile.Tables("ProfileMaster").Rows(0)
                    drNewRow.Item("PartNumber") = newProfile.PartNumber
                    drNewRow.Item("FamilyCode") = newProfile.Family
                    drNewRow.Item("OEMPartNumber") = newProfile.OEMPartNumber
                    drNewRow.Item("Description") = newProfile.Description
                    drNewRow.Item("DateLastPrinted") = newProfile.DateLastPrinted
                    drNewRow.Item("DateAdded") = newProfile.DateAdded
                    odaProfile.Update(dsProfile, "ProfileMaster")
                Else
                    Exit Sub
                End If
            Else
                drNewRow = dsProfile.Tables("ProfileMaster").NewRow

                drNewRow.Item("PartNumber") = newProfile.PartNumber
                drNewRow.Item("FamilyCode") = newProfile.Family
                drNewRow.Item("OEMPartNumber") = newProfile.OEMPartNumber
                drNewRow.Item("Description") = newProfile.Description
                drNewRow.Item("DateLastPrinted") = newProfile.DateLastPrinted
                drNewRow.Item("DateAdded") = newProfile.DateAdded

                dsProfile.Tables("ProfileMaster").Rows.Add(drNewRow)
                odaProfile.Update(dsProfile, "ProfileMaster")
            End If


            MessageBox.Show("P/N: " & newProfile.PartNumber & " Updated", "Updated", MessageBoxButtons.OK, MessageBoxIcon.None)
        Catch ex As Exception
            MessageBox.Show("Error Updating database:" & vbCrLf & ex.ToString, "Error", MessageBoxButtons.OK)
        End Try

        odcUnicompMain.Close()
        'Dim drNewRow As DataRow

        'DsAddProfile1.Clear()
        'selAddProfile.CommandText = "SELECT * FROM  ProfileMaster"
        'odaAddProfile.Fill(DsAddProfile1)
        'drNewRow = DsAddProfile1.Tables("ProfileMaster").NewRow
        ''PartNumber	ProductCode	OEMPartNumber	Description	KeyboardLabelProfile	BOXLabelProfile	DateLastPrinted	DateAdded

        'drNewRow.Item("PartNumber") = newProfile.PartNumber
        'drNewRow.Item("ProductCode") = newProfile.ProductCode
        'drNewRow.Item("OEMPartNumber") = newProfile.OEMPartNumber
        'drNewRow.Item("Description") = newProfile.Description
        'drNewRow.Item("KeyboardLabelProfile") = newProfile.KBProfile
        'drNewRow.Item("BOXLabelProfile") = newProfile.BOXProfile
        'drNewRow.Item("DateLastPrinted") = newProfile.DateLastPrinted
        'drNewRow.Item("DateAdded") = newProfile.DateAdded
        'Try
        '    DsAddProfile1.Tables("ProfileMaster").Rows.Add(drNewRow)
        '    odaAddProfile.Update(DsAddProfile1)
        'Catch ex As Exception
        '    MsgBox(ex.ToString)
        'End Try
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        If VerifyFields() Then
            With newProfile
                .PartNumber = UCase(txtNewPN.Text)
                .Description = txtAllianceDescription.Text
                .OEMPartNumber = UCase(txtNewOEMPN.Text)
                '.ProductCode = UCase(cmbProductFamily.Text)
                .StartingSerial = UCase(txtNewStartSN.Text)
                .KBTemplate = dsKBTemplates.Tables("KBTemplates").Rows(cmbKBProfile.SelectedIndex).Item("KBTemplateID")
                .BOXTemplate = dsBoxTemplates.Tables("BXTemplates").Rows(cmbBoxProfile.SelectedIndex).Item("BXTemplateID")
                '.KBProfile = UCase(cmbKBProfile.Text)
                '.BOXProfile = UCase(cmbBoxProfile.Text)
                .DateAdded = Now
                .DateLastPrinted = "#12/15/1980 12:00:00#"
            End With

            WriteToDatabase()
            'MsgBox("ALL GOOD, OK TO WRITE")
            Me.Close()
        Else
            'MsgBox("FAIL")
        End If


    End Sub

    Private Sub txtNewPN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNewPN.TextChanged
        txtAllianceDescription.Clear()
    End Sub

    Private Sub PopulateCombos()

        Dim connStr As String
        Dim sqlStr As String
        Dim r As DataRow

        connStr = "Integrated Security=SSPI;Packet Size=4096;Data Source=BARCODESERVER;Tag with column collation when possible=False;Initial Catalog=UnicompBarcodeSystem;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False"
        odcUnicompMain = New OleDbConnection(connStr)
        odcUnicompMain.Open()
        dsFamily = New DataSet

        sqlStr = "SELECT * FROM Q_FamilyAndTemplates"
        odaFamily = New OleDbDataAdapter(sqlStr, odcUnicompMain)

        cmdFamily = New OleDbCommandBuilder(odaFamily)
        dsFamily.Clear()

        odaFamily.Fill(dsFamily, "Q_FamilyAndTemplates")

        cmbProductFamily.Items.Clear()
        For Each r In dsFamily.Tables("Q_FamilyAndTemplates").Rows
            cmbProductFamily.Items.Add(RTrim("(" & r.Item("FamilyCode") & ") " & r.Item("FamilyDescription")))
        Next

        'Keyboard
        dsKBTemplates = New DataSet
        sqlStr = "SELECT * FROM KBTemplates"
        odaKBTemplate = New OleDbDataAdapter(sqlStr, odcUnicompMain)
        cmdKBTemplates = New OleDbCommandBuilder(odaKBTemplate)
        dsKBTemplates.Clear()
        odaKBTemplate.Fill(dsKBTemplates, "KBTemplates")

        cmbKBProfile.Items.Clear()
        For Each r In dsKBTemplates.Tables("KBTemplates").Rows
            cmbKBProfile.Items.Add(RTrim("(" & r.Item("KBTemplateID") & ") " & r.Item("KBTemplateDescription")))
        Next

        'BOX
        dsBoxTemplates = New DataSet
        sqlStr = "SELECT * FROM BXTemplates"
        odaBoxTemplate = New OleDbDataAdapter(sqlStr, odcUnicompMain)
        cmdBoxTemplates = New OleDbCommandBuilder(odaBoxTemplate)
        dsBoxTemplates.Clear()
        odaBoxTemplate.Fill(dsBoxTemplates, "BXTemplates")

        cmbBoxProfile.Items.Clear()
        For Each r In dsBoxTemplates.Tables("BXTemplates").Rows
            'Debug.Write(r.Item("KBTemplateID"))
            cmbBoxProfile.Items.Add(RTrim("(" & r.Item("BXTemplateID") & ") " & r.Item("BoxTemplateDescription")))
        Next

        odcUnicompMain.Close()

        'Dim i As Int16
        'Dim iEnd As Int16

        'cmbProductFamily.Items.Clear()
        'cmbKBProfile.Items.Clear()
        'cmbBoxProfile.Items.Clear()

        'DsPC1.Clear()
        ''odaProductCodes.Fill(DsPC1)
        'iEnd = DsPC1.Tables("SerialNumbers").Rows.Count - 1

        'For i = 0 To iEnd
        '    'Debug.Write(DsPC1.Tables("SerialNumbers").Rows(i).Item("ProductCode") & vbCrLf)
        '    cmbProductFamily.Items.Add("(" & DsPC1.Tables("SerialNumbers").Rows(i).Item("ProductCode") & ")" & DsPC1.Tables("SerialNumbers").Rows(i).Item("PCDescription"))
        'Next

        'DsKBProfiles1.Clear()
        ''  odaKeyboardProfiles.Fill(DsKBProfiles1)
        'iEnd = DsKBProfiles1.Tables("KBProfiles").Rows.Count - 1

        'For i = 0 To iEnd
        '    'Debug.Write(DsKBProfiles1.Tables("KBProfiles").Rows(i).Item("KBPDescription") & vbCrLf)
        '    cmbKBProfile.Items.Add("(" & DsKBProfiles1.Tables("KBProfiles").Rows(i).Item("KBProfileName") & ")" & DsKBProfiles1.Tables("KBProfiles").Rows(i).Item("KBPDescription"))
        'Next

        'DsBoxProfiles1.Clear()
        '' odaBoxProfiles.Fill(DsBoxProfiles1)
        'iEnd = DsBoxProfiles1.Tables("BXProfiles").Rows.Count - 1

        'For i = 0 To iEnd
        '    'Debug.Write(DsKBProfiles1.Tables("KBProfiles").Rows(i).Item("KBPDescription") & vbCrLf)
        '    cmbBoxProfile.Items.Add("(" & DsBoxProfiles1.Tables("BXProfiles").Rows(i).Item("BXProfileName") & ")" & DsBoxProfiles1.Tables("BXProfiles").Rows(i).Item("BoxDescription"))
        'Next

    End Sub

    Private Sub cmbProductFamily_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbProductFamily.SelectedIndexChanged

        Dim sKBTemplateCode As String
        Dim sBoxTemplateCode As String
        Dim x As Int16

        txtNewStartSN.Text = dsFamily.Tables("Q_FamilyAndTemplates").Rows(cmbProductFamily.SelectedIndex).Item("NextSerialNumber")

        sKBTemplateCode = dsFamily.Tables("Q_FamilyAndTemplates").Rows(cmbProductFamily.SelectedIndex).Item("KBTemplateID")
        sBoxTemplateCode = dsFamily.Tables("Q_FamilyAndTemplates").Rows(cmbProductFamily.SelectedIndex).Item("BXTemplateID")

        x = 0
        While dsKBTemplates.Tables("KBTemplates").Rows(x).Item("KBTemplateID") <> sKBTemplateCode
            x = x + 1
            If x > 100 Then Exit While
        End While
        cmbKBProfile.SelectedIndex = x
        txtKBLabelFN.Text = dsKBTemplates.Tables("KBTemplates").Rows(x).Item("KBTemplateFilename")

        x = 0
        While dsBoxTemplates.Tables("BXTemplates").Rows(x).Item("BXTemplateID") <> sBoxTemplateCode
            x = x + 1
            If x > 100 Then Exit While
        End While
        cmbBoxProfile.SelectedIndex = x
        txtBoxLabelFN.Text = dsBoxTemplates.Tables("BXTemplates").Rows(x).Item("BoxTemplateFilename")


        newProfile.Family = dsFamily.Tables("Q_FamilyAndTemplates").Rows(cmbProductFamily.SelectedIndex).Item("FamilyCode")
        'newProfile.BOXTemplate = dsFamily.Tables("Q_FamilyAndTemplates").Rows(cmbProductFamily.SelectedIndex).Item("BXTemplateID")
        'newProfile.KBTemplate = dsFamily.Tables("Q_FamilyAndTemplates").Rows(cmbProductFamily.SelectedIndex).Item("KBTemplateID")


        'Debug.Write(cmbProductFamily.Text & vbCrLf)
        'Debug.Write(dsFamily.Tables("Q_FamilyAndTemplates").Rows(cmbProductFamily.SelectedIndex).Item("KeyboardLabelTemplate") & vbCrLf)
        'Dim sNextSerial As String
        '' MsgBox("Changed")
        ''MsgBox(DsPC1.Tables("SerialNumbers").Rows(cmbProductFamily.SelectedIndex).Item("ProductCode"))
        ''MsgBox(cmbProductFamily.SelectedIndex)
        'newProfile.ProductCode = DsPC1.Tables("SerialNumbers").Rows(cmbProductFamily.SelectedIndex).Item("ProductCode")

        'sNextSerial = DsPC1.Tables("SerialNumbers").Rows(cmbProductFamily.SelectedIndex).Item("NextSerialNumber")
        'sNextSerial.PadLeft(5, "0")
        'sNextSerial = DsPC1.Tables("SerialNumbers").Rows(cmbProductFamily.SelectedIndex).Item("ProductCode") & sNextSerial.PadLeft(5, "0")
        'txtNewStartSN.Text = sNextSerial
    End Sub

    Private Sub cmbProductFamily_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbProductFamily.KeyPress
        e.Handled = True
    End Sub

    Private Sub cmbKBProfile_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbKBProfile.SelectedIndexChanged
        'txtKBLabelFN.Text = DsKBProfiles1.Tables("KBProfiles").Rows(cmbKBProfile.SelectedIndex).Item("KeyboardLabelFile")
        'newProfile.KBProfile = DsKBProfiles1.Tables("KBProfiles").Rows(cmbKBProfile.SelectedIndex).Item("KBProfileName")
    End Sub



    Private Sub cmbBoxProfile_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBoxProfile.SelectedIndexChanged
        'txtBoxLabelFN.Text = DsBoxProfiles1.Tables("BXProfiles").Rows(cmbBoxProfile.SelectedIndex).Item("BoxLabelFile")
        'newProfile.BOXProfile = DsBoxProfiles1.Tables("BXProfiles").Rows(cmbBoxProfile.SelectedIndex).Item("BXProfileName")
    End Sub

    Private Function GetAMSDescription(ByVal sPN As String) As String

        Dim odcAlliance As System.Data.OleDb.OleDbConnection
        Dim odaPartMaster As System.Data.OleDb.OleDbDataAdapter
        Dim cmdAlliance As System.Data.OleDb.OleDbCommandBuilder
        Dim dsPartMaster As DataSet
        Dim sqlStr As String
        Dim connAMS As String

        Dim sAMSPN As String

        If Len(sPN) = 7 Then
            sAMSPN = "00" & sPN
        Else
            GetAMSDescription = "Could Not Find in Alliance."
            'MsgBox("Check Partnumber. Should be 7 digits")
            Exit Function
        End If


        'sAMSPN = "0096U1114"

        Try
            connAMS = "User ID=sa;Data Source=""HAL\AllianceMFG"";Tag with column collation when possible=False;Initial Catalog=KBD2002;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False;Packet Size=4096"
            odcAlliance = New OleDbConnection(connAMS)
            odcAlliance.Open()
            dsPartMaster = New DataSet

            sqlStr = "SELECT PartNumber, DescText FROM PartMaster WHERE PartNumber = '" & sAMSPN & "'"
            odaPartMaster = New OleDbDataAdapter(sqlStr, odcAlliance)

            cmdAlliance = New OleDbCommandBuilder(odaPartMaster)
            dsPartMaster.Clear()

            odaPartMaster.Fill(dsPartMaster, "PartMaster")

            If dsPartMaster.Tables("PartMaster").Rows.Count = 1 Then
                GetAMSDescription = dsPartMaster.Tables("PartMaster").Rows(0).Item("DescText")
                'LookUpWorkOrder = dsPartMaster.Tables("WOHeader").Rows(0)
            Else
                GetAMSDescription = "Could Not Find in Alliance"
                'LookUpWorkOrder = Nothing
            End If
            odcAlliance.Close()
        Catch ex As Exception
            MessageBox.Show("Error connecting to Alliance." & vbCrLf & "Check network connection", "Check Network", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub txtNewPN_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNewPN.LostFocus
        If txtNewPN.Text <> "" Then
            txtAllianceDescription.Text = GetAMSDescription(RTrim(LTrim(txtNewPN.Text)))
            If txtNewOEMPN.Text = "" Then
                chkHasOEMPN.Checked = True
                chkHasOEMPN.Checked = False
            End If
        End If
        'Debug.Write(GetAMSDescription("0"))
    End Sub

    Private Sub GuessProfile(ByVal sPN As String)
        'Dim sStructureID As String
        'Dim iPCEnd As Int16
        'Dim sProductCodeToUse As String
        'Dim i As Int16
        'Dim iIndexOfPC As Int16

        'sStructureID = Microsoft.VisualBasic.Left(sPN, 3)
        'DsPNStructureID1.Clear()
        'selStructureID.CommandText = "SELECT * FROM qrySerialNumbers WHERE PNStructureID = '" & sStructureID & "'"
        '' odaStructureID.Fill(DsPNStructureID1)

        'If DsPNStructureID1.Tables("qrySerialNumbers").Rows.Count = 1 Then
        '    sProductCodeToUse = DsPNStructureID1.Tables("qrySerialNumbers").Rows(0).Item("ProductCode")
        'Else
        '    sProductCodeToUse = "GM"
        'End If

        'iPCEnd = DsPC1.Tables("SerialNumbers").Rows.Count - 1
        'For i = 0 To iPCEnd
        '    If DsPC1.Tables("SerialNumbers").Rows(i).Item("ProductCode") = sProductCodeToUse Then
        '        'MsgBox(i)
        '        cmbProductFamily.SelectedIndex = i
        '    End If
        'Next






    End Sub

    Private Sub AddToMaster_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If ABORT Then
            Me.Close()
        End If
        txtNewPN.Focus()
        txtAllianceDescription.Focus()
        txtNewPN.Focus()
    End Sub

    Private Sub txtNewPN_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNewPN.Validated
        '    Dim sPartNumber As String
        '    Dim sAMSPartNumber As String
        '    Dim sDescription As String

        '    sAMSPartNumber = curWorkOrder.PartNumber

        '    sPartNumber = UCase(txtNewPN.Text)

        '    'sAMSPartNumber = sPartNumber
        '    If Len(sPartNumber) = 7 Then
        '        sAMSPartNumber = "00" & sPartNumber
        '    Else
        '        'If sPartNumber = "42H1292U" Then sAMSPartNumber = "042H1292U"
        '    End If

        '    DsAMSPartMaster1.Clear()
        '    selAMSPArtMaster.CommandText = "SELECT * FROM PartMaster WHERE PartNumber = '" & sAMSPartNumber & "'"
        '    'odaAMSPArtMaster.Fill(DsAMSPartMaster1)
        '    odcAlliancePartMaster.Close()

        '    If DsAMSPartMaster1.Tables("PartMaster").Rows.Count = 0 And txtNewPN.Text <> "" Then
        '        If MessageBox.Show("P/N: " & sAMSPartNumber & " no found in Alliance" & vbCrLf & "Would you like to add it to label system anyway?", "Not Found In Alliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
        '            sDescription = InputBox("Cannot find P/N: " & sAMSPartNumber & " in Alliance" & vbCrLf & "Please Enter description manually")
        '            txtAllianceDescription.Text = sDescription
        '        Else
        '            ABORT = True

        '            Return
        '        End If

        '    ElseIf DsAMSPartMaster1.Tables("PartMaster").Rows.Count > 1 Then
        '        sDescription = InputBox("More than on Alliance entry for P/N: " & sPartNumber & vbCrLf & "Please Enter description manually")
        '    ElseIf DsAMSPartMaster1.Tables("PartMaster").Rows.Count = 1 Then
        '        If txtNewPN.Text <> "" Then sDescription = DsAMSPartMaster1.Tables("Partmaster").Rows(0).Item("DescText")
        '        txtAllianceDescription.Text = sDescription
        '    End If

        '    If sDescription = "" And txtNewPN.Text <> "" Then
        '        txtNewPN.Clear()
        '        txtAllianceDescription.Clear()
        '        txtNewPN.Focus()
        '    Else
        '        Call GuessProfile(sPartNumber)
        '    End If
    End Sub

End Class