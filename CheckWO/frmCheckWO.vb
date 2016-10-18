
'v1.1 - Added rev level and failure conditions
'
'v1.0 3/15/05 - Initial Release
'
'


Public Class Form1
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
    Friend WithEvents odaWorkOrders As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents odcAMS As System.Data.OleDb.OleDbConnection
    Friend WithEvents txtPartNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents selWorkOrder As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents DsWorkOrder1 As CheckWO.dsWorkOrder
    Friend WithEvents odaPartMaster As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents selPartMaster As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents DsPartDesc1 As CheckWO.dsPartDesc
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents odaProfileMaster As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbInsertCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents odcProfileMaster As System.Data.OleDb.OleDbConnection
    Friend WithEvents DsProfile1 As CheckWO.dsProfile
    Friend WithEvents selProfile As System.Data.OleDb.OleDbCommand
    Friend WithEvents lblResult As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtOEMPN As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtRevLevel As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.txtWorkOrder = New System.Windows.Forms.TextBox
        Me.odaWorkOrders = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.odcAMS = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.selWorkOrder = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.txtPartNumber = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.DsWorkOrder1 = New CheckWO.dsWorkOrder
        Me.odaPartMaster = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand2 = New System.Data.OleDb.OleDbCommand
        Me.selPartMaster = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand2 = New System.Data.OleDb.OleDbCommand
        Me.DsPartDesc1 = New CheckWO.dsPartDesc
        Me.odaProfileMaster = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand3 = New System.Data.OleDb.OleDbCommand
        Me.odcProfileMaster = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand3 = New System.Data.OleDb.OleDbCommand
        Me.selProfile = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand3 = New System.Data.OleDb.OleDbCommand
        Me.DsProfile1 = New CheckWO.dsProfile
        Me.lblResult = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtOEMPN = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtRevLevel = New System.Windows.Forms.TextBox
        CType(Me.DsWorkOrder1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DsPartDesc1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DsProfile1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtWorkOrder
        '
        Me.txtWorkOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkOrder.Location = New System.Drawing.Point(144, 40)
        Me.txtWorkOrder.Name = "txtWorkOrder"
        Me.txtWorkOrder.TabIndex = 0
        Me.txtWorkOrder.Text = ""
        '
        'odaWorkOrders
        '
        Me.odaWorkOrders.DeleteCommand = Me.OleDbDeleteCommand1
        Me.odaWorkOrders.InsertCommand = Me.OleDbInsertCommand1
        Me.odaWorkOrders.SelectCommand = Me.selWorkOrder
        Me.odaWorkOrders.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "WOHeader", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("WONumber", "WONumber"), New System.Data.Common.DataColumnMapping("QuantityRequired", "QuantityRequired"), New System.Data.Common.DataColumnMapping("PartNumber", "PartNumber")})})
        Me.odaWorkOrders.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM WOHeader WHERE (WONumber = ?) AND (PartNumber = ? OR ? IS NULL AND Pa" & _
        "rtNumber IS NULL) AND (QuantityRequired = ? OR ? IS NULL AND QuantityRequired IS" & _
        " NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.odcAMS
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_WONumber", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "WONumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber1", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_QuantityRequired", System.Data.OleDb.OleDbType.Double, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QuantityRequired", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_QuantityRequired1", System.Data.OleDb.OleDbType.Double, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QuantityRequired", System.Data.DataRowVersion.Original, Nothing))
        '
        'odcAMS
        '
        Me.odcAMS.ConnectionString = "User ID=sa;Data Source=""HAL\AllianceMFG"";Tag with column collation when possible=" & _
        "False;Initial Catalog=KBD2002;Use Procedure for Prepare=1;Auto Translate=True;Pe" & _
        "rsist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encr" & _
        "yption for Data=False;Packet Size=4096"
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO WOHeader(WONumber, QuantityRequired, PartNumber) VALUES (?, ?, ?); SE" & _
        "LECT WONumber, QuantityRequired, PartNumber FROM WOHeader WHERE (WONumber = ?)"
        Me.OleDbInsertCommand1.Connection = Me.odcAMS
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("WONumber", System.Data.OleDb.OleDbType.VarWChar, 10, "WONumber"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("QuantityRequired", System.Data.OleDb.OleDbType.Double, 8, "QuantityRequired"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Select_WONumber", System.Data.OleDb.OleDbType.VarWChar, 10, "WONumber"))
        '
        'selWorkOrder
        '
        Me.selWorkOrder.CommandText = "SELECT WONumber, QuantityRequired, PartNumber FROM WOHeader"
        Me.selWorkOrder.Connection = Me.odcAMS
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE WOHeader SET WONumber = ?, QuantityRequired = ?, PartNumber = ? WHERE (WON" & _
        "umber = ?) AND (PartNumber = ? OR ? IS NULL AND PartNumber IS NULL) AND (Quantit" & _
        "yRequired = ? OR ? IS NULL AND QuantityRequired IS NULL); SELECT WONumber, Quant" & _
        "ityRequired, PartNumber FROM WOHeader WHERE (WONumber = ?)"
        Me.OleDbUpdateCommand1.Connection = Me.odcAMS
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("WONumber", System.Data.OleDb.OleDbType.VarWChar, 10, "WONumber"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("QuantityRequired", System.Data.OleDb.OleDbType.Double, 8, "QuantityRequired"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_WONumber", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "WONumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber1", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_QuantityRequired", System.Data.OleDb.OleDbType.Double, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QuantityRequired", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_QuantityRequired1", System.Data.OleDb.OleDbType.Double, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QuantityRequired", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Select_WONumber", System.Data.OleDb.OleDbType.VarWChar, 10, "WONumber"))
        '
        'txtPartNumber
        '
        Me.txtPartNumber.BackColor = System.Drawing.SystemColors.Control
        Me.txtPartNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartNumber.Location = New System.Drawing.Point(144, 77)
        Me.txtPartNumber.Name = "txtPartNumber"
        Me.txtPartNumber.Size = New System.Drawing.Size(136, 29)
        Me.txtPartNumber.TabIndex = 1
        Me.txtPartNumber.TabStop = False
        Me.txtPartNumber.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 25)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Work Order:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 79)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 25)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Part Number:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 22)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Desc:"
        '
        'txtDesc
        '
        Me.txtDesc.BackColor = System.Drawing.SystemColors.Control
        Me.txtDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesc.Location = New System.Drawing.Point(64, 120)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(416, 20)
        Me.txtDesc.TabIndex = 4
        Me.txtDesc.TabStop = False
        Me.txtDesc.Text = ""
        '
        'DsWorkOrder1
        '
        Me.DsWorkOrder1.DataSetName = "dsWorkOrder"
        Me.DsWorkOrder1.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'odaPartMaster
        '
        Me.odaPartMaster.DeleteCommand = Me.OleDbDeleteCommand2
        Me.odaPartMaster.InsertCommand = Me.OleDbInsertCommand2
        Me.odaPartMaster.SelectCommand = Me.selPartMaster
        Me.odaPartMaster.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "PartMaster", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("PartNumber", "PartNumber"), New System.Data.Common.DataColumnMapping("DescText", "DescText"), New System.Data.Common.DataColumnMapping("Revision", "Revision")})})
        Me.odaPartMaster.UpdateCommand = Me.OleDbUpdateCommand2
        '
        'OleDbDeleteCommand2
        '
        Me.OleDbDeleteCommand2.CommandText = "DELETE FROM PartMaster WHERE (PartNumber = ?) AND (DescText = ? OR ? IS NULL AND " & _
        "DescText IS NULL) AND (Revision = ? OR ? IS NULL AND Revision IS NULL)"
        Me.OleDbDeleteCommand2.Connection = Me.odcAMS
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DescText", System.Data.OleDb.OleDbType.VarWChar, 60, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DescText", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DescText1", System.Data.OleDb.OleDbType.VarWChar, 60, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DescText", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Revision", System.Data.OleDb.OleDbType.VarWChar, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Revision", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Revision1", System.Data.OleDb.OleDbType.VarWChar, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Revision", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand2
        '
        Me.OleDbInsertCommand2.CommandText = "INSERT INTO PartMaster(PartNumber, DescText, Revision) VALUES (?, ?, ?); SELECT P" & _
        "artNumber, DescText, Revision FROM PartMaster WHERE (PartNumber = ?)"
        Me.OleDbInsertCommand2.Connection = Me.odcAMS
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("DescText", System.Data.OleDb.OleDbType.VarWChar, 60, "DescText"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Revision", System.Data.OleDb.OleDbType.VarWChar, 4, "Revision"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Select_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        '
        'selPartMaster
        '
        Me.selPartMaster.CommandText = "SELECT PartNumber, DescText, Revision FROM PartMaster"
        Me.selPartMaster.Connection = Me.odcAMS
        '
        'OleDbUpdateCommand2
        '
        Me.OleDbUpdateCommand2.CommandText = "UPDATE PartMaster SET PartNumber = ?, DescText = ?, Revision = ? WHERE (PartNumbe" & _
        "r = ?) AND (DescText = ? OR ? IS NULL AND DescText IS NULL) AND (Revision = ? OR" & _
        " ? IS NULL AND Revision IS NULL); SELECT PartNumber, DescText, Revision FROM Par" & _
        "tMaster WHERE (PartNumber = ?)"
        Me.OleDbUpdateCommand2.Connection = Me.odcAMS
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("DescText", System.Data.OleDb.OleDbType.VarWChar, 60, "DescText"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Revision", System.Data.OleDb.OleDbType.VarWChar, 4, "Revision"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DescText", System.Data.OleDb.OleDbType.VarWChar, 60, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DescText", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DescText1", System.Data.OleDb.OleDbType.VarWChar, 60, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DescText", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Revision", System.Data.OleDb.OleDbType.VarWChar, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Revision", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Revision1", System.Data.OleDb.OleDbType.VarWChar, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Revision", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Select_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 30, "PartNumber"))
        '
        'DsPartDesc1
        '
        Me.DsPartDesc1.DataSetName = "dsPartDesc"
        Me.DsPartDesc1.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'odaProfileMaster
        '
        Me.odaProfileMaster.DeleteCommand = Me.OleDbDeleteCommand3
        Me.odaProfileMaster.InsertCommand = Me.OleDbInsertCommand3
        Me.odaProfileMaster.SelectCommand = Me.selProfile
        Me.odaProfileMaster.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "ProfileMaster", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("PartNumber", "PartNumber"), New System.Data.Common.DataColumnMapping("ProductCode", "ProductCode"), New System.Data.Common.DataColumnMapping("OEMPartNumber", "OEMPartNumber")})})
        Me.odaProfileMaster.UpdateCommand = Me.OleDbUpdateCommand3
        '
        'OleDbDeleteCommand3
        '
        Me.OleDbDeleteCommand3.CommandText = "DELETE FROM ProfileMaster WHERE (PartNumber = ?) AND (OEMPartNumber = ? OR ? IS N" & _
        "ULL AND OEMPartNumber IS NULL) AND (ProductCode = ? OR ? IS NULL AND ProductCode" & _
        " IS NULL)"
        Me.OleDbDeleteCommand3.Connection = Me.odcProfileMaster
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OEMPartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OEMPartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OEMPartNumber1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OEMPartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        '
        'odcProfileMaster
        '
        Me.odcProfileMaster.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=0;Jet OLEDB:Database Password=;Data Source=""\\Hal\unicomp\Share\KB_B" & _
        "ARCODES\Profiles.mdb"";Password=;Jet OLEDB:Engine Type=5;Jet OLEDB:Global Bulk Tr" & _
        "ansactions=1;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet O" & _
        "LEDB:SFP=False;Extended Properties=;Mode=Read;Jet OLEDB:New Database Password=;J" & _
        "et OLEDB:Create System Database=False;Jet OLEDB:Don't Copy Locale on Compact=Fal" & _
        "se;Jet OLEDB:Compact Without Replica Repair=False;User ID=Admin;Jet OLEDB:Encryp" & _
        "t Database=False"
        '
        'OleDbInsertCommand3
        '
        Me.OleDbInsertCommand3.CommandText = "INSERT INTO ProfileMaster(PartNumber, ProductCode, OEMPartNumber) VALUES (?, ?, ?" & _
        ")"
        Me.OleDbInsertCommand3.Connection = Me.odcProfileMaster
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, "PartNumber"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, "ProductCode"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("OEMPartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, "OEMPartNumber"))
        '
        'selProfile
        '
        Me.selProfile.CommandText = "SELECT PartNumber, ProductCode, OEMPartNumber FROM ProfileMaster"
        Me.selProfile.Connection = Me.odcProfileMaster
        '
        'OleDbUpdateCommand3
        '
        Me.OleDbUpdateCommand3.CommandText = "UPDATE ProfileMaster SET PartNumber = ?, ProductCode = ?, OEMPartNumber = ? WHERE" & _
        " (PartNumber = ?) AND (OEMPartNumber = ? OR ? IS NULL AND OEMPartNumber IS NULL)" & _
        " AND (ProductCode = ? OR ? IS NULL AND ProductCode IS NULL)"
        Me.OleDbUpdateCommand3.Connection = Me.odcProfileMaster
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("PartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, "PartNumber"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, "ProductCode"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("OEMPartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, "OEMPartNumber"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OEMPartNumber", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OEMPartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OEMPartNumber1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OEMPartNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ProductCode1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ProductCode", System.Data.DataRowVersion.Original, Nothing))
        '
        'DsProfile1
        '
        Me.DsProfile1.DataSetName = "dsProfile"
        Me.DsProfile1.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'lblResult
        '
        Me.lblResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblResult.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblResult.ForeColor = System.Drawing.SystemColors.Highlight
        Me.lblResult.Location = New System.Drawing.Point(16, 224)
        Me.lblResult.Name = "lblResult"
        Me.lblResult.Size = New System.Drawing.Size(456, 72)
        Me.lblResult.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 154)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(95, 25)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "OEM P/N:"
        '
        'txtOEMPN
        '
        Me.txtOEMPN.BackColor = System.Drawing.SystemColors.Control
        Me.txtOEMPN.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOEMPN.Location = New System.Drawing.Point(144, 152)
        Me.txtOEMPN.Name = "txtOEMPN"
        Me.txtOEMPN.Size = New System.Drawing.Size(136, 29)
        Me.txtOEMPN.TabIndex = 8
        Me.txtOEMPN.TabStop = False
        Me.txtOEMPN.Text = ""
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(288, 82)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 18)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "OEM P/N:"
        '
        'txtRevLevel
        '
        Me.txtRevLevel.BackColor = System.Drawing.SystemColors.Control
        Me.txtRevLevel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRevLevel.Location = New System.Drawing.Point(360, 80)
        Me.txtRevLevel.Name = "txtRevLevel"
        Me.txtRevLevel.Size = New System.Drawing.Size(80, 22)
        Me.txtRevLevel.TabIndex = 10
        Me.txtRevLevel.TabStop = False
        Me.txtRevLevel.Text = ""
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 302)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtRevLevel)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtOEMPN)
        Me.Controls.Add(Me.lblResult)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDesc)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtPartNumber)
        Me.Controls.Add(Me.txtWorkOrder)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Check Work Order"
        CType(Me.DsWorkOrder1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DsPartDesc1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DsProfile1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    'Dim sWorkOrder As String
    'Dim sDesc As String

    Structure WorkOrder
        Dim WONumber As String
        Dim PartNumber As String
        Dim OEMPN As String
        Dim Description As String
        Dim RevLevel As String
        Dim AMSPN As String
    End Structure
    Dim curWorkOrder As WorkOrder

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Check Work Order v1.1"
    End Sub
    Protected Overrides Function ProcessDialogKey(ByVal keydata As System.Windows.Forms.Keys) As Boolean
        Dim key As System.Windows.Forms.Keys = keydata
        If key = Keys.Tab Then
            key = Keys.Enter
        End If
        If key = Keys.Enter Then
            If txtWorkOrder.Focused Then
                If LookupWorkOrder() Then
                    If LookUpDescription(curWorkOrder.AMSPN) Then

                        Call PopulateFields()
                    Else
                        MessageBox.Show("Part Number: " & curWorkOrder.PartNumber & vbCrLf & "Does not Exsist", "PArt Number not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                Else
                    MessageBox.Show("Work Order: " & curWorkOrder.WONumber & vbCrLf & "Not found in system", "Work Order Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    txtWorkOrder.Clear()
                    Call ResetAll()
                    Exit Function
                End If
            End If
            'If txtCustomPN.Focused Then
            '    Call ProcessCustom()
            'End If
        End If
        'User ID=sa;Data Source="HAL\AllianceMFG";Tag with column collation when possible=False;Initial Catalog=KBD2002;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider="SQLOLEDB.1";Workstation ID=SHAFFER2;Use Encryption for Data=False;Packet Size=4096
        Return MyBase.ProcessDialogKey(keydata)
    End Function

    Private Function LookupWorkOrder() As Boolean
        curWorkOrder.WONumber = txtWorkOrder.Text
        DsWorkOrder1.Clear()
        selWorkOrder.CommandText = "SELECT WONumber, QuantityRequired, PartNumber FROM WOHeader WHERE WONumber = '" & curWorkOrder.WONumber & "'"
        Try
            odaWorkOrders.Fill(DsWorkOrder1)
        Catch ex As Exception
            MsgBox("Problem with Alliance WOHead database")
        End Try
        If DsWorkOrder1.Tables("WOHeader").Rows.Count = 1 Then
            LookupWorkOrder = True
            curWorkOrder.AMSPN = UCase(DsWorkOrder1.Tables("WOHeader").Rows(0).Item("PartNumber"))
            ' sWorkOrder = DsWorkOrder1.Tables("WOHeader").Rows(0).Item("WONumber")
        Else
            LookupWorkOrder = False
        End If
        'Debug.Write(DsWorkOrder1.Tables("WOHeader").Rows(0).Item("PartNumber"))

    End Function

    Private Function LookUpDescription(ByVal sPN As String) As Boolean
        DsPartDesc1.Clear()
        selPartMaster.CommandText = "SELECT PartNumber, DescText, Revision FROM PartMaster WHERE PartNumber = '" & sPN & "'"
        Try
            odaPartMaster.Fill(DsPartDesc1)
        Catch ex As Exception
            MsgBox("Problem with Alliance Part Master")
        End Try
        If DsPartDesc1.Tables("PartMaster").Rows.Count = 1 Then
            LookUpDescription = True
            curWorkOrder.Description = IIf(IsDBNull(DsPartDesc1.Tables("PartMaster").Rows(0).Item("DescText")), "", DsPartDesc1.Tables("PartMaster").Rows(0).Item("DescText"))
            If curWorkOrder.Description = "" Then curWorkOrder.Description = "DESCRIPTION NOT IN ALLIANCE. PLEASE CORRECT"
            'Revision
            curWorkOrder.RevLevel = IIf(IsDBNull(DsPartDesc1.Tables("PartMaster").Rows(0).Item("Revision")), "", DsPartDesc1.Tables("PartMaster").Rows(0).Item("Revision"))

        Else
            LookUpDescription = False
        End If
    End Function
    Private Function LookUpProfile(ByVal sAMSPN As String) As Boolean
        Dim sPN As String
        If Microsoft.VisualBasic.Left(sAMSPN, 2) = "00" Then
            curWorkOrder.PartNumber = Mid(sAMSPN, 3)
        Else
            If sAMSPN = "042H1292U" Then curWorkOrder.PartNumber = "42H1292U"

        End If
        DsProfile1.Clear()
        selProfile.CommandText = "SELECT PartNumber, ProductCode, OEMPartNumber FROM ProfileMaster WHERE PartNumber = '" & curWorkOrder.PartNumber & "'"
        Try
            odaProfileMaster.Fill(DsProfile1)
        Catch ex As Exception
            MsgBox("Problem with Master Profile Database")
        End Try

        If DsProfile1.Tables("ProfileMaster").Rows.Count = 1 Then
            LookUpProfile = True
            curWorkOrder.OEMPN = IIf(IsDBNull(DsProfile1.Tables("ProfileMaster").Rows(0).Item("OEMPartNumber")), "0", DsProfile1.Tables("ProfileMaster").Rows(0).Item("OEMPartNumber"))
        Else
            LookUpProfile = False
        End If
    End Function

    Private Sub PopulateFields()
        Dim GOOD As Boolean

        Dim sMessage As String
        GOOD = True

        If LookUpProfile(curWorkOrder.AMSPN) Then
            sMessage = sMessage & "Profile Loaded in Label Database" & vbCrLf
        Else
            sMessage = sMessage & "Profile NOT Loaded in Label Database. Tell Dan." & vbCrLf
        End If

        txtPartNumber.Text = curWorkOrder.AMSPN
        txtDesc.Text = curWorkOrder.Description
        If curWorkOrder.RevLevel = "" Then
            txtRevLevel.ForeColor = Color.Red
            txtRevLevel.Text = "NONE"
            GOOD = False
            sMessage = sMessage & "Rev Level not set!" & vbCrLf
        Else
            txtRevLevel.ForeColor = Color.Black
            txtRevLevel.Text = curWorkOrder.RevLevel
            sMessage = sMessage & "Rev Level ok" & vbCrLf
        End If


        If curWorkOrder.OEMPN <> curWorkOrder.PartNumber Then
            txtOEMPN.ForeColor = Color.Red
            txtOEMPN.Text = "ERROR"
            GOOD = False
            sMessage = sMessage & "Problem with OEM P/N in Label Database" & vbCrLf
        Else
            txtOEMPN.ForeColor = Color.Black
            txtOEMPN.Text = curWorkOrder.OEMPN
            sMessage = sMessage & "OEM P/N Found" & vbCrLf
        End If

        If GOOD Then
            lblResult.ForeColor = Color.Green
            sMessage = sMessage & "OK TO RUN!"
        Else
            lblResult.ForeColor = Color.Red
            sMessage = sMessage & "PROBLEMS. CONTACT ENGINEERING"
        End If

        lblResult.Text = sMessage
    End Sub

    Private Sub ResetAll()
        ' curWorkOrder.PartNumber = Nothing
        ' sWorkOrder = Nothing
        curWorkOrder = Nothing
        'sDesc = Nothing
        txtDesc.Clear()
        txtPartNumber.Clear()
        lblResult.Text = ""
        DsProfile1.Clear()
        DsPartDesc1.Clear()
        DsWorkOrder1.Clear()
        txtOEMPN.ForeColor = Color.Black
        txtRevLevel.ForeColor = Color.Black
        txtOEMPN.Clear()
        txtRevLevel.Clear()
    End Sub

    Private Sub txtWorkOrder_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWorkOrder.TextChanged
        Call ResetAll()
    End Sub



#Region "Supress Keypresses"
    Private Sub txtOEMPN_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOEMPN.KeyPress
        e.Handled = True
    End Sub
    Private Sub txtRevLevel_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRevLevel.KeyPress
        e.Handled = True
    End Sub
    Private Sub txtDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDesc.KeyPress
        e.Handled = True
    End Sub
    Private Sub txtPartNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPartNumber.KeyPress
        e.Handled = True
    End Sub
#End Region

End Class
