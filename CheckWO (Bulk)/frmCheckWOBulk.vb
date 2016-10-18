
'v1.1 - Added rev level and failure conditions
'
'v1.0 3/15/05 - Initial Release
'
'
Imports System.Data
Imports System.Data.OleDb

Public Class CheckWOBulk
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
    Friend WithEvents txtPartNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label


    Friend WithEvents txtDesc As System.Windows.Forms.TextBox

    Friend WithEvents lblResult As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtOEMPN As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtRevLevel As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CheckWOBulk))
        Me.txtWorkOrder = New System.Windows.Forms.TextBox
        Me.txtPartNumber = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.lblResult = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtOEMPN = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtRevLevel = New System.Windows.Forms.TextBox
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
        'CheckWOBulk
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 302)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtRevLevel)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtOEMPN)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDesc)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtPartNumber)
        Me.Controls.Add(Me.txtWorkOrder)
        Me.Controls.Add(Me.lblResult)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "CheckWOBulk"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Check Work Order"
        Me.ResumeLayout(False)

    End Sub

#End Region


    'Dim sWorkOrder As String
    'Dim sDesc As String

    Private odcLog As System.Data.OleDb.OleDbConnection
    Private odaLog As System.Data.OleDb.OleDbDataAdapter
    Private cmdLog As System.Data.OleDb.OleDbCommandBuilder
    Private myDs As DataSet
    Private sqlStr As String

    Structure WorkOrder
        Dim WONumber As String
        Dim PartNumber As String
        Dim OEMPN As String
        Dim Description As String
        Dim RevLevel As String
        Dim AMSPN As String
    End Structure
    Dim curWorkOrder As WorkOrder

    Private Sub CheckWOBulk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        'curWorkOrder.WONumber = txtWorkOrder.Text
        'DsWorkOrder1.Clear()
        'selWorkOrder.CommandText = "SELECT WONumber, QuantityRequired, PartNumber FROM WOHeader WHERE WONumber = '" & curWorkOrder.WONumber & "'"
        'Try
        '    odaWorkOrders.Fill(DsWorkOrder1)
        'Catch ex As Exception
        '    MsgBox("Problem with Alliance WOHead database")
        'End Try
        'If DsWorkOrder1.Tables("WOHeader").Rows.Count = 1 Then
        '    LookupWorkOrder = True
        '    curWorkOrder.AMSPN = UCase(DsWorkOrder1.Tables("WOHeader").Rows(0).Item("PartNumber"))
        '    ' sWorkOrder = DsWorkOrder1.Tables("WOHeader").Rows(0).Item("WONumber")
        'Else
        '    LookupWorkOrder = False
        'End If
        ''Debug.Write(DsWorkOrder1.Tables("WOHeader").Rows(0).Item("PartNumber"))

    End Function

    Private Function LookUpDescription(ByVal sPN As String) As Boolean
        'DsPartDesc1.Clear()
        'selPartMaster.CommandText = "SELECT PartNumber, DescText, Revision FROM PartMaster WHERE PartNumber = '" & sPN & "'"
        'Try
        '    odaPartMaster.Fill(DsPartDesc1)
        'Catch ex As Exception
        '    MsgBox("Problem with Alliance Part Master")
        'End Try
        'If DsPartDesc1.Tables("PartMaster").Rows.Count = 1 Then
        '    LookUpDescription = True
        '    curWorkOrder.Description = IIf(IsDBNull(DsPartDesc1.Tables("PartMaster").Rows(0).Item("DescText")), "", DsPartDesc1.Tables("PartMaster").Rows(0).Item("DescText"))
        '    If curWorkOrder.Description = "" Then curWorkOrder.Description = "DESCRIPTION NOT IN ALLIANCE. PLEASE CORRECT"
        '    'Revision
        '    curWorkOrder.RevLevel = IIf(IsDBNull(DsPartDesc1.Tables("PartMaster").Rows(0).Item("Revision")), "", DsPartDesc1.Tables("PartMaster").Rows(0).Item("Revision"))

        'Else
        '    LookUpDescription = False
        'End If
    End Function
    Private Function LookUpProfile(ByVal sAMSPN As String) As Boolean
        'Dim sPN As String
        'If Microsoft.VisualBasic.Left(sAMSPN, 2) = "00" Then
        '    curWorkOrder.PartNumber = Mid(sAMSPN, 3)
        'Else
        '    If sAMSPN = "042H1292U" Then curWorkOrder.PartNumber = "42H1292U"

        'End If
        'DsProfile1.Clear()
        'selProfile.CommandText = "SELECT PartNumber, ProductCode, OEMPartNumber FROM ProfileMaster WHERE PartNumber = '" & curWorkOrder.PartNumber & "'"
        'Try
        '    odaProfileMaster.Fill(DsProfile1)
        'Catch ex As Exception
        '    MsgBox("Problem with Master Profile Database")
        'End Try

        'If DsProfile1.Tables("ProfileMaster").Rows.Count = 1 Then
        '    LookUpProfile = True
        '    curWorkOrder.OEMPN = IIf(IsDBNull(DsProfile1.Tables("ProfileMaster").Rows(0).Item("OEMPartNumber")), "0", DsProfile1.Tables("ProfileMaster").Rows(0).Item("OEMPartNumber"))
        'Else
        '    LookUpProfile = False
        'End If
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
        'DsProfile1.Clear()
        'DsPartDesc1.Clear()
        'DsWorkOrder1.Clear()
        txtOEMPN.ForeColor = Color.Black
        txtRevLevel.ForeColor = Color.Black
        txtOEMPN.Clear()
        txtRevLevel.Clear()
    End Sub

    Private Sub txtWorkOrder_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWorkOrder.TextChanged
        Call ResetAll()
    End Sub

    Private Sub OpenAllDatabases()

        'ALLIANCE FOR WORK ORDER LIST WITH DESCRIPTION
        'Dim i As Integer
        Dim connStr As String = "User ID=sa;Data Source=""HAL\AllianceMFG"";Tag with column collation when possible=False;Initial Catalog=KBD2002;Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Workstation ID=SHAFFER2;Use Encryption for Data=False;Packet Size=4096"
        odcLog = New OleDbConnection(connStr)
        odcLog.Open()
        myDs = New DataSet

        sqlStr = "SELECT * FROM LogFile WHERE LexSerial = ''"
        odaLog = New OleDbDataAdapter(sqlStr, odcLog)
        'odaLog.SelectCommand.CommandText = sqlStr

        cmdLog = New OleDbCommandBuilder(odaLog)
        myDs.Clear()
        odaLog.Fill(myDs, "Logfile")


    End Sub ' Open All Databases

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
