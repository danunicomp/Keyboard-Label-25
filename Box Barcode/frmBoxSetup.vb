Public Class frmBoxSetup
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
    Friend WithEvents txtBoxStatonID As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSaveConfig As System.Windows.Forms.Button
    Friend WithEvents btnCancelConfig As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBoxSetup))
        Me.btnSaveConfig = New System.Windows.Forms.Button
        Me.btnCancelConfig = New System.Windows.Forms.Button
        Me.txtBoxStatonID = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'btnSaveConfig
        '
        Me.btnSaveConfig.Location = New System.Drawing.Point(208, 160)
        Me.btnSaveConfig.Name = "btnSaveConfig"
        Me.btnSaveConfig.TabIndex = 0
        Me.btnSaveConfig.Text = "Save"
        '
        'btnCancelConfig
        '
        Me.btnCancelConfig.Location = New System.Drawing.Point(120, 160)
        Me.btnCancelConfig.Name = "btnCancelConfig"
        Me.btnCancelConfig.TabIndex = 1
        Me.btnCancelConfig.Text = "Cancel"
        '
        'txtBoxStatonID
        '
        Me.txtBoxStatonID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBoxStatonID.Location = New System.Drawing.Point(120, 40)
        Me.txtBoxStatonID.Name = "txtBoxStatonID"
        Me.txtBoxStatonID.TabIndex = 2
        Me.txtBoxStatonID.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 22)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Station ID:"
        '
        'frmBoxSetup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 198)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtBoxStatonID)
        Me.Controls.Add(Me.btnCancelConfig)
        Me.Controls.Add(Me.btnSaveConfig)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmBoxSetup"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmBoxSetup"
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private Sub btnCancelConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelConfig.Click
        Me.Close()
    End Sub

    Private Sub btnSaveConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveConfig.Click
        sStationID = UCase(RTrim(LTrim(txtBoxStatonID.Text)))
        Call WriteConfig()
        Me.Close()
    End Sub

    Private Sub frmBoxSetup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtBoxStatonID.Text = sStationID
    End Sub

    Private Sub frmBoxSetup_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        txtBoxStatonID.Focus()
    End Sub
End Class
