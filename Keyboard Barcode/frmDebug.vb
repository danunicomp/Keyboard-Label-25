Public Class frmDebug
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
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents chkAllowPrint As System.Windows.Forms.CheckBox
    Friend WithEvents chkAllowLogging As System.Windows.Forms.CheckBox
    Friend WithEvents chkOverrideExclusions As System.Windows.Forms.CheckBox
    Friend WithEvents chkKeepDump As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.chkAllowPrint = New System.Windows.Forms.CheckBox
        Me.chkAllowLogging = New System.Windows.Forms.CheckBox
        Me.chkOverrideExclusions = New System.Windows.Forms.CheckBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.chkKeepDump = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'chkAllowPrint
        '
        Me.chkAllowPrint.Location = New System.Drawing.Point(56, 40)
        Me.chkAllowPrint.Name = "chkAllowPrint"
        Me.chkAllowPrint.TabIndex = 0
        Me.chkAllowPrint.Text = "No Printing"
        '
        'chkAllowLogging
        '
        Me.chkAllowLogging.Location = New System.Drawing.Point(56, 72)
        Me.chkAllowLogging.Name = "chkAllowLogging"
        Me.chkAllowLogging.Size = New System.Drawing.Size(128, 24)
        Me.chkAllowLogging.TabIndex = 1
        Me.chkAllowLogging.Text = "Do Not Log"
        '
        'chkOverrideExclusions
        '
        Me.chkOverrideExclusions.Location = New System.Drawing.Point(56, 104)
        Me.chkOverrideExclusions.Name = "chkOverrideExclusions"
        Me.chkOverrideExclusions.Size = New System.Drawing.Size(144, 24)
        Me.chkOverrideExclusions.TabIndex = 2
        Me.chkOverrideExclusions.Text = "Override Exclusions"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(56, 216)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 3
        Me.btnSave.Text = "Save"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(176, 216)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        '
        'chkKeepDump
        '
        Me.chkKeepDump.Location = New System.Drawing.Point(56, 136)
        Me.chkKeepDump.Name = "chkKeepDump"
        Me.chkKeepDump.Size = New System.Drawing.Size(144, 32)
        Me.chkKeepDump.TabIndex = 5
        Me.chkKeepDump.Text = "Do Not Delete Dump When Finished"
        '
        'frmDebug
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.chkKeepDump)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.chkOverrideExclusions)
        Me.Controls.Add(Me.chkAllowLogging)
        Me.Controls.Add(Me.chkAllowPrint)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDebug"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmDebug"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub frmDebug_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If NoPrint Then chkAllowPrint.Checked = True
        If OverrideExclusions Then chkOverrideExclusions.Checked = True
        If NoLogging Then chkAllowLogging.Checked = True
        If KeepDump Then chkKeepDump.Checked = True
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        NoPrint = chkAllowPrint.Checked
        OverrideExclusions = chkOverrideExclusions.Checked
        NoLogging = chkAllowLogging.Checked
        KeepDump = chkKeepDump.Checked
        Call WriteConfig()
        Me.Close()

    End Sub

    Private Sub chkAllowPrint_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllowPrint.CheckedChanged

    End Sub

    Private Sub chkAllowLogging_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllowLogging.CheckedChanged

    End Sub
End Class
