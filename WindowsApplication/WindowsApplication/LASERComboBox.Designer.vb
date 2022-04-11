<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class LASERComboBox
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.cmb = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'cmb
        '
        Me.cmb.FormattingEnabled = True
        Me.cmb.Location = New System.Drawing.Point(2, 2)
        Me.cmb.Name = "cmb"
        Me.cmb.Size = New System.Drawing.Size(205, 22)
        Me.cmb.TabIndex = 0
        '
        'LASERComboBox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.cmb)
        Me.Font = New System.Drawing.Font("Calibri", 9.0!)
        Me.Name = "LASERComboBox"
        Me.Size = New System.Drawing.Size(208, 26)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmb As ComboBox
End Class
