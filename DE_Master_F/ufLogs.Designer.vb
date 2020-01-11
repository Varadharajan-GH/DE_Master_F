<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ufLogs
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblLogs = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblLogs
        '
        Me.lblLogs.AutoSize = True
        Me.lblLogs.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLogs.Location = New System.Drawing.Point(5, 4)
        Me.lblLogs.Name = "lblLogs"
        Me.lblLogs.Size = New System.Drawing.Size(32, 15)
        Me.lblLogs.TabIndex = 0
        Me.lblLogs.Text = "Logs"
        '
        'ufLogs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(722, 502)
        Me.Controls.Add(Me.lblLogs)
        Me.Name = "ufLogs"
        Me.Text = "ufLogs"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblLogs As System.Windows.Forms.Label
End Class
