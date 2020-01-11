<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChoose
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
        Me.lblBanner = New System.Windows.Forms.Label()
        Me.lblOp = New System.Windows.Forms.Label()
        Me.lblQC = New System.Windows.Forms.Label()
        Me.lblClose = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblBanner
        '
        Me.lblBanner.AutoSize = True
        Me.lblBanner.Font = New System.Drawing.Font("Arial Narrow", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBanner.ForeColor = System.Drawing.Color.Maroon
        Me.lblBanner.Location = New System.Drawing.Point(118, 56)
        Me.lblBanner.Name = "lblBanner"
        Me.lblBanner.Size = New System.Drawing.Size(223, 46)
        Me.lblBanner.TabIndex = 0
        Me.lblBanner.Text = "DEM for FCR"
        '
        'lblOp
        '
        Me.lblOp.AutoSize = True
        Me.lblOp.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOp.ForeColor = System.Drawing.Color.Green
        Me.lblOp.Location = New System.Drawing.Point(134, 155)
        Me.lblOp.Name = "lblOp"
        Me.lblOp.Size = New System.Drawing.Size(115, 25)
        Me.lblOp.TabIndex = 3
        Me.lblOp.Text = "Production"
        '
        'lblQC
        '
        Me.lblQC.AutoSize = True
        Me.lblQC.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQC.ForeColor = System.Drawing.Color.Green
        Me.lblQC.Location = New System.Drawing.Point(280, 155)
        Me.lblQC.Name = "lblQC"
        Me.lblQC.Size = New System.Drawing.Size(45, 25)
        Me.lblQC.TabIndex = 4
        Me.lblQC.Text = "QC"
        '
        'lblClose
        '
        Me.lblClose.AutoSize = True
        Me.lblClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClose.Location = New System.Drawing.Point(439, 9)
        Me.lblClose.Name = "lblClose"
        Me.lblClose.Size = New System.Drawing.Size(18, 17)
        Me.lblClose.TabIndex = 5
        Me.lblClose.Text = "X"
        '
        'frmChoose
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.ClientSize = New System.Drawing.Size(458, 247)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblClose)
        Me.Controls.Add(Me.lblQC)
        Me.Controls.Add(Me.lblOp)
        Me.Controls.Add(Me.lblBanner)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmChoose"
        Me.ShowIcon = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblBanner As System.Windows.Forms.Label
    Friend WithEvents lblOp As System.Windows.Forms.Label
    Friend WithEvents lblQC As System.Windows.Forms.Label
    Friend WithEvents lblClose As System.Windows.Forms.Label
End Class
