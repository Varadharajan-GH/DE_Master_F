<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGoTo
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtRef = New System.Windows.Forms.TextBox()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.lblTotRefs = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(124, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Enter Reference to goto:"
        '
        'txtRef
        '
        Me.txtRef.Location = New System.Drawing.Point(136, 14)
        Me.txtRef.Name = "txtRef"
        Me.txtRef.Size = New System.Drawing.Size(41, 20)
        Me.txtRef.TabIndex = 1
        '
        'btnGo
        '
        Me.btnGo.Location = New System.Drawing.Point(84, 51)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(75, 23)
        Me.btnGo.TabIndex = 2
        Me.btnGo.Text = "&Go"
        Me.btnGo.UseVisualStyleBackColor = True
        '
        'lblTotRefs
        '
        Me.lblTotRefs.AutoSize = True
        Me.lblTotRefs.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotRefs.Location = New System.Drawing.Point(183, 14)
        Me.lblTotRefs.Name = "lblTotRefs"
        Me.lblTotRefs.Size = New System.Drawing.Size(32, 17)
        Me.lblTotRefs.TabIndex = 3
        Me.lblTotRefs.Text = "/ 00"
        '
        'frmGoTo
        '
        Me.AcceptButton = Me.btnGo
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(242, 97)
        Me.Controls.Add(Me.lblTotRefs)
        Me.Controls.Add(Me.btnGo)
        Me.Controls.Add(Me.txtRef)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmGoTo"
        Me.Text = "GoTo"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtRef As System.Windows.Forms.TextBox
    Friend WithEvents btnGo As System.Windows.Forms.Button
    Friend WithEvents lblTotRefs As System.Windows.Forms.Label
End Class
