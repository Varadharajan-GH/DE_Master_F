<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCharMap
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCharMap))
        Me.CharGrid = New System.Windows.Forms.DataGridView()
        Me.pHolder = New System.Windows.Forms.Panel()
        CType(Me.CharGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pHolder.SuspendLayout()
        Me.SuspendLayout()
        '
        'CharGrid
        '
        Me.CharGrid.AllowUserToAddRows = False
        Me.CharGrid.AllowUserToDeleteRows = False
        Me.CharGrid.AllowUserToResizeColumns = False
        Me.CharGrid.AllowUserToResizeRows = False
        Me.CharGrid.BackgroundColor = System.Drawing.SystemColors.Control
        Me.CharGrid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.CharGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Raised
        Me.CharGrid.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable
        Me.CharGrid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        Me.CharGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.CharGrid.ColumnHeadersVisible = False
        Me.CharGrid.GridColor = System.Drawing.Color.LightGray
        Me.CharGrid.Location = New System.Drawing.Point(2, 2)
        Me.CharGrid.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.CharGrid.MultiSelect = False
        Me.CharGrid.Name = "CharGrid"
        Me.CharGrid.ReadOnly = True
        Me.CharGrid.RowHeadersVisible = False
        Me.CharGrid.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.CharGrid.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.CharGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.CharGrid.ShowCellToolTips = False
        Me.CharGrid.ShowEditingIcon = False
        Me.CharGrid.Size = New System.Drawing.Size(533, 312)
        Me.CharGrid.TabIndex = 2
        Me.chkKeepOnTop = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'chkKeepOnTop
        '
        Me.chkKeepOnTop.AutoSize = True
        Me.chkKeepOnTop.CheckState = CheckState.Checked
        Me.chkKeepOnTop.Location = New System.Drawing.Point(8, 3)
        Me.chkKeepOnTop.Name = "chkKeepOnTop"
        Me.chkKeepOnTop.Size = New System.Drawing.Size(88, 17)
        Me.chkKeepOnTop.TabIndex = 0
        Me.chkKeepOnTop.Text = "Keep on Top"
        Me.chkKeepOnTop.UseVisualStyleBackColor = True
        '
        'frmmain
        '
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.chkKeepOnTop)
        Me.Name = "frmmain"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.ResumeLayout(False)
        Me.PerformLayout()
        '
        'pHolder
        '
        Me.pHolder.BackColor = System.Drawing.SystemColors.Control
        Me.pHolder.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pHolder.AutoScroll = True
        Me.pHolder.Controls.Add(Me.CharGrid)
        Me.pHolder.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.pHolder.Location = New System.Drawing.Point(5, 25)
        Me.pHolder.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.pHolder.Name = "pHolder"
        Me.pHolder.Size = New System.Drawing.Size(540, 318)
        'Me.pHolder.AutoSize = True
        Me.pHolder.TabIndex = 7
        '
        'frmmain
        '
        'Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        'Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(550, 350)
        Me.Controls.Add(Me.pHolder)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.MaximizeBox = False
        Me.Name = "frmmain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Special Characters"
        CType(Me.CharGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pHolder.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CharGrid As System.Windows.Forms.DataGridView
    Friend WithEvents pHolder As System.Windows.Forms.Panel
End Class
