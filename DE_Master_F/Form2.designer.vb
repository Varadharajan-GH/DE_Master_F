<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.txtArtn1 = New System.Windows.Forms.TextBox()
        Me.txtOCR = New System.Windows.Forms.TextBox()
        Me.cmbArtn1 = New System.Windows.Forms.ComboBox()
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.cmdPrev = New System.Windows.Forms.Button()
        Me.cmdDone = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmdHelp = New System.Windows.Forms.Button()
        Me.txtArtn2 = New System.Windows.Forms.TextBox()
        Me.cmbArtn2 = New System.Windows.Forms.ComboBox()
        Me.txtArtn3 = New System.Windows.Forms.TextBox()
        Me.txtArtn4 = New System.Windows.Forms.TextBox()
        Me.cmbArtn3 = New System.Windows.Forms.ComboBox()
        Me.cmbArtn4 = New System.Windows.Forms.ComboBox()
        Me.lblTitChars = New System.Windows.Forms.Label()
        Me.tmTimer = New System.Windows.Forms.Timer(Me.components)
        Me.lvlist = New System.Windows.Forms.ListView()
        Me.No = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Author = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Volume = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Page = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Year = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Title = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cmdSuspend = New System.Windows.Forms.Button()
        Me.menuMain = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddRefToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteRefToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DoneToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrevToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NextToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SuspendToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QueryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ARTNToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ThemeColorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AdminToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmdInter = New System.Windows.Forms.Button()
        Me.ssStatus = New System.Windows.Forms.StatusStrip()
        Me.tsslStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.txtAuthor = New System.Windows.Forms.TextBox()
        Me.lblCurrent = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.chkNTR = New System.Windows.Forms.CheckBox()
        Me.rbYear = New System.Windows.Forms.RadioButton()
        Me.rbAuthor = New System.Windows.Forms.RadioButton()
        Me.lblAutherChars = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtPage = New System.Windows.Forms.TextBox()
        Me.txtVolume = New System.Windows.Forms.TextBox()
        Me.chkDCI = New System.Windows.Forms.CheckBox()
        Me.rbPage = New System.Windows.Forms.RadioButton()
        Me.rbVolume = New System.Windows.Forms.RadioButton()
        Me.rbTitle = New System.Windows.Forms.RadioButton()
        Me.cbField1 = New System.Windows.Forms.CheckBox()
        Me.gbAccn = New System.Windows.Forms.GroupBox()
        Me.lblUName = New System.Windows.Forms.Label()
        Me.lblAccn = New System.Windows.Forms.Label()
        Me.lblItem = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cbField2 = New System.Windows.Forms.CheckBox()
        Me.cbField3 = New System.Windows.Forms.CheckBox()
        Me.gbTime = New System.Windows.Forms.GroupBox()
        Me.lblProdTime = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblSTime = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblSuspTime = New System.Windows.Forms.Label()
        Me.cbField4 = New System.Windows.Forms.CheckBox()
        Me.cbField5 = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lblCompRefs = New System.Windows.Forms.Label()
        Me.lblCompItems = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.gbEntry1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbARTN4 = New System.Windows.Forms.RadioButton()
        Me.rbARTN3 = New System.Windows.Forms.RadioButton()
        Me.rbARTN2 = New System.Windows.Forms.RadioButton()
        Me.rbARTN1 = New System.Windows.Forms.RadioButton()
        Me.cmdGoto = New System.Windows.Forms.Button()
        Me.txtGoto = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.lblBlink = New System.Windows.Forms.Label()
        Me.cmdDelRef = New System.Windows.Forms.Button()
        Me.cmdAddRef = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.pnlDesImage = New System.Windows.Forms.Panel()
        Me.pbImage = New System.Windows.Forms.PictureBox()
        Me.pnlSource = New System.Windows.Forms.Panel()
        Me.pbSource = New System.Windows.Forms.PictureBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.cmdRefresh = New System.Windows.Forms.Button()
        Me.gbZoneTool = New System.Windows.Forms.GroupBox()
        Me.chkMergeImg = New System.Windows.Forms.CheckBox()
        Me.cmdLoadBackup = New System.Windows.Forms.Button()
        Me.cbBackup = New System.Windows.Forms.ComboBox()
        Me.cmdViewAll = New System.Windows.Forms.Button()
        Me.cmdZClose = New System.Windows.Forms.Button()
        Me.chkFullWidth = New System.Windows.Forms.CheckBox()
        Me.lblTotalSeq = New System.Windows.Forms.Label()
        Me.cmdZoning = New System.Windows.Forms.Button()
        Me.cmdAddArtn = New System.Windows.Forms.Button()
        Me.cmdDelArtn = New System.Windows.Forms.Button()
        Me.ttToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.cdColorDlg = New System.Windows.Forms.ColorDialog()
        Me.menuMain.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.ssStatus.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.gbAccn.SuspendLayout()
        Me.gbTime.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.gbEntry1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.pnlDesImage.SuspendLayout()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlSource.SuspendLayout()
        CType(Me.pbSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.gbZoneTool.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.Location = New System.Drawing.Point(738, 11)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(179, 25)
        Me.cmdClose.TabIndex = 4
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'txtTitle
        '
        Me.txtTitle.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTitle.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle.Location = New System.Drawing.Point(79, 125)
        Me.txtTitle.MaxLength = 256
        Me.txtTitle.Multiline = True
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(1222, 79)
        Me.txtTitle.TabIndex = 4
        '
        'txtArtn1
        '
        Me.txtArtn1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtArtn1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtArtn1.Location = New System.Drawing.Point(7, 3)
        Me.txtArtn1.Name = "txtArtn1"
        Me.txtArtn1.Size = New System.Drawing.Size(209, 26)
        Me.txtArtn1.TabIndex = 0
        '
        'txtOCR
        '
        Me.txtOCR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOCR.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOCR.Location = New System.Drawing.Point(11, 845)
        Me.txtOCR.Multiline = True
        Me.txtOCR.Name = "txtOCR"
        Me.txtOCR.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOCR.Size = New System.Drawing.Size(1600, 100)
        Me.txtOCR.TabIndex = 4
        Me.txtOCR.TabStop = False
        '
        'cmbArtn1
        '
        Me.cmbArtn1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbArtn1.FormattingEnabled = True
        Me.cmbArtn1.Items.AddRange(New Object() {"ARTN", "DOI", "PII"})
        Me.cmbArtn1.Location = New System.Drawing.Point(223, 3)
        Me.cmbArtn1.Name = "cmbArtn1"
        Me.cmbArtn1.Size = New System.Drawing.Size(126, 27)
        Me.cmbArtn1.TabIndex = 1
        '
        'cmdNext
        '
        Me.cmdNext.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdNext.Location = New System.Drawing.Point(201, 11)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.Size = New System.Drawing.Size(179, 25)
        Me.cmdNext.TabIndex = 1
        Me.cmdNext.Text = "&Next"
        Me.cmdNext.UseVisualStyleBackColor = True
        '
        'cmdPrev
        '
        Me.cmdPrev.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrev.Location = New System.Drawing.Point(22, 11)
        Me.cmdPrev.Name = "cmdPrev"
        Me.cmdPrev.Size = New System.Drawing.Size(179, 25)
        Me.cmdPrev.TabIndex = 0
        Me.cmdPrev.Text = "&Prev"
        Me.cmdPrev.UseVisualStyleBackColor = True
        '
        'cmdDone
        '
        Me.cmdDone.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDone.Location = New System.Drawing.Point(559, 11)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(179, 25)
        Me.cmdDone.TabIndex = 3
        Me.cmdDone.Text = "&Done"
        Me.cmdDone.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(9, 151)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 20)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Full Title:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(11, 525)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(36, 20)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Artn:"
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.BackColor = System.Drawing.Color.DodgerBlue
        Me.cmdHelp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.cmdHelp.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cmdHelp.Location = New System.Drawing.Point(1096, 11)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(179, 25)
        Me.cmdHelp.TabIndex = 6
        Me.cmdHelp.Text = "&?"
        Me.cmdHelp.UseVisualStyleBackColor = False
        '
        'txtArtn2
        '
        Me.txtArtn2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtArtn2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtArtn2.Location = New System.Drawing.Point(7, 33)
        Me.txtArtn2.Name = "txtArtn2"
        Me.txtArtn2.Size = New System.Drawing.Size(209, 26)
        Me.txtArtn2.TabIndex = 2
        '
        'cmbArtn2
        '
        Me.cmbArtn2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbArtn2.FormattingEnabled = True
        Me.cmbArtn2.Items.AddRange(New Object() {"ARTN", "DOI", "PII"})
        Me.cmbArtn2.Location = New System.Drawing.Point(223, 33)
        Me.cmbArtn2.Name = "cmbArtn2"
        Me.cmbArtn2.Size = New System.Drawing.Size(126, 27)
        Me.cmbArtn2.TabIndex = 3
        '
        'txtArtn3
        '
        Me.txtArtn3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtArtn3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtArtn3.Location = New System.Drawing.Point(7, 63)
        Me.txtArtn3.Name = "txtArtn3"
        Me.txtArtn3.Size = New System.Drawing.Size(209, 26)
        Me.txtArtn3.TabIndex = 4
        '
        'txtArtn4
        '
        Me.txtArtn4.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtArtn4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtArtn4.Location = New System.Drawing.Point(7, 93)
        Me.txtArtn4.Name = "txtArtn4"
        Me.txtArtn4.Size = New System.Drawing.Size(208, 26)
        Me.txtArtn4.TabIndex = 6
        '
        'cmbArtn3
        '
        Me.cmbArtn3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbArtn3.FormattingEnabled = True
        Me.cmbArtn3.Items.AddRange(New Object() {"ARTN", "DOI", "PII"})
        Me.cmbArtn3.Location = New System.Drawing.Point(223, 63)
        Me.cmbArtn3.Name = "cmbArtn3"
        Me.cmbArtn3.Size = New System.Drawing.Size(126, 27)
        Me.cmbArtn3.TabIndex = 5
        '
        'cmbArtn4
        '
        Me.cmbArtn4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbArtn4.FormattingEnabled = True
        Me.cmbArtn4.Items.AddRange(New Object() {"ARTN", "DOI", "PII"})
        Me.cmbArtn4.Location = New System.Drawing.Point(223, 93)
        Me.cmbArtn4.Name = "cmbArtn4"
        Me.cmbArtn4.Size = New System.Drawing.Size(126, 27)
        Me.cmbArtn4.TabIndex = 7
        '
        'lblTitChars
        '
        Me.lblTitChars.AutoSize = True
        Me.lblTitChars.Location = New System.Drawing.Point(24, 178)
        Me.lblTitChars.Name = "lblTitChars"
        Me.lblTitChars.Size = New System.Drawing.Size(22, 16)
        Me.lblTitChars.TabIndex = 24
        Me.lblTitChars.Text = "##"
        '
        'tmTimer
        '
        Me.tmTimer.Enabled = True
        Me.tmTimer.Interval = 1000
        '
        'lvlist
        '
        Me.lvlist.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvlist.BackColor = System.Drawing.Color.Firebrick
        Me.lvlist.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.No, Me.Author, Me.Volume, Me.Page, Me.Year, Me.Title})
        Me.lvlist.Font = New System.Drawing.Font("Lucida Console", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvlist.ForeColor = System.Drawing.Color.Yellow
        Me.lvlist.FullRowSelect = True
        Me.lvlist.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.lvlist.Location = New System.Drawing.Point(14, 9)
        Me.lvlist.MultiSelect = False
        Me.lvlist.Name = "lvlist"
        Me.lvlist.Size = New System.Drawing.Size(1785, 247)
        Me.lvlist.TabIndex = 29
        Me.lvlist.TabStop = False
        Me.lvlist.UseCompatibleStateImageBehavior = False
        Me.lvlist.View = System.Windows.Forms.View.Details
        Me.lvlist.Visible = False
        '
        'No
        '
        Me.No.Text = "No."
        Me.No.Width = 50
        '
        'Author
        '
        Me.Author.Text = "Author"
        Me.Author.Width = 200
        '
        'Volume
        '
        Me.Volume.Text = "Volume"
        Me.Volume.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.Volume.Width = 50
        '
        'Page
        '
        Me.Page.Text = "Page"
        Me.Page.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.Page.Width = 75
        '
        'Year
        '
        Me.Year.Text = "Year"
        Me.Year.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.Year.Width = 50
        '
        'Title
        '
        Me.Title.Text = "Title"
        Me.Title.Width = 800
        '
        'cmdSuspend
        '
        Me.cmdSuspend.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSuspend.Location = New System.Drawing.Point(380, 11)
        Me.cmdSuspend.Name = "cmdSuspend"
        Me.cmdSuspend.Size = New System.Drawing.Size(179, 25)
        Me.cmdSuspend.TabIndex = 2
        Me.cmdSuspend.Text = "Su&spend"
        Me.cmdSuspend.UseVisualStyleBackColor = True
        '
        'menuMain
        '
        Me.menuMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.DoneToolStripMenuItem, Me.PrevToolStripMenuItem, Me.NextToolStripMenuItem, Me.SuspendToolStripMenuItem, Me.QueryToolStripMenuItem, Me.ARTNToolStripMenuItem, Me.OptionsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.menuMain.Location = New System.Drawing.Point(0, 0)
        Me.menuMain.Name = "menuMain"
        Me.menuMain.Padding = New System.Windows.Forms.Padding(7, 2, 0, 2)
        Me.menuMain.Size = New System.Drawing.Size(1813, 24)
        Me.menuMain.TabIndex = 31
        Me.menuMain.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddRefToolStripMenuItem, Me.DeleteRefToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "&Edit"
        '
        'AddRefToolStripMenuItem
        '
        Me.AddRefToolStripMenuItem.Name = "AddRefToolStripMenuItem"
        Me.AddRefToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.AddRefToolStripMenuItem.Text = "Add Ref      (Ctrl+A)"
        '
        'DeleteRefToolStripMenuItem
        '
        Me.DeleteRefToolStripMenuItem.Name = "DeleteRefToolStripMenuItem"
        Me.DeleteRefToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.DeleteRefToolStripMenuItem.Text = "Delete Ref  (Ctrl+D)"
        '
        'DoneToolStripMenuItem
        '
        Me.DoneToolStripMenuItem.Name = "DoneToolStripMenuItem"
        Me.DoneToolStripMenuItem.Size = New System.Drawing.Size(76, 20)
        Me.DoneToolStripMenuItem.Text = "&Done (F10)"
        '
        'PrevToolStripMenuItem
        '
        Me.PrevToolStripMenuItem.Name = "PrevToolStripMenuItem"
        Me.PrevToolStripMenuItem.Size = New System.Drawing.Size(65, 20)
        Me.PrevToolStripMenuItem.Text = "&Prev (F4)"
        '
        'NextToolStripMenuItem
        '
        Me.NextToolStripMenuItem.Name = "NextToolStripMenuItem"
        Me.NextToolStripMenuItem.Size = New System.Drawing.Size(66, 20)
        Me.NextToolStripMenuItem.Text = "&Next (F3)"
        '
        'SuspendToolStripMenuItem
        '
        Me.SuspendToolStripMenuItem.Name = "SuspendToolStripMenuItem"
        Me.SuspendToolStripMenuItem.Size = New System.Drawing.Size(93, 20)
        Me.SuspendToolStripMenuItem.Text = "&Suspend (F11)"
        '
        'QueryToolStripMenuItem
        '
        Me.QueryToolStripMenuItem.Name = "QueryToolStripMenuItem"
        Me.QueryToolStripMenuItem.Size = New System.Drawing.Size(51, 20)
        Me.QueryToolStripMenuItem.Text = "Query"
        '
        'ARTNToolStripMenuItem
        '
        Me.ARTNToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddToolStripMenuItem, Me.DeleteToolStripMenuItem})
        Me.ARTNToolStripMenuItem.Name = "ARTNToolStripMenuItem"
        Me.ARTNToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.ARTNToolStripMenuItem.Text = "ARTN"
        '
        'AddToolStripMenuItem
        '
        Me.AddToolStripMenuItem.Name = "AddToolStripMenuItem"
        Me.AddToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.AddToolStripMenuItem.Text = "Add      (Ctrl+F7)"
        '
        'DeleteToolStripMenuItem
        '
        Me.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem"
        Me.DeleteToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.DeleteToolStripMenuItem.Text = "Delete  (Alt+F7)"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ThemeColorToolStripMenuItem})
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.OptionsToolStripMenuItem.Text = "Options"
        '
        'ThemeColorToolStripMenuItem
        '
        Me.ThemeColorToolStripMenuItem.Name = "ThemeColorToolStripMenuItem"
        Me.ThemeColorToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.ThemeColorToolStripMenuItem.Text = "Theme color"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem, Me.ContentsToolStripMenuItem, Me.AdminToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'ContentsToolStripMenuItem
        '
        Me.ContentsToolStripMenuItem.Name = "ContentsToolStripMenuItem"
        Me.ContentsToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.ContentsToolStripMenuItem.Text = "Contents     (F1)"
        '
        'AdminToolStripMenuItem
        '
        Me.AdminToolStripMenuItem.Name = "AdminToolStripMenuItem"
        Me.AdminToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.AdminToolStripMenuItem.Text = "Admin"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdInter)
        Me.GroupBox1.Controls.Add(Me.cmdClose)
        Me.GroupBox1.Controls.Add(Me.cmdDone)
        Me.GroupBox1.Controls.Add(Me.cmdHelp)
        Me.GroupBox1.Controls.Add(Me.cmdSuspend)
        Me.GroupBox1.Controls.Add(Me.cmdPrev)
        Me.GroupBox1.Controls.Add(Me.cmdNext)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 263)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.Size = New System.Drawing.Size(1302, 41)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        '
        'cmdInter
        '
        Me.cmdInter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdInter.Location = New System.Drawing.Point(917, 11)
        Me.cmdInter.Name = "cmdInter"
        Me.cmdInter.Size = New System.Drawing.Size(179, 25)
        Me.cmdInter.TabIndex = 5
        Me.cmdInter.Text = "&Interrupt"
        Me.cmdInter.UseVisualStyleBackColor = True
        '
        'ssStatus
        '
        Me.ssStatus.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslStatus})
        Me.ssStatus.Location = New System.Drawing.Point(0, 972)
        Me.ssStatus.Name = "ssStatus"
        Me.ssStatus.Size = New System.Drawing.Size(1813, 22)
        Me.ssStatus.TabIndex = 37
        Me.ttToolTip.SetToolTip(Me.ssStatus, "Status bar")
        '
        'tsslStatus
        '
        Me.tsslStatus.Name = "tsslStatus"
        Me.tsslStatus.Size = New System.Drawing.Size(106, 17)
        Me.tsslStatus.Text = "De Master is Ready"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.txtArtn1)
        Me.Panel1.Controls.Add(Me.cmbArtn1)
        Me.Panel1.Controls.Add(Me.txtArtn2)
        Me.Panel1.Controls.Add(Me.cmbArtn4)
        Me.Panel1.Controls.Add(Me.cmbArtn2)
        Me.Panel1.Controls.Add(Me.txtArtn3)
        Me.Panel1.Controls.Add(Me.cmbArtn3)
        Me.Panel1.Controls.Add(Me.txtArtn4)
        Me.Panel1.Location = New System.Drawing.Point(81, 526)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(359, 123)
        Me.Panel1.TabIndex = 1
        '
        'txtYear
        '
        Me.txtYear.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtYear.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtYear.Location = New System.Drawing.Point(309, 89)
        Me.txtYear.MaxLength = 4
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(60, 26)
        Me.txtYear.TabIndex = 3
        '
        'txtAuthor
        '
        Me.txtAuthor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtAuthor.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAuthor.Location = New System.Drawing.Point(79, 51)
        Me.txtAuthor.MaxLength = 18
        Me.txtAuthor.Name = "txtAuthor"
        Me.txtAuthor.Size = New System.Drawing.Size(289, 26)
        Me.txtAuthor.TabIndex = 0
        '
        'lblCurrent
        '
        Me.lblCurrent.AutoSize = True
        Me.lblCurrent.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrent.Location = New System.Drawing.Point(83, 24)
        Me.lblCurrent.Name = "lblCurrent"
        Me.lblCurrent.Size = New System.Drawing.Size(61, 19)
        Me.lblCurrent.TabIndex = 17
        Me.lblCurrent.Text = "00 of 00"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(9, 22)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(70, 20)
        Me.Label10.TabIndex = 17
        Me.Label10.Text = "Referance"
        '
        'chkNTR
        '
        Me.chkNTR.AutoSize = True
        Me.chkNTR.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkNTR.Location = New System.Drawing.Point(402, 54)
        Me.chkNTR.Name = "chkNTR"
        Me.chkNTR.Size = New System.Drawing.Size(50, 21)
        Me.chkNTR.TabIndex = 0
        Me.chkNTR.TabStop = False
        Me.chkNTR.Text = "NTR"
        Me.chkNTR.UseVisualStyleBackColor = True
        '
        'rbYear
        '
        Me.rbYear.AutoSize = True
        Me.rbYear.Location = New System.Drawing.Point(145, 60)
        Me.rbYear.Name = "rbYear"
        Me.rbYear.Size = New System.Drawing.Size(51, 20)
        Me.rbYear.TabIndex = 14
        Me.rbYear.Text = "Year"
        Me.rbYear.UseVisualStyleBackColor = True
        '
        'rbAuthor
        '
        Me.rbAuthor.AutoSize = True
        Me.rbAuthor.Location = New System.Drawing.Point(145, 13)
        Me.rbAuthor.Name = "rbAuthor"
        Me.rbAuthor.Size = New System.Drawing.Size(65, 20)
        Me.rbAuthor.TabIndex = 14
        Me.rbAuthor.Text = "Author"
        Me.rbAuthor.UseVisualStyleBackColor = True
        '
        'lblAutherChars
        '
        Me.lblAutherChars.AutoSize = True
        Me.lblAutherChars.Location = New System.Drawing.Point(377, 56)
        Me.lblAutherChars.Name = "lblAutherChars"
        Me.lblAutherChars.Size = New System.Drawing.Size(0, 16)
        Me.lblAutherChars.TabIndex = 26
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(265, 93)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 20)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Year:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(149, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 20)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Page:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 93)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 20)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Volume:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 20)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Author:"
        '
        'txtPage
        '
        Me.txtPage.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPage.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPage.Location = New System.Drawing.Point(197, 89)
        Me.txtPage.MaxLength = 5
        Me.txtPage.Name = "txtPage"
        Me.txtPage.Size = New System.Drawing.Size(60, 26)
        Me.txtPage.TabIndex = 2
        '
        'txtVolume
        '
        Me.txtVolume.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtVolume.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVolume.Location = New System.Drawing.Point(79, 88)
        Me.txtVolume.MaxLength = 4
        Me.txtVolume.Name = "txtVolume"
        Me.txtVolume.Size = New System.Drawing.Size(60, 26)
        Me.txtVolume.TabIndex = 1
        '
        'chkDCI
        '
        Me.chkDCI.AutoSize = True
        Me.chkDCI.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDCI.Location = New System.Drawing.Point(402, 89)
        Me.chkDCI.Name = "chkDCI"
        Me.chkDCI.Size = New System.Drawing.Size(46, 21)
        Me.chkDCI.TabIndex = 0
        Me.chkDCI.TabStop = False
        Me.chkDCI.Text = "DCI"
        Me.chkDCI.UseVisualStyleBackColor = True
        '
        'rbPage
        '
        Me.rbPage.AutoSize = True
        Me.rbPage.Location = New System.Drawing.Point(145, 44)
        Me.rbPage.Name = "rbPage"
        Me.rbPage.Size = New System.Drawing.Size(53, 20)
        Me.rbPage.TabIndex = 14
        Me.rbPage.Text = "Page"
        Me.rbPage.UseVisualStyleBackColor = True
        '
        'rbVolume
        '
        Me.rbVolume.AutoSize = True
        Me.rbVolume.Location = New System.Drawing.Point(145, 28)
        Me.rbVolume.Name = "rbVolume"
        Me.rbVolume.Size = New System.Drawing.Size(67, 20)
        Me.rbVolume.TabIndex = 14
        Me.rbVolume.Text = "Volume"
        Me.rbVolume.UseVisualStyleBackColor = True
        '
        'rbTitle
        '
        Me.rbTitle.AutoSize = True
        Me.rbTitle.Location = New System.Drawing.Point(145, 76)
        Me.rbTitle.Name = "rbTitle"
        Me.rbTitle.Size = New System.Drawing.Size(51, 20)
        Me.rbTitle.TabIndex = 14
        Me.rbTitle.Text = "Title"
        Me.rbTitle.UseVisualStyleBackColor = True
        '
        'cbField1
        '
        Me.cbField1.AutoSize = True
        Me.cbField1.Location = New System.Drawing.Point(6, 15)
        Me.cbField1.Name = "cbField1"
        Me.cbField1.Size = New System.Drawing.Size(66, 20)
        Me.cbField1.TabIndex = 33
        Me.cbField1.TabStop = False
        Me.cbField1.Text = "Author"
        Me.cbField1.UseVisualStyleBackColor = True
        '
        'gbAccn
        '
        Me.gbAccn.Controls.Add(Me.lblUName)
        Me.gbAccn.Controls.Add(Me.lblAccn)
        Me.gbAccn.Controls.Add(Me.lblItem)
        Me.gbAccn.Controls.Add(Me.Label11)
        Me.gbAccn.Controls.Add(Me.Label12)
        Me.gbAccn.Location = New System.Drawing.Point(7, -6)
        Me.gbAccn.Name = "gbAccn"
        Me.gbAccn.Size = New System.Drawing.Size(175, 77)
        Me.gbAccn.TabIndex = 0
        Me.gbAccn.TabStop = False
        '
        'lblUName
        '
        Me.lblUName.AutoSize = True
        Me.lblUName.Location = New System.Drawing.Point(-15, 15)
        Me.lblUName.Name = "lblUName"
        Me.lblUName.Size = New System.Drawing.Size(0, 16)
        Me.lblUName.TabIndex = 39
        Me.lblUName.Visible = False
        '
        'lblAccn
        '
        Me.lblAccn.AutoSize = True
        Me.lblAccn.Font = New System.Drawing.Font("Arial Narrow", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAccn.Location = New System.Drawing.Point(86, 15)
        Me.lblAccn.Name = "lblAccn"
        Me.lblAccn.Size = New System.Drawing.Size(55, 23)
        Me.lblAccn.TabIndex = 17
        Me.lblAccn.Text = "#####"
        '
        'lblItem
        '
        Me.lblItem.AutoSize = True
        Me.lblItem.Font = New System.Drawing.Font("Arial Narrow", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItem.Location = New System.Drawing.Point(85, 45)
        Me.lblItem.Name = "lblItem"
        Me.lblItem.Size = New System.Drawing.Size(28, 23)
        Me.lblItem.TabIndex = 17
        Me.lblItem.Text = "##"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(8, 46)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(59, 20)
        Me.Label11.TabIndex = 17
        Me.Label11.Text = "Item No:"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(8, 15)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(76, 20)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "Accession:"
        '
        'cbField2
        '
        Me.cbField2.AutoSize = True
        Me.cbField2.Location = New System.Drawing.Point(6, 31)
        Me.cbField2.Name = "cbField2"
        Me.cbField2.Size = New System.Drawing.Size(68, 20)
        Me.cbField2.TabIndex = 33
        Me.cbField2.TabStop = False
        Me.cbField2.Text = "Volume"
        Me.cbField2.UseVisualStyleBackColor = True
        '
        'cbField3
        '
        Me.cbField3.AutoSize = True
        Me.cbField3.Location = New System.Drawing.Point(6, 47)
        Me.cbField3.Name = "cbField3"
        Me.cbField3.Size = New System.Drawing.Size(54, 20)
        Me.cbField3.TabIndex = 33
        Me.cbField3.TabStop = False
        Me.cbField3.Text = "Page"
        Me.cbField3.UseVisualStyleBackColor = True
        '
        'gbTime
        '
        Me.gbTime.Controls.Add(Me.lblProdTime)
        Me.gbTime.Controls.Add(Me.Label7)
        Me.gbTime.Controls.Add(Me.lblSTime)
        Me.gbTime.Controls.Add(Me.Label8)
        Me.gbTime.Controls.Add(Me.Label9)
        Me.gbTime.Controls.Add(Me.lblSuspTime)
        Me.gbTime.Location = New System.Drawing.Point(7, 65)
        Me.gbTime.Name = "gbTime"
        Me.gbTime.Size = New System.Drawing.Size(175, 98)
        Me.gbTime.TabIndex = 33
        Me.gbTime.TabStop = False
        '
        'lblProdTime
        '
        Me.lblProdTime.AutoSize = True
        Me.lblProdTime.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProdTime.Location = New System.Drawing.Point(88, 47)
        Me.lblProdTime.Name = "lblProdTime"
        Me.lblProdTime.Size = New System.Drawing.Size(67, 19)
        Me.lblProdTime.TabIndex = 25
        Me.lblProdTime.Text = "00:00:00"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(8, 18)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 20)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Start Time:"
        '
        'lblSTime
        '
        Me.lblSTime.AutoSize = True
        Me.lblSTime.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSTime.Location = New System.Drawing.Point(88, 20)
        Me.lblSTime.Name = "lblSTime"
        Me.lblSTime.Size = New System.Drawing.Size(67, 19)
        Me.lblSTime.TabIndex = 20
        Me.lblSTime.Text = "00:00:00"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(8, 46)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(77, 20)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "Production:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(8, 73)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(60, 20)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "Supend:"
        '
        'lblSuspTime
        '
        Me.lblSuspTime.AutoSize = True
        Me.lblSuspTime.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSuspTime.Location = New System.Drawing.Point(88, 73)
        Me.lblSuspTime.Name = "lblSuspTime"
        Me.lblSuspTime.Size = New System.Drawing.Size(67, 19)
        Me.lblSuspTime.TabIndex = 28
        Me.lblSuspTime.Text = "00:00:00"
        '
        'cbField4
        '
        Me.cbField4.AutoSize = True
        Me.cbField4.Location = New System.Drawing.Point(6, 63)
        Me.cbField4.Name = "cbField4"
        Me.cbField4.Size = New System.Drawing.Size(52, 20)
        Me.cbField4.TabIndex = 33
        Me.cbField4.TabStop = False
        Me.cbField4.Text = "Year"
        Me.cbField4.UseVisualStyleBackColor = True
        '
        'cbField5
        '
        Me.cbField5.AutoSize = True
        Me.cbField5.Location = New System.Drawing.Point(6, 79)
        Me.cbField5.Name = "cbField5"
        Me.cbField5.Size = New System.Drawing.Size(52, 20)
        Me.cbField5.TabIndex = 33
        Me.cbField5.TabStop = False
        Me.cbField5.Text = "Title"
        Me.cbField5.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lblCompRefs)
        Me.GroupBox3.Controls.Add(Me.lblCompItems)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Location = New System.Drawing.Point(7, 157)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(175, 98)
        Me.GroupBox3.TabIndex = 36
        Me.GroupBox3.TabStop = False
        '
        'lblCompRefs
        '
        Me.lblCompRefs.AutoSize = True
        Me.lblCompRefs.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompRefs.Location = New System.Drawing.Point(65, 66)
        Me.lblCompRefs.Name = "lblCompRefs"
        Me.lblCompRefs.Size = New System.Drawing.Size(30, 20)
        Me.lblCompRefs.TabIndex = 2
        Me.lblCompRefs.Text = "###"
        '
        'lblCompItems
        '
        Me.lblCompItems.AutoSize = True
        Me.lblCompItems.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompItems.Location = New System.Drawing.Point(65, 40)
        Me.lblCompItems.Name = "lblCompItems"
        Me.lblCompItems.Size = New System.Drawing.Size(30, 20)
        Me.lblCompItems.TabIndex = 2
        Me.lblCompItems.Text = "###"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(11, 67)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(40, 20)
        Me.Label15.TabIndex = 1
        Me.Label15.Text = "Refs:"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(11, 40)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(45, 20)
        Me.Label14.TabIndex = 1
        Me.Label14.Text = "Items:"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(6, 11)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(77, 20)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "Completion"
        '
        'gbEntry1
        '
        Me.gbEntry1.Controls.Add(Me.GroupBox2)
        Me.gbEntry1.Controls.Add(Me.cmdGoto)
        Me.gbEntry1.Controls.Add(Me.txtGoto)
        Me.gbEntry1.Controls.Add(Me.Label16)
        Me.gbEntry1.Controls.Add(Me.lblBlink)
        Me.gbEntry1.Controls.Add(Me.cmdDelRef)
        Me.gbEntry1.Controls.Add(Me.cmdAddRef)
        Me.gbEntry1.Controls.Add(Me.lblTitChars)
        Me.gbEntry1.Controls.Add(Me.txtTitle)
        Me.gbEntry1.Controls.Add(Me.Label5)
        Me.gbEntry1.Controls.Add(Me.chkDCI)
        Me.gbEntry1.Controls.Add(Me.txtVolume)
        Me.gbEntry1.Controls.Add(Me.txtPage)
        Me.gbEntry1.Controls.Add(Me.Label1)
        Me.gbEntry1.Controls.Add(Me.Label2)
        Me.gbEntry1.Controls.Add(Me.Label3)
        Me.gbEntry1.Controls.Add(Me.Label4)
        Me.gbEntry1.Controls.Add(Me.lblAutherChars)
        Me.gbEntry1.Controls.Add(Me.chkNTR)
        Me.gbEntry1.Controls.Add(Me.Label10)
        Me.gbEntry1.Controls.Add(Me.lblCurrent)
        Me.gbEntry1.Controls.Add(Me.txtAuthor)
        Me.gbEntry1.Controls.Add(Me.txtYear)
        Me.gbEntry1.Location = New System.Drawing.Point(12, 306)
        Me.gbEntry1.Name = "gbEntry1"
        Me.gbEntry1.Size = New System.Drawing.Size(1307, 208)
        Me.gbEntry1.TabIndex = 0
        Me.gbEntry1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbARTN4)
        Me.GroupBox2.Controls.Add(Me.rbARTN3)
        Me.GroupBox2.Controls.Add(Me.rbARTN2)
        Me.GroupBox2.Controls.Add(Me.rbARTN1)
        Me.GroupBox2.Controls.Add(Me.rbTitle)
        Me.GroupBox2.Controls.Add(Me.rbYear)
        Me.GroupBox2.Controls.Add(Me.cbField1)
        Me.GroupBox2.Controls.Add(Me.rbPage)
        Me.GroupBox2.Controls.Add(Me.rbAuthor)
        Me.GroupBox2.Controls.Add(Me.rbVolume)
        Me.GroupBox2.Controls.Add(Me.cbField2)
        Me.GroupBox2.Controls.Add(Me.cbField5)
        Me.GroupBox2.Controls.Add(Me.cbField3)
        Me.GroupBox2.Controls.Add(Me.cbField4)
        Me.GroupBox2.Location = New System.Drawing.Point(696, 14)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(286, 101)
        Me.GroupBox2.TabIndex = 43
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Visible = False
        '
        'rbARTN4
        '
        Me.rbARTN4.AutoSize = True
        Me.rbARTN4.Location = New System.Drawing.Point(216, 75)
        Me.rbARTN4.Name = "rbARTN4"
        Me.rbARTN4.Size = New System.Drawing.Size(58, 20)
        Me.rbARTN4.TabIndex = 41
        Me.rbARTN4.TabStop = True
        Me.rbARTN4.Text = "Artn4"
        Me.rbARTN4.UseVisualStyleBackColor = True
        '
        'rbARTN3
        '
        Me.rbARTN3.AutoSize = True
        Me.rbARTN3.Location = New System.Drawing.Point(216, 55)
        Me.rbARTN3.Name = "rbARTN3"
        Me.rbARTN3.Size = New System.Drawing.Size(58, 20)
        Me.rbARTN3.TabIndex = 40
        Me.rbARTN3.TabStop = True
        Me.rbARTN3.Text = "Artn3"
        Me.rbARTN3.UseVisualStyleBackColor = True
        '
        'rbARTN2
        '
        Me.rbARTN2.AutoSize = True
        Me.rbARTN2.Location = New System.Drawing.Point(216, 35)
        Me.rbARTN2.Name = "rbARTN2"
        Me.rbARTN2.Size = New System.Drawing.Size(58, 20)
        Me.rbARTN2.TabIndex = 39
        Me.rbARTN2.TabStop = True
        Me.rbARTN2.Text = "Artn2"
        Me.rbARTN2.UseVisualStyleBackColor = True
        '
        'rbARTN1
        '
        Me.rbARTN1.AutoSize = True
        Me.rbARTN1.Location = New System.Drawing.Point(216, 15)
        Me.rbARTN1.Name = "rbARTN1"
        Me.rbARTN1.Size = New System.Drawing.Size(58, 20)
        Me.rbARTN1.TabIndex = 38
        Me.rbARTN1.TabStop = True
        Me.rbARTN1.Text = "Artn1"
        Me.rbARTN1.UseVisualStyleBackColor = True
        '
        'cmdGoto
        '
        Me.cmdGoto.Location = New System.Drawing.Point(461, 23)
        Me.cmdGoto.Name = "cmdGoto"
        Me.cmdGoto.Size = New System.Drawing.Size(34, 23)
        Me.cmdGoto.TabIndex = 42
        Me.cmdGoto.TabStop = False
        Me.cmdGoto.Text = "&Go"
        Me.cmdGoto.UseVisualStyleBackColor = True
        '
        'txtGoto
        '
        Me.txtGoto.Location = New System.Drawing.Point(408, 22)
        Me.txtGoto.Name = "txtGoto"
        Me.txtGoto.Size = New System.Drawing.Size(48, 23)
        Me.txtGoto.TabIndex = 0
        Me.txtGoto.TabStop = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(344, 26)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(60, 16)
        Me.Label16.TabIndex = 40
        Me.Label16.Text = "Goto Ref"
        '
        'lblBlink
        '
        Me.lblBlink.AutoSize = True
        Me.lblBlink.BackColor = System.Drawing.Color.WhiteSmoke
        Me.lblBlink.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBlink.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblBlink.Location = New System.Drawing.Point(10, 127)
        Me.lblBlink.Name = "lblBlink"
        Me.lblBlink.Size = New System.Drawing.Size(34, 13)
        Me.lblBlink.TabIndex = 39
        Me.lblBlink.Text = "blink"
        '
        'cmdDelRef
        '
        Me.cmdDelRef.Location = New System.Drawing.Point(255, 22)
        Me.cmdDelRef.Name = "cmdDelRef"
        Me.cmdDelRef.Size = New System.Drawing.Size(75, 23)
        Me.cmdDelRef.TabIndex = 38
        Me.cmdDelRef.TabStop = False
        Me.cmdDelRef.Text = "&Del Ref"
        Me.cmdDelRef.UseVisualStyleBackColor = True
        '
        'cmdAddRef
        '
        Me.cmdAddRef.Location = New System.Drawing.Point(169, 22)
        Me.cmdAddRef.Name = "cmdAddRef"
        Me.cmdAddRef.Size = New System.Drawing.Size(75, 23)
        Me.cmdAddRef.TabIndex = 37
        Me.cmdAddRef.TabStop = False
        Me.cmdAddRef.Text = "&Add Ref"
        Me.cmdAddRef.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.pnlDesImage)
        Me.Panel2.Controls.Add(Me.Panel1)
        Me.Panel2.Controls.Add(Me.pnlSource)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.cmdRefresh)
        Me.Panel2.Controls.Add(Me.gbZoneTool)
        Me.Panel2.Controls.Add(Me.lblTotalSeq)
        Me.Panel2.Controls.Add(Me.cmdZoning)
        Me.Panel2.Controls.Add(Me.gbEntry1)
        Me.Panel2.Controls.Add(Me.txtOCR)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Controls.Add(Me.cmdAddArtn)
        Me.Panel2.Controls.Add(Me.cmdDelArtn)
        Me.Panel2.Controls.Add(Me.lvlist)
        Me.Panel2.Font = New System.Drawing.Font("Times New Roman", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel2.Location = New System.Drawing.Point(13, 22)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1800, 950)
        Me.Panel2.TabIndex = 38
        '
        'pnlDesImage
        '
        Me.pnlDesImage.AutoScroll = True
        Me.pnlDesImage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlDesImage.Controls.Add(Me.pbImage)
        Me.pnlDesImage.Location = New System.Drawing.Point(11, 559)
        Me.pnlDesImage.Name = "pnlDesImage"
        Me.pnlDesImage.Size = New System.Drawing.Size(1600, 269)
        Me.pnlDesImage.TabIndex = 34
        '
        'pbImage
        '
        Me.pbImage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pbImage.Location = New System.Drawing.Point(2, 3)
        Me.pbImage.Name = "pbImage"
        Me.pbImage.Size = New System.Drawing.Size(200, 165)
        Me.pbImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbImage.TabIndex = 11
        Me.pbImage.TabStop = False
        '
        'pnlSource
        '
        Me.pnlSource.AutoScroll = True
        Me.pnlSource.BackColor = System.Drawing.SystemColors.MenuBar
        Me.pnlSource.Controls.Add(Me.pbSource)
        Me.pnlSource.Location = New System.Drawing.Point(148, 9)
        Me.pnlSource.Name = "pnlSource"
        Me.pnlSource.Size = New System.Drawing.Size(1171, 253)
        Me.pnlSource.TabIndex = 30
        '
        'pbSource
        '
        Me.pbSource.Location = New System.Drawing.Point(5, 9)
        Me.pbSource.Name = "pbSource"
        Me.pbSource.Size = New System.Drawing.Size(100, 50)
        Me.pbSource.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbSource.TabIndex = 0
        Me.pbSource.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.gbAccn)
        Me.Panel3.Controls.Add(Me.gbTime)
        Me.Panel3.Controls.Add(Me.GroupBox3)
        Me.Panel3.Location = New System.Drawing.Point(1344, 268)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(203, 263)
        Me.Panel3.TabIndex = 33
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Location = New System.Drawing.Point(615, 530)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(100, 25)
        Me.cmdRefresh.TabIndex = 32
        Me.cmdRefresh.Text = "R&efresh"
        Me.cmdRefresh.UseVisualStyleBackColor = True
        '
        'gbZoneTool
        '
        Me.gbZoneTool.Controls.Add(Me.chkMergeImg)
        Me.gbZoneTool.Controls.Add(Me.cmdLoadBackup)
        Me.gbZoneTool.Controls.Add(Me.cbBackup)
        Me.gbZoneTool.Controls.Add(Me.cmdViewAll)
        Me.gbZoneTool.Controls.Add(Me.cmdZClose)
        Me.gbZoneTool.Controls.Add(Me.chkFullWidth)
        Me.gbZoneTool.Location = New System.Drawing.Point(12, 6)
        Me.gbZoneTool.Name = "gbZoneTool"
        Me.gbZoneTool.Size = New System.Drawing.Size(136, 253)
        Me.gbZoneTool.TabIndex = 1
        Me.gbZoneTool.TabStop = False
        Me.gbZoneTool.Text = "Zoning Tools"
        '
        'chkMergeImg
        '
        Me.chkMergeImg.AutoSize = True
        Me.chkMergeImg.Location = New System.Drawing.Point(9, 192)
        Me.chkMergeImg.Name = "chkMergeImg"
        Me.chkMergeImg.Size = New System.Drawing.Size(108, 20)
        Me.chkMergeImg.TabIndex = 5
        Me.chkMergeImg.Text = "Merge Images"
        Me.chkMergeImg.UseVisualStyleBackColor = True
        '
        'cmdLoadBackup
        '
        Me.cmdLoadBackup.Location = New System.Drawing.Point(22, 156)
        Me.cmdLoadBackup.Name = "cmdLoadBackup"
        Me.cmdLoadBackup.Size = New System.Drawing.Size(92, 23)
        Me.cmdLoadBackup.TabIndex = 4
        Me.cmdLoadBackup.Text = "Load Backup"
        Me.cmdLoadBackup.UseVisualStyleBackColor = True
        '
        'cbBackup
        '
        Me.cbBackup.FormattingEnabled = True
        Me.cbBackup.Location = New System.Drawing.Point(16, 117)
        Me.cbBackup.Name = "cbBackup"
        Me.cbBackup.Size = New System.Drawing.Size(103, 23)
        Me.cbBackup.TabIndex = 3
        '
        'cmdViewAll
        '
        Me.cmdViewAll.Location = New System.Drawing.Point(22, 36)
        Me.cmdViewAll.Name = "cmdViewAll"
        Me.cmdViewAll.Size = New System.Drawing.Size(92, 23)
        Me.cmdViewAll.TabIndex = 0
        Me.cmdViewAll.TabStop = False
        Me.cmdViewAll.Text = "View All"
        Me.cmdViewAll.UseVisualStyleBackColor = True
        '
        'cmdZClose
        '
        Me.cmdZClose.AutoSize = True
        Me.cmdZClose.Location = New System.Drawing.Point(22, 77)
        Me.cmdZClose.Name = "cmdZClose"
        Me.cmdZClose.Size = New System.Drawing.Size(92, 26)
        Me.cmdZClose.TabIndex = 1
        Me.cmdZClose.TabStop = False
        Me.cmdZClose.Text = "Close Zoning"
        Me.cmdZClose.UseVisualStyleBackColor = True
        '
        'chkFullWidth
        '
        Me.chkFullWidth.AutoSize = True
        Me.chkFullWidth.Checked = True
        Me.chkFullWidth.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFullWidth.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.chkFullWidth.FlatAppearance.MouseDownBackColor = System.Drawing.Color.WhiteSmoke
        Me.chkFullWidth.Location = New System.Drawing.Point(9, 218)
        Me.chkFullWidth.Name = "chkFullWidth"
        Me.chkFullWidth.Size = New System.Drawing.Size(114, 20)
        Me.chkFullWidth.TabIndex = 2
        Me.chkFullWidth.TabStop = False
        Me.chkFullWidth.Text = "Select Full Line"
        Me.ttToolTip.SetToolTip(Me.chkFullWidth, "Select Entire Line when dragged")
        Me.chkFullWidth.UseVisualStyleBackColor = True
        '
        'lblTotalSeq
        '
        Me.lblTotalSeq.AutoSize = True
        Me.lblTotalSeq.Location = New System.Drawing.Point(612, 538)
        Me.lblTotalSeq.Name = "lblTotalSeq"
        Me.lblTotalSeq.Size = New System.Drawing.Size(0, 16)
        Me.lblTotalSeq.TabIndex = 31
        Me.lblTotalSeq.Visible = False
        '
        'cmdZoning
        '
        Me.cmdZoning.Location = New System.Drawing.Point(525, 530)
        Me.cmdZoning.Name = "cmdZoning"
        Me.cmdZoning.Size = New System.Drawing.Size(75, 25)
        Me.cmdZoning.TabIndex = 30
        Me.cmdZoning.Text = "&Zoning"
        Me.cmdZoning.UseVisualStyleBackColor = True
        '
        'cmdAddArtn
        '
        Me.cmdAddArtn.Image = Global.DE_Master.My.Resources.Resources.Add
        Me.cmdAddArtn.Location = New System.Drawing.Point(446, 530)
        Me.cmdAddArtn.Name = "cmdAddArtn"
        Me.cmdAddArtn.Size = New System.Drawing.Size(25, 25)
        Me.cmdAddArtn.TabIndex = 0
        Me.cmdAddArtn.TabStop = False
        Me.cmdAddArtn.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
        Me.cmdAddArtn.UseVisualStyleBackColor = True
        '
        'cmdDelArtn
        '
        Me.cmdDelArtn.Image = CType(resources.GetObject("cmdDelArtn.Image"), System.Drawing.Image)
        Me.cmdDelArtn.Location = New System.Drawing.Point(476, 530)
        Me.cmdDelArtn.Name = "cmdDelArtn"
        Me.cmdDelArtn.Size = New System.Drawing.Size(25, 25)
        Me.cmdDelArtn.TabIndex = 18
        Me.cmdDelArtn.TabStop = False
        Me.cmdDelArtn.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(1152, 680)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.ssStatus)
        Me.Controls.Add(Me.menuMain)
        Me.Font = New System.Drawing.Font("Times New Roman", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.menuMain
        Me.Name = "frmMain"
        Me.Text = "DE Master"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.menuMain.ResumeLayout(False)
        Me.menuMain.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.ssStatus.ResumeLayout(False)
        Me.ssStatus.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.gbAccn.ResumeLayout(False)
        Me.gbAccn.PerformLayout()
        Me.gbTime.ResumeLayout(False)
        Me.gbTime.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.gbEntry1.ResumeLayout(False)
        Me.gbEntry1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.pnlDesImage.ResumeLayout(False)
        Me.pnlDesImage.PerformLayout()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlSource.ResumeLayout(False)
        Me.pnlSource.PerformLayout()
        CType(Me.pbSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.gbZoneTool.ResumeLayout(False)
        Me.gbZoneTool.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents txtTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtArtn1 As System.Windows.Forms.TextBox
    Friend WithEvents txtOCR As System.Windows.Forms.TextBox
    Friend WithEvents cmbArtn1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmdNext As System.Windows.Forms.Button
    Friend WithEvents cmdPrev As System.Windows.Forms.Button
    Friend WithEvents pbImage As System.Windows.Forms.PictureBox
    Friend WithEvents cmdDone As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    Friend WithEvents txtArtn2 As System.Windows.Forms.TextBox
    Friend WithEvents cmbArtn2 As System.Windows.Forms.ComboBox
    Friend WithEvents txtArtn3 As System.Windows.Forms.TextBox
    Friend WithEvents txtArtn4 As System.Windows.Forms.TextBox
    Friend WithEvents cmbArtn3 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbArtn4 As System.Windows.Forms.ComboBox
    Friend WithEvents cmdAddArtn As System.Windows.Forms.Button
    Friend WithEvents cmdDelArtn As System.Windows.Forms.Button
    Friend WithEvents lblTitChars As System.Windows.Forms.Label
    Friend WithEvents tmTimer As System.Windows.Forms.Timer
    Friend WithEvents lvlist As System.Windows.Forms.ListView
    Friend WithEvents No As System.Windows.Forms.ColumnHeader
    Friend WithEvents Author As System.Windows.Forms.ColumnHeader
    Friend WithEvents Volume As System.Windows.Forms.ColumnHeader
    Friend WithEvents Page As System.Windows.Forms.ColumnHeader
    Friend WithEvents Year As System.Windows.Forms.ColumnHeader
    Friend WithEvents Title As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdSuspend As System.Windows.Forms.Button
    Friend WithEvents menuMain As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DoneToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NextToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdInter As System.Windows.Forms.Button
    Friend WithEvents SuspendToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ssStatus As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents txtAuthor As System.Windows.Forms.TextBox
    Friend WithEvents lblCurrent As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents chkNTR As System.Windows.Forms.CheckBox
    Friend WithEvents rbYear As System.Windows.Forms.RadioButton
    Friend WithEvents rbAuthor As System.Windows.Forms.RadioButton
    Friend WithEvents lblAutherChars As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPage As System.Windows.Forms.TextBox
    Friend WithEvents txtVolume As System.Windows.Forms.TextBox
    Friend WithEvents chkDCI As System.Windows.Forms.CheckBox
    Friend WithEvents rbPage As System.Windows.Forms.RadioButton
    Friend WithEvents rbVolume As System.Windows.Forms.RadioButton
    Friend WithEvents rbTitle As System.Windows.Forms.RadioButton
    Friend WithEvents cbField1 As System.Windows.Forms.CheckBox
    Friend WithEvents gbAccn As System.Windows.Forms.GroupBox
    Friend WithEvents lblAccn As System.Windows.Forms.Label
    Friend WithEvents lblItem As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cbField2 As System.Windows.Forms.CheckBox
    Friend WithEvents cbField3 As System.Windows.Forms.CheckBox
    Friend WithEvents gbTime As System.Windows.Forms.GroupBox
    Friend WithEvents lblProdTime As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblSTime As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lblSuspTime As System.Windows.Forms.Label
    Friend WithEvents cbField4 As System.Windows.Forms.CheckBox
    Friend WithEvents cbField5 As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblCompRefs As System.Windows.Forms.Label
    Friend WithEvents lblCompItems As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents gbEntry1 As System.Windows.Forms.GroupBox
    Friend WithEvents ARTNToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrevToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddRefToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteRefToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmdDelRef As System.Windows.Forms.Button
    Friend WithEvents cmdAddRef As System.Windows.Forms.Button
    Friend WithEvents lblUName As System.Windows.Forms.Label
    Friend WithEvents cmdGoto As System.Windows.Forms.Button
    Friend WithEvents txtGoto As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents lblBlink As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents pnlSource As System.Windows.Forms.Panel
    Friend WithEvents pbSource As System.Windows.Forms.PictureBox
    Friend WithEvents cmdZoning As System.Windows.Forms.Button
    Friend WithEvents gbZoneTool As System.Windows.Forms.GroupBox
    Friend WithEvents chkFullWidth As System.Windows.Forms.CheckBox
    Friend WithEvents ttToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents cmdZClose As System.Windows.Forms.Button
    Friend WithEvents cmdViewAll As System.Windows.Forms.Button
    Friend WithEvents lblTotalSeq As System.Windows.Forms.Label
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbARTN4 As System.Windows.Forms.RadioButton
    Friend WithEvents rbARTN3 As System.Windows.Forms.RadioButton
    Friend WithEvents rbARTN2 As System.Windows.Forms.RadioButton
    Friend WithEvents rbARTN1 As System.Windows.Forms.RadioButton
    Friend WithEvents tsslStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents cmdLoadBackup As System.Windows.Forms.Button
    Friend WithEvents cbBackup As System.Windows.Forms.ComboBox
    Friend WithEvents AdminToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents chkMergeImg As System.Windows.Forms.CheckBox
    Friend WithEvents pnlDesImage As System.Windows.Forms.Panel
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ThemeColorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cdColorDlg As System.Windows.Forms.ColorDialog
    Friend WithEvents QueryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
