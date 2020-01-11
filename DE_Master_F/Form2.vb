Imports System.Xml
Imports System.Text.RegularExpressions

Structure Log
    Dim System_Type As String
    Dim System As String
    Dim Session_Start_Date As String
    Dim Session_Start_Time As String
    Dim Session_End_Date As String
    Dim Session_End_Time As String
    Dim Session_Production As String
    Dim Session_Suspend As String
    Dim Item_Accession As String
    Dim Item_Itemno As String
    Dim Status As String
    Dim Referances_Open As String
    Dim Referances_Close As String
    Dim Referances_Inserted As String
    Dim Referances_Deleted As String

End Structure

Public Class frmMain
    Dim srcPath As String = "D:\data\C_Process\"
    'Dim srcPath As String = "\\172.21.15.100\data\"
    Dim xmlDoc As New XmlDocument(), xnlCitations As XmlNodeList
    Dim iNo As Integer, strInFile As String, curRef As Integer
    Dim strInPath As String
    Dim strOutPath As String
    Dim strBackUpPath As String
    Dim strCompPath As String
    Dim strTempPath As String
    Dim strHelpPath As String
    Dim strTifBackUpPath As String
    Dim strLogPath As String
    Dim strQueryPath As String
    Dim dtSTime As DateTime
    Dim bNTR() As Boolean
    Dim bPattent() As Boolean
    Dim swProdTime As Stopwatch = New Stopwatch
    Dim swSuspTime As Stopwatch = New Stopwatch
    Dim swCurProd As Stopwatch = New Stopwatch
    Dim swCurSusp As Stopwatch = New Stopwatch
    Dim strPriorPath As String
    Dim strCurPath As String
    Dim iCompItems As Integer, iCompRefs As Integer
    Dim policyFile As String = srcPath & "Admin\policy.txt"
    Dim mRect As Rectangle
    'Dim 'swLogFile As IO.StreamWriter
    Dim bIsPrior As Boolean
    Dim bIsPatent As Boolean
    Dim LogEntry As New Log

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim i As Integer

        If frmChoose.chosenTool = ToolMode.QC Then
            strInPath = srcPath & "Output\"
            strCurPath = srcPath & "QC\Current\"
            strPriorPath = srcPath & "QC\Priority\"
            strOutPath = srcPath & "QC\Output\"
            strCompPath = srcPath & "QC\Completed\"
            strTempPath = srcPath & "QC\Temp\"
            strHelpPath = srcPath & "Help\"
            strQueryPath = srcPath & "Query\"
            strBackUpPath = "C:\C_Backup\QC\"
            strTifBackUpPath = srcPath & "Input\tifBackup\QC\"
            strLogPath = srcPath & "C_Logs\QC\"
        Else
            strInPath = srcPath & "Input\"
            strCurPath = srcPath & "Current\"
            strPriorPath = srcPath & "Priority\"
            strOutPath = srcPath & "Output\"
            strCompPath = srcPath & "Completed\"
            strTempPath = srcPath & "Temp\"
            strHelpPath = srcPath & "Help\"
            strQueryPath = srcPath & "Query\"
            strBackUpPath = "C:\C_Backup\" & Format(Today, "yyyyMMdd") & "\"
            strTifBackUpPath = srcPath & "Input\tifBackup\"
            strLogPath = srcPath & "C_Logs\" & Format(Today, "yyyyMMdd") & "\"
        End If
        'Create Directories if they doesn't exist
        MakeDirectory(strCurPath)
        MakeDirectory(strPriorPath)
        MakeDirectory(strOutPath)
        MakeDirectory(strCompPath)
        MakeDirectory(strTempPath)
        MakeDirectory(strQueryPath)
        MakeDirectory(strBackUpPath)
        MakeDirectory(strTifBackUpPath)
        MakeDirectory(strLogPath)

        'tester()
        'Application.Exit()

        iCompItems = 0
        iCompRefs = 0
        lblCompItems.Text = iCompItems
        lblCompRefs.Text = iCompRefs
        dtSTime = Now
        swProdTime = Stopwatch.StartNew
        For i = 2 To 4
            Me.Controls.Find("txtArtn" & i, True)(0).Visible = False
            Me.Controls.Find("cmbArtn" & i, True)(0).Visible = False
        Next

        strInFile = fnGetXMLFile(strPriorPath)
        If strInFile = String.Empty Then
            strInFile = fnGetXMLFile(strInPath)
            bIsPrior = False
            If strInFile = "" Then
                MsgBox("No files for input", MsgBoxStyle.Information)
            Else
                '--------------------------------
                'swLogFile = New IO.StreamWriter(strLogPath & frmLogin.txtUserName.Text & ".txt", True)
                'swLogFile.WriteLine()
                'swLogFile.WriteLine("=============================================================================================")

                MakeLogFile(strLogPath & frmLogin.txtUserName.Text & ".xml")
                LogEntry.System_Type = ""
                LogEntry.System = ""
                If Net.Dns.GetHostEntry(Net.Dns.GetHostName()).AddressList.Count > 3 Then
                    'swLogFile.WriteLine("User " & frmLogin.txtUserName.Text & " logged in. IP address : " & Net.Dns.GetHostEntry(Net.Dns.GetHostName()).AddressList(3).ToString())
                    LogEntry.System_Type = "IP_Address"
                    LogEntry.System = Net.Dns.GetHostEntry(Net.Dns.GetHostName()).AddressList(3).ToString()
                Else
                    'swLogFile.WriteLine("User " & frmLogin.txtUserName.Text & " logged in to the system " & Net.Dns.GetHostName())
                    LogEntry.System_Type = "Name"
                    LogEntry.System = Net.Dns.GetHostName()
                End If
                'swLogFile.WriteLine("=============================================================================================")
                'swLogFile.WriteLine("Date" & vbTab & vbTab & "  Time" & vbTab & "  Accession" & vbTab & "Item_No" & vbTab &
                '"Refs_Open" & vbTab & "Refs_Close" & vbTab & "Production" & vbTab & "Suspend")
                'swLogFile.Close()
                '--------------------------------
                fnLoadFile(strInFile)
            End If
        Else
            '------------------------------------------
            'swLogFile = New IO.StreamWriter(strLogPath & frmLogin.txtUserName.Text & ".xml", True)
            'swLogFile.WriteLine()
            'swLogFile.WriteLine("=============================================================================================")

            MakeLogFile(strLogPath & frmLogin.txtUserName.Text & ".xml")
            LogEntry.System_Type = ""
            LogEntry.System = ""
            If Net.Dns.GetHostEntry(Net.Dns.GetHostName()).AddressList.Count > 3 Then
                'swLogFile.WriteLine("User " & frmLogin.txtUserName.Text & " logged in. IP address : " & Net.Dns.GetHostEntry(Net.Dns.GetHostName()).AddressList(3).ToString())
                LogEntry.System_Type = "IP_Address"
                LogEntry.System = Net.Dns.GetHostEntry(Net.Dns.GetHostName()).AddressList(3).ToString()
            Else
                'swLogFile.WriteLine("User " & frmLogin.txtUserName.Text & " logged in to the system " & Net.Dns.GetHostName())
                LogEntry.System_Type = "Name"
                LogEntry.System = Net.Dns.GetHostName()
            End If
            'swLogFile.WriteLine("=============================================================================================")
            'swLogFile.WriteLine("Date" & vbTab & vbTab & "  Time" & vbTab & "  Accession" & vbTab & "Item_No" & vbTab &
            '"Refs_Open" & vbTab & "Refs_Close" & vbTab & "Production" & vbTab & "Suspend")
            'swLogFile.Close()
            '--------------------------------
            bIsPrior = True
            fnLoadFile(strInFile)
        End If
    End Sub

    Private Sub frmMain_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        focusToBox()
    End Sub

    Public Sub New()
        InitializeComponent()
        'Me.DoubleBuffered = True
    End Sub

    Private Sub frmMain_EnabledChanged(sender As Object, e As EventArgs) Handles Me.EnabledChanged
        If Me.Enabled = True Then
            swSuspTime.Stop()
            swProdTime.Start()

            swCurSusp.Stop()
            swCurProd.Start()
        Else
            swProdTime.Stop()
            swSuspTime.Start()

            swCurProd.Stop()
            swCurSusp.Start()
        End If
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            MsgBox("Use close button instead")
            e.Cancel = True
        End If
    End Sub

    Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F1
                Help.ShowHelp(Me, strHelpPath & "ISIHELP.hlp.12-1-11.HLP")
                e.Handled = True
            Case Keys.F4
                cmdPrev.PerformClick()
                e.Handled = True
            Case Keys.F3
                cmdNext.PerformClick()
                e.Handled = True
            Case Keys.F10
                cmdDone.PerformClick()
                e.Handled = True
            Case Keys.F11
                cmdSuspend.PerformClick()
                e.Handled = True
            Case Keys.F7
                If (e.Modifiers = Keys.Control) Then
                    cmdAddArtn.PerformClick()
                    e.Handled = True
                ElseIf (e.Modifiers = Keys.Alt) Then
                    cmdDelArtn.PerformClick()
                    e.Handled = True
                ElseIf (e.Modifiers = Keys.None) Then
                    If iNo = 0 Then Exit Sub
                    Dim nod As XmlNode
                    nod = xnlCitations(iNo - 1).SelectSingleNode("CI_INFO/CI_JOURNAL")
                    If nod Is Nothing Then nod = xnlCitations(iNo - 1).SelectSingleNode("CI_INFO/CI_PATENT")
                    If nod Is Nothing Then Exit Sub
                    On Error Resume Next
                    Dim tstr As String
                    If rbAuthor.Checked Then
                        tstr = StrConv((nod.SelectSingleNode("CI_AUTHOR").InnerText), VbStrConv.Uppercase)
                        If tstr.Contains(".") Then
                            Dim atstr() As String = Split(tstr, ".")
                            tstr = Regex.Replace(atstr(0), "[^A-Z0-9* ]+", String.Empty) & "." & Regex.Replace(atstr(1), "[^A-Z0-9* ]+", String.Empty)
                        Else
                            tstr = Regex.Replace(tstr, "[^A-Z0-9* ]+", String.Empty)
                        End If
                        txtAuthor.Text = tstr

                        'txtAuthor.Text = Regex.Replace(StrConv((nod.SelectSingleNode("CI_AUTHOR").InnerText), VbStrConv.Uppercase), "[^A-Z0-9 ]+", String.Empty)
                        txtAuthor.Focus()
                        txtAuthor.SelectionStart = txtAuthor.Text.Length
                        txtAuthor.SelectionLength = 0
                    ElseIf rbVolume.Checked Then
                        tstr = Regex.Replace(StrConv(nod.SelectSingleNode("CI_VOLUME").InnerText, VbStrConv.Uppercase), "[^a-zA-Z0-9 ]+", String.Empty).Trim
                        If tstr.Length > 4 Then tstr = tstr.Substring(0, 4)
                        txtVolume.Text = tstr
                        If IsNumeric(txtVolume.Text) Then
                            txtVolume.Text = CInt(txtVolume.Text)
                        End If
                        txtVolume.Focus()
                        txtVolume.SelectionStart = txtVolume.Text.Length
                        txtVolume.SelectionLength = 0
                    ElseIf rbPage.Checked Then
                        tstr = Regex.Replace(StrConv(nod.SelectSingleNode("CI_PAGE").InnerText, VbStrConv.Uppercase), "[^A-Z0-9 ]+", String.Empty)
                        tstr = Split(tstr, "-")(0)
                        tstr = Regex.Replace(tstr, "[^a-zA-Z0-9 ]", String.Empty).Trim
                        If tstr.Length > 5 Then tstr = tstr.Substring(0, 5)
                        If IsNumeric(txtPage.Text) Then
                            txtPage.Text = CInt(txtPage.Text)
                        End If
                        txtPage.Text = tstr
                        txtPage.Focus()
                        txtPage.SelectionStart = txtPage.Text.Length
                        txtPage.SelectionLength = 0
                    ElseIf rbYear.Checked Then
                        tstr = StrConv(nod.SelectSingleNode("CI_YEAR").InnerText, VbStrConv.Uppercase)

                        tstr = Regex.Replace(tstr, "[^0-9]", String.Empty).Trim
                        If tstr.Length > 4 Then tstr = tstr.Substring(0, 4)
                        'If Regex.Match(tStr, "\d").Success Then
                        txtYear.Text = CInt(tstr)
                        txtYear.Focus()
                        txtYear.SelectionStart = txtYear.Text.Length
                        txtYear.SelectionLength = 0
                    ElseIf rbTitle.Checked Then
                        tstr = StrConv(xnlCitations(iNo - 1).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE").InnerText, VbStrConv.Uppercase)
                        If tstr.Length > 256 Then tstr = tstr.Substring(0, 256)
                        txtTitle.Text = tstr
                        txtTitle.Focus()
                        txtTitle.SelectionStart = txtTitle.Text.Length
                        txtTitle.SelectionLength = 0
                    End If
                    e.Handled = True
                    On Error GoTo 0
                End If
            Case Keys.L
                If e.Modifiers = (Keys.Control Or Keys.Alt) Then
                    lvlist.Visible = Not lvlist.Visible
                    pnlSource.Visible = Not lvlist.Visible
                    gbZoneTool.Visible = Not lvlist.Visible
                    e.Handled = True
                End If
        End Select
    End Sub


    Private Sub cmdPrev_Click(sender As Object, e As EventArgs) Handles cmdPrev.Click
        If fnAllVisited() Then
            subViewRef(iNo - 1, "Prev")
        ElseIf MsgBox("All fields are not visited.Do you want to leave anyway?", MsgBoxStyle.YesNo, "Confirm goto previous") = MsgBoxResult.Yes Then
            subViewRef(iNo - 1, "Prev")
        End If
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        If fnAllVisited() Then
            subViewRef(iNo + 1)
        ElseIf MsgBox("All fields are not visited.Do you want to leave anyway?", MsgBoxStyle.YesNo, "Confirm goto next") = MsgBoxResult.Yes Then
            subViewRef(iNo + 1)
        End If
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        If MsgBox("Do you want to exit?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
            Try
                'swLogFile.Close()
            Catch ex As Exception
            End Try

            If strInFile = "" Then
                Application.Exit()
            End If

            If MsgBox("Do you want to exit without complete current item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                If My.Computer.FileSystem.FileExists(strCurPath & System.IO.Path.GetFileName(strInFile)) Then
                    If bIsPrior Then
                        My.Computer.FileSystem.MoveFile(strCurPath & System.IO.Path.GetFileName(strInFile), strPriorPath & System.IO.Path.GetFileName(strInFile), True)
                    Else
                        My.Computer.FileSystem.MoveFile(strCurPath & System.IO.Path.GetFileName(strInFile), strInPath & System.IO.Path.GetFileName(strInFile), True)
                    End If
                End If
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub txtOCR_MouseUp(sender As Object, e As MouseEventArgs) Handles txtOCR.MouseUp
        Dim strST As String = StrConv(txtOCR.SelectedText, VbStrConv.Uppercase)
        If chkNTR.CheckState = CheckState.Checked Then Exit Sub
        If Trim(strST) = vbNullString Then Exit Sub
        Dim tStr As String
        tStr = Trim(strST)
        If rbAuthor.Checked Then
            If tStr.Length > 18 Then tStr = tStr.Substring(0, 18)
            txtAuthor.Text = tStr
            txtAuthor.Focus()
            txtAuthor.SelectionStart = txtAuthor.Text.Length
            txtAuthor.SelectionLength = 0
        ElseIf rbVolume.Checked Then
            tStr = Trim(Regex.Replace(tStr, "[^a-zA-Z0-9]", String.Empty))
            If tStr.Length > 4 Then tStr = tStr.Substring(0, 4)
            'If Regex.Match(tStr, "^[a-zA-Z0-9]+$").Success Then
            txtVolume.Text = tStr
            'End If
            txtVolume.Focus()
            txtVolume.SelectionStart = txtVolume.Text.Length
            txtVolume.SelectionLength = 0
        ElseIf rbPage.Checked Then
            tStr = Split(tStr, "-")(0)
            tStr = Trim(Regex.Replace(tStr, "[^a-zA-Z0-9 ]", String.Empty))
            If tStr.Length > 5 Then tStr = tStr.Substring(0, 5)
            'If Regex.Match(tStr, "^[a-zA-Z0-9]+$").Success Then
            txtPage.Text = tStr
            'End If
            txtPage.Focus()
            txtPage.SelectionStart = txtPage.Text.Length
            txtPage.SelectionLength = 0
        ElseIf rbYear.Checked Then
            tStr = Trim(Regex.Replace(tStr, "[^0-9]", String.Empty))
            If tStr.Length > 4 Then tStr = tStr.Substring(0, 4)
            'If Regex.Match(tStr, "\d").Success Then
            txtYear.Text = CInt(tStr)
            'End If
            txtYear.Focus()
            txtYear.SelectionStart = txtYear.Text.Length
            txtYear.SelectionLength = 0
        ElseIf rbTitle.Checked Then
            If tStr.Length > 256 Then tStr = tStr.Substring(0, 256)
            txtTitle.Text = tStr
            txtTitle.Focus()
            txtTitle.SelectionStart = txtTitle.Text.Length
            txtTitle.SelectionLength = 0
        ElseIf rbARTN1.Checked Then
            'tStr = Trim(Regex.Replace(tStr, "[^a-zA-Z0-9 ]", String.Empty))
            txtArtn1.Text = tStr
            txtArtn1.Focus()
            txtArtn1.SelectionStart = txtArtn1.Text.Length
            txtArtn1.SelectionLength = 0
        ElseIf rbARTN2.Checked Then
            'tStr = Trim(Regex.Replace(tStr, "[^a-zA-Z0-9 ]", String.Empty))
            txtArtn2.Text = tStr
            txtArtn2.Focus()
            txtArtn2.SelectionStart = txtArtn2.Text.Length
            txtArtn2.SelectionLength = 0
        ElseIf rbARTN3.Checked Then
            'tStr = Trim(Regex.Replace(tStr, "[^a-zA-Z0-9 ]", String.Empty))
            txtArtn3.Text = tStr
            txtArtn3.Focus()
            txtArtn3.SelectionStart = txtArtn3.Text.Length
            txtArtn3.SelectionLength = 0
        ElseIf rbARTN4.Checked Then
            'tStr = Trim(Regex.Replace(tStr, "[^a-zA-Z0-9 ]", String.Empty))
            txtArtn4.Text = tStr
            txtArtn4.Focus()
            txtArtn4.SelectionStart = txtArtn4.Text.Length
            txtArtn4.SelectionLength = 0
        End If
    End Sub

    Private Sub cmdHelp_Click(sender As Object, e As EventArgs) Handles cmdHelp.Click
        Help.ShowHelp(Me, strHelpPath & "ISIHELP.hlp.12-1-11.HLP")
    End Sub

    Private Sub txtAuthor_GotFocus(sender As Object, e As EventArgs) Handles txtAuthor.GotFocus
        rbAuthor.Checked = True
        cbField1.Checked = True
    End Sub

    Private Sub txtVolume_GotFocus(sender As Object, e As EventArgs) Handles txtVolume.GotFocus
        rbVolume.Checked = True
        cbField2.Checked = True
    End Sub

    Private Sub txtPage_GotFocus(sender As Object, e As EventArgs) Handles txtPage.GotFocus
        rbPage.Checked = True
        cbField3.Checked = True
    End Sub

    Private Sub txtYear_GotFocus(sender As Object, e As EventArgs) Handles txtYear.GotFocus
        rbYear.Checked = True
        cbField4.Checked = True
    End Sub

    Private Sub txtTitle_GotFocus(sender As Object, e As EventArgs) Handles txtTitle.GotFocus
        rbTitle.Checked = True
        cbField5.Checked = True
        txtTitle.SelectionStart = txtTitle.Text.Length
        txtTitle.SelectionLength = 0
    End Sub

    Private Sub chkNTR_CheckedChanged(sender As Object, e As EventArgs) Handles chkNTR.CheckedChanged
        If chkNTR.CheckState = CheckState.Checked Then
            'fnClearFields()
            'txtTitle.Text = "Non traditional reference"
            disableFields()
            'bNTR(iNo) = True
        ElseIf chkNTR.CheckState = CheckState.Unchecked Then
            enableFields()
            'bNTR(iNo) = False
        End If
    End Sub

    'Private Sub txtAuthor_KeyDown(sender As Object, e As KeyEventArgs) Handles txtAuthor.KeyDown
    'If e.KeyCode = Keys.Return Then
    '    txtVolume.Focus()
    '    e.SuppressKeyPress = True
    '    Exit Sub
    'ElseIf e.KeyCode = Keys.OemMinus Then
    '    txtAuthor.Clear()
    '    e.SuppressKeyPress = True
    '    Exit Sub
    'ElseIf e.KeyCode = Keys.F7 Then
    '    txtAuthor.Clear()
    '    e.Handled = True
    '    Exit Sub
    'End If
    'If (Char.IsControl(ChrW(e.KeyCode)) = False) Then
    '    If Char.IsPunctuation(ChrW(e.KeyCode)) Then
    '        If e.KeyCode = Asc("*") Or e.KeyCode = ("&") Then
    '            If txtAuthor.SelectionStart > 0 Then
    '                MsgBox(" */& allowed only at beginning", MsgBoxStyle.Information, "Verify")
    '                e.Handled = True
    '                txtAuthor.Focus()
    '                txtAuthor.SelectionStart = txtAuthor.Text.Length
    '                txtAuthor.SelectionLength = 0
    '                Exit Sub
    '            End If
    '        ElseIf e.KeyCode = "." Then
    '            If txtAuthor.Text.Length <> 15 Then
    '                MsgBox("Dot Allowed only after 15 chars!", MsgBoxStyle.Information, "Verify")
    '                e.Handled = True
    '                txtAuthor.Focus()
    '                txtAuthor.SelectionStart = txtAuthor.Text.Length
    '                txtAuthor.SelectionLength = 0
    '                Exit Sub
    '            End If
    '        Else
    '            MsgBox("Illegal character", MsgBoxStyle.Information, "Verify")
    '            e.Handled = True
    '            txtAuthor.Focus()
    '            txtAuthor.SelectionStart = txtAuthor.Text.Length
    '            txtAuthor.SelectionLength = 0
    '            Exit Sub
    '        End If
    '    ElseIf Char.IsSymbol(ChrW(e.KeyCode)) Then
    '        MsgBox("Illegal character", MsgBoxStyle.Information, "Verify")
    '        e.Handled = True
    '        txtAuthor.Focus()
    '        txtAuthor.SelectionStart = txtAuthor.Text.Length
    '        txtAuthor.SelectionLength = 0
    '        Exit Sub
    '    End If
    'End If
    'End Sub

    Private Sub txtAuthor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAuthor.KeyPress
        If Asc(e.KeyChar) = Keys.Space Then
            If txtAuthor.SelectionStart = 0 Then Exit Sub
            If txtAuthor.Text.Chars(txtAuthor.SelectionStart - 1) = " " Then
                txtAuthor.SelectionStart = txtAuthor.SelectionStart - 1
            End If
            If Not txtAuthor.Text.Trim.StartsWith("*") Then
                If txtAuthor.Text.Trim.Contains(" ") Then
                    If txtAuthor.Text.Chars(txtAuthor.SelectionStart - 1) = " " Then
                        txtAuthor.SelectionStart = txtAuthor.SelectionStart '- 1
                    End If
                    e.Handled = True
                End If
            End If
        End If
        If Asc(e.KeyChar) = Keys.Return Then
            txtVolume.Focus()
            e.Handled = True
            Exit Sub
        ElseIf e.KeyChar = "-" Then
            txtAuthor.Clear()
            e.Handled = True
            Exit Sub
        End If
        If (Char.IsControl(e.KeyChar) = False) Then
            If Char.IsPunctuation(e.KeyChar) Then
                If e.KeyChar = "*" Or e.KeyChar = "&" Then
                    If txtAuthor.SelectionStart > 0 Then
                        MsgBox(" */& allowed only at beginning", MsgBoxStyle.Information, "Verify")
                        e.Handled = True
                        txtAuthor.Focus()
                        txtAuthor.SelectionLength = 0
                        Exit Sub
                    Else
                        If txtAuthor.Text.StartsWith("*") Or txtAuthor.Text.StartsWith("&") Then
                            MsgBox(" */& allowed one time only at beginning", MsgBoxStyle.Information, "Verify")
                            e.Handled = True
                            txtAuthor.Focus()
                            txtAuthor.SelectionLength = 0
                            Exit Sub
                        End If
                    End If
                ElseIf e.KeyChar = "." Then
                    If txtAuthor.SelectionStart <> 15 Then
                        MsgBox("Dot Allowed only after 15 chars!", MsgBoxStyle.Information, "Verify")
                        e.Handled = True
                        'txtAuthor.Focus()
                        txtAuthor.SelectionStart = txtAuthor.Text.Length
                        txtAuthor.SelectionLength = 0
                        Exit Sub
                    End If
                Else
                    MsgBox("Illegal character", MsgBoxStyle.Information, "Verify")
                    e.Handled = True
                    'txtAuthor.Focus()
                    txtAuthor.SelectionStart = txtAuthor.Text.Length
                    txtAuthor.SelectionLength = 0
                    Exit Sub
                End If
            ElseIf Char.IsSymbol(e.KeyChar) Then
                MsgBox("Illegal character", MsgBoxStyle.Information, "Verify")
                e.Handled = True
                'txtAuthor.Focus()
                txtAuthor.SelectionStart = txtAuthor.Text.Length
                'txtAuthor.SelectionLength = 0
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtVolume_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtVolume.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txtPage.Focus()
            e.Handled = True
            Exit Sub
        ElseIf e.KeyChar = "-" Then
            txtVolume.Clear()
            e.Handled = True
            Exit Sub
        End If
        If (Char.IsControl(e.KeyChar) = False) Then
            If Not Char.IsLetterOrDigit(e.KeyChar) Then
                MsgBox("Only Letters & Digits Allowed!!", MsgBoxStyle.Information, "Verify")
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPage_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPage.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txtYear.Focus()
            e.Handled = True
            Exit Sub
        ElseIf e.KeyChar = "-" Then
            txtPage.Clear()
            e.Handled = True
            Exit Sub
        End If
        If (Char.IsControl(e.KeyChar) = False) Then
            If Not Char.IsLetterOrDigit(e.KeyChar) Then
                MsgBox("Only Letters & Digits Allowed!!", MsgBoxStyle.Information, "Verify")
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtYear_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtYear.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txtTitle.Focus()
            e.Handled = True
            Exit Sub
        ElseIf e.KeyChar = "-" Then
            txtYear.Clear()
            e.Handled = True
            Exit Sub
        End If
        If (Char.IsControl(e.KeyChar) = False) Then
            If Not Char.IsDigit(e.KeyChar) Then
                MsgBox("Only Digits Allowed!!", MsgBoxStyle.Information, "Verify")
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtTitle_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTitle.KeyPress
        If Asc(e.KeyChar) = Keys.Space Then
            If txtTitle.SelectionStart = 0 Then Exit Sub
            If txtTitle.Text.Chars(txtTitle.SelectionStart - 1) = " " Then
                txtTitle.SelectionStart = txtTitle.SelectionStart - 1
            End If
        End If
        If Asc(e.KeyChar) = Keys.Enter Then
            cmdNext.Focus()
            cmdNext.PerformClick()
            e.Handled = True
            Exit Sub
        ElseIf e.KeyChar = "-" Then
            txtTitle.Clear()
            e.Handled = True
            Exit Sub
        ElseIf Asc(e.KeyChar) = Keys.Tab Then
            txtArtn1.Focus()
            e.Handled = True
            Exit Sub
        End If
        If (Char.IsControl(e.KeyChar) = False) Then
            If Char.IsPunctuation(e.KeyChar) Then
                MsgBox("Punctuation not Allowed!!", MsgBoxStyle.Information, "Verify")
                e.Handled = True
            ElseIf (Char.GetUnicodeCategory(e.KeyChar) = Globalization.UnicodeCategory.MathSymbol) Then
                MsgBox("Symbol not Allowed!!", MsgBoxStyle.Information, "Verify")
                e.Handled = True
            End If
        End If

    End Sub

    Private Sub cmdAddArtn_Click(sender As Object, e As EventArgs) Handles cmdAddArtn.Click
        Dim i As Integer
        For i = 2 To 4
            If Me.Controls.Find("txtArtn" & i, True)(0).Visible = False Then
                Me.Controls.Find("txtArtn" & i, True)(0).Visible = True
                Me.Controls.Find("cmbArtn" & i, True)(0).Visible = True
                If i = 4 Then cmdAddArtn.Visible = False
                If cmdDelArtn.Visible = False Then cmdDelArtn.Visible = True
                subResizePB() '("Add")
                'pbImage.Height = pbImage.Height - 30
                'pbImage.Location = New Point(pbImage.Location.X, Me.Controls.Find("cmbArtn" & i, True)(0).Location.Y + 30)
                Exit For
            End If
        Next
    End Sub

    Private Sub cmdDelArtn_Click(sender As Object, e As EventArgs) Handles cmdDelArtn.Click
        Dim i As Integer
        For i = 4 To 2 Step -1
            If Me.Controls.Find("txtArtn" & i, True)(0).Visible = True Then
                Me.Controls.Find("txtArtn" & i, True)(0).Visible = False
                Me.Controls.Find("cmbArtn" & i, True)(0).Visible = False
                If cmdAddArtn.Visible = False Then cmdAddArtn.Visible = True
                If i = 2 Then cmdDelArtn.Visible = False
                subResizePB() '("Del")
                'pbImage.Height = pbImage.Height + 30
                'pbImage.Location = New Point(pbImage.Location.X, Me.Controls.Find("cmbArtn" & i, True)(0).Location.Y)
                Exit For
            End If
        Next
    End Sub

    Private Sub strFindString(strOCR As String)
        Dim bFound As Boolean = False, TextLine As String = ""

        If System.IO.File.Exists(policyFile) = True Then
            Dim objReader As New IO.StreamReader(policyFile)
            lblBlink.Text = ""
            lblBlink.BackColor = Color.WhiteSmoke
            Do While (objReader.Peek() <> -1 And bFound = False)
                TextLine = objReader.ReadLine()
                If InStr(strOCR, " " & Trim(TextLine) & " ", CompareMethod.Text) Then
                    lblBlink.Text = StrConv(Trim(TextLine), VbStrConv.Uppercase)
                    lblBlink.BackColor = Color.Red
                    Flash()
                    bFound = True
                    Exit Sub
                End If
            Loop
        End If
    End Sub

    Private Async Sub Flash()
        While True
            Await Task.Delay(1000)
            lblBlink.Visible = Not lblBlink.Visible
        End While
    End Sub

    Private Sub cmdDone_Click(sender As Object, e As EventArgs) Handles cmdDone.Click
        CheckDeleted()
        If xnlCitations Is Nothing Then Exit Sub
        If txtAuthor.Text.StartsWith("&") Then
            bIsPatent = True
        Else
            bIsPatent = False
        End If
        If chkNTR.Checked Then
            SaveNTR()
        ElseIf bIsPatent Then
            SavePatent()
        Else
            subSaveRef()
        End If

        If MsgBox("Do you want to done this item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
            If strInFile Is String.Empty Then Exit Sub
            xmlDoc.Save(strOutPath & System.IO.Path.GetFileName(strInFile))
            xmlDoc.Save(strBackUpPath & System.IO.Path.GetFileName(strInFile))

            On Error Resume Next
            My.Computer.FileSystem.DeleteFile(strCompPath & System.IO.Path.GetFileName(strInFile))
            My.Computer.FileSystem.DeleteFile(strTempPath & System.IO.Path.GetFileName(strInFile))
            My.Computer.FileSystem.MoveFile(strCurPath & System.IO.Path.GetFileName(strInFile), strCompPath & System.IO.Path.GetFileName(strInFile))

            For Each file As String In My.Computer.FileSystem.GetFiles(strInPath)
                If file.Contains(IO.Path.GetFileNameWithoutExtension(strInFile)) Then
                    If Not file.Contains("backup") Then
                        IO.File.Copy(file, strOutPath & System.IO.Path.GetFileName(file), True)
                        IO.File.Copy(file, strBackUpPath & System.IO.Path.GetFileName(file), True)
                        'My.Computer.FileSystem.MoveFile(file, strCompPath & System.IO.Path.GetFileName(file))
                        IO.File.Copy(file, strCompPath & System.IO.Path.GetFileName(file))
                        IO.File.Delete(file)
                    Else
                        IO.File.Copy(file, strTifBackUpPath & System.IO.Path.GetFileName(file), True)
                        IO.File.Delete(file)
                    End If
                End If
            Next

            iCompItems = iCompItems + 1
            iCompRefs = iCompRefs + xnlCitations.Count
            swCurProd.Stop()
            swCurSusp.Stop()

            'swLogFile.WriteLine(vbTab & xnlCitations.Count & vbTab & swCurProd.Elapsed.ToString("hh\:mm\:ss") & vbTab & vbTab & _
            'swCurSusp.Elapsed.ToString("hh\:mm\:ss"))
            'swLogFile.Close()

            Dim TotRefs As Integer = xnlCitations.Count
            Dim DelRefs As Integer = xnlCitations(0).ParentNode.SelectNodes("CI_CITATION[@D='Y']").Count
            Dim InsRefs As Integer = xnlCitations(0).ParentNode.SelectNodes("CI_CITATION[@I='Y']").Count
            LogEntry.Referances_Close = TotRefs + InsRefs - DelRefs
            LogEntry.Referances_Deleted = DelRefs
            LogEntry.Referances_Inserted = InsRefs
            LogEntry.Session_Production = swCurProd.Elapsed.ToString("hh\:mm\:ss")
            LogEntry.Session_Suspend = swCurSusp.Elapsed.ToString("hh\:mm\:ss")
            LogEntry.Session_End_Date = Format(Today, "dd-MM-yyyy")
            LogEntry.Session_End_Time = Now.ToShortTimeString
            LogEntry.Status = "CO"
            'MakeLogEntry(strLogPath & frmLogin.txtUserName.Text & ".xml")
            UpdateLogEntry(strLogPath & frmLogin.txtUserName.Text & ".xml", LogEntry)
            On Error GoTo 0
            lblCompItems.Text = iCompItems
            lblCompRefs.Text = iCompRefs
            fnClearFields()
            pbSource.Image = Nothing
            lblCurrent.Text = " of "
            xnlCitations = Nothing
            'System.IO.File.Move(strCurPath & System.IO.Path.GetFileName(strInFile), strCompPath & System.IO.Path.GetFileName(strInFile))
            If MsgBox("Do you want to load new item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                strInFile = fnGetXMLFile(strPriorPath)

                If strInFile = "" Then
                    bIsPrior = False
                    strInFile = fnGetXMLFile(strInPath)
                End If
                If strInFile = "" Then
                    MsgBox("No files for input", MsgBoxStyle.Information)
                Else
                    bIsPrior = False
                    fnLoadFile(strInFile)
                End If
            Else
                strInFile = ""
                lvlist.Items.Clear()
            End If
        End If
    End Sub

    Private Sub CheckDeleted()              ' This will remove FULL_CI_INFO tag from deleted nodes
        If xnlCitations Is Nothing Then Exit Sub
        For Each nod As XmlNode In xnlCitations
            If nod.Attributes("D") IsNot Nothing Then
                If nod.Attributes("D").Value = "Y" Then
                    Dim nodFCI As XmlNode
                    nodFCI = nod.SelectSingleNode("FULL_CI_INFO")
                    If nodFCI IsNot Nothing Then
                        nodFCI.ParentNode.RemoveChild(nodFCI)
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub disableFields()
        txtAuthor.Enabled = False
        txtVolume.Enabled = False
        txtPage.Enabled = False
        txtYear.Enabled = False
        txtTitle.Enabled = False
    End Sub

    Private Sub enableFields()
        txtAuthor.Enabled = True
        txtVolume.Enabled = True
        txtPage.Enabled = True
        txtYear.Enabled = True
        txtTitle.Enabled = True
    End Sub

    Private Sub txtTitle_Leave(sender As Object, e As EventArgs) Handles txtTitle.Leave
        On Error GoTo addtitle
        txtTitle.Text = bumpSpace(txtTitle.Text)
        If lvlist.Items.Count > iNo Then lvlist.Items(curRef).SubItems(5).Text = txtTitle.Text
        On Error GoTo 0
        Exit Sub
addtitle:
        For i As Integer = lvlist.Items(curRef).SubItems.Count To 6
            lvlist.Items(curRef).SubItems.Add("")
        Next
        lvlist.Items(curRef).SubItems(5).Text = txtTitle.Text
        Resume Next
    End Sub

    Private Sub txtTitle_TextChanged(sender As Object, e As EventArgs) Handles txtTitle.TextChanged
        Dim selst As Integer = txtTitle.SelectionStart
        txtTitle.Text = bumpSpace(Regex.Replace(txtTitle.Text, "[^a-zA-Z0-9* ]+", String.Empty))
        'txtTitle.Text = txtTitle.Text.Replace("-", String.Empty)
        txtTitle.SelectionStart = selst
        lblTitChars.Text = txtTitle.TextLength
    End Sub

    Private Sub txtAuthor_Leave(sender As Object, e As EventArgs) Handles txtAuthor.Leave
        'If lvlist.Items.Count > curRef Then lvlist.Items(curRef).SubItems(1).Text = txtAuthor.Text
        If Trim(txtAuthor.Text) = String.Empty Then Exit Sub
        If Not txtAuthor.Text.StartsWith("&") Then
            If Regex.Match(txtAuthor.Text, "\d").Success Then
                MsgBox("Digits not allowed here", MsgBoxStyle.Information, "Verify")
                txtAuthor.Focus()
                Exit Sub
            End If
        End If
        If txtAuthor.Text.StartsWith("*") Then
            If txtAuthor.Text.Contains(".") Then
                MsgBox("Dot not allowed here", MsgBoxStyle.Information, "Verify")
                txtAuthor.Focus()
                Exit Sub
            End If
        ElseIf txtAuthor.Text.StartsWith("&") Then

        Else
            If Not Regex.Match(txtAuthor.Text, "^[a-zA-Z .]+$").Success Then
                MsgBox("Only Alphabets allowed", MsgBoxStyle.Information, "Verify")
                txtAuthor.Focus()
                Exit Sub
            End If
            If txtAuthor.Text.Length < 17 Then
                If txtAuthor.Text.Contains(".") Then
                    MsgBox("Dot not allowed", MsgBoxStyle.Information, "Verify")
                    txtAuthor.Focus()
                    Exit Sub
                End If
            Else
                Dim tname As String()
                If txtAuthor.Text.Contains(".") Then
                    tname = Split(txtAuthor.Text, ".")
                    If tname(0).Length <> 15 Then
                        MsgBox("Name should be 15 chars", MsgBoxStyle.Information, "Verify")
                        txtAuthor.Focus()
                        Exit Sub
                    End If
                    If (tname.Length <> 1) And (tname.Length <> 2) Then
                        MsgBox("Atmost 2 initials allowed", MsgBoxStyle.Information, "Verify")
                        txtAuthor.Focus()
                        Exit Sub
                    End If
                Else
                    If txtAuthor.Text.Chars(15) <> " " Then
                        MsgBox("Name in 15<space>2 or 15<dot>2 format allowed", MsgBoxStyle.Information, "Verify")
                    End If
                End If
            End If
        End If
        'Dim sTemp As String = txtAuthor.Text.Trim
        'Dim bNoSpace As Boolean = False
        'While Not bNoSpace
        '    If sTemp.Contains("  ") Then
        '        sTemp = sTemp.Replace("  ", " ")
        '        bNoSpace = False
        '    Else
        '        bNoSpace = True
        '    End If
        'End While
        txtAuthor.Text = bumpSpace(txtAuthor.Text)
        On Error GoTo addtitle
        If lvlist.Items.Count > curRef Then lvlist.Items(curRef).SubItems(1).Text = txtAuthor.Text
        On Error GoTo 0
        Exit Sub
addtitle:
        For i As Integer = lvlist.Items(iNo).SubItems.Count To 2
            lvlist.Items(curRef).SubItems.Add("")
        Next
        lvlist.Items(curRef).SubItems(1).Text = txtAuthor.Text
        Resume Next
    End Sub

    Private Sub txtYear_Leave(sender As Object, e As EventArgs) Handles txtYear.Leave
        If Trim(txtYear.Text) = String.Empty Then Exit Sub

        If Not IsNumeric(txtYear.Text) Then
            MsgBox("Year must be numeric", MsgBoxStyle.Information, "Verify")
            txtYear.Focus()
            Exit Sub
        End If
        If CInt(txtYear.Text) <= 20 Then
            txtYear.Text = CInt(txtYear.Text) + 2000
        ElseIf txtYear.Text.Length <= 2 Then
            txtYear.Text = CInt(txtYear.Text) + 1900
        ElseIf txtYear.Text.Length <> 4 Then
            MsgBox("Year must be 4 digits or empty", MsgBoxStyle.Information, "Verify")
            txtYear.Focus()
        End If

        On Error GoTo addvol
        If lvlist.Items.Count > curRef Then lvlist.Items(curRef).SubItems(4).Text = txtYear.Text
        On Error GoTo 0
        Exit Sub
addvol:
        For i As Integer = lvlist.Items(curRef).SubItems.Count To 5
            lvlist.Items(curRef).SubItems.Add("")
        Next
        lvlist.Items(curRef).SubItems(4).Text = txtYear.Text
        Resume Next
    End Sub

    Private Sub txtPage_Leave(sender As Object, e As EventArgs) Handles txtPage.Leave
        If Trim(txtPage.Text) = String.Empty Then Exit Sub
        Dim spaces As String, pText As String, NewStr As String
        pText = Split(txtPage.Text, "-")(0)
        NewStr = pText
        'If pText.Length = 5 Then Exit Sub
        If pText.Length > 5 Then
            MsgBox("page must have atmost 5 chars", MsgBoxStyle.Information, "Verify")
            txtPage.Focus()
            Exit Sub
        End If
        If Not IsNumeric(NewStr) Then
            If Not Regex.Match(NewStr, "^[a-zA-Z0-9 ]+$").Success Then
                MsgBox("page must be alphanumeric", MsgBoxStyle.Information, "Verify")
                txtPage.Focus()
                Exit Sub
            End If
            spaces = Space(5 - pText.Length)
            For iCnt = 1 To pText.Length
                If Char.IsDigit((Mid(pText, iCnt, 1))) Then
                    NewStr = Mid(pText, 1, iCnt - 1) & spaces & Mid(pText, iCnt, pText.Length - iCnt + 1)
                    Exit For
                End If
            Next
            txtPage.Text = Trim(NewStr)
        Else
            txtPage.Text = CInt(NewStr)
        End If

        On Error GoTo addtitle
        If lvlist.Items.Count > curRef Then lvlist.Items(curRef).SubItems(3).Text = txtPage.Text
        On Error GoTo 0
        Exit Sub
addtitle:
        For i As Integer = lvlist.Items(curRef).SubItems.Count To 4
            lvlist.Items(curRef).SubItems.Add("")
        Next
        lvlist.Items(curRef).SubItems(3).Text = txtPage.Text
        Resume Next
    End Sub

    Private Sub txtAuthor_TextChanged(sender As Object, e As EventArgs) Handles txtAuthor.TextChanged
        Dim selst As Integer = txtAuthor.SelectionStart

        txtAuthor.Text = bumpSpace(Regex.Replace(txtAuthor.Text, "[^a-zA-Z0-9.&* ]+", String.Empty))
        txtAuthor.SelectionStart = selst
        lblAutherChars.Text = txtAuthor.Text.Length
        If txtAuthor.Text.Length >= 15 Then
            txtAuthor.BackColor = Color.Aqua
        Else
            txtAuthor.BackColor = Color.White
        End If
        If Not txtAuthor.Text.Trim.StartsWith("*") Then
            If txtAuthor.Text.Trim.Length - txtAuthor.Text.Trim.Replace(" ", String.Empty).Length > 1 Then
                txtAuthor.BackColor = Color.Red
            End If
        End If
    End Sub

    Private Sub txtPage_TextChanged(sender As Object, e As EventArgs) Handles txtPage.TextChanged
        'txtPage.Text = Split(txtPage.Text, "-")(0)
        txtPage.Text = Regex.Replace(txtPage.Text, "[^a-zA-Z0-9 ]+", String.Empty)
    End Sub
    Private Sub fnClearFields()
        txtAuthor.Clear()
        txtVolume.Clear()
        txtPage.Clear()
        txtYear.Clear()
        txtTitle.Clear()
        pbImage.ImageLocation = Nothing
        'pbSource.ImageLocation = Nothing
        txtOCR.Clear()
        Me.Controls.Find("txtArtn1", True)(0).Text = String.Empty
        Me.Controls.Find("cmbArtn1", True)(0).Text = String.Empty
        For i As Integer = 2 To 4
            Me.Controls.Find("txtArtn" & i, True)(0).Text = String.Empty
            Me.Controls.Find("cmbArtn" & i, True)(0).Text = String.Empty
            Me.Controls.Find("txtArtn" & i, True)(0).Visible = False
            Me.Controls.Find("cmbArtn" & i, True)(0).Visible = False
        Next
        chkNTR.Checked = False
        chkDCI.Checked = False
        For i As Integer = 1 To 5
            CType(Me.Controls.Find("cbField" & i, True)(0), CheckBox).Checked = False
        Next
        'lvlist.Items.Clear()
    End Sub

    Private Sub fnLoadFields(cits As XmlNodeList, iNum As Integer)
        Dim nod As XmlNode
        If cits Is Nothing Then
            MsgBox("Nothing to load")
            Exit Sub
        End If
        If iNum >= cits.Count Then Exit Sub
        nod = cits(iNum).Attributes("D")
        If nod IsNot Nothing Then
            If nod.Value = "Y" Then disableFields()
        Else
            enableFields()
        End If

        nod = cits(iNum).SelectSingleNode("CI_INFO/CI_JOURNAL")
        If nod IsNot Nothing Then
            checkNTR(nod)
            On Error Resume Next
            Dim tstr As String = StrConv((cits(iNum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_AUTHOR").InnerText), VbStrConv.Uppercase)
            If tstr.Contains(".") Then
                Dim atstr() As String = Split(tstr, ".")
                tstr = Regex.Replace(atstr(0), "[^A-Z0-9* ]+", String.Empty) & "." & Regex.Replace(atstr(1), "[^A-Z0-9* ]+", String.Empty)
            Else
                tstr = Regex.Replace(tstr, "[^A-Z0-9* ]+", String.Empty)
            End If
            txtAuthor.Text = tstr

            txtVolume.Text = Regex.Replace(StrConv(cits(iNum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_VOLUME").InnerText, VbStrConv.Uppercase), "[^A-Z0-9 ]+", String.Empty)
            If IsNumeric(txtVolume.Text) Then
                txtVolume.Text = CInt(txtVolume.Text)
            End If

            txtPage.Text = Regex.Replace(StrConv(cits(iNum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_PAGE").InnerText, VbStrConv.Uppercase), "[^A-Z0-9 ]+", String.Empty)
            If txtPage.Text.Contains("-") Then txtPage.Text = Split(txtPage.Text, "-")(0)
            If IsNumeric(txtPage.Text) Then
                txtPage.Text = CInt(txtPage.Text)
            End If

            txtYear.Text = StrConv(cits(iNum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_YEAR").InnerText, VbStrConv.Uppercase)
            If IsNumeric(txtYear.Text) Then
                txtYear.Text = CInt(txtYear.Text)
            End If
            txtTitle.Text = StrConv(cits(iNum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE").InnerText, VbStrConv.Uppercase)
            On Error GoTo 0
        Else
            nod = cits(iNum).SelectSingleNode("CI_INFO/CI_PATENT")
            If nod IsNot Nothing Then
                On Error Resume Next
                txtTitle.Text = "&" & Regex.Replace(StrConv((cits(iNum).SelectSingleNode("CI_INFO/CI_PATENT/PATENT_ASSIGNEE").InnerText), VbStrConv.Uppercase), "[^A-Z0-9 ]+", String.Empty)

                txtVolume.Text = Regex.Replace(StrConv(cits(iNum).SelectSingleNode("CI_INFO/CI_PATENT/PATENT_VOLUME").InnerText, VbStrConv.Uppercase), "[^A-Z0-9 ]+", String.Empty)
                If IsNumeric(txtVolume.Text) Then
                    txtVolume.Text = CInt(txtVolume.Text)
                End If

                txtPage.Text = Regex.Replace(StrConv(cits(iNum).SelectSingleNode("CI_INFO/CI_PATENT/PATENT_PAGE").InnerText, VbStrConv.Uppercase), "[^A-Z0-9 ]+", String.Empty)
                If txtPage.Text.Contains("-") Then txtPage.Text = Split(txtPage.Text, "-")(0)
                If IsNumeric(txtPage.Text) Then
                    txtPage.Text = CInt(txtPage.Text)
                End If

                txtYear.Text = StrConv(cits(iNum).SelectSingleNode("CI_INFO/CI_PATENT/PATENT_YEAR").InnerText, VbStrConv.Uppercase)
                If IsNumeric(txtYear.Text) Then
                    txtYear.Text = CInt(txtYear.Text)
                End If
                txtAuthor.Text = "&" & StrConv(cits(iNum).SelectSingleNode("CI_INFO/CI_PATENT").Attributes("ID").InnerText, VbStrConv.Uppercase)
                On Error GoTo 0
            End If
        End If

        Dim inClipName As String
        On Error Resume Next
        inClipName = xnlCitations(iNum).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText
        On Error GoTo 0
        If inClipName = String.Empty Then
            makeTree(xnlCitations(iNum), "CI_CAPTURE/CI_IMAGE_CLIP_NAME")
            inClipName = lblAccn.Text & CInt(xmlDoc.SelectSingleNode("//ITEM").Attributes("ITEMNO").Value).ToString("D3") & "C_" & xnlCitations(iNum).Attributes("seq").InnerText & ".TIF"
        Else
            inClipName = Replace(inClipName, Split(inClipName, "_").Last, xnlCitations(iNum).Attributes("seq").InnerText & ".TIF")
        End If

        xnlCitations(iNum).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText = inClipName
        pbImage.ImageLocation = strInPath & Split(inClipName, "\").Last
        If My.Computer.FileSystem.FileExists(strInPath & Split(cits(iNum).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText, "\").Last) Then
            pbSource.ImageLocation = strInPath & Split(cits(iNum).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText, "\").Last
        End If
        txtOCR.Text = cits(iNum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_BLURB").InnerText
        Dim strARTN As String
        For i As Integer = 1 To cits(iNum).SelectNodes("RI_CITATIONIDENTIFIER").Count
            Me.Controls.Find("txtArtn" & i, True)(0).Text = cits(iNum).SelectSingleNode("RI_CITATIONIDENTIFIER[@seq='" & i & "']").InnerText
            strARTN = cits(iNum).SelectSingleNode("RI_CITATIONIDENTIFIER[@seq='" & i & "']").Attributes("type").Value
            For Each itm In cmbArtn1.Items
                If itm.ToString = strARTN Then
                    Me.Controls.Find("cmbArtn" & i, True)(0).Text = strARTN
                    Exit For
                End If
            Next

            If Trim(Me.Controls.Find("txtArtn" & i, True)(0).Text) = String.Empty And (i <> 1) Then
                Me.Controls.Find("txtArtn" & i, True)(0).Visible = False
                Me.Controls.Find("cmbArtn" & i, True)(0).Visible = False
            Else
                Me.Controls.Find("txtArtn" & i, True)(0).Visible = True
                Me.Controls.Find("cmbArtn" & i, True)(0).Visible = True
            End If
        Next
        curRef = iNum
        subResizePB()
    End Sub

    Private Sub tmTimer_Tick(sender As Object, e As EventArgs) Handles tmTimer.Tick
        lblProdTime.Text = swProdTime.Elapsed.ToString("hh\:mm\:ss")
        lblSuspTime.Text = swSuspTime.Elapsed.ToString("hh\:mm\:ss")
    End Sub

    Private Sub cmdSuspend_Click(sender As Object, e As EventArgs) Handles cmdSuspend.Click
        Me.Enabled = False
        frmLogin.txtUserName.Text = lblUName.Text
        frmLogin.txtUserName.Enabled = False
        focusToBox()
        frmLogin.Show(Me)
    End Sub

    Private Sub ContentsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContentsToolStripMenuItem.Click
        Help.ShowHelp(Me, strHelpPath & "ISIHELP.hlp.12-1-11.HLP")
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("DE Master VIL V 1")
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub lvlist_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvlist.SelectedIndexChanged
        If lvlist.SelectedIndices.Count Then
            iNo = lvlist.SelectedIndices(0)
            subViewRef(iNo)
        End If
    End Sub

    Private Sub subViewRef(iNum As Integer, Optional sBut As String = "")

        Dim i As Integer


        If xnlCitations Is Nothing Then Exit Sub
        If txtAuthor.Text.StartsWith("&") Then
            bIsPatent = True
        Else
            bIsPatent = False
        End If

        Dim cCit As XmlNode = xnlCitations(iNum)
        While cCit IsNot Nothing
            Dim stest As String = ""
            Dim itest As String = ""
            Try
                stest = cCit.Attributes("D").Value
                'itest = cCit.Attributes("I").Value
            Catch ex As Exception
            End Try
            Try
                itest = cCit.Attributes("I").Value
            Catch ex As Exception
            End Try
            If stest <> "Y" Then Exit While
            If itest = "Y" Then
                iNum = CInt(lblCurrent.Text.Split("of").First.Trim + 1)
            End If
            If sBut = "Prev" Then
                iNum = iNum - 1
            Else
                iNum = iNum + 1
            End If
            cCit = xnlCitations(iNum)
        End While

        If iNum <= xnlCitations.Count Then
            If chkNTR.Checked Then
                SaveNTR()
            ElseIf bIsPatent Then
                SavePatent()
            Else
                subSaveRef(curRef)
            End If
        End If
        If iNum < 0 Then
            MsgBox("No more reference to view")
            Exit Sub
        ElseIf iNum >= xnlCitations.Count Then
            MsgBox("No more reference to view")
            fnClearFields()
            lblCurrent.Text = " " & xnlCitations.Count & " of " & xnlCitations.Count
            iNo = xnlCitations.Count
            Exit Sub
        End If

        iNo = iNum
        lblCurrent.Text = " " & iNum & " of " & xnlCitations.Count
        fnClearFields()
        For i = 1 To 5
            CType(Me.Controls.Find("cbField" & i, True)(0), CheckBox).Checked = False
        Next
        fnLoadFields(xnlCitations, iNum)
        subResizePB() '("Add")
        strFindString(txtOCR.Text)
        txtAuthor.Focus()
        txtAuthor.SelectionStart = 0
        txtAuthor.SelectionLength = 0
        'chkNTR.CheckState = CheckState.Unchecked
        'checkNTR(xnlCitations(iNum))        
    End Sub

    Public Sub subSaveRef(Optional inum As Integer = -1)
        Dim nod As XmlNode = Nothing, stest As String = ""
        If inum = -1 Then inum = CInt(Split(lblCurrent.Text, "of").First.Trim)

        If txtTitle.Text.Trim <> vbNullString Then
15:         xnlCitations(inum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE").InnerText = txtTitle.Text.Trim
        Else
            MsgBox("Title should not be empty", MsgBoxStyle.Information, "Verify")
            Exit Sub
        End If
        nod = Nothing
        Try
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_TITLE")
            nod.ParentNode.RemoveChild(nod)
        Catch ex As Exception
        End Try
        nod = Nothing
        Try
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_AUTHOR")
            nod.InnerText = txtAuthor.Text.Trim
        Catch ex As Exception
            makeTree(xnlCitations(inum), "CI_INFO/CI_JOURNAL/CI_AUTHOR", "")
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_AUTHOR")
            nod.InnerText = txtAuthor.Text.Trim
        End Try
        nod = Nothing
        Try
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_VOLUME")
            nod.InnerText = txtVolume.Text.Trim
        Catch ex As Exception
            makeTree(xnlCitations(inum), "CI_INFO/CI_JOURNAL/CI_VOLUME", "")
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_VOLUME")
            nod.InnerText = txtVolume.Text.Trim
        End Try
        nod = Nothing
        Try
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_PAGE")
            nod.InnerText = txtPage.Text.Trim
        Catch ex As Exception
            makeTree(xnlCitations(inum), "CI_INFO/CI_JOURNAL/CI_PAGE", "")
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_PAGE")
            nod.InnerText = txtPage.Text.Trim
        End Try
        nod = Nothing
        Try
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_YEAR")
            nod.InnerText = txtYear.Text.Trim
        Catch ex As Exception
            makeTree(xnlCitations(inum), "CI_INFO/CI_JOURNAL/CI_YEAR", "")
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_YEAR")
            nod.InnerText = txtYear.Text.Trim
        End Try

        Dim inClipName As String = String.Empty
        Try
            inClipName = xnlCitations(inum).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText
        Catch ex As Exception
        End Try
        If inClipName = String.Empty Then
            makeTree(xnlCitations(inum), "CI_CAPTURE/CI_IMAGE_CLIP_NAME")
            inClipName = lblAccn.Text & CInt(xmlDoc.SelectSingleNode("//ITEM").Attributes("ITEMNO").Value).ToString("D3") & "_" & xnlCitations(inum).Attributes("seq").InnerText & ".TIF"
        Else
            inClipName = Replace(inClipName, Split(inClipName, "_").Last, xnlCitations(inum).Attributes("seq").InnerText & ".TIF")
        End If
        If Not inClipName.StartsWith(".\") Then inClipName = ".\" & inClipName
        xnlCitations(inum).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText = inClipName
        For i = 1 To 4
            If Me.Controls.Find("txtArtn" & i, True)(0).Visible = True Then
                If (Me.Controls.Find("txtArtn" & i, True)(0).Text.Trim = vbNullString) Xor (Me.Controls.Find("cmbArtn" & i, True)(0).Text = vbNullString) Then
                    MsgBox("Check Artn fields", MsgBoxStyle.Information, "Verify")
                    Exit Sub
                ElseIf (Me.Controls.Find("txtArtn" & i, True)(0).Text.Trim <> vbNullString) And (Me.Controls.Find("cmbArtn" & i, True)(0).Text <> vbNullString) Then
                    Dim nd As XmlNode = xnlCitations(inum).SelectSingleNode("RI_CITATIONIDENTIFIER[@seq='" & i & "']")
                    If nd Is Nothing Then
                        makeTreeAttr(xnlCitations(inum), "RI_CITATIONIDENTIFIER", "seq", "type", Trim(Me.Controls.Find("txtArtn" & i, True)(0).Text), i, Me.Controls.Find("cmbArtn" & i, True)(0).Text)
                    Else
                        xnlCitations(inum).SelectSingleNode("RI_CITATIONIDENTIFIER[@seq='" & i & "']").Attributes("type").Value = Me.Controls.Find("cmbArtn" & i, True)(0).Text
                        xnlCitations(inum).SelectSingleNode("RI_CITATIONIDENTIFIER[@seq='" & i & "']").InnerText = Me.Controls.Find("txtArtn" & i, True)(0).Text.Trim
                    End If
                Else
                    Dim nd As XmlNode = xnlCitations(inum).SelectSingleNode("RI_CITATIONIDENTIFIER[@seq='" & i & "']")
                    Try
                        nd.ParentNode.RemoveChild(nd)
                    Catch ex As Exception
                    End Try
                End If

            Else
                Try
                    Dim tnod As XmlNode
                    tnod = xnlCitations(inum).SelectSingleNode("RI_CITATIONIDENTIFIER[@seq='" & i & "']")
                    tnod.ParentNode.RemoveChild(tnod)
                Catch ex As Exception
                End Try
            End If
        Next
        nod = Nothing
        Try
            nod = xnlCitations(inum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_BLURB")
            Dim CData As XmlCDataSection
            CData = xmlDoc.CreateCDataSection(txtOCR.Text.Trim)
            nod.InnerText = ""
            nod.AppendChild(CData)
            'nod.InnerText = txtOCR.Text
            'cits(inum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_BLURB").InnerText()
        Catch ex As Exception
            makeTree(xnlCitations(inum), "CI_CAPTURE/CI_CAPTURE_BLURB", "")
            nod = xnlCitations(inum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_BLURB")
            'nod.InnerText = txtOCR.Text
            Dim CData As XmlCDataSection
            CData = xmlDoc.CreateCDataSection(txtOCR.Text.Trim)
            nod.InnerText = ""
            nod.AppendChild(CData)
        End Try
        nod = Nothing
        Try
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_PATENT")
            nod.ParentNode.RemoveChild(nod)
        Catch ex As Exception
        End Try
        Dim attr As XmlAttribute = xnlCitations(inum).Attributes("D")
        If attr IsNot Nothing Then
            Try
                xnlCitations(inum).RemoveChild(xnlCitations(inum).SelectSingleNode("FULL_CI_INFO"))
            Catch ex As Exception
            End Try
        End If
        'xmlDoc.Save(Split(xmlDoc.BaseURI, "///").Last)
        xmlDoc.Save(strTempPath & System.IO.Path.GetFileName(strInFile))
    End Sub

    Public Function fnGetXMLFile(strDir As String) As String
        Try
            Return System.IO.Directory.GetFiles(strDir, "*.xml", IO.SearchOption.TopDirectoryOnly)(0)
        Catch e As System.IndexOutOfRangeException
            Return String.Empty
        End Try
    End Function

    Private Sub fnLoadFile(strFile As String)
        iNo = 0
        On Error Resume Next
        System.IO.File.Delete(strCurPath & System.IO.Path.GetFileName(strFile))

        My.Computer.FileSystem.MoveFile(strFile, strCurPath & System.IO.Path.GetFileName(strFile))
        strFile = strCurPath & System.IO.Path.GetFileName(strFile)
        xmlDoc.Load(strFile)
        cbBackup.Items.Clear()
        frmImages.cbBackup.Items.Clear()
        swCurProd.Restart()
        swCurSusp.Reset()
        tsslStatus.Text = "Loaded file : " & strFile

        lblAccn.Text = xmlDoc.SelectSingleNode("//ID_ACCESSION").InnerText
        lblItem.Text = xmlDoc.SelectSingleNode("//ITEM").Attributes.GetNamedItem("ITEMNO").InnerText
        xnlCitations = xmlDoc.DocumentElement.SelectNodes("//CI_CITATION")
        lblTotalSeq.Text = xnlCitations.Count
        lblCurrent.Text = " 0 of " & xnlCitations.Count
        lvlist.Items.Clear()

        'swLogFile = New IO.StreamWriter(strLogPath & frmLogin.txtUserName.Text & ".txt", True)
        'swLogFile.Write(Format(Today, "yyyyMMdd") & vbTab & Now.ToShortTimeString & vbTab & vbTab)
        'swLogFile.Write(lblAccn.Text & vbTab & vbTab & lblItem.Text & vbTab & vbTab & lblTotalSeq.Text & vbTab & vbTab & vbTab & vbTab)
        InitLog(LogEntry)
        LogEntry.Session_Start_Date = Format(Today, "dd-MM-yyyy")
        LogEntry.Session_Start_Time = Now.ToShortTimeString
        LogEntry.Item_Accession = lblAccn.Text
        LogEntry.Item_Itemno = lblItem.Text
        LogEntry.Referances_Open = lblTotalSeq.Text
        LogEntry.Status = "IP"
        MakeLogEntry(strLogPath & frmLogin.txtUserName.Text & ".xml")

        Dim nodej As XmlNode, nodep As XmlNode
        ReDim bNTR(0 To xnlCitations.Count - 1)
        'Try
        For i = 0 To xnlCitations.Count - 1
            bNTR(i) = False

            nodej = xnlCitations(i).SelectSingleNode("CI_INFO/CI_JOURNAL")
            nodep = xnlCitations(i).SelectSingleNode("CI_INFO/CI_PATENT")

            If (nodej Is Nothing) And (nodep Is Nothing) Then
                makeTree(xnlCitations(i), "CI_INFO")
                makeTree(xnlCitations(i), "CI_JOURNAL")
                makeTree(xnlCitations(i), "CI_INFO/CI_JOURNAL/CI_AUTHOR")
                makeTree(xnlCitations(i), "CI_INFO/CI_JOURNAL/CI_VOLUME")
                makeTree(xnlCitations(i), "CI_INFO/CI_JOURNAL/CI_PAGE")
                makeTree(xnlCitations(i), "CI_INFO/CI_JOURNAL/CI_YEAR")
                makeTree(xnlCitations(i), "CI_CAPTURE/CI_CAPTURE_TITLE")

                Dim auth As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_AUTHOR").InnerText
                Dim vol As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_VOLUME").InnerText
                Dim page As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_PAGE").InnerText
                Dim year As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_YEAR").InnerText
                Dim title As String = xnlCitations(i).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE").InnerText

                With lvlist.Items.Add(i + 1)
                    .SubItems.Add(auth)
                    .SubItems.Add(vol)
                    .SubItems.Add(page)
                    .SubItems.Add(year)
                    .SubItems.Add(title)
                End With
            Else
                If (nodep Is Nothing) Then
                    makeTree(xnlCitations(i), "CI_INFO")
                    makeTree(xnlCitations(i), "CI_JOURNAL")
                    makeTree(xnlCitations(i), "CI_INFO/CI_JOURNAL/CI_AUTHOR")
                    makeTree(xnlCitations(i), "CI_INFO/CI_JOURNAL/CI_VOLUME")
                    makeTree(xnlCitations(i), "CI_INFO/CI_JOURNAL/CI_PAGE")
                    makeTree(xnlCitations(i), "CI_INFO/CI_JOURNAL/CI_YEAR")
                    makeTree(xnlCitations(i), "CI_CAPTURE/CI_CAPTURE_TITLE")

                    Dim auth As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_AUTHOR").InnerText
                    Dim vol As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_VOLUME").InnerText
                    Dim page As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_PAGE").InnerText
                    Dim year As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_YEAR").InnerText
                    Dim title As String = xnlCitations(i).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE").InnerText

                    With lvlist.Items.Add(i + 1)
                        .SubItems.Add(auth)
                        .SubItems.Add(xnlCitations(i).SelectSingleNode(vol).InnerText)
                        .SubItems.Add(xnlCitations(i).SelectSingleNode(page).InnerText)
                        .SubItems.Add(xnlCitations(i).SelectSingleNode(year).InnerText)
                        .SubItems.Add(xnlCitations(i).SelectSingleNode(title).InnerText)
                    End With
                Else
                    Dim auth As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_PATENT/PATENT_ASSIGNEE").InnerText
                    Dim vol As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_PATENT/PATENT_VOLUME").InnerText
                    Dim page As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_PATENT/PATENT_PAGE").InnerText
                    Dim year As String = xnlCitations(i).SelectSingleNode("CI_INFO/CI_PATENT/PATENT_YEAR").InnerText
                    Dim title As String = "&" & xnlCitations(i).SelectSingleNode("CI_INFO/CI_PATENT").Attributes("ID").Value

                    With lvlist.Items.Add(i + 1)
                        .SubItems.Add(auth)
                        .SubItems.Add(vol)
                        .SubItems.Add(page)
                        .SubItems.Add(year)
                        .SubItems.Add(title)
                    End With
                End If
            End If
        Next
        'Catch ex As Exception
        'End Try
        On Error GoTo 0
        cmdDelArtn.Visible = False
        fnLoadFields(xnlCitations, iNo)
        Dim tempInPath As String

        If bIsPrior Then
            tempInPath = strPriorPath
        Else
            tempInPath = strInPath
        End If

        For Each tif As String In IO.Directory.GetFiles(tempInPath,
                                    System.IO.Path.GetFileNameWithoutExtension(strFile) & "_*.tif", IO.SearchOption.TopDirectoryOnly)
            If Not tif.Contains("backup") Then
                If Not IO.File.Exists(Replace(tif, ".tif", "-backup.tif", Compare:=CompareMethod.Text)) Then
                    IO.File.Copy(tif, Replace(tif, ".tif", "-backup.tif", Compare:=CompareMethod.Text))
                End If
                cbBackup.Items.Add(IO.Path.GetFileNameWithoutExtension(tif))
            End If
        Next
        lblSTime.Text = Now.ToShortTimeString
        strFindString(txtOCR.Text)
        txtAuthor.Focus()
        txtAuthor.SelectionStart = 0
        txtAuthor.SelectionLength = 0

    End Sub

    Private Sub cmdInter_Click(sender As Object, e As EventArgs) Handles cmdInter.Click
        If MsgBox("Do you want to interrupt this item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
            xmlDoc.Save(strPriorPath & System.IO.Path.GetFileName(strInFile))
            fnClearFields()
            Try
                My.Computer.FileSystem.DeleteFile(strCurPath & System.IO.Path.GetFileName(strInFile))
                My.Computer.FileSystem.DeleteFile(strTempPath & System.IO.Path.GetFileName(strInFile))
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub txtVolume_Leave(sender As Object, e As EventArgs) Handles txtVolume.Leave
        If IsNumeric(txtVolume.Text) Then
            txtVolume.Text = CInt(txtVolume.Text)
        Else
            If Regex.IsMatch(txtVolume.Text, "\d") Then
                MsgBox("Volume must either be Alpha or numeric", MsgBoxStyle.Information, "Verify")
                txtVolume.Focus()
            End If
        End If
        txtVolume.Text = bumpSpace(txtVolume.Text)
        On Error GoTo addvol
        If lvlist.Items.Count > curRef Then lvlist.Items(curRef).SubItems(2).Text = txtVolume.Text
        On Error GoTo 0
        Exit Sub
addvol:
        For i As Integer = lvlist.Items(curRef).SubItems.Count To 3
            lvlist.Items(curRef).SubItems.Add("")
        Next
        lvlist.Items(curRef).SubItems(2).Text = txtVolume.Text
        Resume Next
    End Sub

    Private Sub subResizePB()
        For i As Integer = 4 To 1 Step -1
            If Me.Controls.Find("txtArtn" & i, True)(0).Visible = True Then
                pnlDesImage.Location = New Point(pnlDesImage.Location.X, Me.Controls.Find("txtArtn" & i, True)(0).Location.Y + 560)
                pnlDesImage.Height = txtOCR.Top - pnlDesImage.Top - 10
                Exit For
            End If
        Next
    End Sub

    Private Function fnAllVisited() As Boolean
        For i As Integer = 1 To 5
            If Not CType(Me.Controls.Find("cbField" & i, True)(0), CheckBox).Checked Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Sub cmdDelRef_Click(sender As Object, e As EventArgs) Handles cmdDelRef.Click
        If MsgBox("Do you want to Delete this reference?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
            DelRef(iNo)
        End If
    End Sub

    Private Sub cmdAddRef_Click(sender As Object, e As EventArgs) Handles cmdAddRef.Click
        subSaveRef()
        If MsgBox("Do you want to Add reference?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
            AddRef()
            'iNo = xnlCitations.Count - 1
            Dim lvitem As New ListViewItem(xnlCitations.Count)
            With lvitem
                .SubItems.Add(xnlCitations(xnlCitations.Count - 1).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_AUTHOR").InnerText)
                .SubItems.Add(xnlCitations(xnlCitations.Count - 1).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_VOLUME").InnerText)
                .SubItems.Add(xnlCitations(xnlCitations.Count - 1).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_PAGE").InnerText)
                .SubItems.Add(xnlCitations(xnlCitations.Count - 1).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_YEAR").InnerText)
                .SubItems.Add("")
                '.SubItems.Add(xnlCitations(iNo).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE").InnerText)
            End With
            'lvlist.Items.Insert(iNo, lvitem)

            lvlist.Items.Add(lvitem)
            lvlist.Items(xnlCitations.Count - 1).ForeColor = Color.White
            'lblCurrent.Text = iNo & " of " & xnlCitations.Count
            lblCurrent.Text = CInt(lblCurrent.Text.Split("of").First.Trim) & " of " & xnlCitations.Count
            txtTitle.Clear()
            txtAuthor.Focus()
            txtAuthor.SelectionStart = txtAuthor.Text.Length
            txtAuthor.SelectionLength = 0
            subSaveRef(xnlCitations.Count - 1)
            curRef = xnlCitations.Count - 1
            'subViewRef(iNo)
        End If
    End Sub

    Private Sub DelRef(iNum As Integer)
        Dim xaAttr As XmlAttribute, attr As XmlAttribute
        If xnlCitations(iNum) Is Nothing Then Exit Sub
        attr = xnlCitations(iNum).Attributes("I")
        If attr IsNot Nothing Then
            xnlCitations(iNum).Attributes.Remove(attr)
        End If
        xaAttr = xmlDoc.CreateAttribute("D")
        xaAttr.Value = "Y"
        xnlCitations(iNum).Attributes.Append(xaAttr)
        disableFields()
        lvlist.Items(iNum).ForeColor = Color.Silver
        lblCurrent.Text = CInt(lblCurrent.Text.Split("of").First.Trim) & " of " & xnlCitations.Count
        MsgBox("Successfully Deleted")
        'subSaveRef()
        subViewRef(iNum + 1)
    End Sub

    Private Sub AddRef()
        Dim cr As String = Environment.NewLine
        Dim newCitation As String
        curRef = xnlCitations.Count - 1
        newCitation = cr &
                "   <CI_CITATION seq='" & xnlCitations(curRef).Attributes("seq").Value + 1 & "' I='Y'>" & cr &
                "       <CI_JOURNAL/>" & cr &
                "       <CI_INFO>" & cr &
                "           <CI_JOURNAL>" & cr &
                "               <CI_AUTHOR></CI_AUTHOR>" & cr &
                "               <CI_VOLUME></CI_VOLUME>" & cr &
                "               <CI_PAGE></CI_PAGE>" & cr &
                "               <CI_YEAR></CI_YEAR>" & cr &
                "               <CI_TITLE></CI_TITLE>" & cr &
                "           </CI_JOURNAL>" & cr &
                "       </CI_INFO>" & cr &
                "       <CI_CAPTURE>" & cr &
                "           <CI_CAPTURE_BLURB></CI_CAPTURE_BLURB>" & cr &
                "           <CI_CAPTURE_TITLE></CI_CAPTURE_TITLE>" & cr &
                "           <CI_CAPTURE_TITLE_CONF_IND></CI_CAPTURE_TITLE_CONF_IND>" & cr &
                "           <CI_IMAGE_CLIP_NAME>.\" & IO.Path.GetFileNameWithoutExtension(strInFile) & "_" & xnlCitations(curRef).Attributes("seq").Value + 1 & "*.tif</CI_IMAGE_CLIP_NAME>" & cr &
                "       </CI_CAPTURE>" & cr &
                "   </CI_CITATION>"
        Try
            IO.File.Copy(strInFile.Replace(IO.Path.GetExtension(strInFile), "_" & xnlCitations(iNo).Attributes("seq").Value & ".tif"), strInFile.Replace(IO.Path.GetExtension(strInFile), "_" & xnlCitations(curRef).Attributes("seq").Value + 1 & ".tif")) '& "_" & xnlCitations(iNo).Attributes("seq").Value + 1 & ".tif</CI_IMAGE_CLIP_NAME>" & cr &
        Catch ex As Exception
        End Try

        ' & xnlCitations(iNo).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_BLURB").InnerXml & 
        Dim docFrag As XmlDocumentFragment = xmlDoc.CreateDocumentFragment()
        docFrag.InnerXml = newCitation
        'xnlCitations(iNo).ParentNode.InsertAfter(docFrag, xnlCitations(iNo))
        xnlCitations(iNo).ParentNode.InsertAfter(docFrag, xnlCitations(iNo).ParentNode.LastChild)
        xnlCitations = xmlDoc.DocumentElement.SelectNodes("/ISSUE/ITEM/ITEM_CONTENT/CI_CITATION")
    End Sub

    Private Sub AddRefToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddRefToolStripMenuItem.Click
        cmdAddRef.PerformClick()
    End Sub

    Private Sub DeleteRefToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteRefToolStripMenuItem.Click
        cmdDelRef.PerformClick()
    End Sub

    Private Sub txtArtn2_GotFocus(sender As Object, e As EventArgs) Handles txtArtn2.GotFocus
        rbARTN2.Checked = True
    End Sub

    Private Sub txtArtn2_VisibleChanged(sender As Object, e As EventArgs) Handles txtArtn2.VisibleChanged
        If txtArtn2.Visible = False Then
            cmdDelArtn.Visible = False
        Else
            cmdDelArtn.Visible = True
        End If
    End Sub

    Private Sub SaveNTR()
        Dim ref As Integer
        ref = CInt(Split(lblCurrent.Text, "of").First.Trim)
        makeTree(xnlCitations(ref), "CI_INFO/CI_JOURNAL/CI_AUTHOR", "")
        makeTree(xnlCitations(ref), "CI_INFO/CI_JOURNAL/CI_VOLUME", "")
        makeTree(xnlCitations(ref), "CI_INFO/CI_JOURNAL/CI_PAGE", "")
        makeTree(xnlCitations(ref), "CI_INFO/CI_JOURNAL/CI_YEAR", "")
        makeTree(xnlCitations(ref), "CI_INFO/CI_JOURNAL/CI_TITLE", "")
        Dim cdTitle As XmlCDataSection
        cdTitle = xmlDoc.CreateCDataSection("**NON-TRADITIONAL**")
        xnlCitations(ref).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_TITLE").AppendChild(cdTitle)
        makeTree(xnlCitations(ref), "CI_CAPTURE/CI_CAPTURE_TITLE", "NON TRADITIONAL REF")
    End Sub

    Public Shared Sub makeTree(xnNode As XmlNode, strTree As String, Optional value As String = "!@!@!")
        Dim nod As XmlNode, nodnew As XmlNode, xeElement As XmlElement
        nod = xnNode
        For Each str As String In Split(strTree, "/")
            If Trim(str) <> String.Empty Then
                nodnew = nod.SelectSingleNode(str)
                If nodnew Is Nothing Then
                    xeElement = xnNode.OwnerDocument.CreateElement(str)
                    nod = nod.InsertBefore(xeElement, nod.FirstChild)    'todo: correct possition
                Else
                    nod = nodnew
                End If
            End If
        Next
        If value <> "!@!@!" Then
            nod.InnerText = value
        End If
    End Sub

    Private Shared Sub makeTreeAttr(xnNode As XmlNode, strTree As String, strAttr1 As String, strAttr2 As String, Optional nodVal As String = "!@!@!", Optional attrVal1 As String = "!@!@!", Optional attrVal2 As String = "!@!@!")
        Dim nod As XmlNode, nodnew As XmlNode, xeElement As XmlElement, xaAttr As XmlAttribute
        nod = xnNode
        For Each str As String In Split(strTree, "/")
            If Trim(str) <> String.Empty Then
                nodnew = nod.SelectSingleNode(str & "[" & strAttr1 & "='" & attrVal1 & "']")
                If nodnew Is Nothing Then
                    xeElement = xnNode.OwnerDocument.CreateElement(str)
                    xaAttr = xnNode.OwnerDocument.CreateAttribute(strAttr1)
                    If attrVal1 <> "!@!@!" Then
                        xaAttr.Value = attrVal1
                    End If
                    xeElement.Attributes.Append(xaAttr)
                    xaAttr = xnNode.OwnerDocument.CreateAttribute(strAttr2)
                    If attrVal2 <> "!@!@!" Then
                        xaAttr.Value = attrVal2
                    End If
                    xeElement.Attributes.Append(xaAttr)
                    nod = nod.InsertBefore(xeElement, nod.SelectNodes("CI_CITATION")(0))    'todo: correct possition
                Else
                    nod = nodnew
                End If
            End If
        Next
        If nodVal <> "!@!@!" Then
            nod.InnerText = nodVal
        End If
    End Sub

    Private Sub focusToBox()
        Select Case True
            Case rbAuthor.Checked
                frmLogin.lblLField.Text = "txtAuthor"
            Case rbVolume.Checked
                txtVolume.Focus()
                frmLogin.lblLField.Text = "txtVolume"
            Case rbPage.Checked
                txtPage.Focus()
                frmLogin.lblLField.Text = "txtPage"
            Case rbYear.Checked
                txtYear.Focus()
                frmLogin.lblLField.Text = "txtYear"
            Case rbTitle.Checked
                txtTitle.Focus()
                frmLogin.lblLField.Text = "txtTitle"
        End Select
    End Sub

    Private Sub cmdGoto_Click(sender As Object, e As EventArgs) Handles cmdGoto.Click
        If Not IsNumeric(txtGoto.Text) Then
            MsgBox("Input number only", MsgBoxStyle.Information, "Verify")
            Exit Sub
        Else
            If CInt(txtGoto.Text) < 1 Or CInt(txtGoto.Text) > xnlCitations.Count Then
                MsgBox("Out of range", MsgBoxStyle.Information, "Verify")
                Exit Sub
            Else
                subViewRef(CInt(txtGoto.Text) - 1)
            End If
        End If
    End Sub

    Private Sub txtGoto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtGoto.KeyPress
        If Asc(e.KeyChar) = Keys.Return Then
            cmdGoto.PerformClick()
            e.Handled = True
        End If
    End Sub

    Private Sub cmdZoning_Click(sender As Object, e As EventArgs) Handles cmdZoning.Click
        pbSource.ImageLocation = strInPath & Split(xnlCitations(iNo).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText, "\")(1)
        subToggleZoning()
    End Sub

    Private Sub pbSource_MouseDown(sender As Object, e As MouseEventArgs) Handles pbSource.MouseDown
        If pbSource.ImageLocation = String.Empty Then Exit Sub
        If chkFullWidth.Checked Then
            mRect = New Rectangle(0, e.Y, pbSource.Width, 0)
        Else
            mRect = New Rectangle(e.X, e.Y, 0, 0)
        End If
        pbSource.Invalidate()
    End Sub

    Private Sub pbSource_MouseMove(sender As Object, e As MouseEventArgs) Handles pbSource.MouseMove
        If pbSource.ImageLocation = String.Empty Then Exit Sub
        If e.Button = Windows.Forms.MouseButtons.Left Then
            If chkFullWidth.Checked Then
                mRect = New Rectangle(0, mRect.Top, pbSource.Width, e.Y - mRect.Top)
            Else
                mRect = New Rectangle(mRect.Left, mRect.Top, e.X - mRect.Left, e.Y - mRect.Top)
            End If
            pbSource.Invalidate()
        End If
    End Sub

    Private Sub pbSource_MouseUp(sender As Object, e As MouseEventArgs) Handles pbSource.MouseUp
        If pbSource.ImageLocation = String.Empty Then Exit Sub
        If mRect.IsEmpty Then Exit Sub
        CropImage()
    End Sub

    Private Sub pbSource_Paint(sender As Object, e As PaintEventArgs) Handles pbSource.Paint
        If mRect.IsEmpty Then Exit Sub
        Using pen As New Pen(Color.Red, 1)
            e.Graphics.DrawRectangle(pen, mRect)
        End Using
    End Sub

    'Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
    '    If keyData = Keys.Up Then
    '        If Me.pbSource.Bounds.Contains(Me.PointToClient(Cursor.Position)) Then
    '            mRect.Location = New Point(mRect.Location.X, mRect.Location.Y - 1)
    '            pbSource.Invalidate()
    '            CropImage()
    '            Return True
    '        End If
    '    End If
    '    If keyData = Keys.Down Then
    '        If Me.pbSource.Bounds.Contains(Me.PointToClient(Cursor.Position)) Then
    '            mRect.Location = New Point(mRect.Location.X, mRect.Location.Y + 1)
    '            pbSource.Invalidate()
    '            CropImage()
    '            Return True
    '        End If
    '    End If
    '    If keyData = Keys.Left Then
    '        If Me.pbSource.Bounds.Contains(Me.PointToClient(Cursor.Position)) Then
    '            mRect.Location = New Point(mRect.Location.X - 1, mRect.Location.Y)
    '            pbSource.Invalidate()
    '            CropImage()
    '            Return True
    '        End If
    '    End If
    '    If keyData = Keys.Right Then
    '        If Me.pbSource.Bounds.Contains(Me.PointToClient(Cursor.Position)) Then
    '            mRect.Location = New Point(mRect.Location.X + 1, mRect.Location.Y)
    '            pbSource.Invalidate()
    '            CropImage()
    '            Return True
    '        End If
    '    End If
    '    Return MyBase.ProcessCmdKey(msg, keyData)
    'End Function

    Private Sub CropImage()
        Dim fileName = pbSource.ImageLocation
        If fileName Is Nothing Then Exit Sub
        If mRect.IsEmpty Then Exit Sub
        If mRect.Width * mRect.Height = 0 Then Exit Sub
        Dim CropRect As New Rectangle(mRect.Left, mRect.Top, mRect.Width, mRect.Height)
        Dim OriginalImage = Image.FromFile(fileName)
        Dim CropImage = New Bitmap(CropRect.Width, CropRect.Height)
        Using grp = Graphics.FromImage(CropImage)
            grp.DrawImage(OriginalImage, New Rectangle(0, 0, CropRect.Width, CropRect.Height), CropRect, GraphicsUnit.Pixel)
            OriginalImage.Dispose()
            'CropImage.Save(IO.Path.GetDirectoryName(fileName) & "\temp.tif")
            If chkMergeImg.Checked = True Then
                pbImage.Image = CombineImages(pbImage.Image, CropImage)
            Else
                pbImage.Image = CropImage
            End If
            pbImage.Image.Save(IO.Path.GetDirectoryName(fileName) & "\temp.tif")
            If Not fileName.Contains("-backup.tif") Then
                If Not IO.File.Exists(Replace(fileName, ".tif", "-backup.tif", Compare:=CompareMethod.Text)) Then
                    My.Computer.FileSystem.RenameFile(fileName, Replace(IO.Path.GetFileName(fileName), ".tif", "-backup.tif", Compare:=CompareMethod.Text))
                End If
                pbSource.ImageLocation = IO.Path.GetDirectoryName(fileName) & "\" & Replace(IO.Path.GetFileName(fileName), ".tif", "-backup.tif", Compare:=CompareMethod.Text)
            End If

            If IO.File.Exists(IO.Path.GetDirectoryName(fileName) & "\" & Split(xnlCitations(iNo).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText, "\").Last) Then
                IO.File.Delete(IO.Path.GetDirectoryName(fileName) & "\" & Split(xnlCitations(iNo).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText, "\").Last)
            End If
            Try
                My.Computer.FileSystem.RenameFile(IO.Path.GetDirectoryName(fileName) & "\temp.tif", Split(xnlCitations(iNo).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText, "\").Last)
            Catch ex As Exception
                My.Computer.FileSystem.RenameFile(IO.Path.GetDirectoryName(fileName) & "\temp.tif", lblAccn.Text & CInt(xmlDoc.SelectSingleNode("//ITEM").Attributes("ITEMNO").Value).ToString("D3") & "C_" & xnlCitations.Count & ".TIF")
            End Try
            ' My.Computer.FileSystem.RenameFile(IO.Path.GetDirectoryName(fileName) & "\temp.tif", Split(xnlCitations(iNo).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText, "\").Last)

            'IO.File.Replace(IO.Path.GetDirectoryName(fileName) & "\temp.tif", fileName, Replace(fileName, ".tif", "-backup.tif", Compare:=CompareMethod.Text))
        End Using
    End Sub

    Private Sub subToggleZoning()
        pnlSource.Visible = Not pnlSource.Visible
        gbZoneTool.Visible = Not gbZoneTool.Visible
    End Sub

    Private Sub cmdZClose_Click(sender As Object, e As EventArgs) Handles cmdZClose.Click
        subToggleZoning()
        pbSource.Image = Nothing
    End Sub

    Shared Function getCitations() As XmlNodeList
        Return frmMain.xnlCitations
    End Function

    Shared Function getInPath() As String
        If frmMain.bIsPrior Then
            Return frmMain.strPriorPath
        Else
            Return frmMain.strInPath
        End If
    End Function

    Shared Function getInFile() As String
        Return frmMain.strInFile
    End Function

    Private Sub cmdViewAll_Click(sender As Object, e As EventArgs) Handles cmdViewAll.Click
        frmImages.Show()
    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        pbImage.ImageLocation = pbImage.ImageLocation
    End Sub

    Private Sub txtArtn1_GotFocus(sender As Object, e As EventArgs) Handles txtArtn1.GotFocus
        rbARTN1.Checked = True
    End Sub

    Private Sub txtArtn3_GotFocus(sender As Object, e As EventArgs) Handles txtArtn3.GotFocus
        rbARTN3.Checked = True
    End Sub

    Private Sub txtArtn4_GotFocus(sender As Object, e As EventArgs) Handles txtArtn4.GotFocus
        rbARTN4.Checked = True
    End Sub

    Private Sub txtArtn4_VisibleChanged(sender As Object, e As EventArgs) Handles txtArtn4.VisibleChanged
        If txtArtn4.Visible = True Then
            cmdAddArtn.Visible = False
        Else
            cmdAddArtn.Visible = True
        End If
    End Sub

    Private Sub cmdLoadBackup_Click(sender As Object, e As EventArgs) Handles cmdLoadBackup.Click
        If cbBackup.SelectedItem = vbNullString Then
            MsgBox("Select file")
            Exit Sub
        Else
            Dim temppath As String
            If bIsPrior Then
                temppath = strPriorPath
            Else
                temppath = strInPath
            End If
            If IO.File.Exists(temppath & cbBackup.SelectedItem & "-backup.tif") Then
                pbSource.ImageLocation = temppath & cbBackup.SelectedItem & "-backup.tif"
            Else
                MsgBox("Invalid file name")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub cmdDone_ClientSizeChanged(sender As Object, e As EventArgs) Handles cmdDone.ClientSizeChanged

    End Sub

    Private Sub SavePatent()

        Dim inum As Integer = CInt(Trim(Split(lblCurrent.Text, "of")(0)))
        'Dim nod As XmlNode
        If Trim(txtTitle.Text) = vbNullString Then
            MsgBox("Title should not be empty", MsgBoxStyle.Information, "Verify")
            Exit Sub
        End If

        Dim root As XmlNode = xmlDoc.DocumentElement
        Dim elem As XmlElement = xmlDoc.CreateElement("CI_PATENT")
        Dim attr As XmlAttribute = xmlDoc.CreateAttribute("ID")

        attr.Value = txtAuthor.Text.Replace("&", String.Empty)
        elem.Attributes.Append(attr)
        Dim strInXMl As String = String.Empty

        If Trim(txtVolume.Text) <> vbNullString Then strInXMl = strInXMl & "<PATENT_VOLUME>" & txtVolume.Text & "</PATENT_VOLUME>" & vbNewLine
        If Trim(txtPage.Text) <> vbNullString Then strInXMl = strInXMl & "<PATENT_PAGE>" & txtPage.Text & "</PATENT_PAGE>" & vbNewLine
        If Trim(txtYear.Text) <> vbNullString Then strInXMl = strInXMl & "<PATENT_YEAR>" & txtYear.Text & "</PATENT_YEAR>" & vbNewLine
        If Trim(txtTitle.Text) <> vbNullString Then strInXMl = strInXMl & "<PATENT_ASSIGNEE>" & txtTitle.Text & "</PATENT_ASSIGNEE>" & vbNewLine

        elem.InnerXml = Replace(strInXMl, "&", "")
        'elem.InnerXml = strInXMl
        Dim node As XmlNode = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL")

        If node Is Nothing Then
            node = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_PATENT")
            If node Is Nothing Then
                xnlCitations(inum).SelectSingleNode("CI_INFO").AppendChild(elem)
            Else
                node.ParentNode.ReplaceChild(elem, node)
            End If
        Else
            node.ParentNode.ReplaceChild(elem, node)
        End If

        Dim inClipName As String = String.Empty
        Try
            inClipName = xnlCitations(inum).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText
            inClipName = Replace(inClipName, Split(inClipName, "_").Last, xnlCitations(inum).Attributes("seq").InnerText & ".TIF")
        Catch ex As Exception
            makeTree(xnlCitations(inum), "CI_CAPTURE/CI_IMAGE_CLIP_NAME")
            inClipName = lblAccn.Text & CInt(xmlDoc.SelectSingleNode("//ITEM").Attributes("ITEMNO").Value).ToString("D3") & "_" & xnlCitations(inum).Attributes("seq").InnerText & ".TIF"
        End Try

        'If inClipName = String.Empty Then
        '    makeTree(xnlCitations(inum), "CI_CAPTURE/CI_IMAGE_CLIP_NAME")
        '    inClipName = lblAccn.Text & CInt(xmlDoc.SelectSingleNode("//ITEM").Attributes("ITEMNO").Value).ToString("D3") & "_" & xnlCitations(inum).Attributes("seq").InnerText & ".TIF"
        'Else
        '    inClipName = Replace(inClipName, Split(inClipName, "_").Last, xnlCitations(inum).Attributes("seq").InnerText & ".TIF")
        'End If
        xnlCitations(inum).SelectSingleNode("CI_CAPTURE/CI_IMAGE_CLIP_NAME").InnerText = inClipName
        xnlCitations(inum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_BLURB").InnerText = txtOCR.Text
        For i = 1 To 4
            If Me.Controls.Find("txtArtn" & i, True)(0).Visible = True Then
                If (Trim(Me.Controls.Find("txtArtn" & i, True)(0).Text) = vbNullString) Xor (Me.Controls.Find("cmbArtn" & i, True)(0).Text = vbNullString) Then
                    MsgBox("Check Artn fields", MsgBoxStyle.Information, "Verify")
                    Exit Sub
                ElseIf Trim(Me.Controls.Find("txtArtn" & i, True)(0).Text) <> vbNullString Then
16:                 Dim nd As XmlNode = xnlCitations(inum).SelectSingleNode("RI_CITATIONIDENTIFIER[@seq='" & i & "']")
                    If nd Is Nothing Then
                        makeTreeAttr(xnlCitations(inum), "RI_CITATIONIDENTIFIER", "seq", "type", Trim(Me.Controls.Find("txtArtn" & i, True)(0).Text), i, Me.Controls.Find("cmbArtn" & i, True)(0).Text)
                    End If
                End If
            Else
                Exit For
            End If
        Next
        Try
            Dim nod As XmlNode
            nod = xnlCitations(inum).SelectSingleNode("CI_INFO/CI_JOURNAL")
            nod.ParentNode.RemoveChild(nod)
        Catch ex As Exception
        End Try
        attr = xnlCitations(inum).Attributes("D")
        If attr IsNot Nothing Then
            Try
                xnlCitations(inum).RemoveChild(xnlCitations(inum).SelectSingleNode("FULL_CI_INFO"))
            Catch ex As Exception
            End Try
        End If
        xmlDoc.Save(Split(xmlDoc.BaseURI, "///").Last)
    End Sub

    Public Shared Function CombineImages(ByVal img1 As Image, ByVal img2 As Image) As Image
        Dim bmp As New Bitmap(Math.Max(img1.Width, img2.Width), img1.Height + img2.Height)
        Dim g As Graphics = Graphics.FromImage(bmp)

        g.DrawImage(img1, 0, 0, img1.Width, img1.Height)
        g.DrawImage(img2, 0, img1.Height, img2.Width, img2.Height)
        g.Dispose()
        Return bmp
    End Function

    Private Sub ThemeColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThemeColorToolStripMenuItem.Click
        If cdColorDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.BackColor = cdColorDlg.Color
        End If
    End Sub

    Private Sub txtArtn1_KeyDown(sender As Object, e As KeyEventArgs) Handles txtArtn1.KeyDown
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.Down Then
                Try
                    cmbArtn1.SelectedIndex = (cmbArtn1.SelectedIndex + 1)
                Catch ex As Exception
                    cmbArtn1.SelectedIndex = (0)
                End Try
                e.Handled = True
            ElseIf e.KeyCode = Keys.Up Then
                Try
                    cmbArtn1.SelectedIndex = (cmbArtn1.SelectedIndex - 1)
                Catch ex As Exception
                    cmbArtn1.SelectedIndex = (cmbArtn1.Items.Count - 1)
                End Try
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtArtn2_KeyDown(sender As Object, e As KeyEventArgs) Handles txtArtn2.KeyDown
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.Down Then
                Try
                    cmbArtn2.SelectedIndex = (cmbArtn2.SelectedIndex + 1)
                Catch ex As Exception
                    cmbArtn2.SelectedIndex = (0)
                End Try
                e.Handled = True
            ElseIf e.KeyCode = Keys.Up Then
                Try
                    cmbArtn2.SelectedIndex = (cmbArtn2.SelectedIndex - 1)
                Catch ex As Exception
                    cmbArtn2.SelectedIndex = (cmbArtn2.Items.Count - 1)
                End Try
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtArtn3_KeyDown(sender As Object, e As KeyEventArgs) Handles txtArtn3.KeyDown
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.Down Then
                Try
                    cmbArtn3.SelectedIndex = (cmbArtn3.SelectedIndex + 1)
                Catch ex As Exception
                    cmbArtn3.SelectedIndex = (0)
                End Try
                e.Handled = True
            ElseIf e.KeyCode = Keys.Up Then
                Try
                    cmbArtn3.SelectedIndex = (cmbArtn3.SelectedIndex - 1)
                Catch ex As Exception
                    cmbArtn3.SelectedIndex = (cmbArtn3.Items.Count - 1)
                End Try
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtArtn4_KeyDown(sender As Object, e As KeyEventArgs) Handles txtArtn4.KeyDown
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.Down Then
                Try
                    cmbArtn4.SelectedIndex = (cmbArtn4.SelectedIndex + 1)
                Catch ex As Exception
                    cmbArtn4.SelectedIndex = (0)
                End Try
                e.Handled = True
            ElseIf e.KeyCode = Keys.Up Then
                Try
                    cmbArtn4.SelectedIndex = (cmbArtn4.SelectedIndex - 1)
                Catch ex As Exception
                    cmbArtn4.SelectedIndex = (cmbArtn4.Items.Count - 1)
                End Try
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub QueryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QueryToolStripMenuItem.Click
        If MsgBox("Do you want to reject this item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
            If strInFile Is String.Empty Then Exit Sub

            On Error Resume Next
            My.Computer.FileSystem.DeleteFile(strQueryPath & System.IO.Path.GetFileName(strInFile))

            My.Computer.FileSystem.MoveFile(strCurPath & System.IO.Path.GetFileName(strInFile), strQueryPath & System.IO.Path.GetFileName(strInFile))

            For Each file As String In My.Computer.FileSystem.GetFiles(strInPath)
                If file.Contains(IO.Path.GetFileNameWithoutExtension(strInFile)) Then
                    'If Not file.Contains("backup") Then
                    'IO.File.Copy(file, strOutPath & System.IO.Path.GetFileName(file))
                    'IO.File.Copy(file, strBackUpPath & Format(Today, "yyyyMMdd") & "\" & System.IO.Path.GetFileName(file))
                    My.Computer.FileSystem.MoveFile(file, strQueryPath & System.IO.Path.GetFileName(file))
                    'Else
                    'IO.File.Move(file, strTifBackUpPath & System.IO.Path.GetFileName(file))
                    'End If
                End If
            Next
            On Error GoTo 0
            'iCompItems = iCompItems + 1
            'iCompRefs = iCompRefs + xnlCitations.Count
            swCurProd.Stop()
            swCurSusp.Stop()
            'swLogFile.WriteLine(vbTab & xnlCitations.Count & vbTab & swCurProd.Elapsed.ToString("hh\:mm\:ss") & vbTab & vbTab & swCurSusp.Elapsed.ToString("hh\:mm\:ss"))
            'swLogFile.Close()

            'Dim TotRefs As Integer = xnlCitations.Count
            'Dim DelRefs As Integer = xnlCitations(0).ParentNode.SelectNodes("CI_CITATION[@D='Y']").Count
            'Dim InsRefs As Integer = xnlCitations(0).ParentNode.SelectNodes("CI_CITATION[@D='Y']").Count
            'LogEntry.Referances_Close = TotRefs + InsRefs - DelRefs
            'LogEntry.Referances_Deleted = DelRefs
            'LogEntry.Referances_Inserted = InsRefs
            LogEntry.Session_Production = swCurProd.Elapsed.ToString("hh\:mm\:ss")
            LogEntry.Session_Suspend = swCurSusp.Elapsed.ToString("hh\:mm\:ss")
            LogEntry.Session_End_Date = Format(Today, "dd-MM-yyyy")
            LogEntry.Session_End_Time = Now.ToShortTimeString
            LogEntry.Status = "QR"
            'MakeLogEntry(strLogPath & frmLogin.txtUserName.Text & ".xml")
            UpdateLogEntry(strLogPath & frmLogin.txtUserName.Text & ".xml", LogEntry)
            lblCompItems.Text = iCompItems
            lblCompRefs.Text = iCompRefs
            fnClearFields()
            pbSource.Image = Nothing
            lblCurrent.Text = " of "

            'System.IO.File.Move(strCurPath & System.IO.Path.GetFileName(strInFile), strCompPath & System.IO.Path.GetFileName(strInFile))
            If MsgBox("Do you want to load new item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                strInFile = fnGetXMLFile(strPriorPath)

                If strInFile = "" Then
                    bIsPrior = False
                    strInFile = fnGetXMLFile(strInPath)
                End If
                If strInFile = "" Then
                    MsgBox("No files for input", MsgBoxStyle.Information)
                Else
                    bIsPrior = False
                    fnLoadFile(strInFile)
                End If
            Else
                strInFile = ""
                lvlist.Items.Clear()
            End If
        End If
    End Sub

    Private Function bumpSpace(sTemp As String) As String

        Dim bBumped As Boolean = False
        While Not bBumped
            If sTemp.Contains("  ") Then
                sTemp = sTemp.Replace("  ", " ")
                bBumped = False
            Else
                bBumped = True
            End If
        End While
        Return sTemp
    End Function

    Private Sub txtArtn1_TextChanged(sender As Object, e As EventArgs) Handles txtArtn1.TextChanged
        txtArtn1.Text = txtArtn1.Text.Replace(" ", String.Empty).Trim
    End Sub

    Private Sub txtArtn2_TextChanged(sender As Object, e As EventArgs) Handles txtArtn2.TextChanged
        txtArtn2.Text = txtArtn2.Text.Replace(" ", String.Empty).Trim
    End Sub

    Private Sub txtArtn3_TextChanged(sender As Object, e As EventArgs) Handles txtArtn3.TextChanged
        txtArtn3.Text = txtArtn3.Text.Replace(" ", String.Empty).Trim
    End Sub

    Private Sub txtArtn4_TextChanged(sender As Object, e As EventArgs) Handles txtArtn4.TextChanged
        txtArtn4.Text = txtArtn4.Text.Replace(" ", String.Empty).Trim
    End Sub

    Private Function MakeDirectory(strDir As String) As Boolean
        Try
            If Not IO.Directory.Exists(strDir) Then
                IO.Directory.CreateDirectory(strDir)
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub DoneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DoneToolStripMenuItem.Click
        cmdDone.PerformClick()
    End Sub

    Private Sub checkNTR(ndNode As XmlNode)
        Try
            If ndNode.SelectSingleNode("CI_TITLE").InnerText = "**NON-TRADITIONAL**" Then
                chkNTR.Checked = True
            Else
                chkNTR.Checked = False
            End If
        Catch ex As Exception
            chkNTR.Checked = False
        End Try
    End Sub

    'Private Sub AdminToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AdminToolStripMenuItem.Click
    '    MakeLogFile(strLogPath & frmLogin.txtUserName.Text & ".xml")
    'End Sub

    Private Sub MakeLogFile(filename As String)
        If Not IO.File.Exists(filename) Then
            Dim writer As New XmlTextWriter(filename, System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 3
            writer.WriteStartElement("Logs")
            writer.WriteEndElement()                                    'Logs
            writer.WriteEndDocument()
            writer.Close()
        Else
            Dim tdoc As New XmlDocument
            Try
                tdoc.Load(filename)
            Catch ex As Exception
                Dim tsb As New System.Text.StringBuilder
                tsb.AppendLine("DB file is either corrupted or invalid")
                tsb.AppendLine("It will be backed up and New DB file will be created")
                MsgBox(tsb.ToString)
                Try
                    My.Computer.FileSystem.RenameFile(filename, IO.Path.GetFileName(filename) & ".backup")
                Catch ex1 As Exception
                    Dim i As Integer = 1
                    While (IO.File.Exists(IO.Path.GetFileName(filename) & ".backup" & i))
                        i = i + 1
                    End While
                    My.Computer.FileSystem.RenameFile(filename, IO.Path.GetFileName(filename) & ".backup" & i)
                End Try

                Dim writer As New XmlTextWriter(filename, System.Text.Encoding.UTF8)
                writer.WriteStartDocument(True)
                writer.Formatting = Formatting.Indented
                writer.Indentation = 3
                writer.WriteStartElement("Logs")
                writer.WriteEndElement()                                    'Logs
                writer.WriteEndDocument()
                writer.Close()
            End Try
        End If
        'MakeLogEntry(filename)
        tsslStatus.Text = "Created"
    End Sub

    Private Sub InitLog(ByRef LE As Log)
        LE.Session_Start_Date = ""
        LE.Session_Start_Time = ""
        LE.Session_End_Date = ""
        LE.Session_End_Time = ""
        LE.Session_Production = ""
        LE.Session_Suspend = ""
        LE.Item_Accession = ""
        LE.Item_Itemno = ""
        LE.Status = ""
        LE.Referances_Open = ""
        LE.Referances_Close = ""
        LE.Referances_Inserted = ""
        LE.Referances_Deleted = ""
    End Sub

    Private Sub MakeLogEntry(ByVal filename As String)
        Dim xdDoc As New XmlDocument
        Dim nod1 As XmlNode, nod2 As XmlNode, attr As XmlAttribute

        Try
            xdDoc.Load(strLogPath & frmLogin.txtUserName.Text & ".xml")
        Catch ex As Exception
            MsgBox("Could not open DB file")
            Exit Sub
        End Try

        nod1 = xdDoc.DocumentElement
        nod2 = xdDoc.CreateElement("Log")
        attr = xdDoc.CreateAttribute("seq")
        attr.Value = xdDoc.DocumentElement.SelectNodes("Log").Count + 1
        nod2.Attributes.Append(attr)
        nod1.AppendChild(nod2)
        nod1 = nod2

        nod2 = xdDoc.CreateElement("System")
        attr = xdDoc.CreateAttribute("type")
        attr.Value = LogEntry.System_Type
        nod2.Attributes.Append(attr)
        nod2.InnerText = LogEntry.System
        nod1.AppendChild(nod2)

        nod2 = xdDoc.CreateElement("Session")
        nod1.AppendChild(nod2)
        nod1 = nod2

        nod2 = xdDoc.CreateElement("Start_Time")
        nod1.AppendChild(nod2)
        nod1 = nod2
        nod2 = xdDoc.CreateElement("Date")
        nod2.InnerText = LogEntry.Session_Start_Date
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("Time")
        nod2.InnerText = LogEntry.Session_Start_Time
        nod1.AppendChild(nod2)

        nod1 = nod1.ParentNode
        nod2 = xdDoc.CreateElement("End_Time")
        nod1.AppendChild(nod2)
        nod1 = nod2
        nod2 = xdDoc.CreateElement("Date")
        nod2.InnerText = LogEntry.Session_End_Date
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("Time")
        nod2.InnerText = LogEntry.Session_End_Time
        nod1.AppendChild(nod2)
        'nod1.ParentNode.AppendChild(nod1)

        nod1 = nod1.ParentNode
        nod2 = xdDoc.CreateElement("Production")
        nod2.InnerText = LogEntry.Session_Production
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("Suspend")
        nod2.InnerText = LogEntry.Session_Suspend
        nod1.AppendChild(nod2)

        nod1 = nod1.ParentNode
        nod2 = xdDoc.CreateElement("Item")
        nod1.AppendChild(nod2)
        nod1 = nod2
        nod2 = xdDoc.CreateElement("Accession")
        nod2.InnerText = LogEntry.Item_Accession
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("ItemNo")
        nod2.InnerText = LogEntry.Item_Itemno
        nod1.AppendChild(nod2)
        'nod1.ParentNode.AppendChild(nod1)

        nod1 = nod1.ParentNode
        nod2 = xdDoc.CreateElement("Status")
        nod2.InnerText = LogEntry.Status
        nod1.AppendChild(nod2)

        nod2 = xdDoc.CreateElement("Referances")
        nod1.AppendChild(nod2)
        nod1 = nod2
        nod2 = xdDoc.CreateElement("Open")
        nod2.InnerText = LogEntry.Referances_Open
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("Close")
        nod2.InnerText = LogEntry.Referances_Close
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("Inserted")
        nod2.InnerText = LogEntry.Referances_Inserted
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("Deleted")
        nod2.InnerText = LogEntry.Referances_Deleted
        nod1.AppendChild(nod2)

        xdDoc.Save(filename)
    End Sub

    Private Sub UpdateLogEntry(ByVal filename As String, LE As Log)
        Dim xddoc As New XmlDocument, lChild As XmlNode
        Try
            xddoc.Load(filename)
        Catch ex As Exception
            MsgBox("Could not open DB file")
            Exit Sub
        End Try

        lChild = xddoc.DocumentElement.LastChild   '.SelectSingleNode(tagname).InnerText = newval
        If lChild Is Nothing Then
            MsgBox("Could not update. Check DB file")
            Exit Sub
        End If
        lChild.SelectSingleNode("System").InnerText = LE.System
        lChild.SelectSingleNode("System").Attributes("type").Value = LE.System_Type
        lChild.SelectSingleNode("Session/Start_Time/Date").InnerText = LE.Session_Start_Date
        lChild.SelectSingleNode("Session/Start_Time/Time").InnerText = LE.Session_Start_Time
        lChild.SelectSingleNode("Session/End_Time/Date").InnerText = LE.Session_End_Date
        lChild.SelectSingleNode("Session/End_Time/Time").InnerText = LE.Session_End_Time
        lChild.SelectSingleNode("Session/Production").InnerText = LE.Session_Production
        lChild.SelectSingleNode("Session/Suspend").InnerText = LE.Session_Suspend
        lChild.SelectSingleNode("Item/Accession").InnerText = LE.Item_Accession
        lChild.SelectSingleNode("Item/ItemNo").InnerText = LE.Item_Itemno
        lChild.SelectSingleNode("Status").InnerText = LE.Status
        lChild.SelectSingleNode("Referances/Open").InnerText = LE.Referances_Open
        lChild.SelectSingleNode("Referances/Close").InnerText = LE.Referances_Close
        lChild.SelectSingleNode("Referances/Inserted").InnerText = LE.Referances_Inserted
        lChild.SelectSingleNode("Referances/Deleted").InnerText = LE.Referances_Deleted
        xddoc.Save(filename)
    End Sub



    'Private Sub EditLogEntry(ByVal filename As String, ByVal tagname As String, ByVal newval As String)
    '    Dim xddoc As New XmlDocument
    '    xddoc.Load(filename)
    '    xddoc.DocumentElement.LastChild.SelectSingleNode(tagname).InnerText = newval
    '    xddoc.Save(filename)
    'End Sub

    'Public Shared Sub tester()
    '    frmMain.MakeLogFile(frmMain.strLogPath & frmLogin.txtUserName.Text & ".xml")
    '    frmMain.EditLogEntry(frmMain.strLogPath & frmLogin.txtUserName.Text & ".xml", "Referances/Open", "150")
    'End Sub

End Class
