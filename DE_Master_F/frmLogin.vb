Imports System.ComponentModel
Imports System.IO
Imports DEM_FCR_Library

Public Class frmLogin

    Public UserName As String
    'Private ConfigFile As String = Path.GetFullPath(".\") & "App.cfg"

    Private usersFile As String
    Private sourcePath As String

    'Public Const SOURCEPATH As String = "D:\data\F_Process\"
    'Public Const SOURCEPATH As String = ".\"


    'Dim usersFile As String = SOURCEPATH & "Admin\users.txt"
    'Dim usersFile As String = "\\172.21.15.100\data\Admin\users.txt"
    'Dim usersFile As String = ".\Admin\users.txt"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        sourcePath = SettingsReader.ReadSetting("SOURCE_PATH")
        If String.IsNullOrWhiteSpace(sourcePath) Then
            MsgBox("Source path is empty")
            Application.Exit()
        ElseIf Not Directory.Exists(sourcepath) Then
            MsgBox("Source path does not exist")
            Application.Exit()
        End If
        If Not sourcePath.EndsWith(Path.DirectorySeparatorChar) Then
            sourcePath &= (Path.DirectorySeparatorChar)
        End If

        usersFile = sourcePath & SettingsReader.ReadSetting("USERS_FILE")
        If Not File.Exists(usersFile) Then
            MsgBox("Users file does not exist")
            Application.Exit()
        End If

        txtUserName.Focus()
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs)
        Application.Exit()
    End Sub

    Private Sub cmdLogin_Click(sender As Object, e As EventArgs) Handles cmdLogin.Click

        Dim bLoggedIn As Boolean = False
        Dim TextLine As String

        If File.Exists(usersFile) Then
            Using objReader As New StreamReader(usersFile)
                Do While (objReader.Peek() <> -1 And bLoggedIn = False)
                    TextLine = objReader.ReadLine()
                    If TextLine = txtUserName.Text & " " & txtPassword.Text Then
                        bLoggedIn = True
                        txtPassword.Clear()
                        txtPassword.Focus()
                        Me.Hide()
                        UserName = txtUserName.Text
                        frmFMain.Enabled = True
                        frmFMain.Show()
                    End If
                Loop
            End Using
            If Not bLoggedIn Then MsgBox("Wrong credentials")
            txtPassword.Focus()
            txtPassword.SelectAll()
        Else
            MessageBox.Show("Admin File Does Not Exist")
        End If
    End Sub



    Private Sub txtUserName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUserName.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txtPassword.Focus()
            txtPassword.SelectAll()
            e.Handled = True
        End If
    End Sub

    Private Sub txtPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPassword.KeyPress
        If Asc(e.KeyChar) = 13 Then
            cmdLogin.Focus()
            e.Handled = True
            cmdLogin.PerformClick()
        End If
    End Sub

    Private Sub frmLogin_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If frmFMain Is Nothing Then
            Application.Exit()
        Else
            MsgBox("You must Login.")
            e.Cancel = True
        End If
    End Sub
End Class
