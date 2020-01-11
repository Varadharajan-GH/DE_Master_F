Imports System.IO

Public Class SettingsReader
    'Private ReadOnly ConfigFile As String = Path.GetFullPath(".\") & "App.cfg"

    Public Shared Function ReadSetting(setting As String, Optional filename As String = "") As String
        If String.IsNullOrWhiteSpace(filename) Then
            filename = Path.GetFullPath(".\") & "App.cfg"
        End If

        If Not File.Exists(filename) Then
            MsgBox("Configuration file missing")
            Return String.Empty
        End If

        Using reader As New StreamReader(filename)
            Do While reader.Peek <> -1
                Dim TextLine As String = reader.ReadLine().Trim
                If String.IsNullOrWhiteSpace(TextLine) Or TextLine.StartsWith("#") Or (Not TextLine.Contains("=")) Then
                    Continue Do
                End If
                Dim key As String = TextLine.Split("="c).First.Trim
                Dim value As String = TextLine.Split("="c).Last.Trim

                If key = setting Then
                    Return value
                End If
            Loop
        End Using
        MsgBox("Unable to find setting : " & setting)
        Return String.Empty
    End Function




End Class
