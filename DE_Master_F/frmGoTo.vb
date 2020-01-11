Public Class frmGoTo

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            If CInt(txtRef.Text) < 1 Or CInt(txtRef.Text) > frmFMain.lstRefNum.Count Then
                MsgBox("Invalid Reference")
            Else
                If frmFMain.SaveReference() = -1 Then
                    Me.Close()
                    Exit Sub
                End If
                frmFMain.subLoadRefs(frmFMain.getCitations, frmFMain.lstRefNum(CInt(txtRef.Text) - 1), "Next")
                Me.Close()
            End If
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub

    Private Sub frmGoTo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            lblTotRefs.Text = " / " & frmFMain.lstRefNum.Count
            txtRef.Focus()
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub
End Class