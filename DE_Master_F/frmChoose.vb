Public Enum ToolMode
    OP = 1
    QC = 2
End Enum

Public Class frmChoose

    Public Shared chosenTool As ToolMode

    Private Sub lblOp_Click(sender As Object, e As EventArgs) Handles lblOp.Click
        chosenTool = ToolMode.OP
        Me.Hide()
        frmlogin.Show()
    End Sub

    Private Sub lblQC_Click(sender As Object, e As EventArgs) Handles lblQC.Click
        chosenTool = ToolMode.QC
        Me.Hide()
        frmlogin.Show()
    End Sub

    Private Sub lblOp_MouseEnter(sender As Object, e As EventArgs) Handles lblOp.MouseEnter
        Me.Cursor = Cursors.Hand
        Me.Refresh()
    End Sub

    Private Sub lblOp_MouseLeave(sender As Object, e As EventArgs) Handles lblOp.MouseLeave
        Me.Cursor = Cursors.Default
        Me.Refresh()
    End Sub

    Private Sub lblQC_MouseEnter(sender As Object, e As EventArgs) Handles lblQC.MouseEnter
        Me.Cursor = Cursors.Hand
        Me.Refresh()
    End Sub

    Private Sub lblQC_MouseLeave(sender As Object, e As EventArgs) Handles lblQC.MouseLeave
        Me.Cursor = Cursors.Default
        Me.Refresh()
    End Sub

    Private Sub lblClose_Click(sender As Object, e As EventArgs) Handles lblClose.Click
        Me.Close()
    End Sub
End Class