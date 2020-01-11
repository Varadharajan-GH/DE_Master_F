Public Class Resizer

    Private Structure ControlInfo
        Public name As String
        Public parentName As String
        Public leftOffsetPercent As Double
        Public topOffsetPercent As Double
        Public heightPercent As Double
        Public originalHeight As Integer
        Public originalWidth As Integer
        Public widthPercent As Double
        Public originalFontSize As Single
    End Structure

    
    Private ctrlDict As Dictionary(Of String, ControlInfo) = New Dictionary(Of String, ControlInfo)
    Public Sub FindAllControls(thisCtrl As Control)

        For Each ctl As Control In thisCtrl.Controls
            Try
                If Not IsNothing(ctl.Parent) Then
                    Dim parentHeight = ctl.Parent.Height
                    Dim parentWidth = ctl.Parent.Width

                    Dim c As New ControlInfo
                    c.name = ctl.Name
                    c.parentName = ctl.Parent.Name
                    c.topOffsetPercent = Convert.ToDouble(ctl.Top) / Convert.ToDouble(parentHeight)
                    c.leftOffsetPercent = Convert.ToDouble(ctl.Left) / Convert.ToDouble(parentWidth)
                    c.heightPercent = Convert.ToDouble(ctl.Height) / Convert.ToDouble(parentHeight)
                    c.widthPercent = Convert.ToDouble(ctl.Width) / Convert.ToDouble(parentWidth)
                    c.originalFontSize = ctl.Font.Size
                    c.originalHeight = ctl.Height
                    c.originalWidth = ctl.Width
                    ctrlDict.Add(c.name, c)
                End If

            Catch ex As Exception
                'Debug.Print(ex.Message)
            End Try

            If ctl.Controls.Count > 0 Then
                FindAllControls(ctl)
            End If

        Next

    End Sub

    Public Sub ResizeAllControls(thisCtrl As Control, Optional wform As Form = Nothing)

        Dim fontRatioW As Single
        Dim fontRatioH As Single
        Dim fontRatio As Single
        Dim f As Font
        If IsNothing(wform) = False Then wform.Opacity = 0

        For Each ctl As Control In thisCtrl.Controls
            Try
                If TypeOf ctl Is MenuStrip Then
                    Exit Try
                ElseIf TypeOf ctl Is TabControl Then
                    Exit Try
                ElseIf TypeOf ctl Is StatusStrip Then
                    Exit Try
                ElseIf Not IsNothing(ctl.Parent) Then
                    Dim parentHeight = ctl.Parent.Height
                    Dim parentWidth = ctl.Parent.Width

                    Dim c As New ControlInfo

                    Dim ret As Boolean = False
                    Try
                        ret = ctrlDict.TryGetValue(ctl.Name, c)
                        If (ret) Then
                            ctl.SetBounds(Int(parentWidth * c.leftOffsetPercent),
                                          Int(parentHeight * c.topOffsetPercent),
                                          Int(parentWidth * c.widthPercent),
                                          Int(parentHeight * c.heightPercent))
                            f = ctl.Font
                            fontRatioW = Int(parentWidth * c.widthPercent) / c.originalWidth
                            fontRatioH = Int(parentHeight * c.heightPercent) / c.originalHeight
                            fontRatio = (fontRatioW + fontRatioH) / 2
                            ctl.Font = New Font(f.FontFamily, c.originalFontSize * fontRatio, f.Style)

                        End If
                    Catch
                    End Try
                End If
            Catch ex As Exception
            End Try
            If ctl.Controls.Count > 0 Then
                ResizeAllControls(ctl)
            End If

        Next
        If IsNothing(wform) = False Then wform.Opacity = 1
    End Sub

End Class