Public Class Form1
    Declare Function SetCursorPos Lib "user32" (ByVal p As Point) As Long
    Declare Sub mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Integer, ByVal dx As Integer， ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Public selected As Boolean = False
    Public selectpoint As Point
    Dim a As Integer = 0
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        For i1 = 0 To TrackBar1.Value
            For i2 = 0 To 10
                For i3 = 0 To 100
                    SetCursorPos(selectpoint)
                    mouse_event(&H2, 0, 0, 0, 0)
                    mouse_event(&H4, 0, 0, 0, 0)
                Next
                System.Threading.Thread.Sleep(TrackBar2.Value)
            Next
        Next
        Me.Show()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label7.Text = TrackBar1.Value * 1000
        Label4.Text = TrackBar1.Value
        Label9.Text = TrackBar2.Value
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Label7.Text = TrackBar1.Value * 1000
        Label4.Text = TrackBar1.Value
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        Label9.Text = TrackBar2.Value
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
       
    End Sub
End Class
