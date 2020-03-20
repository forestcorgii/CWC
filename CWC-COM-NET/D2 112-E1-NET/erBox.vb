Imports System.Windows.Forms

Public Class erBox
    Private Sub erBox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.Close()
    End Sub
    Private Sub erBox_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'OK_Button.Focus()
    End Sub
End Class
