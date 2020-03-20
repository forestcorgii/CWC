Imports System.Windows.Forms

Public Class messageBox
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Private Sub Dialog2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Cancel_Button.Select()
    End Sub

    Private Sub OK_Button_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub OK_Button_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles OK_Button.GotFocus
        OK_Button.BackColor = Color.FromArgb(47, 21, 65)
        OK_Button.ForeColor = Color.White

    End Sub

    Private Sub Cancel_Button_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cancel_Button.GotFocus
        Cancel_Button.BackColor = Color.FromArgb(47, 21, 65)
        Cancel_Button.ForeColor = Color.White

    End Sub

    Private Sub Cancel_Button_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cancel_Button.LostFocus
        Cancel_Button.BackColor = Color.White
        Cancel_Button.ForeColor = Color.FromArgb(47, 21, 65)

    End Sub

    Private Sub OK_Button_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles OK_Button.LostFocus
        OK_Button.BackColor = Color.White
        OK_Button.ForeColor = Color.FromArgb(47, 21, 65)
    End Sub
End Class
