Imports System.Windows.Forms

Public Class Dialog1
    Private mouseDownLoc As Point
    Public valID As String
    Public Function valBox2(ByVal text As String)
        valBox2 = ""
        erBox.Label1.Text = text
        erBox.ShowDialog()
    End Function
    Private Sub Dialog1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Not TextBox1.Text.Length = 0 Then
                If Not TextBox1.Text.Length = 4 Then
                    Dim bool As Boolean = valBox("Invalid Length....." & vbNewLine & " Do you want to Continue??", "Stop!!")
                    If bool = True Then
                        valID = TextBox1.Text
                        TextBox1.Text = ""
                        Me.Close()
                    End If
                Else
                    valID = TextBox1.Text
                    TextBox1.Text = ""
                    Me.Close()
                End If
            Else
                valBox2("Operator's ID Cannot be Blank")
            End If
        End If
        If e.KeyCode = Keys.Escape Then
            TextBox1.Text = ""
            Me.Close()
        End If
    End Sub
    Public Function valBox(ByVal validation As String, Optional ByVal caption As String = "") As Boolean
        messageBox.Label2.Text = caption
        messageBox.Label1.Text = validation
        Dim ask As DialogResult
        ask = messageBox.ShowDialog
        If ask = Windows.Forms.DialogResult.Cancel Then
            Return False
        ElseIf ask = Windows.Forms.DialogResult.OK Then
            Return True
        End If
    End Function
    Private Sub Dialog1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub
    Private Sub Panel1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        mouseDownLoc.X = e.X
        mouseDownLoc.Y = e.Y
    End Sub
    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Const DROPSHADOW = &H30000
            Dim cParam As CreateParams = MyBase.CreateParams
            cParam.ClassStyle = cParam.ClassStyle Or DROPSHADOW
            Return cParam
        End Get
    End Property
    Private Sub Panel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Me.Left += e.X - mouseDownLoc.X
            Me.Top += e.Y - mouseDownLoc.Y
        End If
    End Sub

End Class
