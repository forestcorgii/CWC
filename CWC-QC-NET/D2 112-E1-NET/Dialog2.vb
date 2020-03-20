Imports System.Windows.Forms

Public Class Dialog2
    Private datetime As String
    Private val2 As Boolean
    Private uID, locID As String
    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Const DROPSHADOW = &H30000
            Dim cParam As CreateParams = MyBase.CreateParams
            cParam.ClassStyle = cParam.ClassStyle Or DROPSHADOW
            Return cParam
        End Get
    End Property
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Hide()
    End Sub
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Hide()
    End Sub

    Private Sub Dialog2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
        If e.KeyCode = Keys.Up Then
            txtLocID.Focus()
        End If
    End Sub

    Private Sub Dialog2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dateTime = Now.ToString("G")
        Form1.valBox2("System DateTime: " & datetime)
        txtLocID.Select()
    End Sub

    Private Sub txtLocID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLocID.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If

        val2 = True
        If e.KeyCode = Keys.Enter Then
            If val2 = False Then
                Exit Sub
            End If
            If txtLocID.Text = "" Then
                MsgBox("Location ID cannot be Blank")
                txtLocID.Focus()
                txtLocID.Select(txtLocID.Text.Length, 0)

            ElseIf Not txtLocID.Text.Length = 4 Then
                Dim ask1 As Boolean
                ask1 = Form1.valBox("Location ID length is not equal to 4", "Error")
                If ask1 = True Then
                    locID = txtLocID.Text
                    Form1.txtLocID.Text = locID
                    txtUserID.Select()
                    val2 = False
                ElseIf ask1 = False Then
                    txtLocID.Focus()
                    txtLocID.Select(txtLocID.Text.Length, 0)

                    val2 = False
                    Exit Sub
                End If
            Else
                locID = txtLocID.Text
                Form1.txtLocID.Text = locID
                txtUserID.Select()
            End If
        End If
    End Sub

 
    Private Sub txtUserID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserID.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If

        If e.KeyCode = Keys.Up Then
            txtLocID.Focus()
        End If
        If e.KeyCode = Keys.Enter Then
            If txtUserID.Text = "" Then
                MsgBox("User ID cannot be Blank")
                txtUserID.Text = ""
                txtUserID.Focus()
                txtUserID.Select(txtLocID.Text.Length, 0)
            ElseIf Not txtUserID.Text.Length = 4 Then
                Dim ask1 As Boolean
                ask1 = Form1.valBox("User ID length is not equal to 4", "Error")
                If ask1 = True Then
                    uID = txtUserID.Text
                    Form1.txtUserID.Text = uID
                    If Not Form1.Visible = True Then
                        txtUserID.Text = ""
                        txtLocID.Text = ""
                        Form1.Show()
                        Me.Close()
                    Else
                        txtUserID.Text = ""
                        txtLocID.Text = ""
                        Form1.Show()
                        Me.Close()
                    End If
                ElseIf ask1 = False Then
                    txtUserID.Focus()
                    txtUserID.Select(txtLocID.Text.Length, 0)
                    Exit Sub
                End If
            Else
                val2 = True
                uID = txtUserID.Text
                Form1.txtUserID.Text = uID
                If Not Form1.Visible = True Then
                    txtUserID.Text = ""
                    txtLocID.Text = ""
                    Form1.Show()
                    Me.Close()
                Else
                    txtUserID.Text = ""
                    txtLocID.Text = ""
                    Form1.Show()
                    Me.Close()
                End If
            End If
        End If
    End Sub

 
End Class
