Imports System.Windows.Forms
Imports System.Data.OleDb
Public Class dataDialog
    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Const DROPSHADOW = &H30000
            Dim cParam As CreateParams = MyBase.CreateParams
            cParam.ClassStyle = cParam.ClassStyle Or DROPSHADOW
            Return cParam
        End Get
    End Property
    Private ds As DataSet
    Private recordnumber
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtLocID.Select()
    End Sub

    Private Sub txtLocID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLocID.KeyDown
        If e.KeyCode = Keys.Escape Then
            txtLocID.Text = ""
            Me.Close()
        End If

        If e.KeyCode = Keys.Enter Then
            If Not txtLocID.Text = "" Then
                If Microsoft.VisualBasic.Val(txtLocID.Text) <= Microsoft.VisualBasic.Val(Form1.txtTotalRec.Text) Then
                    Form1.curRec = txtLocID.Text - 1
                    Form1.txtCrntRec.Text = Form1.curRec + 1
                    Form1.loadDataFields()
                    Form1.getImages()
                    txtLocID.Text = ""
                    Me.Close()
                Else
                    MsgBox("Out of Bound")
                End If
            End If
        End If
    End Sub
    'Public Function recordNumberDisplay(ByVal rec As Integer) As Integer
    '    recordNumberDisplay = 0
    '    ds = New DataSet
    '    Dim da As New OleDbDataAdapter("select recnum from data001", Form1.myconn)
    '    da.Fill(ds)
    '    For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
    '        If i = rec - 1 Then
    '            Return ds.Tables(0).Rows(i).Item(0)
    '        End If
    '    Next
    'End Function
End Class
