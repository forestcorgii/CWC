Imports System.IO
Imports System.Collections
Imports System.Data.OleDb
Imports System.ComponentModel.Component
Imports System.Object
Imports System.Diagnostics.Process
Imports System.Windows.Forms
Imports System.Text.RegularExpressions

Public Structure validation
    Public text As String

    Public Function email() As Boolean
        '''' EMAIL ADDRESS'''
            Form1.txtValue.Text = Replace(Form1.txtValue.Text, Chr(32), "")
        Form1.txtValue.Text = Replace(Form1.txtValue.Text, "@@", "@")
        Form1.txtValue.Text = Replace(Form1.txtValue.Text, "..", ".")
        If Not Form1.txtValue.Text = "" Then
            If Form1.txtValue.Text.Contains("@") = False Or Form1.txtValue.Text.Contains(".") = False Then
                Dim ask As Boolean
                ask = Form1.valBox("Email Address Doesn't contain any @/. Do You want to proceed??", "Error")
                If ask = True Then
                    Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = Form1.txtValue.Text
                    Form1.nextLine()
                ElseIf ask = False Then
                    Return False
                End If
            Else
                Return True
            End If
        Else
            Return True
        End If
         ''EMAIL ADDRESS '' END''
    End Function
    Public Function fname() As Boolean
        ''''FIRSTNAME''''
             Form1.txtValue.Text = Regex.Replace(Form1.txtValue.Text, " {2,}", " ")
        Return True
        ''FIRSTNAME'' END''

    End Function
    Public Function lname() As Boolean
        Dim ask As Boolean
        ''''SURNAME''''
          If Form1.txtValue.Text = "" Then
            ask = Form1.valBox("Surname is Blank" & vbNewLine & "Continue Anyway??", "Blank")
            If ask = True Then
                Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = Form1.txtValue.Text
                Form1.nextLine()
            ElseIf ask = False Then
                Return False
            End If
        Else
            If Form1.txtValue.Text.Length = 1 Then
                ask = Form1.valBox("Check Length.." & vbNewLine & "Continue Anyway??", "Blank")
                If ask = True Then
                    Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = Form1.txtValue.Text
                    Form1.nextLine()
                ElseIf ask = False Then
                    Return False
                End If
            Else
                Return True
            End If
        End If
        ''SURNAME'' END ''

    End Function
    Public Function mname() As Boolean
        ''''FIRSTNAME''''
         Form1.txtValue.Text = Regex.Replace(Form1.txtValue.Text, " {2,}", " ")
        Return True
        ''FIRSTNAME'' END''

    End Function
    Public Function state() As Boolean
        ''''STATE''''
            Form1.txtValue.Text = StrConv(Form1.txtValue.Text, VbStrConv.Uppercase)
        If Not Form1.txtValue.Text = "" Then
            If Form1.txtValue.Text.Contains("  ") = True Then
                Form1.valBox2("Text contains double spaces")
                Return False
            Else
                Return True
            End If
        Else
            Return True
         End If
        ''STATE'' END''
    End Function
    Public Function zipCode() As Boolean
        ''''POSTCODE''''
        Form1.txtValue.Text = StrConv(Form1.txtValue.Text, VbStrConv.Uppercase)
        Return True
        ''POSTCODE'' END''
    End Function

    Public Function phone() As Boolean
        Dim ask As Boolean
        ''''CONTACTNUMBER''''
             If Not Form1.txtValue.Text = "" Then
            If Not Form1.txtValue.Text.Length <= 10 Then
                ask = Form1.valBox("Character Length is not Equals 10 " & vbNewLine & "Do you want to proceed", "Error")
                If ask = False Then
                    Return False
                ElseIf ask = True Then
                    Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = Form1.txtValue.Text
                    Form1.nextLine()
                End If
            Else
                Return True
            End If
        Else
            Return True
        End If
         ''CONTACTNUMBER'' END ''
    End Function
    Public Function title(Optional ByVal e As KeyPressEventArgs = Nothing, Optional ByVal pick As Integer = 0) As Boolean
        ''''TITLE''
        If pick = 1 Then
            Dim keystring As New KeysConverter
            If IsNumeric(e.KeyChar) = True Then
                Select Case e.KeyChar
                    Case "1"
                        Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = "MR"
                        Form1.nextLine()
                        Return e.Handled = True
                    Case "2"
                        Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = "MRS"
                        Form1.nextLine()
                        Return e.Handled = True
                    Case "3"
                        Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = "MS"
                        Form1.nextLine()
                        Return e.Handled = True
                    Case "4"
                        Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = "MISS"
                        Form1.nextLine()
                        Return e.Handled = True
                End Select
            Else
                Select Case e.KeyChar
                    Case "*"
ask:                    Dim ask As Boolean
                        ask = Form1.valBox("Are You Sure You Want To Pullout??")
                        If ask = True Then
                            For i As Integer = 0 To Form1.DataGridView1.Rows.Count - 1
                                Form1.DataGridView1.Rows(i).Cells(1).Value = ""
                            Next
                            Form1.DataGridView1.Rows(1).Cells(1).Value = e.KeyChar
                            If Form1.mymode = Form1.mode.editmode Then
                                Dim ask2 As Boolean
                                ask2 = Form1.valBox("Save Changes??")
                                If ask2 = True Then
                                    Form1.saveData()
                                    Form1.mymode = Form1.mode.browsemode
                                    Form1.changemode()
                                ElseIf ask2 = False Then
                                    Form1.mymode = Form1.mode.browsemode
                                    Form1.changemode()
                                End If
                            ElseIf Form1.mymode = Form1.mode.entrymode Then
                                Form1.saveData()
                            End If
                        End If
                End Select
            End If
        Else
            Select Case Form1.txtValue.Text
                Case "MR", "MRS", "MISS", "MS", "MASTER", "LADY", "REV", "DR", "MAJOR", "SIR", "CAPT", "PROF", "BRIG", "CPL", ""
                    Return True
                Case "*"
                    GoTo ask
                Case Else
                    Form1.valBox2("Invalid Title" & vbNewLine & "valid Titles:" & vbNewLine & "MR, MRS, MISS, MS, MASTER, LADY, REV, DR, MAJOR, SIR, CAPT, PROF, BRIG, CPL")
                    Return False
            End Select
        End If
        ''TITLE'' END''
    End Function
    Public Function qasflag(ByVal e As KeyEventArgs) As Boolean
        If Not e.KeyCode = Keys.Y And Not e.KeyCode = Keys.N And Not e.KeyCode = Keys.Enter Then
            Return False
        End If
        If e.KeyCode = Keys.Enter Then
            If Form1.txtValue.Text.Contains("Y") Or Form1.txtValue.Text.Contains("N") Then
                Return True
            Else
                Form1.valBox2("QAS Flag Cannot be Blank..")
                Return False
            End If
        Else
            Return True
        End If
    End Function
    Public Function address() As Boolean
        ''''STREET''''
        Dim ask As Boolean
        If Form1.txtValue.Text = "" Then
            ask = Form1.valBox("Address is Blank" & vbNewLine & "Continue Anyway??", "Blank")
            If ask = True Then
                Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = Form1.txtValue.Text
                Form1.nextLine()
            ElseIf ask = False Then
                Return False
            End If
        Else
            If Form1.txtValue.Text.Contains("  ") = True Then
                Form1.valBox2("Text contains double spaces")
                Return False
            Else
                Return True
            End If
        End If
        ''STREET'' END''
    End Function
    Public Function include()
        If Integer.Parse(Form1.txtValue.Text) > 15 Then
            Form1.valBox2("Invalid Type")
            Return False
        Else
            Return True
        End If
    End Function
    Public Function yn(ByVal text As String) As Boolean
        Select Case text
            Case "Y", "y", "N", "n", ""
                Return True
            Case Else
                Return False
        End Select
    End Function
    Public Function city() As Boolean
        ''''STREET''''
        Dim ask As Boolean
        If Form1.txtValue.Text = "" Then
            ask = Form1.valBox("City is Blank" & vbNewLine & "Continue Anyway??", "Blank")
            If ask = True Then
                Form1.DataGridView1.Rows(Form1.getCurrentRow).Cells(1).Value = Form1.txtValue.Text
                Form1.nextLine()
            ElseIf ask = False Then
                Return False
            End If
        Else
            If Form1.txtValue.Text.Contains("  ") = True Then
                Form1.valBox2("Text contains double spaces")
                Return False
            Else
                Return True
            End If
         End If
        ''STREET'' END''
    End Function
    Public Function commonReason() As Boolean
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        Dim d As Integer
        Dim e As Integer
        Dim f As Integer
        Dim g As Integer
        Dim h As Integer
        Dim j As Integer
        For i As Integer = 0 To Form1.txtValue.Text.Length - 1
            Select Case Form1.txtValue.Text.Substring(i, 1)
                Case "1"
                    a += 1
                Case "2"
                    b += 1
                Case "3"
                    c += 1
                Case "4"
                    d += 1
                Case "5"
                    e += 1
                Case "6"
                    f += 1
                Case "7"
                    g += 1
                Case "8"
                    h += 1
                Case "9"
                    j += 1
            End Select
        Next
        If a > 1 Or b > 1 Or c > 1 Or d > 1 Or e > 1 Or f > 1 Or g > 1 Or h > 1 Or j > 1 Then
            Form1.valBox2("Contains Duplicate")
            Return False
        Else
            Return True
        End If
    End Function
End Structure

Public Class Form1
    Private mouseDownLoc As Point
    Private da As OleDb.OleDbDataAdapter
    Private ds As DataSet
    Private fieldDS As DataSet
    Private mdbDS As DataSet
    Private clientname As String = "CWC"
    Private jobName As String = "YYF_RSG"
    Private sqlQry As String = ""
    Private dataBy As String = ""
    Public mymode As mode
    Private recordnumber As Integer = 0
    Public myconn As OleDb.OleDbConnection
    Public conn As OleDb.OleDbConnection
    Private dbLoc As String
    Private img1, OImg As String
    Public curRec As Integer = 0
    Private totalrec As Integer
    Private imagePath As String
    Private currentRows As Integer
    Private datetime As String = Now.ToString("G")
    Private valid As Boolean
    Private ask As DialogResult
    Private validation As New validation



    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Const DROPSHADOW = &H30000
            Dim cParam As CreateParams = MyBase.CreateParams
            cParam.ClassStyle = cParam.ClassStyle Or DROPSHADOW
            Return cParam
        End Get
    End Property
    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        Me.Close()
    End Sub
    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub Label1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label1.MouseDown
        mouseDownLoc.X = e.X
        mouseDownLoc.Y = e.Y
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
    Public Function valBox2(ByVal text As String)
        valBox2 = ""
        erBox.Label1.Text = text
        erBox.ShowDialog()
    End Function
    Private Sub Label1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label1.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Me.Left += e.X - mouseDownLoc.X
            Me.Top += e.Y - mouseDownLoc.Y
        End If
    End Sub
    Function getCurrentRow() As Integer
        Try
            Return DataGridView1.CurrentCell.RowIndex
        Catch ex As Exception
            Return -1
        End Try
    End Function
    Function setCurrentRow(ByVal ndx As Integer) As Boolean
        Try
            DataGridView1.CurrentCell = DataGridView1.Item(0, ndx)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Sub openmdb(ByVal connectionpath As String)
        If Not myconn Is Nothing Then
            If myconn.State = ConnectionState.Open Then
                myconn.Close()
            End If
        End If
        myconn = New OleDbConnection
        myconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & connectionpath
        myconn.Open()
     
    End Sub
    Public Function filldtg()
        filldtg = ""
        DataGridView1.Rows.Add("Code")
        mdbDS = New DataSet
        da = New OleDbDataAdapter("select * from Data001", myconn)
        da.Fill(mdbDS)
        totalrec = mdbDS.Tables(0).Rows.Count
        txtTotalRec.Text = totalrec
        txtCrntRec.Text = curRec + 1
    End Function
    Public Sub getCode(ByVal code As String)
        conn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "  Data Source=" & Application.StartupPath & "\Codings-E1.mdb" & ";User Id=admin;Password=;")
        If Not conn.State = ConnectionState.Open Then
            conn.Open()
        End If
        fieldDS = New DataSet
        Dim da As New OleDbDataAdapter("Select * FROM [CWC_tbl] WHERE Code = '" & code & "'", conn)
        da.Fill(fieldDS)
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add("Code")
        DataGridView1.Rows(getCurrentRow).Cells(1).Value = code
        For i As Integer = 0 To fieldDS.Tables(0).Columns.Count - 1
            If fieldDS.Tables(0).Rows(0).Item(i).ToString = "1" Then
                DataGridView1.Rows.Add(fieldDS.Tables(0).Columns(i).ColumnName)
            End If
        Next
        '  loadDataFields()
    End Sub
    Public Function getDesc()
        getDesc = ""
        Dim valDS As New DataSet
        Dim valDA As New OleDbDataAdapter("SELECT * FROM CWCval_tbl WHERE FieldName = '" & DataGridView1.Rows(getCurrentRow).Cells(0).Value & "'", conn)
        valDA.Fill(valDS)
        Try
            lblDesc.Text = valDS.Tables(0).Rows(0).Item("Description")
            txtValue.MaxLength = valDS.Tables(0).Rows(0).Item("MaxLength")
            txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(1).Value
        Catch ex As Exception
            '  MsgBox(ex.Message)
            lblDesc.Text = ""
        End Try
       End Function
    Public Sub loadDataFields(Optional ByVal code As String = "")

             getRec()
        da = New OleDbDataAdapter("select * from Data001 where RecNum = " & recordnumber.ToString, myconn)
        ds = New DataSet
        da.Fill(ds)

        Try
            If Not code = "" Then
                getCode(code)
            Else
                getCode(ds.Tables(0).Rows(0).Item("code").ToString)
            End If
            If ds.Tables(0).Rows.Count > 0 Then
                Dim temp As Integer = 1
                For i As Integer = 1 To DataGridView1.Rows.Count - 1
                    DataGridView1.Rows(i).Cells(1).Value = ds.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value).ToString
                Next
            End If
            If Not mymode = mode.entrymode And DataGridView1.Rows(0).Cells(1).Value = "" Then
                setCurrentRow(getCurrentRow) 'DataGridView1.Rows(currentRows).Selected = True
                txtValue.Text = DataGridView1.Rows(currentRows).Cells(1).Value
                lblFieldName.Text = DataGridView1.Rows(currentRows).Cells(0).Value
            Else
                setCurrentRow(0) 'DataGridView1.Rows(currentRows).Selected = True
                lblFieldName.Text = DataGridView1.Rows(0).Cells(0).Value
                txtValue.Text = DataGridView1.Rows(currentRows).Cells(1).Value
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try
    End Sub
    Public Enum mode
        browsemode = 0
        entrymode = 1
        editmode = 2
    End Enum
    Public Sub changemode()
        Select Case mymode
            Case mode.browsemode
                BROWSEMODEToolStripMenuItem.Enabled = False
                SAVEToolStripMenuItem.Enabled = False
                EDITRECORDToolStripMenuItem.Enabled = True
                STARTENTRYToolStripMenuItem.Enabled = True
                GOTOToolStripMenuItem.Enabled = True
                OPENToolStripMenuItem.Enabled = True
                lblMode.Text = "BROWSE MODE"
                txtValue.Enabled() = False
                Me.Focus()
                checkFlag()
                lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value 'lblField.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value
            Case mode.editmode
                EDITRECORDToolStripMenuItem.Enabled = False
                STARTENTRYToolStripMenuItem.Enabled = False
                GOTOToolStripMenuItem.Enabled = False
                OPENToolStripMenuItem.Enabled = False
                BROWSEMODEToolStripMenuItem.Enabled = True
                SAVEToolStripMenuItem.Enabled = True
                lblMode.Text = "EDIT MODE"
                txtValue.Enabled() = True
                txtValue.Focus()
                Dim bool As Boolean = False
                bool = checkFlag()
                If bool = False Then
                    valBox2("Record Has Not Been Keyed")
                    mymode = mode.browsemode
                    changemode()
                End If
                getDesc()
            Case mode.entrymode
                EDITRECORDToolStripMenuItem.Enabled = False
                STARTENTRYToolStripMenuItem.Enabled = False
                GOTOToolStripMenuItem.Enabled = False
                OPENToolStripMenuItem.Enabled = False
                BROWSEMODEToolStripMenuItem.Enabled = True
                SAVEToolStripMenuItem.Enabled = True
                lblMode.Text = "ENTRY MODE"
                txtValue.Enabled() = True
                 txtValue.Focus()
                selectTop()
                checkFlag()
                setCurrentRow(0)
                getDesc()
                lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
        End Select
    End Sub  
    Private Sub OPENToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPENToolStripMenuItem.Click

        If dbLoc = "" Then
openagain:  OpenFileDialog1.FileName = "00000001"
            ask = OpenFileDialog1.ShowDialog()
            dbLoc = OpenFileDialog1.FileName
        End If
        Dim val As New DirectoryInfo(OpenFileDialog1.FileName())
        Dim parent As String = val.Parent.Name

        If ask = Windows.Forms.DialogResult.OK Then
            If Not parent = "Entry1" Then
                valBox2("Please Select Database From Entry1 Folder")
                GoTo openAgain
            Else
                Label1.Text = Label1.Text & " - " & dbLoc
                imagePath = val.Parent.Parent.FullName & "\Images\"
                openmdb(dbLoc)
                filldtg()
                loadDataFields()
                getImages()
                mymode = mode.browsemode
                changemode()
            End If
        End If
        dbLoc = ""
    End Sub
    Public Function getImages()
        getImages = ""
        Dim cmd As New OleDbCommand("SELECT OImage001, Image001 FROM " & jobName & " WHERE RecNum = " & recordnumber, myconn)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        Do While (reader.Read())
            img1 = reader("Image001")
            OImg = reader("OImage001")
         Loop
        imgvwer.ImagePath = imagePath & img1
        txtCrntImg.Text = img1
        reader.Close()
    End Function
    Public Function getRec()
        Try
            recordnumber = mdbDS.Tables(0).Rows(curRec).Item("RecNum")
        Catch ex As Exception
        End Try
        Return recordnumber
    End Function
    Public Function countImg(ByVal tbl As DataSet)
        Dim recnum As Integer
        recnum = tbl.Tables(0).Rows.Count
        Return recnum
    End Function

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If mymode = mode.browsemode Then
         Else
            Dim bool2 As Boolean = False
            DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If Not DataGridView1.Rows(i).Cells(0).Value = "receive Sample" Then
                    If Not DataGridView1.Rows(i).Cells(1).Value Is ds.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value) Then
                        bool2 = True
                    End If
                Else
                    If Not DataGridView1.Rows(i).Cells(1).Value Is ds.Tables(0).Rows(0).Item("receiveSamples") Then
                        bool2 = True
                    End If
                End If
            Next
            If bool2 = True Then
                Dim bool As Boolean
                bool = valBox("Do You Want To Save??")
                If bool = True Then
                    saveData()
                  End If
             End If
        End If
    End Sub
    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim crntRow As Integer
        If e.KeyCode = Keys.Up Then
            If Not getCurrentRow() = 0 Then
                DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
                setCurrentRow(getCurrentRow() - 1)
                getDesc()
                lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
                txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(1).Value
                txtValue.Focus()
                txtValue.SelectAll()
                e.Handled = True
            End If
        End If
        If mymode = mode.browsemode Then
            If e.KeyCode = Keys.Enter Then
                If Not getCurrentRow() = DataGridView1.Rows.Count Then
                    setCurrentRow(getCurrentRow() + 1)
                    getDesc()
                    lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
                    txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(1).Value
                    txtValue.Focus()
                    txtValue.SelectAll()
                End If
            End If
        End If

        If mymode = mode.browsemode Then
            'move to previous entry
            'function: move left
            'validations: If record 1 as equals to one then exit sub
            If e.KeyCode = Keys.Left Then
                If mymode = mode.browsemode Then
                    If txtCrntRec.Text = 1 Then
                        Exit Sub
                    Else
                        curRec -= 1
                        txtCrntRec.Text = curRec + 1
                        crntRow = getCurrentRow()
                        loadDataFields()
                        getImages()
                        setCurrentRow(crntRow)
                        checkFlag()
                        getDesc()
                        Application.DoEvents()
                    End If
                End If
            End If
            If e.KeyCode = Keys.Right Then
                If mymode = mode.browsemode Then
                    'move to next entry
                    'function: move right
                    'validations: If record 1 as equals to one then exit sub
                    If txtCrntRec.Text = txtTotalRec.Text Then
                        Exit Sub
                    Else
                        curRec += 1
                        txtCrntRec.Text = curRec + 1
                        crntRow = getCurrentRow()
                        loadDataFields()
                        getImages()
                        setCurrentRow(crntRow)
                        checkFlag()
                        getDesc()
                        Application.DoEvents()
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error Resume Next
        Dialog2.Close()
        If myconn.State = ConnectionState.Open Then
            myconn.Close()
        End If
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If


    End Sub
    Private Sub GOTOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GOTOToolStripMenuItem.Click
        dataDialog.ShowDialog()
    End Sub
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
         DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        OpenFileDialog1.FileName = "00000001"
        ask = OpenFileDialog1.ShowDialog
        dbLoc = OpenFileDialog1.FileName
        OPENToolStripMenuItem_Click(New Object, e)
    End Sub

    Private Sub txtValue_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtValue.KeyDown
        valid = False
        If txtValue.MaxLength > 1 Then
            If e.KeyCode = Keys.Enter Then
                Select Case lblFieldName.Text
                    Case "address"
                        valid = validation.address
                    Case "include"
                        valid = validation.include
                    Case "city"
                        valid = validation.city()
                    Case "commonReason", "visit", "issue", "milk", "topic", "healthCare", "include1", "include2", "include3", "include4", "include5"
                        valid = validation.commonReason()
                    Case "dob", "dob1", "dob2", "dob3", "edd"
                        valid = CheckDate(txtValue.Text, "MMddyy")
                    Case "emailadd"
                        valid = validation.email()
                    Case "fname", "suffix", "other"
                        valid = validation.fname()
                    Case "lname"
                        valid = validation.lname()
                    Case "mname"
                        valid = validation.mname()
                    Case "phone"
                        valid = validation.phone()
                    Case "state"
                        valid = validation.state()
                    Case "title"
                        valid = validation.title
                    Case "zipcode"
                        valid = validation.zipCode()
                End Select
                If getCurrentRow() = DataGridView1.Rows.Count - 2 Then
                    DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
                    saveData()
                End If
            Else
                Select Case lblFieldName.Text
                    Case "dob1", "dob2", "dob3"
                        If e.KeyCode = Keys.F6 Then
                            For i As Integer = getCurrentRow() To DataGridView1.Rows.Count - 1
                                If Not DataGridView1.Rows(i).Cells(0).Value = "dob3" Then
                                    DataGridView1.Rows(i).Cells(1).Value = ""
                                ElseIf DataGridView1.Rows(i).Cells(0).Value = "dob3" Then
                                    DataGridView1.Rows(i).Cells(1).Value = ""
                                    valid = False
                                    setCurrentRow(i + 1)
                                    getDesc()
                                    lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
                                    txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(1).Value
                                    txtValue.Focus()
                                    txtValue.SelectAll()
                                    Exit Sub
                                End If
                            Next
                        End If
                End Select
            End If
        ElseIf txtValue.MaxLength = 1 Then
            Select Case lblFieldName.Text
                Case "Code"
                    Select Case Chr(e.KeyCode)
                        Case "E", "F", "G", "H", "I", "J", "K", "L"
                            e.Handled = True
                            DataGridView1.Rows(getCurrentRow).Cells(1).Value = e.KeyCode.ToString
                            loadDataFields(DataGridView1.Rows(getCurrentRow).Cells(1).Value)
                            valid = True
                    End Select
                    If e.KeyCode = Keys.Enter Then
                        If txtValue.Text = "*" Then
                            Dim ask As Boolean
                            ask = valBox("Are You Sure You Want To Pullout??")
                            If ask = True Then
                                For i As Integer = 0 To DataGridView1.Rows.Count - 3
                                    DataGridView1.Rows(i).Cells(1).Value = ""
                                Next
                                If mymode = mode.editmode Then
                                    Dim ask2 As Boolean
                                    ask2 = valBox("Save Changes??")
                                    If ask2 = True Then
                                        valid = False
                                        DataGridView1.Rows(0).Cells(1).Value = "*"
                                        saveData()
                                        mymode = mode.browsemode
                                        changemode()
                                    ElseIf ask2 = False Then
                                        valid = False
                                        mymode = mode.browsemode
                                        changemode()
                                    End If
                                ElseIf mymode = mode.entrymode Then
                                    valid = False
                                    DataGridView1.Rows(0).Cells(1).Value = "*"
                                    saveData()
                                End If
                                getDesc()
                                Exit Sub
                            ElseIf ask = False Then
                                valid = False
                                Exit Sub
                                txtValue.Text = ""
                            End If
                        ElseIf txtValue.Text = "" Then
                            valBox2("Code Cannot Be Blank")
                            valid = False
                            getDesc()
                            Exit Sub
                        End If
                        setCurrentRow(getCurrentRow() + 1)
                        getDesc()
                        lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
                        txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(1).Value
                        txtValue.Focus()
                        txtValue.SelectAll()
                    End If
                Case "qasFlag"
                    valid = validation.qasflag(e)
                Case "advertisement", "1stBaby", "HeartSmart", "hibiclens", "qasFlag", "adt", "advetori", "axa", "BOBA", "consider", "content", "coupon", "dhaO3", "editorial", "factsUp", "financialfuture", "hemange", "hemangeolC", "influence", "makeAPurchase", "notice", "offer", "receiveSample", "receiveSamples", "review", "review2", "stonyfie"
                    valid = True
                Case "healthCare", "receive Sample"
                    valid = True
                Case "title"
                    valid = True
                Case "gender1", "gender2", "gender3"
                    If e.KeyCode = Keys.F6 Then
                        For i As Integer = getCurrentRow() To DataGridView1.Rows.Count - 1
                            If Not DataGridView1.Rows(i).Cells(0).Value = "dob3" Then
                                DataGridView1.Rows(i).Cells(1).Value = ""
                            ElseIf DataGridView1.Rows(i).Cells(0).Value = "dob3" Then
                                DataGridView1.Rows(i).Cells(1).Value = ""
                                valid = False
                                setCurrentRow(i + 1)
                                getDesc()
                                lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
                                txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(1).Value
                                txtValue.Focus()
                                txtValue.SelectAll()
                                Exit Sub
                            End If
                        Next
                    End If
                    valid = True
            End Select
        End If
    End Sub
  


    Private Sub ChangeIDToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeIDToolStripMenuItem.Click
        Dialog2.ShowDialog()
    End Sub
    Private Sub ROTATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ROTATEToolStripMenuItem.Click
        imgvwer.RotateFlip(RotateFlipType.Rotate90FlipXY)
        Application.DoEvents()
    End Sub
    Private Sub ZOOMINToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZOOMINToolStripMenuItem.Click
        imgvwer.ZoomIn()
    End Sub
    Private Sub ZOOMOUTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZOOMOUTToolStripMenuItem.Click
        imgvwer.ZoomOut()
    End Sub
    Private Sub MOVEIMAGETOLEFTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MOVEIMAGETOLEFTToolStripMenuItem.Click
        imgvwer.MoveLeft()
    End Sub
    Private Sub MOVEIMAGETORIGHTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MOVEIMAGETORIGHTToolStripMenuItem.Click
        imgvwer.MoveRight()
    End Sub
    Private Sub MOVEIMAGEUPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MOVEIMAGEUPToolStripMenuItem.Click
        imgvwer.MoveUp()
    End Sub
    Private Sub MOVEIMAGEDOWNToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MOVEIMAGEDOWNToolStripMenuItem.Click
        imgvwer.MoveDown()
    End Sub

    Private Sub txtValue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtValue.KeyPress
        If e.KeyChar = ChrW(8) Then
            e.Handled = False
            Exit Sub
        End If

          Select Case lblFieldName.Text
            Case "fname", "lname", "mname", "other", "city"
                e.Handled = alphaOrNum(e, 1, , "-'")
            Case "suffix"
                e.Handled = alphaOrNum(e, 1)
            Case "advertisement", "1stBaby", "HeartSmart", "gender1", "gender2", "gender3", "hibiclens", "adt", "advetori", "axa", "BOBA", "consider", "content", "coupon", "dhaO3", "editorial", "factsUp", "financialfuture", "hemange", "hemangeolC", "influence", "makeAPurchase", "notice", "offer", "receiveSample", "receiveSamples", "review", "review2", "stonyfie"
                e.Handled = alphaOrNum(e, 2, "12")
            Case "qasFlag"
                e.Handled = alphaOrNum(e, 2)
            Case "zipcode", "dob", "dob1", "dob2", "dob3", "include"
                e.Handled = alphaOrNum(e, 0)
            Case "receive Sample"
                e.Handled = alphaOrNum(e, 2, "1")
            Case "phone"
                e.Handled = alphaOrNum(e, 0, , "-()& ")
            Case "gender1", "gender2", "gender3", "include5"
                e.Handled = alphaOrNum(e, 2, "12")
            Case "commonReason", "visit"
                e.Handled = alphaOrNum(e, 2, "12345")
            Case "healthCare", "include1", "issue"
                e.Handled = alphaOrNum(e, 2, "1234")
            Case "include2", "include3", "include4"
                e.Handled = alphaOrNum(e, 2, "123")
            Case "milk", "topic"
                e.Handled = alphaOrNum(e, 2, "1234567")
            Case "title"
                e.Handled = validation.title(e, 1)
                e.Handled = alphaOrNum(e, 1, , "*")
            Case "Code"
                e.Handled = alphaOrNum(e, 2, "EFGHIJKLefghijkl*")
                If e.KeyChar = "*" Then
                    Dim ask As Boolean
                    txtValue.Text = e.KeyChar
                    ask = valBox("Are You Sure You Want To Pullout??")
                    If ask = True Then
                        For i As Integer = 0 To DataGridView1.Rows.Count - 3
                            DataGridView1.Rows(i).Cells(1).Value = ""
                        Next

                        If mymode = mode.editmode Then
                            Dim ask2 As Boolean
                            ask2 = valBox("Save Changes??")
                            If ask2 = True Then
                                e.Handled = True
                                DataGridView1.Rows(0).Cells(1).Value = "*"
                                saveData()
                                mymode = mode.browsemode
                                changemode()
                            ElseIf ask2 = False Then
                                e.Handled = True
                                mymode = mode.browsemode
                                changemode()
                            End If
                        ElseIf mymode = mode.entrymode Then
                            e.Handled = True
                            DataGridView1.Rows(0).Cells(1).Value = "*"
                            saveData()
                        End If
                    ElseIf ask = False Then
                        e.Handled = True
                        txtValue.Text = ""
                    End If
                End If
            Case "emailadd"
                e.Handled = alphaOrNum(e, 3, , "@&()_'./")
            Case "address"
                e.Handled = alphaOrNum(e, 3, , "#&()-'")
        End Select
         If txtValue.MaxLength > 1 Then
            If e.Handled = False Then
                If Not lblFieldName.Text = "emailadd" Then
                    If Char.IsLower(e.KeyChar) Then
                        txtValue.SelectedText = Char.ToUpper(e.KeyChar)
                        e.Handled = True
                    End If
                End If
            End If
            If e.KeyChar = ChrW(10) Or e.KeyChar = ChrW(11) Or e.KeyChar = ChrW(12) Or e.KeyChar = ChrW(13) Then
                If valid = True Then
                    DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
                    If lblFieldName.Text = "include5" Then
                        If Not txtValue.Text.Contains("3") Then
                            setCurrentRow(getCurrentRow() + 1)
                        End If
                    End If
                    nextLine()
                End If
                e.Handled = True
            End If
        ElseIf txtValue.MaxLength = 1 Then
            If e.KeyChar = ChrW(8) Then
                e.Handled = False
            End If
            If e.Handled = False And valid = True Then
                e.Handled = True
                If e.KeyChar = ChrW(10) Or e.KeyChar = ChrW(11) Or e.KeyChar = ChrW(12) Or e.KeyChar = ChrW(13) Then
                Else
                    txtValue.Text = Char.ToUpper(e.KeyChar)
                End If
                If getCurrentRow() = DataGridView1.Rows.Count - 3 Then
                    DataGridView1.Rows(getCurrentRow).Cells(1).Value = Char.ToUpper(e.KeyChar)
                    saveData()
                Else
                    DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
                    nextLine()
                End If
            End If
        End If
    End Sub

    Public Function insert()
        Dim cmd As New OleDb.OleDbCommand
        'code by erwin
        Dim sql As String = "update Data001 set "
        For i As Integer = 0 To DataGridView1.Rows.Count - 3
            If DataGridView1.Rows(i).Cells(1).Value IsNot Nothing And DataGridView1.Rows(i).Cells(1).Value <> "" Then
                If i = DataGridView1.Rows.Count - 1 Then
                    sql = sql & "[" & DataGridView1.Rows(i).Cells(0).Value.ToString & "]" & " = '" & DataGridView1.Rows(i).Cells(1).Value.ToString.Replace("'", "''") & "'"
                Else
                    sql = sql & "[" & DataGridView1.Rows(i).Cells(0).Value.ToString & "]" & " = '" & DataGridView1.Rows(i).Cells(1).Value.ToString.Replace("'", "''") & "',"
                End If
            End If
        Next
        If DataGridView1.Rows.Count = 1 Then
            sql = sql & "[" & DataGridView1.Rows(0).Cells(0).Value.ToString & "]" & " = '" & DataGridView1.Rows(0).Cells(1).Value.ToString.Replace("'", "''") & "', "
        End If
        sql = sql.Replace("receive Sample", "receiveSamples")
        sql = sql & "[" & "dataCaptu" & "]" & " = '" & datetime.Replace("'", "''") & "',"
        sql = sql & "[" & "image" & "]" & " = '" & OImg.Replace("'", "''") & "'"
        sql = sql + " where recNum = " & recordnumber
        cmd = New OleDbCommand(sql, myconn)
        cmd.ExecuteNonQuery()
        'end code
        sqlQry = "Update " & jobName & " Set Keyflag = @keyflag, KeyID = @keyid, KeyDate = @keydate, LocID1 = @locid1  where RecNum=@recnum"
        cmd = New OleDbCommand(sqlQry, myconn)
        cmd.Parameters.AddWithValue("@keyflag", "^")
        cmd.Parameters.AddWithValue("@keyid", txtUserID.Text)
        cmd.Parameters.AddWithValue("@keydate", datetime)
        cmd.Parameters.AddWithValue("@keydate", txtLocID.Text)
        cmd.Parameters.AddWithValue("@recnum", recordnumber)
        cmd.ExecuteNonQuery()
        Return curRec
    End Function
    
    Public Function editRecord()
        Dim cmd As New OleDb.OleDbCommand
        'code by erwin
        Dim sql As String = "update Data001 set "
        For i As Integer = 0 To DataGridView1.Rows.Count - 3
            If DataGridView1.Rows(i).Cells(1).Value IsNot Nothing And DataGridView1.Rows(i).Cells(1).Value <> "" Then
                If i = DataGridView1.Rows.Count - 3 Then
                    sql = sql & "[" & DataGridView1.Rows(i).Cells(0).Value.ToString & "]" & " = '" & DataGridView1.Rows(i).Cells(1).Value.ToString.Replace("'", "''") & "' "
                Else
                    sql = sql & "[" & DataGridView1.Rows(i).Cells(0).Value.ToString & "]" & " = '" & DataGridView1.Rows(i).Cells(1).Value.ToString.Replace("'", "''") & "', "
                End If
            Else
                If i = DataGridView1.Rows.Count - 3 Then
                    sql = sql.Remove(sql.Length - 2, 2) & " "
                End If
            End If
        Next
        'sql = sql & "[" & "dataCaptu" & "]" & " = '" & datetime.Replace("'", "''") & "',"
        'sql = sql & "[" & "image" & "]" & " = '" & img1.Replace("'", "''") & "'"
        sql = sql.Replace("receive Sample", "receiveSamples")
        sql = sql + " where RecNum = " & recordnumber
        cmd = New OleDbCommand(sql, myconn)
        cmd.ExecuteNonQuery()
        'end code
        Dim updFlagCnt As Integer
        sqlQry = "SELECT Updflag FROM " & jobName & " WHERE RecNum = " & recordnumber
        cmd = New OleDbCommand(sqlQry, myconn)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        Do While (reader.Read())
            updFlagCnt = CInt(Int(reader("Updflag")))
        Loop
        reader.Close()
        updFlagCnt += 1
        sqlQry = "Update " & jobName & " Set Updflag = @keyflag, UpdID = @keyid, UpdDate = @keydate  where RecNum=@recnum"
        cmd = New OleDbCommand(sqlQry, myconn)
        cmd.Parameters.AddWithValue("@keyflag", updFlagCnt)
        cmd.Parameters.AddWithValue("@keyid", txtUserID.Text)
        cmd.Parameters.AddWithValue("@keydate", datetime)
        cmd.Parameters.AddWithValue("@recnum", recordnumber)
        cmd.ExecuteNonQuery()
        Return True
    End Function
    Public Function checkFlag()
        checkFlag = ""
        Dim cmd As New OleDb.OleDbCommand
        cmd = New OleDbCommand("SELECT KeyID FROM " & jobName & " Where recNum = @rec and KeyFlag = '^'", myconn)
        cmd.Parameters.Add(New OleDbParameter("@rec", recordnumber))
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        Do While (reader.Read())
            dataBy = reader("KeyID").ToString
            If Not dataBy = "" Then
                reader.Close()
                txtDataBy.Text = dataBy
                Return True
                Exit Do
            Else
                txtDataBy.Text = dataBy
                reader.Close()
                Return False
            End If
        Loop
    End Function

    Public Function selectTop(Optional ByVal frombrowse As Boolean = False) As Integer
        Dim ds2 As New DataSet
        Dim da2 As New OleDbDataAdapter
        Dim temp As Integer

        Try
            ds2 = New DataSet
            da2 = New OleDbDataAdapter(IIf(frombrowse, "SELECT TOP 1 RecNum FROM " & jobName & " order by recnum", "SELECT TOP 1 RecNum FROM " & jobName & " WHERE  KeyFlag = '' OR ISNULL(KeyFlag)"), myconn)
            da2.Fill(ds2)
            temp = ds2.Tables(0).Rows(0).Item(0)
            For i As Integer = 0 To mdbDS.Tables(0).Rows.Count - 1
                If mdbDS.Tables(0).Rows(i).Item(0) = temp Then
                    curRec = i
                    txtCrntRec.Text = curRec + 1
                    Exit For
                End If
            Next
            loadDataFields()
            getImages()
        Catch ex As Exception
            ' MsgBox(ex.Message)
            valBox2("All data has been keyed!")
            mymode = mode.browsemode
            changemode()
        End Try
        Return True
    End Function
    Public Function saveData()
        saveData = ""
        If mymode = mode.entrymode Then
            If txtCrntRec.Text = txtTotalRec.Text Then
                insert()
                txtValue.Text = ""
                valBox2("All record Has Been Keyed")
                mymode = mode.browsemode
                changemode()
            Else
                insert()
                txtValue.Text = ""
                curRec += 1
                txtCrntRec.Text = curRec + 1
                loadDataFields()
                getImages()
                lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
                txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(1).Value
                getDesc()
            End If
        ElseIf mymode = mode.editmode Then
            editRecord()
            txtValue.Text = ""
            valBox2("Record Has Been Saved!")
            mymode = mode.browsemode
            changemode()
        End If
    End Function
   
    Private Sub STARTENTRYToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles STARTENTRYToolStripMenuItem.Click
        mymode = mode.entrymode
        changemode()

    End Sub

    Private Sub EDITRECORDToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EDITRECORDToolStripMenuItem.Click
        mymode = mode.editmode
        changemode()
    End Sub
    Private Sub BROWSEMODEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BROWSEMODEToolStripMenuItem.Click
        SAVEToolStripMenuItem_Click(New Object, e)
        mymode = mode.browsemode
        changemode()
    End Sub

    Private Sub DataGridView1_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.CurrentCellChanged
        If mymode = mode.browsemode Then
            Try
                lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
                txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(1).Value
                getDesc()

            Catch ex As Exception
                ' MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub EXITToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EXITToolStripMenuItem.Click
        If mymode = mode.browsemode Then
            Me.Close()
        Else
            Dim bool2 As Boolean = False
            DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If Not DataGridView1.Rows(i).Cells(1).Value Is ds.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value) Then
                    bool2 = True
                End If
            Next
            If bool2 = True Then
                Dim bool As Boolean
                bool = valBox("Do You Want To Save??")
                If bool = True Then
                    saveData()
                    Me.Close()
                Else
                    Me.Close()
                End If
            Else
                Me.Close()
            End If
        End If
    End Sub

    Private Sub SAVEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SAVEToolStripMenuItem.Click
        Dim bool2 As Boolean = False
        DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
        For i As Integer = 0 To DataGridView1.Rows.Count - 3
            If Not DataGridView1.Rows(i).Cells(0).Value = "receive Sample" Then
                If Not DataGridView1.Rows(i).Cells(1).Value Is ds.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value) Then
                    bool2 = True
                End If
            Else
                If Not DataGridView1.Rows(i).Cells(1).Value Is ds.Tables(0).Rows(0).Item("receiveSamples") Then
                    bool2 = True
                End If
            End If
        Next
        If bool2 = True Then
            Dim bool As Boolean
            bool = valBox("Do You Want To Save??")
            If bool = True Then
                saveData()
                mymode = mode.browsemode
                changemode()
            Else
                mymode = mode.browsemode
                changemode()
            End If
        Else
            mymode = mode.browsemode
            changemode()
        End If
    End Sub

    Public Function checkSpChar(ByVal text As String, ByVal additionalValidChars As String) As Boolean
        text = text.ToUpper
        Dim alphaNum As String = "ABCEFGHIJKLMNOPQRSTUVWXYZ0123456789" & additionalValidChars
        Dim match As Integer = 0
        For i As Integer = 0 To text.Length - 1
            If alphaNum.Contains(text.Substring(i, 1)) Then
                match += match
            End If
        Next
        If match = text.Length Then
            Return False
        Else
            Return True
        End If

    End Function
    Public Function checklength(ByVal textlength As String, ByVal validlength As String) As Boolean
        If Not textlength = checklength Then
            Dim bool As Boolean = valBox("Invalid Length" & vbNewLine & "Do you want to Proceed??")
            If bool = True Then
                Return False
            Else
                Return True
            End If
        Else
            Return False
        End If
    End Function
    'Public Function checkDate(ByVal valstr As String) As Boolean
    '    'Dim mm As String
    '    'Dim dd As String
    '    'Dim yy As String

    '    Dim date1 As Date
    '    If Not Date.TryParse(valstr, date1) Then
    '        valBox2("Invalid Date")
    '    End If
    'End Function
    Public Function checkBlank(ByVal text As String) As Boolean
        If text = "" Or text Is Nothing Then
            Dim bool As Boolean = valBox("Text is Blank" & vbNewLine & "Do you want to Proceed??")
            If bool = True Then
                Return False
            Else
                Return True
            End If
        Else
            Return False
        End If
       
    End Function
    Public Function alphaOrNum(ByVal e As KeyPressEventArgs, ByVal type As Integer, Optional ByVal pick As String = "ynYN", Optional ByVal additionalChar As String = "")
        ''[0] - Numeric
        ''[1] - Alpha
        ''[2] - Specify
        ''[3] - All
        Dim alpha As String = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & Chr(10) & Chr(11) & Chr(12) & Chr(13) & additionalChar
        Dim num As String = "0123456789" & Chr(10) & Chr(11) & Chr(12) & Chr(13) & additionalChar
        Dim sp As String = pick & Chr(10) & Chr(11) & Chr(12) & Chr(13) & additionalChar
        Dim all As String = "0123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & Chr(10) & Chr(11) & Chr(12) & Chr(13) & additionalChar
        Select Case type
            Case 0
                If num.Contains(e.KeyChar) = True Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case 1
                If alpha.Contains(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case 2
                If sp.Contains(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case 3
                If all.Contains(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
        End Select
        Return e.Handled
    End Function

    Public Function nextLine()
        nextLine = ""
        setCurrentRow(getCurrentRow() + 1)
        getDesc()
        lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
        txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(1).Value
        txtValue.Focus()
        txtValue.SelectAll()
    End Function
    Public Function CheckDate(ByVal strval As String, ByVal formats As String, Optional ByVal Rng As Integer = 100)
        If txtValue.Text = "" Then
            Return True
        End If
        Dim isValOk As String = ""
        Dim yyyy As String = ""

        Dim dd As String = ""
        Dim mm As String = ""
        Dim yy As String = ""

        If strval.Length = 6 Then
            mm = strval.Substring(0, 2)
            dd = strval.Substring(2, 2)
            yy = strval.Substring(4, 2)

            If (dd.Trim.Equals("") And mm.Trim.Equals("") And yy.Trim.Equals("")) Or _
              (Val(dd) = 0 Or Val(mm) = 0 Or Val(yy) = 0) Then
                valBox2("Date is invalid.")
                isValOk = False
                Return False
            Else
                'Month
                If Val(mm) > 12 Then
                    valBox2("Invalid Month Detected.")
                    isValOk = False
                    Return False
                Else
                    'Day
                    Select Case Val(mm)
                        Case 1, 3, 5, 7, 8, 10, 12
                            If Val(dd) > 31 Then
                                valBox2("Invalid Day Detected.")
                                isValOk = False
                                Return False
                            End If
                        Case 2
                            Dim leaf = Val(mm) Mod 4
                            If leaf = 0 Then
                                If Val(dd) > 29 Then
                                    valBox2("Invalid Day Detected.")
                                    isValOk = False
                                    Return False
                                End If
                            Else
                                If Val(dd) > 28 Then
                                    valBox2("Invalid Day Detected.")
                                    isValOk = False
                                    Return False
                                End If
                            End If
                        Case Else
                            If Val(dd) > 30 Then
                                valBox2("Invalid Day Detected.")
                                isValOk = False
                                Return False
                            End If
                    End Select
                    'Year
                    If yy > Val(Now.ToString("yy")) Then
                        yyyy = "19" & yy
                    Else
                        yyyy = "20" & yy
                    End If
                    If Val(yyyy) <= Val(Now.ToString("yyyy") - Rng) Then
                        If Val(yyyy) = Val(Now.ToString("yyyy") - Rng) Then
                            If Val(mm) >= Val(Now.ToString("MM")) Then
                            End If
                            If Val(mm) = Val(Now.ToString("MM")) Then
                            End If
                            If Val(dd) > Val(Now.ToString("dd")) Then
                            End If
                        End If
                        ask = MsgBox("Age is Less than " & Rng & "...do you want to overwrite???", _
                                MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, _
                                "Date")
                        If ask = MsgBoxResult.Yes Then
                            isValOk = True
                            Return True
                        Else
                            isValOk = False
                            Return False
                        End If
                    ElseIf mm.Trim <> "" And Val(mm) <> 0 And dd.Trim <> "" And Val(dd) <> 0 Then
                        Dim date1 As Date = Date.ParseExact(Format(Now, "MMddyy"), "MMddyy", Nothing)
                        Dim date2 As Date = Date.ParseExact(Val(mm & dd & yy), "MMddyy", Nothing)
                        If date2.Date > date1.Date Then
                            valBox2("Future Date Detected")
                            isValOk = False
                            Return False
                        End If
                    End If
                End If
                    End If
        Else
                    valBox2("Invalid date format")
                    isValOk = False
                    Return False
        End If
        Return True
    End Function

  
    Private Sub txtValue_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtValue.KeyUp
       
    End Sub
End Class
