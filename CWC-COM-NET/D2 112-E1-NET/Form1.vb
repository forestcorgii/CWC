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
                    Return True
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
            If Not Form1.txtValue.Text.Length = 10 Then
                ask = Form1.valBox("Character Length is not Equals 10 " & vbNewLine & "Do you want to proceed", "Error")
                If ask = False Then
                    Return False
                ElseIf ask = True Then
                    Return True
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
                                Form1.curRec += 1
                                Form1.txtCrntRec.Text += 1
                                Form1.loadDataFields()
                                Form1.getImages()
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
                Return True
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
                Return True
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
            End Select
        Next
        If a > 1 Or b > 1 Or c > 1 Or d > 1 Or e > 1 Then
            Form1.valBox2("Contains Duplicate")
            Return False
        Else
            Return True
        End If
    End Function
End Structure
Public Structure comLog
    Private bool As Boolean
    Public Function addLength(ByVal str As String, ByVal len As Integer) As String
        For i As Integer = str.Length To len - 1
            str &= " "
        Next
        Return str
    End Function
    Public Function startLog(ByVal logPath As String, ByVal text As String)
        startLog = ""
        If Not My.Computer.FileSystem.FileExists(logPath) Then
            Dim logStr As String
            text = addLength(text, 196)
            logStr = "┌────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐" & vbNewLine & _
                     "│" & text & "│" & vbNewLine & _
                     "└────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘"
            My.Computer.FileSystem.WriteAllText(logPath, logStr, False)
        End If
    End Function
    Public Function ifCompareStarted(ByVal logPath As String, ByVal text As String)
        ifCompareStarted = ""
        If My.Computer.FileSystem.FileExists(logPath) Then
            Dim log As String = My.Computer.FileSystem.ReadAllText(logPath)
            Dim logStr As String
            text = addLength(text, 196)
            logStr = log & vbNewLine & "┌────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐" & vbNewLine & _
                                       "│" & text & "│" & vbNewLine & _
                                       "├─────────┬──────────────────────────┬────────────────────────────────────────────┬────────────────────────────────────────────┬────────────────────────────────────────────┬────────────────────────┤" & vbNewLine & _
                                       "│ Record  │ Field                    │ Entry1                                     │ Entry2                                     │ Compare                                    │ Result                 │" & vbNewLine & _
                                       "├─────────┼──────────────────────────┼────────────────────────────────────────────┼────────────────────────────────────────────┼────────────────────────────────────────────┼────────────────────────┤"

            My.Computer.FileSystem.WriteAllText(logPath, logStr, False)
        End If
    End Function
    Public Function ifMismatchAppear(ByVal logPath As String, ByVal record As String, ByVal field As String, ByVal entry1 As String, ByVal entry2 As String, ByVal compare As String, ByVal result As String)
        ifMismatchAppear = ""
        If My.Computer.FileSystem.FileExists(logPath) Then
            Dim log As String = My.Computer.FileSystem.ReadAllText(logPath)
            Dim logStr As String
            record = addLength(record, 9)
            field = addLength(field, 26)
            entry1 = addLength(entry1, 44)
            entry2 = addLength(entry2, 44)
            compare = addLength(compare, 44)
            result = addLength(result, 24)
            logStr = log & vbNewLine & "│" & record & "│" & field & "│" & entry1 & "│" & entry2 & "│" & compare & "│" & result & "│"
            My.Computer.FileSystem.WriteAllText(logPath, logStr, False)
        End If
    End Function
    Public Function compareFooter(ByVal logPath As String, ByVal text As String)
        compareFooter = ""

        If My.Computer.FileSystem.FileExists(logPath) Then
            Dim log As String = My.Computer.FileSystem.ReadAllText(logPath)
            Dim logStr As String
            text = addLength(text, 196)
            logStr = log & vbNewLine & "├─────────┴──────────────────────────┴────────────────────────────────────────────┴────────────────────────────────────────────┴────────────────────────────────────────────┴────────────────────────┤" & vbNewLine & _
                                       "│" & text & "│" & vbNewLine & _
                                       "└────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘"
            My.Computer.FileSystem.WriteAllText(logPath, logStr, False)
        End If
    End Function
    Public Function ifNoRecordsFound(ByVal logPath As String, ByVal text As String)
        ifNoRecordsFound = ""

        If My.Computer.FileSystem.FileExists(logPath) Then
            Dim log As String = My.Computer.FileSystem.ReadAllText(logPath)
            Dim logStr As String
            Dim text1 As String = ""
            Dim text3 As String = ""
            Dim text4 As String = ""
            Dim text5 As String = ""
            Dim text6 As String = ""

            text1 = addLength(text1, 9)
            text = addLength(text, 26)
            text3 = addLength(text1, 44)
            text4 = addLength(text1, 44)
            text5 = addLength(text1, 44)
            text6 = addLength(text1, 24)
            logStr = log & vbNewLine & "│" & text1 & "│" & text & "│" & text3 & "│" & text4 & "│" & text5 & "│" & text6 & "│"

            My.Computer.FileSystem.WriteAllText(logPath, logStr, False)
        End If
    End Function
End Structure

Public Class Form1
    Private mouseDownLoc As Point
    Private ComDa, E1Da, E2Da As OleDb.OleDbDataAdapter
    Private comDs, E1Ds, E2Ds As DataSet
    Private fieldDS As DataSet
    Private mdbDS As DataSet
    Private clientname As String = "CWC"
    Private jobName As String = "YYF_RSG"
    Private sqlQry As String = ""
    Private dataBy As String = ""
    Public mymode As mode
    Private recordnumber As Integer = 0
     Private dbLoc As String
    Private img1, OImg As String
    Public curRec As Integer = 0
    Private totalrec As Integer
    Private imagePath, comlogpath As String
    Private currentRows As Integer
    Private datetime As String = Now.ToString("G")
    Private valid, val2, comCnt2 As Boolean
    Private ask As DialogResult
    Private validation As New validation
    Private comLog As New comLog
    Private E1Loc, E2Loc, ComLoc As String
    Private e1Databy As String = ""
    Private e2Databy As String = ""
    Private comDataby As String = ""
    Private myComconn, myE1conn, myE2conn, conn As New OleDb.OleDbConnection
    Private currentRows2 As Integer


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
    Public Sub openComMdb(ByVal connectionpath As String)

        If Not myComconn Is Nothing Then
            If myComconn.State = ConnectionState.Open Then
                myComconn.Close()
            End If
        End If
        myComconn = New OleDbConnection
        myComconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & connectionpath
        myComconn.Open()
    End Sub
    Public Sub openE1Mdb(ByVal connectionpath As String)

        If Not myE1conn Is Nothing Then
            If myE1conn.State = ConnectionState.Open Then
                myE1conn.Close()
            End If
        End If
        myE1conn = New OleDbConnection
        myE1conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & connectionpath
        myE1conn.Open()

    End Sub
    Public Sub openE2Mdb(ByVal connectionpath As String)

        If Not myE2conn Is Nothing Then
            If myE2conn.State = ConnectionState.Open Then
                myE2conn.Close()
            End If
        End If
        myE2conn = New OleDbConnection
        myE2conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & connectionpath
        myE2conn.Open()

    End Sub
    Public Function filldtg()
        filldtg = ""
        mdbDS = New DataSet
        ComDa = New OleDbDataAdapter("select * from Data001", myComconn)
        ComDa.Fill(mdbDS)
        totalrec = mdbDS.Tables(0).Rows.Count
        txtTotalRec.Text = totalrec
        txtCrntRec.Text = curRec + 1
    End Function
    Public Function getCode(ByVal code As String)
        getCode = ""
        myComconn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "  Data Source=" & Application.StartupPath & "\Codings.mdb" & ";User Id=admin;Password=;")
        If Not myComconn.State = ConnectionState.Open Then
            myComconn.Open()
        End If
        fieldDS = New DataSet
        Dim da As New OleDbDataAdapter("Select * FROM [CWC_tbl] WHERE Code = '" & code & "'", myComconn)
        da.Fill(fieldDS)
        DataGridView1.Rows.Clear()
        For i As Integer = 0 To fieldDS.Tables(0).Columns.Count - 1
            If fieldDS.Tables(0).Rows(0).Item(i).ToString = "1" Then
                DataGridView1.Rows.Add(fieldDS.Tables(0).Columns(i).ColumnName)
            End If
        Next
        '  loadDataFields()
    End Function
    Public Function getDesc()
        getDesc = ""
        conn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "  Data Source=" & Application.StartupPath & "\Codings.mdb" & ";User Id=admin;Password=;")
        If Not myComconn.State = ConnectionState.Open Then
            conn.Open()
        End If
        Dim valDS As New DataSet
        Dim valDA As New OleDbDataAdapter("SELECT * FROM CWCval_tbl WHERE FieldName = '" & DataGridView1.Rows(getCurrentRow).Cells(0).Value & "'", conn)
        valDA.Fill(valDS)
        Try
            lblDesc.Text = valDS.Tables(0).Rows(0).Item("Description")
            txtValue.MaxLength = valDS.Tables(0).Rows(0).Item("MaxLength")
            txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(2).Value
            lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
        Catch ex As Exception
            ' MsgBox(ex.Message)
            lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value
            lblDesc.Text = ""
            txtValue.MaxLength = 60
            txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(2).Value
        End Try
       End Function
    Public Sub loadDataFields(Optional ByVal code As String = "")
        getRec()
        ComDa = New OleDbDataAdapter("select Code,  phone,emailadd, dataCaptu, image  from data001 where RecNum = " & recordnumber, myComconn)
        comDs = New DataSet
        ComDa.Fill(comDs)

        E1Da = New OleDbDataAdapter("select Code,  phone,emailadd, dataCaptu, image  from data001 where RecNum = " & recordnumber, myE1conn)
        E1Ds = New DataSet
        E1Da.Fill(E1Ds)

        E2Da = New OleDbDataAdapter("select Code,  phone,emailadd, dataCaptu, image  from data001 where RecNum = " & recordnumber, myE2conn)
        E2Ds = New DataSet
        E2Da.Fill(E2Ds)

        Try
            If DataGridView1.Rows.Count < 1 Then
                If comDs.Tables(0).Rows.Count > 0 Then
                    For i As Integer = 0 To comDs.Tables(0).Columns.Count - 1
                        DataGridView1.Rows.Add(comDs.Tables(0).Columns(i).ColumnName)
                    Next
                End If
            End If
            If comDs.Tables(0).Rows.Count > 0 Then
                checkFlag()
                DataGridView1.Columns(1).HeaderText = "Compared BY: " & comDataby
                DataGridView1.Columns(2).HeaderText = "Entry1 BY: " & e1Databy & " | Enter/F10 = E1 Data"
                DataGridView1.Columns(3).HeaderText = "Entry2 BY: " & e2Databy & " | F11 = E2 Data"
                e1Databy = ""
                e2Databy = ""
                comDataby = ""
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    DataGridView1.Rows(i).Cells(1).Value = comDs.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value).ToString
                    DataGridView1.Rows(i).Cells(2).Value = E1Ds.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value).ToString
                    DataGridView1.Rows(i).Cells(3).Value = E2Ds.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value).ToString
                Next
            End If
            If Not mymode = mode.entrymode And DataGridView1.Rows(0).Cells(1).Value = "" Then
                setCurrentRow(getCurrentRow) 'DataGridView1.Rows(currentRows).Selected = True
                txtValue.Text = DataGridView1.Rows(currentRows).Cells(1).Value
                lblFieldName.Text = DataGridView1.Rows(currentRows).Cells(0).Value
            Else
                setCurrentRow(0) 'DataGridView1.Rows(currentRows).Selected = True
                txtCrntImg.Text = DataGridView1.Rows(0).Cells(1).Value
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
        recompareAll = 3
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
                DataGridView1.Enabled = True
                DataGridView1.Focus()
                checkFlag()
                loadDataFields()
                getImages()
                lblFieldName.Text = DataGridView1.Rows(getCurrentRow).Cells(0).Value 'lblField.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value
            Case mode.editmode
                EDITRECORDToolStripMenuItem.Enabled = False
                STARTENTRYToolStripMenuItem.Enabled = False
                GOTOToolStripMenuItem.Enabled = False
                OPENToolStripMenuItem.Enabled = False
                BROWSEMODEToolStripMenuItem.Enabled = True
                SAVEToolStripMenuItem.Enabled = True
                lblMode.Text = "RECOMPARE MODE(SPECIFIC)"
                txtValue.Enabled() = True
                DataGridView1.Enabled = True
                txtValue.Focus()
                Dim bool As Boolean = False
                bool = checkFlag()
                If bool = False Then
                    valBox2("Record Has Not Been Keyed")
                    mymode = mode.browsemode
                    changemode()
                End If
            Case mode.entrymode
                EDITRECORDToolStripMenuItem.Enabled = False
                STARTENTRYToolStripMenuItem.Enabled = False
                GOTOToolStripMenuItem.Enabled = False
                OPENToolStripMenuItem.Enabled = False
                BROWSEMODEToolStripMenuItem.Enabled = True
                SAVEToolStripMenuItem.Enabled = True
                lblMode.Text = "COMPARE MODE"
                comLog.ifCompareStarted(comlogpath, "Compare Started At " & datetime & " By: " & txtUserID.Text)
                txtValue.Enabled() = True
                DataGridView1.Enabled = True
                txtValue.Focus()
                selectTop()
                setCurrentRow(0)
                recordnumber = compareEntry(getCurrentRow, 0)
                getDesc()
            Case mode.recompareAll
                EDITRECORDToolStripMenuItem.Enabled = False
                STARTENTRYToolStripMenuItem.Enabled = False
                GOTOToolStripMenuItem.Enabled = False
                OPENToolStripMenuItem.Enabled = False
                BROWSEMODEToolStripMenuItem.Enabled = True
                SAVEToolStripMenuItem.Enabled = True
                lblMode.Text = "RECOMPARE MODE(ALL)"
                txtValue.ReadOnly() = False
                recordnumber = getRec()
                loadDataFields()
                getImages()
                recordnumber = compareEntry(getCurrentRow(), 0)
                comLog.ifCompareStarted(comlogpath, "Recompare(ALL) Started At " & datetime & " By: " & txtUserID.Text)
                getDesc()
                txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(2).Value
                txtValue.Focus()
                txtValue.SelectAll()
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
            If Not parent = "Compare" Then
                valBox2("Please Select Database From Compare Folder")
                GoTo openAgain
            Else
                Dim child As String = val.Name
                Dim par As String = val.Parent.Parent.FullName
                E1Loc = par & "\" & "Entry1\" & child
                E2Loc = par & "\" & "Entry2\" & child
                ComLoc = OpenFileDialog1.FileName()
                openComMdb(ComLoc)
                openE1Mdb(E1Loc)
                openE2Mdb(E2Loc)
                Label1.Text = Label1.Text & " - " & dbLoc
                imagePath = val.Parent.Parent.FullName & "\Images\"
                filldtg()
                mymode = mode.browsemode
                changemode()
                comlogpath = val.Parent.FullName & "\" & Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName) & ".log"
                comLog.startLog(comlogpath, comlogpath & " Opened At " & datetime & " By: " & txtUserID.Text)
            End If
        End If
        dbLoc = ""
    End Sub
    Public Function getImages()
        getImages = ""
        Dim cmd As New OleDbCommand("SELECT OImage001, Image001 FROM " & jobName & " WHERE RecNum = " & recordnumber, myComconn)
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
                If Not DataGridView1.Rows(i).Cells(1).Value Is comDs.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value) Then
                    bool2 = True
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
    Public Function compareEntry(ByVal ndx As Integer, ByVal ch As Integer)
        compareEntry = ""
        Dim bool As Boolean
        Dim i As Integer = 0
        ndx += ch
        For i = txtCrntRec.Text To txtTotalRec.Text
            Dim comCnt As Integer
            For ndx = ndx To DataGridView1.Rows.Count - 3

                Dim e1Data As String = DataGridView1.Rows(ndx).Cells(2).Value
                Dim e2Data As String = DataGridView1.Rows(ndx).Cells(3).Value
                If Not e1Data = e2Data Then
                    bool = True
                    comCnt2 = True
                    Exit For
                Else
                    DataGridView1.Rows(ndx).Cells(1).Value = e1Data
                    comCnt += 1
                    bool = False
                End If
            Next ndx
            If bool = True Then
                If ndx = DataGridView1.Rows.Count Then
                    If Not comCnt = DataGridView1.Rows.Count Then
                        If mymode = mode.entrymode Or mymode = mode.recompareAll Then
                            insert()
                        End If
                        If txtCrntRec.Text = txtTotalRec.Text Then
                            If mymode = mode.entrymode Then
                                comLog.compareFooter(comlogpath, "Finish Compare At " & " By: " & txtUserID.Text)
                                valBox2("Finish Compare At " & vbNewLine & datetime & " By: " & txtUserID.Text)
                            ElseIf mymode = mode.recompareAll Then
                                comLog.compareFooter(comlogpath, "Finish Re-Compare At " & datetime & " By: " & txtUserID.Text)
                                valBox2("Finish Re-Compare At " & vbNewLine & datetime & " By: " & txtUserID.Text)
                            End If
                            mymode = mode.browsemode
                            changemode()
                            setCurrentRow(ndx)
                            val2 = False
                            Return recordnumber
                        Else
                            recordnumber += 1
                            curRec += 1
                            txtCrntRec.Text += 1
                            getImages()
                            loadDataFields()
                            ndx = 0
                        End If
                    End If
                End If
                Exit For
            Else
                If ndx = DataGridView1.Rows.Count - 2 Then
                    If mymode = mode.entrymode Or mymode = mode.recompareAll Then
                        insert()
                    End If
                    If txtCrntRec.Text = txtTotalRec.Text Then
                        If comCnt2 = False Then
                            comLog.ifNoRecordsFound(comlogpath, "**No Records Found**")
                        End If
                        If mymode = mode.entrymode Then
                            comLog.compareFooter(comlogpath, "Finish Compare At " & " By: " & txtUserID.Text)
                            valBox2("Finish Compare At " & vbNewLine & datetime & " By: " & txtUserID.Text)
                        ElseIf mymode = mode.recompareAll Then
                            comLog.compareFooter(comlogpath, "Finish Re-Compare At " & datetime & " By: " & txtUserID.Text)
                            valBox2("Finish Re-Compare At " & vbNewLine & datetime & " By: " & txtUserID.Text)
                        End If
                        mymode = mode.browsemode
                        changemode()
                        setCurrentRow(ndx)
                        val2 = False
                        Return recordnumber
                    Else
                        If mymode = mode.editmode Then
                            If comCnt = DataGridView1.Rows.Count - 2 Then
                                valBox2("There is no Mismatch")
                                comLog.ifNoRecordsFound(comlogpath, "**NO RECORDS FOUND**")
                                comLog.compareFooter(comlogpath, "Finish Re-Compare At " & datetime & " By: " & txtUserID.Text)
                                mymode = mode.browsemode
                                changemode()
                                txtValue.Focus()
                                val2 = False
                                Return recordnumber
                            Else
                                Dim ask As Boolean
                                ask = valBox("Save Changes??")
                                If ask = True Then
                                    comLog.compareFooter(comlogpath, "Finish Re-Compare At " & datetime & " By: " & txtUserID.Text)
                                    editRecord()
                                    valBox2("Finish Re-Compare At " & vbNewLine & datetime & " By: " & txtUserID.Text)
                                    mymode = mode.browsemode
                                    changemode()
                                    setCurrentRow(ndx)
                                    txtValue.Focus()
                                    val2 = False
                                    Return recordnumber
                                ElseIf ask = False Then
                                    comLog.compareFooter(comlogpath, "Finish Re-Compare At " & datetime & " By: " & txtUserID.Text)
                                    mymode = mode.browsemode
                                    changemode()
                                    valBox2("Record has not been saved")
                                    txtValue.Focus()
                                    valid = False
                                    Return recordnumber
                                End If
                            End If
                        Else
                            recordnumber += 1
                            curRec += 1
                            txtCrntRec.Text += 1
                            getImages()
                            loadDataFields()
                            ndx = 0
                        End If
                    End If
                End If
            End If
        Next i
        setCurrentRow(ndx)
        currentRows2 = ndx
        Return recordnumber
    End Function
    Public Function updateEntry(ByVal val1 As Boolean)
        updateEntry = ""
        If val1 = True Then
             Dim id As Integer = txtCrntRec.Text
            If mymode = mode.entrymode Then
                If txtCrntRec.Text = txtTotalRec.Text Then
                    insert()
                    valBox2("Finish Compare At " & datetime & " By: " & txtUserID.Text)
                    mymode = mode.browsemode
                    changemode()
                    getDesc()
                    txtValue.Text = ""
                    Exit Function
                Else
                    recordnumber = compareEntry(getCurrentRow(), 1)
                    txtCrntRec.Text = curRec + 1
                    getImages()
                    loadDataFields()
                    getDesc()
                    txtValue.Text = ""
                End If
            ElseIf mymode = mode.editmode Or mymode = mode.recompareAll Then
                If mymode = mode.recompareAll Then
                    If txtCrntRec.Text = txtTotalRec.Text Then
                        insert()
                        valBox2("Finish Re-Compare At " & datetime & " By: " & txtUserID.Text)
                        mymode = mode.browsemode
                        changemode()
                        getDesc()
                        txtValue.Text = ""
                        Exit Function
                    Else
                        recordnumber = compareEntry(getCurrentRow(), 1)
                        txtCrntRec.Text = curRec + 1
                        getImages()
                        loadDataFields()
                        getDesc()
                        txtValue.Text = ""
                    End If
                ElseIf mymode = mode.editmode Then
                    comLog.compareFooter(comlogpath, "Finished At " & datetime & " By: " & txtUserID.Text)
                    valBox2("Finish Re-Compare At " & datetime & " By: " & txtUserID.Text)
                    mymode = mode.browsemode
                    changemode()
                    setCurrentRow(0)
                    getDesc()
                End If
            End If
        End If
    End Function
    Public Function continueCom(ByVal val As Boolean, ByVal e As KeyEventArgs)
        continueCom = ""
        If Not mymode = mode.browsemode Then
            If e.KeyCode = Keys.F10 Or e.KeyCode = Keys.Enter Or e.KeyCode = Keys.F11 Or e.KeyCode = Keys.Up Or e.KeyCode = Keys.Down Then
                If val = True Then
                    If (getCurrentRow() + 1) = DataGridView1.Rows.Count - 2 Then
                        If mymode = mode.entrymode Or mymode = mode.recompareAll Then
                            updateEntry(val)
                        Else
                            saveData()
                        End If
                    Else
                        If Not getCurrentRow() = DataGridView1.Rows.Count - 3 Then
                            recordnumber = compareEntry(getCurrentRow(), 1)
                            txtCrntRec.Text = curRec + 1
                        End If
                    End If
                    getDesc()
                    txtValue.Focus()
                    txtValue.SelectAll()
                Else
                    txtValue.Focus()
                    txtValue.SelectAll()
                    Exit Function
                End If
            End If
        End If
    End Function
    Public Function validations()
        If txtValue.MaxLength > 1 Then
            Select Case lblFieldName.Text
                Case "emailadd"
                    valid = validation.email()
                Case "phone"
                    valid = validation.phone()
            End Select
        ElseIf txtValue.MaxLength = 1 Then
            Select Case lblFieldName.Text
                Case "Code"
                    valid = True
            End Select
        End If
        Return valid
    End Function
    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim crntRow As Integer
        If Not mymode = mode.browsemode Then
            If e.KeyCode = Keys.Up Then
                If Not getCurrentRow() = 0 Then
                    DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
                    setCurrentRow(getCurrentRow() - 1)
                    getDesc()
                    txtValue.Focus()
                    txtValue.SelectAll()
                    e.Handled = True
                End If
            End If
            If e.KeyCode = Keys.F10 Then
                DataGridView1.Rows(getCurrentRow).Cells(1).Value() = DataGridView1.Rows(getCurrentRow).Cells(2).Value()
                txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(2).Value()
                comLog.ifMismatchAppear(comlogpath, txtCrntRec.Text, DataGridView1.Rows(getCurrentRow).Cells(0).Value, DataGridView1.Rows(getCurrentRow).Cells(2).Value, DataGridView1.Rows(getCurrentRow).Cells(3).Value, DataGridView1.Rows(getCurrentRow).Cells(1).Value, "Entry1 is Correct")
                valid = validations()
                continueCom(valid, e)
                setCurrentRow(currentRows2)
               getdesc
            End If

            If e.KeyCode = Keys.F11 Then
                DataGridView1.Rows(getCurrentRow).Cells(1).Value() = DataGridView1.Rows(getCurrentRow).Cells(3).Value()
                txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(3).Value()
                comLog.ifMismatchAppear(comlogpath, txtCrntRec.Text, DataGridView1.Rows(getCurrentRow).Cells(0).Value, DataGridView1.Rows(getCurrentRow).Cells(2).Value, DataGridView1.Rows(getCurrentRow).Cells(3).Value, DataGridView1.Rows(getCurrentRow).Cells(1).Value, "Entry2 is Correct")
                valid = validations()
                continueCom(valid, e)
                setCurrentRow(currentRows2)
                getDesc()
            End If
            If e.KeyCode = Keys.Enter Then
                If Not txtValue.Modified Then
                    DataGridView1.Rows(getCurrentRow).Cells(1).Value() = DataGridView1.Rows(getCurrentRow).Cells(2).Value()
                    txtValue.Text = DataGridView1.Rows(getCurrentRow).Cells(2).Value()
                    comLog.ifMismatchAppear(comlogpath, txtCrntRec.Text, DataGridView1.Rows(getCurrentRow).Cells(0).Value, DataGridView1.Rows(getCurrentRow).Cells(2).Value, DataGridView1.Rows(getCurrentRow).Cells(3).Value, txtValue.Text, "Entry1 is Correct")
                    '  setCurrentRow(currentRows2)
                    'TextBox3.Text = DataGridView1.Rows(getCurrentRow).Cells(2).Value()
                    getDesc()
                Else
                    DataGridView1.Rows(getCurrentRow).Cells(1).Value() = txtValue.Text
                    comLog.ifMismatchAppear(comlogpath, txtCrntRec.Text, DataGridView1.Rows(getCurrentRow).Cells(0).Value, DataGridView1.Rows(getCurrentRow).Cells(2).Value, DataGridView1.Rows(getCurrentRow).Cells(3).Value, txtValue.Text, "Fully Editted")
                    ' setCurrentRow(currentRows2)
                    'TextBox3.Text = DataGridView1.Rows(getCurrentRow).Cells(2).Value()
                    getDesc()
                End If
            End If
        End If

        If mymode = mode.browsemode Then
            If e.KeyCode = Keys.Enter Then
                If Not getCurrentRow() = DataGridView1.Rows.Count Then
                    setCurrentRow(getCurrentRow() + 1)
                    getDesc()
                    txtValue.Focus()
                    txtValue.SelectAll()
                End If
            End If

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
                        Application.DoEvents()
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error Resume Next
        Dialog2.Close()
        If myComconn.State = ConnectionState.Open Then
            myComconn.Close()
        End If
        If myE1conn.State = ConnectionState.Open Then
            myE1conn.Close()
        End If
        If myE2conn.State = ConnectionState.Open Then
            myE2conn.Close()
        End If
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        If Not mymode = mode.browsemode Then
            comLog.compareFooter(comlogpath, "<~Interrupted~> ")
        End If
    End Sub
    Private Sub GOTOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GOTOToolStripMenuItem.Click
        dataDialog.ShowDialog()
    End Sub
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DataGridView1.EnableHeadersVisualStyles = False
        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkCyan
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        datetime = Now.ToString("G")
        Label8.Text = datetime
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
                    Case "emailadd"
                        valid = validation.email()
                    Case "phone"
                        valid = validation.phone()
                End Select
                If valid = True Then
                    valid = False
                    continueCom(True, e)
                End If
            End If
        ElseIf txtValue.MaxLength = 1 Then
            If getCurrentRow() = DataGridView1.Rows.Count - 3 Then
                DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
                saveData()
            Else
                Select Case lblFieldName.Text
                    Case "Code"
                        Select Case Chr(e.KeyCode)
                            Case "E", "F", "G", "H", "I", "J", "K", "L"
                                e.Handled = True
                                DataGridView1.Rows(getCurrentRow).Cells(1).Value = Chr(e.KeyCode)
                                valid = True
                        End Select
                        If e.KeyCode = Keys.Enter Then
                            Select Case txtValue.Text
                                Case "E", "F", "G", "H", "I", "J", "K", "L"
                                    e.Handled = True
                                    DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
                                    valid = True
                            End Select
                            If txtValue.Text = "*" Then
                                Dim ask As Boolean
                                txtValue.Text = "*"
                                ask = valBox("Are You Sure You Want To Pullout??")
                                If ask = True Then
                                    For i As Integer = 0 To DataGridView1.Rows.Count - 1
                                        DataGridView1.Rows(i).Cells(1).Value = ""
                                    Next
                                    DataGridView1.Rows(0).Cells(1).Value = "*"
                                    If mymode = mode.editmode Then
                                        Dim ask2 As Boolean
                                        ask2 = valBox("Save Changes??")
                                        If ask2 = True Then
                                            e.Handled = True
                                            saveData()
                                            mymode = mode.browsemode
                                            changemode()
                                        ElseIf ask2 = False Then
                                            e.Handled = True
                                            mymode = mode.browsemode
                                            changemode()
                                        End If
                                    ElseIf mymode = mode.entrymode Or mymode = mode.recompareAll Then
                                        e.Handled = True
                                        saveData()
                                        recordnumber = compareEntry(getCurrentRow(), 0)
                                        getDesc()
                                        Exit Sub
                                    End If
                                ElseIf ask = False Then
                                    valid = False
                                    e.Handled = True
                                    txtValue.Text = ""
                                End If
                            End If
                            If valid = True Then
                                valid = False
                                continueCom(True, e)
                            End If
                        End If
                End Select
            End If
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
            Case "phone"
                e.Handled = alphaOrNum(e, 0, , "- ()")
            Case "Code"
                e.Handled = alphaOrNum(e, 2, "EFGHIJKLefghijkl*")
                If e.KeyChar = "*" Then
                    Dim ask As Boolean
                    txtValue.Text = e.KeyChar
                    ask = valBox("Are You Sure You Want To Pullout??")
                    If ask = True Then
                        For i As Integer = 0 To DataGridView1.Rows.Count - 1
                            DataGridView1.Rows(i).Cells(1).Value = ""
                        Next
                        DataGridView1.Rows(0).Cells(1).Value = "*"
                        If mymode = mode.editmode Then
                            Dim ask2 As Boolean
                            ask2 = valBox("Save Changes??")
                            If ask2 = True Then
                                e.Handled = True
                                saveData()
                                mymode = mode.browsemode
                                changemode()
                            ElseIf ask2 = False Then
                                e.Handled = True
                                mymode = mode.browsemode
                                changemode()
                            End If
                        ElseIf mymode = mode.entrymode Or mymode = mode.recompareAll Then
                            e.Handled = True
                            saveData()
                            recordnumber = compareEntry(getCurrentRow(), 0)
                            getDesc()
                        End If
                    ElseIf ask = False Then
                        e.Handled = True
                        txtValue.Text = ""
                    End If
                End If
            Case "emailadd"
                e.Handled = alphaOrNum(e, 3, , "@&()_'./")
            End Select
        'If e.Handled = False Then
        '    valid = True
        'ElseIf e.Handled = True Then
        '    valid = False
        'End If
        ''End Select
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
                    If (getCurrentRow() + 1) = DataGridView1.Rows.Count - 2 Then
                        If mymode = mode.entrymode Or mymode = mode.recompareAll Then
                            updateEntry(valid)
                        Else
                            saveData()
                        End If
                    Else
                        If Not getCurrentRow() = DataGridView1.Rows.Count - 3 Then
                            recordnumber = compareEntry(getCurrentRow(), 1)
                            txtCrntRec.Text = curRec + 1
                        End If
                    End If
                    getDesc()
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
                DataGridView1.Rows(getCurrentRow).Cells(1).Value = txtValue.Text
                If (getCurrentRow() + 1) = DataGridView1.Rows.Count - 2 Then
                    If mymode = mode.entrymode Or mymode = mode.recompareAll Then
                        updateEntry(valid)
                    Else
                        saveData()
                    End If
                Else
                    If Not getCurrentRow() = DataGridView1.Rows.Count - 3 Then
                        recordnumber = compareEntry(getCurrentRow(), 1)
                        txtCrntRec.Text = curRec + 1
                    End If
                End If
             getDesc()
                txtValue.Focus()
                txtValue.SelectAll()

            End If
        End If
    End Sub
    Public Function addE1Data() As String
        Dim str As String = ""
        Try
            conn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "  Data Source=" & Application.StartupPath & "\e1Codes.mdb" & ";User Id=admin;Password=;")
            If Not myComconn.State = ConnectionState.Open Then
                conn.Open()
            End If
            Dim tempds As New DataSet
            Dim tempda As New OleDb.OleDbDataAdapter("SELECT * FROM CWC_tbl WHERE Code = '" & DataGridView1.Rows(0).Cells(1).Value & "'", conn)
            tempda.Fill(tempds)
            Dim sql As String = "SELECT "
            Debug.Print("werwerw" & tempds.Tables(0).Columns.Count - 1)
            For i As Integer = 0 To tempds.Tables(0).Columns.Count - 1
                If tempds.Tables(0).Rows(0).Item(i).ToString = "1" Then
                    sql = sql & "[" & tempds.Tables(0).Columns(i).ColumnName & "], "
                End If
            Next
            sql = sql.Remove(sql.Length - 2, 2) & " "
            sql = sql & "FROM data001 WHERE recNum = " & recordnumber

            E1Da = New OleDbDataAdapter(sql, myE1conn)
            E1Ds = New DataSet
            E1Da.Fill(E1Ds)

            For i As Integer = 0 To E1Ds.Tables(0).Columns.Count - 1
                str = str & "[" & E1Ds.Tables(0).Columns(i).ColumnName & "] = '" & E1Ds.Tables(0).Rows(0).Item(i).ToString.Replace("'", "''") & "',"
            Next
        Catch
            Str = ""
        End Try
        Return str
    End Function
    Public Function insert()
        Dim cmd As New OleDb.OleDbCommand
        Dim sql As String = "update Data001 set "
        sql = sql & addE1Data()
        For i As Integer = 0 To DataGridView1.Rows.Count - 3
            '  If DataGridView1.Rows(i).Cells(1).Value IsNot Nothing And DataGridView1.Rows(i).Cells(1).Value <> "" Then
            If i = DataGridView1.Rows.Count - 1 Then
                sql = sql & "[" & DataGridView1.Rows(i).Cells(0).Value.ToString & "]" & " = '" & DataGridView1.Rows(i).Cells(1).Value.ToString.Replace("'", "''") & "'"
            Else
                sql = sql & "[" & DataGridView1.Rows(i).Cells(0).Value.ToString & "]" & " = '" & DataGridView1.Rows(i).Cells(1).Value.ToString.Replace("'", "''") & "',"
            End If
            'End If
        Next
        sql = sql & "[" & "dataCaptu" & "]" & " = '" & datetime.Replace("'", "''") & "',"
        sql = sql & "[" & "image" & "]" & " = '" & OImg.Replace("'", "''") & "'"
        sql = sql + " where recNum = " & recordnumber
        cmd = New OleDbCommand(sql, myComconn)
        cmd.ExecuteNonQuery()
        'end code
        Dim ComFlagCnt As Integer
        sqlQry = "SELECT Comflag FROM " & jobName & " WHERE RecNum = " & recordnumber
        cmd = New OleDbCommand(sqlQry, myComconn)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        Do While (reader.Read())
            ComFlagCnt = CInt(Int(reader("Comflag")))
        Loop
        reader.Close()
        ComFlagCnt += 1
        sqlQry = "Update " & jobName & " Set Comflag = @keyflag, ComID = @keyid, ComDate = @keydate, LocID1 = @locid1  where RecNum=@recnum"
        cmd = New OleDbCommand(sqlQry, myComconn)
        cmd.Parameters.AddWithValue("@keyflag", ComFlagCnt)
        cmd.Parameters.AddWithValue("@keyid", txtUserID.Text)
        cmd.Parameters.AddWithValue("@keydate", datetime)
        cmd.Parameters.AddWithValue("@locid1", txtLocID.Text)
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
        sql = sql + " where RecNum = " & recordnumber
        cmd = New OleDbCommand(sql, myComconn)
        cmd.ExecuteNonQuery()
        'end code
        Dim updFlagCnt As Integer
        sqlQry = "SELECT Updflag FROM " & jobName & " WHERE RecNum = " & recordnumber
        cmd = New OleDbCommand(sqlQry, myComconn)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        Do While (reader.Read())
            updFlagCnt = CInt(Int(reader("Updflag")))
        Loop
        reader.Close()
        updFlagCnt += 1
        sqlQry = "Update " & jobName & " Set Updflag = @keyflag, UpdID = @keyid, UpdDate = @keydate  where RecNum=@recnum"
        cmd = New OleDbCommand(sqlQry, myComconn)
        cmd.Parameters.AddWithValue("@keyflag", updFlagCnt)
        cmd.Parameters.AddWithValue("@keyid", txtUserID.Text)
        cmd.Parameters.AddWithValue("@keydate", datetime)
        cmd.Parameters.AddWithValue("@recnum", recordnumber)
        cmd.ExecuteNonQuery()
        Return True
    End Function
    Public Function checkFlag() As Boolean
        On Error GoTo e1
        Dim cmdCom = New OleDbCommand("SELECT ComID FROM " & jobName & " Where recNum = @rec ", myComconn)
        cmdCom.Parameters.Add(New OleDbParameter("@rec", recordnumber))
        Dim comreader As OleDbDataReader
        comreader = cmdCom.ExecuteReader()
        Do While (comreader.Read())
            comdataBy = comreader("ComID")
            If Not comdataBy = "" Then
                Exit Do
            Else
                Return False
            End If
        Loop
        comreader.Close()

e1:     Dim cmdE1 = New OleDbCommand("SELECT KeyID FROM " & jobName & " Where recNum = @rec and KeyFlag = '^'", myE1conn)
        cmdE1.Parameters.Add(New OleDbParameter("@rec", recordnumber))
        Dim e1reader As OleDbDataReader
        e1reader = cmdE1.ExecuteReader()
        Do While (e1reader.Read())
            e1dataBy = e1reader("KeyID")
            If Not e1dataBy = "" Then
                Exit Do
            Else
                Return False
            End If
        Loop
        e1reader.Close()
        Dim cmdE2 = New OleDbCommand("SELECT KeyID FROM " & jobName & " Where recNum = @rec and KeyFlag = '~'", myE2conn)
        cmdE2.Parameters.Add(New OleDbParameter("@rec", recordnumber))
        Dim e2reader As OleDbDataReader
        e2reader = cmdE2.ExecuteReader()
        Do While (e2reader.Read())
            e2dataBy = e2reader("KeyID")
            If Not e1dataBy = "" Then
                Exit Do
            Else
                Return False
            End If
        Loop
        e2reader.Close()
        Return True
    End Function
    Public Function checkAllFlag() As Boolean
        checkAllFlag = False
        Dim ch As Boolean
        Try
            E1Ds = New DataSet
            E1Da = New OleDbDataAdapter("SELECT TOP 1 RecNum FROM " & jobName & " WHERE  KeyFlag = '' OR ISNULL(KeyFlag)", myE1conn)
            E1Da.Fill(E1Ds)
            If E1Ds.Tables(0).Rows(0).Item(0).ToString = "" Or E1Ds.Tables(0).Rows(0).Item(0).ToString Is Nothing Then
                valBox2(E1Ds.Tables(0).Rows(0).Item(0).ToString)
                Exit Try
            End If
            ch = False
            If ch = False Then
                valBox2("Not yet Finish in Entry1")
                mymode = mode.browsemode
                changemode()
            End If
        Catch ex As Exception
            ch = True
        End Try
        Try
            E2Ds = New DataSet
            E2Da = New OleDbDataAdapter("SELECT TOP 1 RecNum FROM " & jobName & " WHERE  KeyFlag = '' OR ISNULL(KeyFlag)", myE2conn)
            E2Da.Fill(E2Ds)
            If E2Ds.Tables(0).Rows(0).Item(0).ToString = "" Or E2Ds.Tables(0).Rows(0).Item(0).ToString Is Nothing Then
                valBox2(E1Ds.Tables(0).Rows(0).Item(0).ToString)
                Exit Try
            End If
            ch = False
            If ch = False Then
                valBox2("Not yet Finish in Entry2")
                mymode = mode.browsemode
                changemode()
            End If
        Catch ex As Exception
            ch = True
        End Try
        Return ch
    End Function
    Public Function selectTop(Optional ByVal frombrowse As Boolean = False) As Integer
        Dim ds2 As New DataSet
        Dim da2 As New OleDbDataAdapter
        Dim temp As Integer

        Try
            ds2 = New DataSet
            da2 = New OleDbDataAdapter("SELECT TOP 1 RecNum FROM " & jobName & " WHERE  ComFlag = '0'", myComconn)
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
            Dim ask As Boolean
            ask = valBox("All data has been keyed!" & "Do You Want To Recompare All??")
            If ask = True Then
                mymode = mode.recompareAll
                changemode()
            Else
                mymode = mode.browsemode
                changemode()
            End If
           
        End Try
        Return True
    End Function
    Public Function saveData(Optional ByVal n As Integer = 1)
        saveData = ""
        If mymode = mode.entrymode Or mymode = mode.recompareAll Then
            If txtCrntRec.Text = txtTotalRec.Text Then
                insert()
                txtValue.Text = ""
                valBox2("All record Has Been Keyed")
                mymode = mode.browsemode
                changemode()
            Else
                insert()
                txtValue.Text = ""
                curRec += n
                txtCrntRec.Text = curRec + 1
                loadDataFields()
                getImages()
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
        Dim bool As Boolean = checkAllFlag()
        If bool = False Then
            mymode = mode.browsemode
            changemode()
        Else
            mymode = mode.entrymode
            changemode()
        End If
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
                If Not DataGridView1.Rows(i).Cells(1).Value Is comDs.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value) Then
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
            If Not DataGridView1.Rows(i).Cells(1).Value Is comDs.Tables(0).Rows(0).Item(DataGridView1.Rows(i).Cells(0).Value) Then
                bool2 = True
            End If
        Next
        If bool2 = True Then
            Dim bool As Boolean
            bool = valBox("Do You Want To Save??")
            If bool = True Then
                saveData(0)
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
        txtValue.Focus()
        txtValue.SelectAll()
    End Function
    Private Sub RECOMPAREALLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        mymode = mode.editmode
        changemode()
    End Sub

End Class
