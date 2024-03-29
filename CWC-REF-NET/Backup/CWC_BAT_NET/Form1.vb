Imports System.IO
Imports System.Collections
Imports System.Data.OleDb
Imports System.ComponentModel.Component
Imports System.Object
Imports System.IO.StreamReader
Imports System.IO.StreamWriter
Imports System.text
Imports System.Diagnostics.Process
Imports System.Windows.Forms
Imports System.Text.RegularExpressions



Public Class Form1
    Private mouseDownLoc As Point
    Private currentPath As String = ""
    Private clientname As String = "CWC"
    Private jobName As String = "YYF_RSG"
    Private folders() As String
    Private valID As String
    Private totalPullout As Integer
    Private totalRecord As Integer
    Private conn, imgconn As OleDb.OleDbConnection
    Private dateTime As String
    Private targetDir As String
    Private oXL(10) As Excel.Application
    Private oWB(10) As Excel.Workbook
    Private oSheet(10) As Excel.Worksheet
    Private oRng(10) As Excel.Range
    Private filename(10) As String
      Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        Me.Close()
    End Sub
    Private Sub Label15_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label15.MouseHover
        Label15.ForeColor = Color.White
        Label15.BackColor = Color.Firebrick
    End Sub
    Private Sub Label15_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label15.MouseLeave
        Label15.ForeColor = Color.White
        Label15.BackColor = Color.Transparent
    End Sub 
    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Const DROPSHADOW = &H30000
            Dim cParam As CreateParams = MyBase.CreateParams
            cParam.ClassStyle = cParam.ClassStyle Or DROPSHADOW
            Return cParam
        End Get
    End Property
    Private Sub Panel1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        mouseDownLoc.X = e.X
        mouseDownLoc.Y = e.Y
    End Sub
    Private Sub Panel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Me.Left += e.X - mouseDownLoc.X
            Me.Top += e.Y - mouseDownLoc.Y
        End If
    End Sub
    Private Sub DirListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox1.SelectedIndexChanged
        Application.DoEvents()
        currentPath = DirListBox1.DirList(DirListBox1.DirListIndex)
        DirListBox1.Path = currentPath
        getList()
    End Sub
    Public Function getList()
        On Error Resume Next
        getList = ""
        ListView1.Items.Clear()
        ListView3.Items.Clear()
       folders = IO.Directory.GetDirectories(DirListBox1.Path)
        For i As Integer = 0 To folders.Length - 1
            ListView1.Items.Add(folders(i))
        Next
        Dim mdbCount() As String
        mdbCount = IO.Directory.GetFiles(DirListBox1.Path, "*.mdb")
        For j As Integer = 0 To mdbCount.Length - 1
            Dim name As New DirectoryInfo(mdbCount(j))
            ListView3.Items.Add(name.Name)
        Next
    End Function
    Private Sub DriveListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DriveListBox1.SelectedIndexChanged
        DirListBox1.Path = DriveListBox1.Drive
    End Sub 
    Private Sub cmdAddAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddAll.Click
        For i As Integer = 0 To ListView1.Items.Count - 1
            ListView1.Items(i).Selected = True
            ListView1.Select()
        Next
        cmdAdd.PerformClick()
    End Sub
    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not ListView2.SelectedItems.Count = 0 Then
            If ListView2.SelectedItems.Count > 0 Then
                For i As Integer = 0 To ListView2.SelectedItems.Count - 1
                    Dim li1 As ListViewItem.ListViewSubItem = ListView2.SelectedItems(i).SubItems(0)
                    Dim li As ListViewItem = ListView1.Items.Add(li1.Text)
                Next
                For Each i As ListViewItem In ListView2.SelectedItems
                    ListView2.Items.Remove(i)
                Next
            Else
                valBox2("No Item selected")
            End If
        Else
            valBox2("No Items to be Removed")
        End If
    End Sub
    Private Sub cmdClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearAll.Click
        If Not ListView2.Items.Count = 0 Then
            For i As Integer = 0 To ListView2.Items.Count - 1
                Dim li1 As ListViewItem.ListViewSubItem = ListView2.Items(i).SubItems(0)
                Dim li As ListViewItem = ListView1.Items.Add(li1.Text)
            Next
            ListView2.Items.Clear()
        Else
            valBox2("No Items to be Cleared")
        End If
    End Sub
    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        For j As Integer = 0 To ListView1.SelectedItems.Count - 1
            For i As Integer = 0 To ListView2.Items.Count - 1
                If (ListView1.SelectedItems.Item(j).SubItems(0)).Text = (ListView2.Items(i).SubItems(0)).Text Then
                    valBox2(" Selected Item is alreday selected")
                    Exit Sub
                End If
            Next i
        Next j
        For i As Integer = 0 To ListView1.SelectedItems.Count - 1
            Dim li1 As ListViewItem.ListViewSubItem = ListView1.SelectedItems(i).SubItems(0)

            Dim li As ListViewItem = ListView2.Items.Add(li1.Text)

        Next
        For Each i As ListViewItem In ListView1.SelectedItems
            ListView1.Items.Remove(i)
        Next
    End Sub
    Private Sub ListView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView2.DoubleClick
        For i As Integer = 0 To ListView2.SelectedItems.Count - 1
            Dim li1 As ListViewItem.ListViewSubItem = ListView2.SelectedItems(i).SubItems(0)
            Dim li As ListViewItem = ListView1.Items.Add(li1.Text)
        Next
        For Each i As ListViewItem In ListView2.SelectedItems
            ListView2.Items.Remove(i)
        Next
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
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        For j As Integer = 0 To ListView1.SelectedItems.Count - 1
            For i As Integer = 0 To ListView2.Items.Count - 1
                If (ListView1.SelectedItems.Item(j).SubItems(0)).Text = (ListView2.Items(i).SubItems(0)).Text Then
                    valBox2(" Selected Item is already selected")
                    Exit Sub
                End If
            Next i
        Next j
        If Not ListView1.SelectedItems.Count = 0 Then
            If ListView1.SelectedItems.Count > 0 Then
                For i As Integer = 0 To ListView1.SelectedItems.Count - 1
                    Dim li1 As ListViewItem.ListViewSubItem = ListView1.SelectedItems(i).SubItems(0)
                    Dim li As ListViewItem = ListView2.Items.Add(li1.Text)
                Next
                For Each i As ListViewItem In ListView1.SelectedItems
                    ListView1.Items.Remove(i)
                Next
            Else
                valBox2("No Item selected")
            End If
        Else
            valBox2("No Available Items")
        End If

    End Sub 
    Public Function createConnection(ByVal DatabaseFullPath As String)
        conn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & DatabaseFullPath & ";User Id=admin;Password=;")
        Return conn
    End Function
    Public Function createImgConnection(ByVal DatabaseFullPath As String)
        imgconn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & DatabaseFullPath & ";User Id=admin;Password=;")
        Return imgconn
    End Function
    Public Function insertFlag(ByVal mdbpath As String, ByVal recNum As Integer, ByVal valDate As String, ByVal valid As String)
        createConnection(mdbpath)
        Dim sqlQry As String = "Update " & jobName & " Set Refflag = @keyflag, RefID = @keyid, RefDate = @keydate where RecNum=@recnum"
        Dim cmd As New OleDbCommand(sqlQry, conn)
        cmd.Parameters.AddWithValue("@keyflag", "<>")
        cmd.Parameters.AddWithValue("@keyid", valid)
        cmd.Parameters.AddWithValue("@keydate", valDate)
        cmd.Parameters.AddWithValue("@recnum", recNum)
        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()
        Return recNum
    End Function
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If codeConn.State = ConnectionState.Open Then
            codeConn.Close()
        End If
    End Sub
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        dateTime = Now.ToString("MMddyyyy")
        If Not IO.Directory.Exists("C:\" & clientname & "\" & jobName & "\OUT") Then
            System.IO.Directory.CreateDirectory("C:\" & clientname & "\" & jobName & "\OUT")
            targetDir = "C:\" & clientname & "\" & jobName & "\OUT"
        Else
            targetDir = "C:\" & clientname & "\" & jobName & "\OUT"
        End If
        TextBox3.Text = targetDir
        createCodeConnection(Application.StartupPath & "\" & "Codings-REF.mdb")
    End Sub
    Public Function getFileName(ByVal code As String) As String
        Dim ds As New DataSet
        Dim da As New OleDb.OleDbDataAdapter("SELECT filename FROM CWC_tbl WHERE Code = '" & code & "'", codeConn)
        da.Fill(ds)
        getFileName = ds.Tables(0).Rows(0).Item(0).ToString
        Return getFileName
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub cmdRef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRef.Click
        If Not ListView2.Items.Count = 0 Then
            Dialog1.ShowDialog()
            valID = Dialog1.valID
            If valID = "" Then
                Exit Sub
            End If
            Dim mdb() As String
           For n As Integer = 0 To ListView2.Items.Count - 1
                Dim dirlist() As String
                dirlist = IO.Directory.GetDirectories(ListView2.Items(n).Text, "Compare")
                For m As Integer = 0 To dirlist.Length - 1
                    '  On Error GoTo exitproc
                    mdb = IO.Directory.GetFiles(dirlist(m), "*.mdb")
                    For i As Integer = 0 To mdb.Length - 1
                        Dim pullout As Integer
                        Dim record As Integer
                        createConnection(mdb(i))
                        Dim mdbpath As New DirectoryInfo(mdb(i))
                        Dim ds, ids As New DataSet
                        Dim da As New OleDbDataAdapter("SELECT * FROM Data001", conn)
                        Dim ida As New OleDbDataAdapter("SELECT OImage001 FROM " & jobName, conn)
                        da.Fill(ds)
                        ida.Fill(ids)
                        Dim rowcount As Integer = 1
                        ProgressBar1.Value = 0
                        ProgressBar1.Minimum = 0
                        ProgressBar1.Maximum = ds.Tables(0).Rows.Count
                        Label5.Text = i + 1 & "/" & mdb.Length
                        Application.DoEvents()
                        For k As Integer = 0 To ds.Tables(0).Rows.Count - 1
                            Dim code As String
                            If Not ds.Tables(0).Rows(k).Item("code").ToString = "*" Then
                                totalRecord += 1
                                record += 1
                                code = ds.Tables(0).Rows(k).Item("code").ToString
                                Dim c As Integer
                                Select Case code
                                    Case "E"
                                        c = 0
                                        tc(0) += 1
                                        filename(0) = targetDir & "\" & getFileName(code) & dateTime & ".xlsx"
                                    Case "F"
                                        c = 1
                                        filename(1) = targetDir & "\" & getFileName(code) & dateTime & ".xlsx"
                                        tc(1) += 1
                                    Case "G"
                                        c = 2
                                        filename(2) = targetDir & "\" & getFileName(code) & dateTime & ".xlsx"
                                        tc(2) += 1
                                    Case "H"
                                        c = 3
                                        filename(3) = targetDir & "\" & getFileName(code) & dateTime & ".xlsx"
                                        tc(3) += 1
                                    Case "I"
                                        c = 4
                                        filename(4) = targetDir & "\" & getFileName(code) & dateTime & ".xlsx"
                                        tc(4) += 1
                                    Case "J"
                                        If ds.Tables(0).Rows(k).Item("review2").ToString = "1" Then
                                            c = 5
                                            filename(5) = targetDir & "\" & getFileName(code) & "A_" & dateTime & ".xlsx"
                                            tc(5) += 1
                                        ElseIf ds.Tables(0).Rows(k).Item("review2").ToString = "2" Then
                                            c = 6
                                            filename(6) = targetDir & "\" & getFileName(code) & "B_" & dateTime & ".xlsx"
                                            tc(6) += 1
                                        End If
                                    Case "K"
                                        c = 7
                                        filename(7) = targetDir & "\" & getFileName(code) & dateTime & ".xlsx"
                                        tc(7) += 1
                                    Case "L"
                                        c = 8
                                        filename(8) = targetDir & "\" & getFileName(code) & dateTime & ".xlsx"
                                        tc(8) += 1
                                    Case Else
                                        GoTo exitproc
                                End Select
                                If oWB(c) Is Nothing Then
                                    tc(c) += getrowCnt(filename(c))
                                    If IO.File.Exists(filename(c)) = False Then
                                        oXL(c) = CreateObject("Excel.Application")
                                        oWB(c) = oXL(c).Workbooks.Add
                                    Else
                                        oXL(c) = CreateObject("Excel.Application")
                                        oWB(c) = oXL(c).Workbooks.Open(Filename:=filename(c))
                                    End If
                                    oSheet(c) = oWB(c).ActiveSheet
                                    oRng(c) = oSheet(c).Range("A1", "AZ1")
                                End If
                                Dim codeDS, e1DS As New DataSet
                                Dim codeDA As New OleDbDataAdapter("SELECT * FROM CWC_tbl WHERE code = '" & code & "'", codeConn)
                                Dim e1DA As New OleDbDataAdapter("SELECT * FROM e1_tbl WHERE code = '" & code & "'", codeConn)
                                codeDA.Fill(codeDS)
                                e1DA.Fill(e1DS)
                                If Not IO.File.Exists(filename(c)) Then
                                    Dim colCount As Integer = 1
                                    For col As Integer = 0 To codeDS.Tables(0).Columns.Count - 1
                                        Dim colheader As String = codeDS.Tables(0).Columns(col).ColumnName
                                        If codeDS.Tables(0).Rows(0).Item(col).ToString = "1" Then
                                            oRng(c).ColumnWidth = 12
                                            If colheader = "hemangeolC" Then
                                                colheader = "hemangeol"
                                            ElseIf colheader = "receive Samples" Then
                                                colheader = "receive Sample"
                                            End If
                                            If codeDS.Tables(0).Columns(col).ColumnName = "telephone number" Then
                                                oSheet(c).Cells(1, colCount).value = codeDS.Tables(0).Columns(col + 1).ColumnName.ToUpper
                                                oSheet(c).Cells(1, colCount + 1).value = colheader.ToUpper
                                                oSheet(c).Cells(1, colCount + 1).NumberFormat = "@"
                                                colCount += 2
                                                col += 1
                                            Else
                                                oSheet(c).Cells(1, colCount).value = colheader.ToUpper
                                                colCount += 1
                                            End If
                                        End If
                                    Next
                                    Dim colcount2 As Integer = 1
                                    For col As Integer = 0 To e1DS.Tables(0).Columns.Count - 1
                                        If e1DS.Tables(0).Rows(0).Item(col).ToString = "1" Then
                                            Dim colname As String = e1DS.Tables(0).Columns(col).ColumnName
                                            oSheet(c).Cells(2, colcount2).NumberFormat = "@"
                                            If colname = "dataCaptu" Then
                                                oSheet(c).Cells(2, colcount2).value = formatDate(ds.Tables(0).Rows(k).Item(col).ToString)
                                            ElseIf colname = "image" Then
                                                oSheet(c).Cells(2, colcount2).value = ids.Tables(0).Rows(k).Item(0).ToString
                                            ElseIf colname = "influence" Then
                                                Select Case c
                                                    Case 0
                                                        oSheet(c).Cells(tc(c), colcount2).value = getconv(colname, "")
                                                    Case Else
                                                        oSheet(c).Cells(tc(c), colcount2).value = getconv(colname, ds.Tables(0).Rows(k).Item(colname).ToString)
                                                End Select
                                            ElseIf colname = "include1" Then
                                                With ds.Tables(0)
                                                    Dim str As String = ""
                                                    str += getconv("include1", .Rows(k).Item("include1").ToString())
                                                    str += getconv("include2", .Rows(k).Item("include2").ToString())
                                                    str += getconv("include3", .Rows(k).Item("include3").ToString())
                                                    str += getconv("include4", .Rows(k).Item("include4").ToString())
                                                    str += getconv("include5", .Rows(k).Item("include5").ToString())
                                                    oSheet(c).Cells(2, colcount2).value = str
                                                End With
                                            ElseIf colname = "phone" Then
                                                With ds.Tables(0).Rows(k)
                                                    oSheet(c).Cells(2, colcount2).value = .Item("emailadd").ToString()
                                                    oSheet(c).Cells(2, colcount2 + 1).value = .Item("phone").ToString()
                                                    oSheet(c).Cells(2, colcount2 + 1).NumberFormat = "@"
                                                End With
                                                colcount2 += 1
                                                col += 1
                                            Else
                                                oSheet(c).Cells(2, colcount2).value = getconv(colname, ds.Tables(0).Rows(k).Item(colname).ToString)
                                            End If
                                            colcount2 += 1
                                        End If
                                    Next
                                    oWB(c).SaveAs(filename(c))
                                Else
                                    Dim colcount2 As Integer = 1
                                    For col As Integer = 0 To e1DS.Tables(0).Columns.Count - 1
                                        If e1DS.Tables(0).Rows(0).Item(col).ToString = "1" Then
                                            Dim colname As String = e1DS.Tables(0).Columns(col).ColumnName
                                            oSheet(c).Cells(tc(c), colcount2).NumberFormat = "@"
                                            If colname = "dataCaptu" Then
                                                oSheet(c).Cells(tc(c), colcount2).value = formatDate(ds.Tables(0).Rows(k).Item("dataCaptu").ToString)
                                            ElseIf colname = "image" Then
                                                oSheet(c).Cells(tc(c), colcount2).value = ids.Tables(0).Rows(k).Item(0).ToString
                                            ElseIf colname = "influence" Then
                                                Select Case c
                                                    Case 0
                                                        oSheet(c).Cells(tc(c), colcount2).value = getconv(colname, "")
                                                    Case 1
                                                        oSheet(c).Cells(tc(c), colcount2).value = getconv(colname, ds.Tables(0).Rows(k).Item(colname).ToString)
                                                End Select
                                            ElseIf colname = "include1" Then
                                                With ds.Tables(0)
                                                    Dim str As String = ""
                                                    str += getconv("include1", .Rows(k).Item("include1").ToString())
                                                    str += getconv("include2", .Rows(k).Item("include2").ToString())
                                                    str += getconv("include3", .Rows(k).Item("include3").ToString())
                                                    str += getconv("include4", .Rows(k).Item("include4").ToString())
                                                    str += getconv("include5", .Rows(k).Item("include5").ToString())
                                                    oSheet(c).Cells(tc(c), colcount2).value = str
                                                End With
                                            ElseIf colname = "phone" Then
                                                With ds.Tables(0).Rows(k)
                                                    oSheet(c).Cells(tc(c), colcount2).value = .Item("emailadd").ToString()
                                                    oSheet(c).Cells(tc(c), colcount2 + 1).value = .Item("phone").ToString()
                                                    oSheet(c).Cells(tc(c), colcount2 + 1).NumberFormat = "@"
                                                End With
                                                colcount2 += 1
                                                col += 1
                                            Else
                                                oSheet(c).Cells(tc(c), colcount2).value = getconv(colname, ds.Tables(0).Rows(k).Item(colname).ToString)
                                            End If
                                            colcount2 += 1
                                        End If
                                    Next
                                End If
                                insertFlag(mdb(i), ds.Tables(0).Rows(k).Item(0).ToString, dateTime, valID)
                            Else
                                totalPullout += 1
                                pullout += 1
                            End If
                            ProgressBar1.Value += 1
                            Application.DoEvents()
                        Next
                        record = 0
                        pullout = 0
                        save()
                    Next
                Next
exitproc:   Next
            killExcel()
            valBox2("DONE")
            TextBox1.Text = totalPullout
            totalRecord = 0
            totalPullout = 0
            ' Make sure that you release object references.
        Else
            MsgBox("No Batches To Validate")
        End If
    End Sub
    Public Function getconv(ByVal fieldname As String, ByVal str As String) As String
        getconv = ""
        Select Case fieldname
            Case "commonReason", "visit", "healthCare"
                Dim tmparr(str.Length) As String
                For j As Integer = 0 To str.Length - 1
                    Select Case str.Substring(j, 1)
                        Case "1"
                            tmparr(j) = str.Substring(j, 1).Replace("1", "1 ")
                        Case "2"
                            tmparr(j) = str.Substring(j, 1).Replace("2", "3 ")
                        Case "3"
                            tmparr(j) = str.Substring(j, 1).Replace("3", "2 ")
                        Case "4"
                            tmparr(j) = str.Substring(j, 1).Replace("4", "4 ")
                        Case "5"
                            tmparr(j) = str.Substring(j, 1).Replace("5", "5 ")
                    End Select
                Next
                Array.Sort(tmparr)
                getconv = String.Join("", tmparr).Replace(" ", "")
                Return getconv
            Case "include1"
                str = str.Replace("1", "01")
                str = str.Replace("2", "02")
                str = str.Replace("3", "03")
                str = str.Replace("4", "04")
                getconv = str
            Case "include2"
                str = str.Replace("1", "05")
                str = str.Replace("2", "06")
                str = str.Replace("3", "07")
                getconv = str
            Case "include3"
                str = str.Replace("1", "08")
                str = str.Replace("2", "09")
                str = str.Replace("3", "10")
                getconv = str
            Case "include4"
                str = str.Replace("1", "11")
                str = str.Replace("2", "12")
                str = str.Replace("3", "13")
                getconv = str
            Case "include5"
                str = str.Replace("1", "14")
                str = str.Replace("2", "15")
                str = str.Replace("3", "")
                getconv = str
            Case "issue"
                 Dim tmparr(str.Length) As String
                For j As Integer = 0 To str.Length - 1
                    Select Case str.Substring(j, 1)
                        Case "1"
                            tmparr(j) = str.Substring(j, 1).Replace("1", "1 ")
                        Case "2"
                            tmparr(j) = str.Substring(j, 1).Replace("2", "4 ")
                        Case "3"
                            tmparr(j) = str.Substring(j, 1).Replace("3", "2 ")
                        Case "4"
                            tmparr(j) = str.Substring(j, 1).Replace("4", "3 ")
                    End Select
                Next
                Array.Sort(tmparr)
                getconv = String.Join("", tmparr).Replace(" ", "")
                Return getconv
            Case "topic"
                Dim tmparr(str.Length) As String
                For j As Integer = 0 To str.Length - 1
                    Select Case str.Substring(j, 1)
                        Case "1"
                            tmparr(j) = str.Substring(j, 1).Replace("1", "1 ")
                        Case "2"
                            tmparr(j) = str.Substring(j, 1).Replace("2", "6 ")
                        Case "3"
                            tmparr(j) = str.Substring(j, 1).Replace("3", "2 ")
                        Case "4"
                            tmparr(j) = str.Substring(j, 1).Replace("4", "3 ")
                        Case "5"
                            tmparr(j) = str.Substring(j, 1).Replace("5", "7 ")
                        Case "6"
                            tmparr(j) = str.Substring(j, 1).Replace("6", "4 ")
                        Case "7"
                            tmparr(j) = str.Substring(j, 1).Replace("7", "5 ")
                    End Select
                Next
                Array.Sort(tmparr)
                getconv = String.Join("", tmparr).Replace(" ", "")
                Return getconv
            Case "title"
                Select Case str
                    Case "1"
                        getconv = "MR"
                    Case "2"
                        getconv = "MRS"
                    Case "3"
                        getconv = "MS"
                    Case "4"
                        getconv = "MISS"
                    Case Else
                        getconv = str
                End Select
            Case "advertisement", "1stBaby", "HeartSmart", "hibiclens", "adt", "advetori", "axa", "BOBA", "consider", "content", "coupon", "dhaO3", "editorial", "factsUp", "financialfuture", "hemange", "hemangeolC", "influence", "makeAPurchase", "notice", "offer", "receiveSample", "receiveSamples", "review", "review2", "stonyfie"
                Select Case str
                    Case "1"
                        getconv = "Y"
                    Case "2"
                        getconv = "N"
                End Select
            Case "gender1", "gender2", "gender3"
                Select Case str
                    Case ""
                        getconv = "U"
                    Case "1"
                        getconv = "B"
                    Case "2"
                        getconv = "G"
                End Select
            Case Else
                getconv = str
        End Select
        Return getconv
    End Function
    Public Function formatDate(ByVal str As String) As String
        formatDate = str
        Try
            formatDate = System.DateTime.Parse(str).ToString("MMddyy")
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try
        Return formatDate
    End Function
    Public Sub killExcel()
        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        For Each Process As Process In xlp
            Process.Kill()
            If xlp.Length = 0 Then
                Exit For
            End If
        Next
    End Sub
    Private tc(10) As Integer
    Public Sub reformat(ByVal tmpXL As Excel.Application, ByVal tmpWB As Excel.Workbook, ByVal tmpSheet As Excel.Worksheet, ByVal tmpRng As Excel.Range, ByVal ds As DataSet, ByVal filename As String, ByVal code As String, ByVal rCnt As Integer)
        Dim codeDS As New DataSet
        Dim codeDA As New OleDbDataAdapter("SELECT * FROM CWC_tbl WHERE code = '" & code & "'", codeConn)
        codeDA.Fill(codeDS)
        Dim rowCnt As Integer = 0
        Select Case code
            Case "E"
                tc(0) += 1
                rowCnt += tc(0)
            Case "F"
                tc(1) += 1
                rowCnt += tc(1)
            Case "G"
                tc(2) += 1
                rowCnt += tc(2)
            Case "H"
                tc(3) += 1
                rowCnt += tc(3)
            Case "I"
                tc(4) += 1
                rowCnt += tc(4)
            Case "J"
                tc(5) += 1
                rowCnt += tc(5)
            Case "K"
                tc(6) += 1
                rowCnt += tc(6)
            Case "L"
                tc(7) += 1
                rowCnt += tc(7)
        End Select
        rowCnt += getrowCnt(filename)
        If Not IO.File.Exists(filename) Then
            Dim colCount As Integer = 1
            For col As Integer = 0 To ds.Tables(0).Columns.Count - 1
                If codeDS.Tables(0).Rows(0).Item(col).ToString = "1" Then
                    tmpSheet.Cells(1, colCount).Value = codeDS.Tables(0).Columns(col).ColumnName
                    colCount += 1
                End If
            Next
            Dim colcount2 As Integer = 1
            For col As Integer = 0 To ds.Tables(0).Columns.Count - 1
                If codeDS.Tables(0).Rows(0).Item(col).ToString = "1" Then
                    tmpSheet.Cells(2, colcount2).value = ds.Tables(0).Rows(rCnt).Item(col).ToString
                    colcount2 += 1
                End If
            Next
            With tmpSheet.Range("A1", "CC1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            tmpWB.SaveAs(filename)
        Else
            Dim colcount2 As Integer = 1
            For col As Integer = 0 To ds.Tables(0).Columns.Count - 1
                If codeDS.Tables(0).Rows(0).Item(col).ToString = "1" Then
                    tmpSheet.Cells(rowCnt, colcount2).value = ds.Tables(0).Rows(rCnt).Item(col).ToString
                    colcount2 += 1
                End If
            Next
        End If
    End Sub
    Public Sub save()
        For i As Integer = 0 To oXL.Length - 1
            If Not filename(i) = "" Then
                If IO.File.Exists(filename(i)) Then
                    oWB(i).Save()
                Else
                    oWB(i).SaveAs(filename(i))
                End If
                oWB(i).Close()
                oWB(i) = Nothing
                oXL(i).Quit()
                oXL(i) = Nothing
                tc(i) = 0
            End If
        Next
    End Sub
    Private xlsConn, codeConn As OleDb.OleDbConnection
    Public Function createCodeConnection(ByVal DatabaseFullPath As String)
        codeConn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & DatabaseFullPath & ";User Id=admin;Password=;")
        Return codeConn
    End Function
    Public Function getrowCnt(ByVal DatabaseFullPath As String) As Integer
        Try
            xlsConn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "  Data Source=" & DatabaseFullPath & ";Extended Properties=Excel 8.0")
            xlsConn.Open()
            Dim xlsDS As New DataSet
            Dim xlsDA As New OleDbDataAdapter("SELECT * FROM [Sheet1$]", xlsConn)
            xlsDA.Fill(xlsDS)
            getrowCnt = xlsDS.Tables(0).Rows.Count + 1
             xlsConn.Close()
            xlsConn.Dispose()
        Catch ex As Exception
            getrowCnt = 1
            If xlsConn.State = ConnectionState.Open Then
                xlsConn.Close()
                xlsConn.Dispose()
            End If
        End Try
        Return getrowCnt
    End Function
  
 
End Class

