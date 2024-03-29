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
    Private dblog As String
    Private batchCount As String
    Private batchSize As String
    Private batchLocation As String
    Private valID As String
    Private totalPullout As Integer
    Private totalRecord As Integer
    Private conn, cConn As OleDb.OleDbConnection
    Private dateTime As String
    Public targetDir As String
    Private cDS As DataSet
    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        Me.Close()
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
        If rbPerBatch.Checked Then
            ListView1.Items.Clear()
            Dim folders() As String
            folders = IO.Directory.GetDirectories(DirListBox1.Path)
            For i As Integer = 0 To folders.Length - 1
                ListView1.Items.Add(folders(i))
            Next
        ElseIf rbPerMDB.Checked Then
            ListView1.Items.Clear()
            Dim mdbCount() As String
            mdbCount = IO.Directory.GetFiles(DirListBox1.Path, "*.mdb")
            For j As Integer = 0 To mdbCount.Length - 1
                Dim name As New DirectoryInfo(mdbCount(j))
                ListView1.Items.Add(name.Name)
            Next
        End If
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
    Public Function countDB(ByVal imgCount As Integer, ByVal batchSize As Integer)
        Dim dbcount As Integer
        If batchSize > imgCount Then
            dbcount = 1
            Return dbcount
        End If
        If Not imgCount Mod 2 = 0 Then
            imgCount += 1
        End If
        dbcount = imgCount \ batchSize
        If Not imgCount Mod batchSize = 0 Then
            dbcount += 1
        End If
        Return dbcount
    End Function
    Public Function countAllImages()
        Dim imgcount As Integer
        For i As Integer = 0 To ListView2.Items.Count - 1
            imgcount += ListView2.Items(i).SubItems(1).Text
        Next
        Return imgcount
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
    Function addLength(ByVal str As String, ByVal len As Integer) As String
        For i As Integer = str.Length To len - 1
            str &= " "
        Next
        Return str
    End Function
    Public Function startLog(ByVal logPath As String, ByVal validID As String, ByVal companyName As String)
        startLog = ""
        Dim logStr As String
        logStr = "Error Logfile Found In: " & companyName & vbNewLine & logPath & vbNewLine & "Validation ID:  " & validID
        My.Computer.FileSystem.WriteAllText(logPath & ".err", logStr, False)
    End Function
    Public Function ifValidationAppear(ByVal logPath As String, ByVal record As String, ByVal field As String, ByVal description As String, ByVal value As String)
        ifValidationAppear = ""
        If My.Computer.FileSystem.FileExists(logPath & ".err") Then
            Dim log As String = My.Computer.FileSystem.ReadAllText(logPath & ".err")
            Dim logStr As String
            record = addLength(record, 9)
            field = addLength(field, 26)
            description = addLength(description, 60)
            value = addLength(value, 50)
            logStr = log & vbNewLine & record & field & description & value
            My.Computer.FileSystem.WriteAllText(logPath & ".err", logStr, False)
        End If
    End Function  
    Public Function createConnection(ByVal DatabaseFullPath As String)
        conn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & DatabaseFullPath & ";User Id=admin;Password=;")
        Return conn
    End Function  
    Public Function insertCount(ByVal kscount As String, ByVal recnum As Integer)
        insertCount = ""
        Dim sqlQry As String = "Update " & jobName & " Set  KSCount = @keydate where RecNum=@recnum"
        Dim cmd As New OleDbCommand(sqlQry, conn)
        cmd.Parameters.AddWithValue("@keyDate", kscount)
        cmd.Parameters.AddWithValue("@recnum", recnum)
        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()
    End Function
    Public Function startWriteCNT(ByVal logPath As String)
        startWriteCNT = ""
        Dim log As String
        log = "Batch No,Total Count,Total KeyStroke,Total Average" & vbNewLine
        My.Computer.FileSystem.WriteAllText(logPath & ".CNT", log, False)
    End Function
    Public Function writeCNT(ByVal logPath As String, ByVal batch As String, ByVal totalCnt As String, ByVal totalKeyStroke As Integer, ByVal totalAverage As Integer)
        writeCNT = ""
        If My.Computer.FileSystem.FileExists(logPath & ".CNT") Then
            Dim log As String = My.Computer.FileSystem.ReadAllText(logPath & ".CNT")
            log = log & batch & "," & totalCnt & "," & totalKeyStroke & "," & totalAverage & vbNewLine
            My.Computer.FileSystem.WriteAllText(logPath & ".CNT", log, False)
        End If
    End Function
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        dateTime = Now.ToString("MMddyyyy")
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Public Function finishWriteCNT(ByVal logPath As String, ByVal totalRecord As Integer, ByVal validRecord As Integer, ByVal totalPullout As Integer, ByVal aveKeyStroke As String, ByVal totalKS As String)
        finishWriteCNT = ""
        If My.Computer.FileSystem.FileExists(logPath & ".CNT") Then
            Dim log As String = My.Computer.FileSystem.ReadAllText(logPath & ".CNT")
            log = log & vbNewLine & vbNewLine & _
            "================================================" & vbNewLine & _
            "Total Pullout: " & totalPullout & vbNewLine & _
            "================================================" & vbNewLine & _
            "Total Record: " & totalRecord & vbNewLine & _
            "Total Valid Record: " & validRecord & vbNewLine & _
            "Total KeyStroke: " & totalKS & vbNewLine & _
            "Average KeyStroke: " & aveKeyStroke & vbNewLine & _
            "================================================"


            My.Computer.FileSystem.WriteAllText(logPath & ".CNT", log, False)
        End If
    End Function 
    Private Sub cmdKS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKS.Click
        If Not ListView2.Items.Count = 0 Then
            Dim batchname As String = ""
            Dim totalPullout As Integer
            Dim totalrecord As Integer
            Dim totalKS As Integer
            Dim aveKS As Double
            Dim validRecord As Integer
            Dim database As String = ""
            Dim mdbType As String = ""
            If rbCom.Checked Then
                database = "Compare"
                mdbType = "CWC4E1"
            ElseIf rbE1.Checked Then
                database = "Entry1"
                mdbType = "CWC4E1"
            ElseIf rbE2.Checked Then
                database = "Entry2"
                mdbType = "CWC4E2"
            Else
                valBox2("Database Is Not Selected")
                Exit Sub
            End If
            FolderBrowserDialog1.ShowDialog()
            Dim KSloc As String = FolderBrowserDialog1.SelectedPath
            If KSloc = "" Then
                valBox2("No Path Selected")
                Exit Sub
            End If
            startWriteCNT(KSloc & "\" & jobName & "_" & dateTime & "_" & database)
            If rbPerBatch.Checked Then
                Dim mdb() As String
                ProgressBar1.Value = 0
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = ListView2.Items.Count
                For n As Integer = 0 To ListView2.Items.Count - 1
                    Dim inRec As Integer
                    Dim inKey As Integer
                    Dim inaveKey As Double
                    Dim dirlist() As String
                    dirlist = IO.Directory.GetDirectories(ListView2.Items(n).Text, database)
                    For m As Integer = 0 To dirlist.Length - 1
                        mdb = IO.Directory.GetFiles(dirlist(m), "*.mdb")
                        ProgressBar2.Value = 0
                        ProgressBar2.Minimum = 0
                        ProgressBar2.Maximum = mdb.Length
                        For i As Integer = 0 To mdb.Length - 1
                            createConnection(mdb(i))
                            Dim mdbpath As New DirectoryInfo(mdb(i))
                            Dim ds, ds2 As New DataSet
                            Dim da2 As New OleDbDataAdapter("SELECT * FROM Data001", conn)
                            Dim da As New OleDbDataAdapter("SELECT ZipFile FROM " & jobName, conn)
                            da.Fill(ds)
                            da2.Fill(ds2)
                            batchname = ds.Tables(0).Rows(0).Item("ZipFile").ToString
                            For k As Integer = 0 To ds.Tables(0).Rows.Count - 1
                                Dim recLength As Integer = 0
                                With ds2.Tables(0)
                                    getCode(.Rows(k).Item("Code").ToString, mdbType)
                                    If .Rows(k).Item("Code").ToString = "*" Then
                                        totalPullout += 1
                                         GoTo skip1
                                    Else
                                        If database = "Compare" Then
                                            recLength += 1 * 2
                                        Else
                                            recLength += 1
                                        End If
                                        validRecord += 1
                                    End If
                                    Select Case database
                                        Case "Entry1"
                                            For l As Integer = 1 To cDS.Tables(0).Columns.Count - 1
                                                 If cDS.Tables(0).Rows(0).Item(l).ToString = "1" Then
                                                    Dim dataCount As Integer = .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString.Length
                                                    Select Case cDS.Tables(0).Columns(l).ColumnName
                                                        Case "qasFlag", "advertisement", "1stBaby", "HeartSmart", "hibiclens", "qasFlag", "adt", "advetori", "axa", "BOBA", "consider", "content", "coupon", "dhaO3", "editorial", "factsUp", "financialfuture", "hemange", "hemangeolC", "influence", "makeAPurchase", "notice", "offer", "receiveSample", "receiveSamples", "review", "review2", "stonyfie", "receive Sample", "title", "gender1", "gender2", "gender3"
                                                                recLength += 1
                                                            Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " 1" & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                          Case "state", "dob", "dob1", "dob2", "dob3", "edd"
                                                            recLength += dataCount
                                                            Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString.Length & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                        Case "issue", "healthCare", "include1"
                                                            If dataCount >= 4 Then
                                                                recLength += .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString.Length
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString.Length & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            ElseIf dataCount > 0 And dataCount < 4 Then
                                                                recLength += dataCount + 1
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount + 1 & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            Else
                                                                recLength += 1
                                                            End If
                                                        Case "topic"
                                                            If dataCount >= 7 Then
                                                                recLength += dataCount
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            ElseIf dataCount > 0 And dataCount < 7 Then
                                                                recLength += dataCount + 1
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount + 1 & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            Else
                                                                recLength += 1
                                                            End If
                                                        Case "zipCode", "visit", "commonReason"
                                                            If dataCount >= 5 Then
                                                                recLength += dataCount
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            ElseIf dataCount > 0 And dataCount < 5 Then
                                                                recLength += dataCount + 1
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount + 1 & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            Else
                                                                recLength += 1
                                                            End If
                                                        Case "include2", "include3", "include4", "include5"
                                                            If dataCount >= 3 Then
                                                                recLength += dataCount
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            ElseIf dataCount > 0 And dataCount < 3 Then
                                                                recLength += dataCount + 1
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount + 1 & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            Else
                                                                recLength += 1
                                                            End If
                                                        Case Else
                                                            recLength += dataCount + 1
                                                            Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount + 1 & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                    End Select
                                                End If
                                            Next
                                        Case "Entry2"
                                            For l As Integer = 1 To cDS.Tables(0).Columns.Count - 1
                                                 If cDS.Tables(0).Rows(0).Item(l).ToString = "1" Then
                                                    Dim dataCount As Integer = .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString.Length
                                                    recLength += dataCount + 1
                                                    Debug.Print(cDS.Tables(0).Columns(l).ColumnName)
                                                End If
                                             Next
                                        Case "Compare"
                                            For l As Integer = 1 To cDS.Tables(0).Columns.Count - 1
                                                If cDS.Tables(0).Rows(0).Item(l).ToString = "1" Then
                                                    Dim dataCount As Integer = .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString.Length
                                                    Select Case cDS.Tables(0).Columns(l).ColumnName
                                                        Case "state", "dob", "dob1", "dob2", "dob3", "edd", "qasFlag", "commonReason", "healthCare", "topic", "issue", "include1", "include2", "include3", "include4", "include5", "advertisement", "1stBaby", "HeartSmart", "hibiclens", "qasFlag", "adt", "advetori", "axa", "BOBA", "consider", "content", "coupon", "dhaO3", "editorial", "factsUp", "financialfuture", "hemange", "hemangeolC", "influence", "makeAPurchase", "notice", "offer", "receiveSample", "receiveSamples", "review", "review2", "stonyfie", "receive Sample", "title", "gender1", "gender2", "gender3"
                                                            recLength += dataCount
                                                        Case "phone", "emailadd"
                                                            recLength += (dataCount + 1) * 2
                                                        Case "issue", "healthCare", "include1"
                                                            If dataCount >= 4 Then
                                                                recLength += .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString.Length
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString.Length & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            ElseIf dataCount > 0 And dataCount < 4 Then
                                                                recLength += dataCount + 1
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount + 1 & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            Else
                                                                recLength += 1
                                                            End If
                                                        Case "topic"
                                                            If dataCount >= 7 Then
                                                                recLength += dataCount
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            ElseIf dataCount > 0 And dataCount < 7 Then
                                                                recLength += dataCount + 1
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount + 1 & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            Else
                                                                recLength += 1
                                                            End If
                                                        Case "zipCode", "visit", "commonReason"
                                                            If dataCount >= 5 Then
                                                                recLength += dataCount
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            ElseIf dataCount > 0 And dataCount < 5 Then
                                                                recLength += dataCount + 1
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount + 1 & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            Else
                                                                recLength += 1
                                                            End If
                                                        Case "include2", "include3", "include4", "include5"
                                                            If dataCount >= 3 Then
                                                                recLength += dataCount
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            ElseIf dataCount > 0 And dataCount < 3 Then
                                                                recLength += dataCount + 1
                                                                Debug.Print(cDS.Tables(0).Columns(l).ColumnName & " " & dataCount + 1 & " " & .Rows(k).Item(cDS.Tables(0).Columns(l).ColumnName).ToString)
                                                            Else
                                                                recLength += 1
                                                            End If
                                                        Case Else
                                                            recLength += dataCount + 1
                                                    End Select
                                                    Debug.Print(cDS.Tables(0).Columns(l).ColumnName)
                                                End If
                                            Next
                                    End Select
                                    insertCount(recLength, k + 1)
skip1:                          End With
                                inRec += 1
                                inKey += recLength
                                totalrecord += 1
                                totalKS += recLength
                            Next
                            ProgressBar2.Value += 1
                            Application.DoEvents()
                        Next
                    Next
                    ProgressBar1.Value += 1
                    Application.DoEvents()
                    inaveKey = (inKey / inRec).ToString("F2")
                    writeCNT(KSloc & "\" & jobName & "_" & dateTime & "_" & database, batchname, inRec, inKey, inaveKey)
                Next
                aveKS = (totalKS / validRecord).ToString("F2")
                txtAveKeyStroke.Text = aveKS.ToString("N2")
                txtTotalValRec.Text = validRecord
                txtTotalRec.Text = totalrecord
                txtTotalKeyStroke.Text = totalKS.ToString("N2")
                txtTotalPullout.Text = totalPullout
                finishWriteCNT(KSloc & "\" & jobName & "_" & dateTime & "_" & database, totalrecord, validRecord, totalPullout, aveKS.ToString("N2"), totalKS.ToString("N2"))
        End If
            valBox2("Done!!" & vbNewLine & "FileName: " & jobName & "_" & dateTime & "_" & database & vbNewLine)
        Else
            valBox2("No Selected Batches or MDB")
        End If
    End Sub
    Public Sub getCode(ByVal code As String, ByVal mdbtype As String)      
        cConn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "  Data Source=" & Application.StartupPath & "\CWC Codes.mdb" & ";User Id=admin;Password=;")
        If Not conn.State = ConnectionState.Open Then
            cConn.Open()
        End If
        cDS = New DataSet
        Dim da As New OleDbDataAdapter("Select * FROM [" & mdbtype & "] WHERE Code = '" & code & "'", cConn)
        da.Fill(cDS)
    End Sub

    Private Sub rbPerBatch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPerBatch.CheckedChanged

    End Sub
End Class
