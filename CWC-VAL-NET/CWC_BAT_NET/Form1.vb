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
    Private dblog As String
    Private batchCount As String
    Private batchSize As String
    Private batchLocation As String
    Private valID As String
    Private totalPullout As Integer
    Private conn As OleDb.OleDbConnection
    Private dateTime As String

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        Me.Close()
    End Sub
    Private Sub Label15_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label15.MouseHover
        Label15.ForeColor = Color.FromArgb(63, 28, 88)
        Label15.BackColor = Color.White
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
        getList(LV1, DirListBox1.Path)
    End Sub
    Public folders, subFolders As String()
    Public Sub getList(ByRef list1 As ListView, ByVal dirListPath As String)
        On Error Resume Next
        list1.Items.Clear()
        folders = IO.Directory.GetDirectories(dirListPath, "**", SearchOption.TopDirectoryOnly)
        For Each j As String In folders
            subFolders = IO.Directory.GetDirectories(j, "Compare", SearchOption.TopDirectoryOnly)
            For Each i As String In subFolders
                Dim mdb() = IO.Directory.GetFiles(i, "0*.mdb")
                ' For j As Integer = 0 To mdb.Length
                ' Dim bool As Boolean = selectTop(mdb(j))
                If mdb.Length > 0 Then
                    Dim li As ListViewItem = list1.Items.Add(i)
                End If
                '  Next
            Next
        Next
    End Sub
  
    Private Sub DriveListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DriveListBox1.SelectedIndexChanged
        DirListBox1.Path = DriveListBox1.Drive
    End Sub
    'Public Function CreateAccessDatabase(ByVal DatabaseFullPath As String) As Boolean
    '    Dim bAns As Boolean
    '    Dim cat As New ADOX.Catalog
    '    Try
    '        Dim sCreateString As String
    '        sCreateString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    '                "Data Source=" & DatabaseFullPath & ";" & _
    '                "Jet OLEDB:Engine Type=5"
    '        cat.Create(sCreateString)
    '        bAns = True
    '    Catch Excep As System.Runtime.InteropServices.COMException
    '        bAns = False
    '    Finally
    '        cat = Nothing
    '    End Try
    '    Return bAns
    'End Function
    Private Sub cmdAddAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddAll.Click
        For i As Integer = 0 To LV1.Items.Count - 1
            LV1.Items(i).Selected = True
            LV1.Select()
        Next
        cmdAdd.PerformClick()
    End Sub
    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not LV2.SelectedItems.Count = 0 Then
            If LV2.SelectedItems.Count > 0 Then
                For i As Integer = 0 To LV2.SelectedItems.Count - 1
                    Dim li1 As ListViewItem.ListViewSubItem = LV2.SelectedItems(i).SubItems(0)
                    Dim li As ListViewItem = LV1.Items.Add(li1.Text)
                Next
                For Each i As ListViewItem In LV2.SelectedItems
                    LV2.Items.Remove(i)
                Next
            Else
                valBox2("No Item selected")
            End If
        Else
            valBox2("No Items to be Removed")
        End If
    End Sub
    Private Sub cmdClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearAll.Click
        If Not LV2.Items.Count = 0 Then
            For i As Integer = 0 To LV2.Items.Count - 1
                Dim li1 As ListViewItem.ListViewSubItem = LV2.Items(i).SubItems(0)
                Dim li As ListViewItem = LV1.Items.Add(li1.Text)
            Next
            LV2.Items.Clear()
        Else
            valBox2("No Items to be Cleared")
        End If
    End Sub
    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LV1.DoubleClick
        For j As Integer = 0 To LV1.SelectedItems.Count - 1
            For i As Integer = 0 To LV2.Items.Count - 1
                If (LV1.SelectedItems.Item(j).SubItems(0)).Text = (LV2.Items(i).SubItems(0)).Text Then
                    valBox2(" Selected Item is alreday selected")
                    Exit Sub
                End If
            Next i
        Next j
        For i As Integer = 0 To LV1.SelectedItems.Count - 1
            Dim li1 As ListViewItem.ListViewSubItem = LV1.SelectedItems(i).SubItems(0)

            Dim li As ListViewItem = LV2.Items.Add(li1.Text)

        Next
        For Each i As ListViewItem In LV1.SelectedItems
            LV1.Items.Remove(i)
        Next
    End Sub
    Private Sub ListView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LV2.DoubleClick
        For i As Integer = 0 To LV2.SelectedItems.Count - 1
            Dim li1 As ListViewItem.ListViewSubItem = LV2.SelectedItems(i).SubItems(0)
            Dim li As ListViewItem = LV1.Items.Add(li1.Text)
        Next
        For Each i As ListViewItem In LV2.SelectedItems
            LV2.Items.Remove(i)
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
        For i As Integer = 0 To LV2.Items.Count - 1
            imgcount += LV2.Items(i).SubItems(1).Text
        Next
        Return imgcount
    End Function
    Public Function valBox2(ByVal text As String)
        valBox2 = ""
        erBox.Label1.Text = text
        erBox.ShowDialog()
    End Function
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        For j As Integer = 0 To LV1.SelectedItems.Count - 1
            For i As Integer = 0 To LV2.Items.Count - 1
                If (LV1.SelectedItems.Item(j).SubItems(0)).Text = (LV2.Items(i).SubItems(0)).Text Then
                    valBox2(" Selected Item is already selected")
                    Exit Sub
                End If
            Next i
        Next j
        If Not LV1.SelectedItems.Count = 0 Then
            If LV1.SelectedItems.Count > 0 Then
                For i As Integer = 0 To LV1.SelectedItems.Count - 1
                    Dim li1 As ListViewItem.ListViewSubItem = LV1.SelectedItems(i).SubItems(0)
                    Dim li As ListViewItem = LV2.Items.Add(li1.Text)
                Next
                For Each i As ListViewItem In LV1.SelectedItems
                    LV1.Items.Remove(i)
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
  
    Public Function selectTop(ByVal path As String, Optional ByVal frombrowse As Boolean = False) As Integer
        Dim ds2 As New DataSet
        Dim da2 As New OleDbDataAdapter
        Dim temp As Integer
        openmdb(path)
        Try
            ds2 = New DataSet
            da2 = New OleDbDataAdapter("SELECT TOP 1 RecNum FROM " & jobName & " WHERE  QCFlag = '0'", conn)
            da2.Fill(ds2)
            temp = ds2.Tables(0).Rows(0).Item(0)
        Catch ex As Exception
            conn.Close()
            Return True
        End Try
        conn.Close()
        Return False
    End Function
    Public Sub openmdb(ByVal connectionpath As String)
        If Not conn Is Nothing Then
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End If
        conn = New OleDbConnection
        conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & connectionpath
        conn.Open()
    End Sub

    Public Function createConnection(ByVal DatabaseFullPath As String)
        conn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & DatabaseFullPath & ";User Id=admin;Password=;")
        Return conn
    End Function
    Public Sub insertFlag(ByVal mdbpath As String, ByVal recNum As Integer, ByVal valDate As String, ByVal valid As String)
        '  createConnection(mdbpath)
        Dim sqlQry As String = "Update " & jobName & " Set Valflag = @keyflag, ValID = @keyid, ValDate = @keydate where RecNum=@recnum"
        Dim cmd As New OleDbCommand(sqlQry, conn)
        cmd.Parameters.AddWithValue("@keyflag", "*")
        cmd.Parameters.AddWithValue("@keyid", valid)
        cmd.Parameters.AddWithValue("@keydate", valDate)
        cmd.Parameters.AddWithValue("@recnum", recNum)
        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()
    End Sub
    Public Function checkValFlag() As Integer
        Try
            Dim comDs As New DataSet
            Dim comDa As New OleDbDataAdapter("SELECT TOP 1 RecNum FROM " & jobName & " WHERE  ValFlag = '*'", conn)
            comDa.Fill(comDs)
            If comDs.Tables(0).Rows(0).Item(0).ToString = "" Then
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Private Sub cmdVal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVal.Click
        If Not LV2.Items.Count = 0 Then
            Dialog1.ShowDialog()
            valID = Dialog1.valID
            Dim totalRecord As Integer
            PB1.Value = 0
            PB1.Minimum = 0
            PB1.Maximum = LV2.Items.Count
            totalRecord = 0
            totalPullout = 0
            For batch As Integer = 0 To LV2.Items.Count - 1
                Dim mdbs As String() = Directory.GetFiles(LV2.Items(batch).Text, "*.mdb")
                PB2.Value = 0
                PB2.Minimum = 0
                PB2.Maximum = mdbs.Length
                For i As Integer = 0 To mdbs.Length - 1
                    createConnection(mdbs(i))
                    Dim mdbpath As New DirectoryInfo(mdbs(i))
                    Dim errpath As String = mdbpath.Parent.FullName & "\" & IO.Path.GetFileNameWithoutExtension(mdbs(i))
                    'Dim bool As Boolean = checkValFlag()
                    'If bool = True Then
                    '    Dim ask As Boolean
                    '    ask = valBox("File: " & LV2.Items(i).Text & " Has Been Validated" & vbNewLine & "Do You want To Re-Validate")
                    '    If ask = False Then
                    '        Exit Sub
                    '    End If
                    'End If
                    startLog(errpath, valID, "CWC(YYF_RSG)")
                    Dim ds As New DataSet
                    Dim da As New OleDbDataAdapter("select * from Data001", conn)
                    da.Fill(ds)
                       Label6.Text = 0 & " / " & ds.Tables(0).Rows.Count
                    For k As Integer = 0 To ds.Tables(0).Rows.Count - 1
                        Dim PO1 As String = ds.Tables(0).Rows(k).Item("Code").ToString
                        Dim PO2 As String = ds.Tables(0).Rows(k).Item("Title").ToString
                        For j As Integer = 1 To ds.Tables(0).Columns.Count - 1
                            Dim fieldname As String = ds.Tables(0).Columns(j).ColumnName
                            Dim fieldValue As String = ds.Tables(0).Rows(k).Item(j).ToString
                            validation(errpath, k + 1, fieldname, fieldValue)
                            If PO1 = "*" Then
                                totalPullout += 1
                                validation(errpath, k + 1, "Code", PO1)
                                Exit For
                            ElseIf PO2 = "*" Then
                                totalPullout += 1
                                validation(errpath, k + 1, "title", PO2)
                                Exit For
                            End If
                        Next
                        '  insertFlag(mdbs(i), ds.Tables(0).Rows(k).Item(0).ToString, dateTime, valID)
                        totalRecord += 1
                    Next
                        If Not errtxt = "" Then ifValidationAppear2(errpath, errtxt)
                        errtxt = ""
                        PB2.Value += 1 '(ds.Tables(0).Rows.Count / 100)
                        Label6.Text = i + 1 & " / " & mdbs.Length
                    Next
                    PB1.Value += 1 '(ListView2.Items.Count / 100)
                    Label5.Text = batch + 1 & " / " & LV2.Items.Count
                Next
                valBox2("Done")
                TextBox1.Text = totalRecord
                TextBox2.Text = totalPullout
        Else
                valBox2("No Batches To Validate")
        End If
    End Sub
    Public errtxt As String
    Public Sub ifValidationAppear(ByVal logPath As String, ByVal record As String, ByVal field As String, ByVal description As String, ByVal value As String)
        ' If My.Computer.FileSystem.FileExists(logPath & ".err") Then
        'Dim log As String = My.Computer.FileSystem.ReadAllText(logPath & ".err")
        'Dim logStr As String
        record = addLength(record, 9)
        field = addLength(field, 26)
        description = addLength(description, 60)
        value = addLength(value, 50)
        errtxt &= record & field & description & value & vbNewLine
        '   My.Computer.FileSystem.WriteAllText(logPath & ".err", logStr, False)
        ' End If
    End Sub
    Public Sub ifValidationAppear2(ByVal logPath As String, ByVal txt As String)
        If My.Computer.FileSystem.FileExists(logPath & ".err") Then
            My.Computer.FileSystem.WriteAllText(logPath & ".err", txt, False)
        End If
    End Sub
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        dateTime = Now.ToString("G")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Public Function validation(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal value As String) As Boolean
        Select Case fieldname
            Case "address"
                address(errpath, rec, fieldname, value)
            Case "city"
                city(errpath, rec, fieldname, value)
            Case "commonReason", "visit", "issue", "milk", "topic", "healthCare", "include1", "include2", "include3", "include4", "include5"
                commonReason(errpath, rec, fieldname, value)
            Case "dob", "dob1", "dob2", "dob3"
                CheckDate(errpath, rec, fieldname, value, "MMddyy")
            Case "edd"
                CheckDate(errpath, rec, fieldname, value, "MMddyy", , , , , , , , 1)
              Case "emailadd"
                email(errpath, rec, fieldname, value)
            Case "fname"
                fname(errpath, rec, fieldname, value)
            Case "other"
                other(errpath, rec, fieldname, value)
            Case "lname"
                lname(errpath, rec, fieldname, value)
            Case "suffix"
                suffix(errpath, rec, fieldname, value)
            Case "mname"
                mname(errpath, rec, fieldname, value)
            Case "phone"
                phone(errpath, rec, fieldname, value)
            Case "state"
                state(errpath, rec, fieldname, value)
            Case "title"
                title(errpath, rec, fieldname, value)
            Case "zipcode"
                zipCode(errpath, rec, fieldname, value)
            Case "Code"
                code(errpath, rec, fieldname, value)
            Case "qasFlag"
                qasflag(errpath, rec, fieldname, value)
            Case "advertisement", "1stBaby", "HeartSmart", "hibiclens", "qasFlag", "adt", "advetori", "axa", "BOBA", "consider", "content", "coupon", "dhaO3", "editorial", "factsUp", "financialfuture", "hemange", "hemangeolC", "influence", "makeAPurchase", "notice", "offer", "receiveSample", "receiveSamples", "review", "review2", "stonyfie"
                alphaOrNum(errpath, rec, fieldname, value, 2, "12")
            Case "healthCare", "include1"
                alphaOrNum(errpath, rec, fieldname, value, 2, "1234")
            Case "issue", "include2", "include3", "include4", "include5"
                alphaOrNum(errpath, rec, fieldname, value, 2, "123")
            Case "gender1", "gender2", "gender3"
                alphaOrNum(errpath, rec, fieldname, value, 2, "12")
        End Select
    End Function
    Public Function alphaOrNum(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String, ByVal type As Integer, Optional ByVal pick As String = "ynYN", Optional ByVal additionalChar As String = "") As Boolean
        ''[0] - Numeric
        ''[1] - Alpha
        ''[2] - Specify
        ''[3] - All
        Dim alpha As String = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & Chr(10) & Chr(11) & Chr(12) & Chr(13) & additionalChar
        Dim num As String = "0123456789" & Chr(10) & Chr(11) & Chr(12) & Chr(13) & additionalChar
        Dim sp As String = pick & Chr(10) & Chr(11) & Chr(12) & Chr(13) & additionalChar
        Dim all As String = "0123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & Chr(10) & Chr(11) & Chr(12) & Chr(13) & additionalChar
        For i As Integer = 0 To txt.Length - 1
            Select Case type
                Case 0
                    If num.Contains(txt.Substring(i, 1)) = True Then
                        Return True
                    Else
                        ifValidationAppear(errpath, rec, fieldname, "  ", txt)
                    End If
                Case 1
                    If alpha.Contains(txt.Substring(i, 1)) Then
                        Return True
                    Else
                        ifValidationAppear(errpath, rec, fieldname, "  ", txt)
                    End If
                Case 2
                    If sp.Contains(txt.Substring(i, 1)) Then
                        Return True
                    Else
                        ifValidationAppear(errpath, rec, fieldname, "  ", txt)
                    End If
                Case 3
                    If all.Contains(txt.Substring(i, 1)) Then
                        Return True
                    Else
                        ifValidationAppear(errpath, rec, fieldname, "  ", txt)
                    End If
            End Select
        Next
    End Function
    Public Function code(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String)
        code = ""
        If txt = "*" Then
            ifValidationAppear(errpath, rec, fieldname, "Pullout", txt)
        Else
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            If Char.IsUpper(tmp) = False Then ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            If txt.Length > 1 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
            alphaOrNum(errpath, rec, fieldname, txt, 2, "EFGHIJKLefghijkl*")
        End If
    End Function
    Public Function email(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        '''' EMAIL ADDRESS'''
        If Not txt = "" Then
            If Not Regex.IsMatch(txt, "\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase) Then ifValidationAppear(errpath, rec, fieldname, "Invalid Email", txt)
            'If txt.Contains("  ") Then ifValidationAppear(errpath, rec, fieldname, "Invalid Email", txt)
            'If txt.Contains("..") Then ifValidationAppear(errpath, rec, fieldname, "Invalid Email", txt)
            'If txt.Contains("@@") Then ifValidationAppear(errpath, rec, fieldname, "Invalid Email", txt)
            'If txt.Length > 50 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Email", txt)
            alphaOrNum(errpath, rec, fieldname, txt, 3, , "@&()_'./")
        Else
            Return True
        End If
        ''EMAIL ADDRESS '' END''
    End Function
    Public Function fname(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''FIRSTNAME''''
        If Not txt = "" Then
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            If Char.IsUpper(tmp) = False Then ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            If txt.Contains("  ") Then ifValidationAppear(errpath, rec, fieldname, "Contains Multiple Spaces", txt)
            If txt.Length > 70 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
        End If
        alphaOrNum(errpath, rec, fieldname, txt, 1, , "-'")
        ''FIRSTNAME'' END''
    End Function
    Public Function other(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''FIRSTNAME''''
        If Not txt = "" Then
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            If Char.IsUpper(tmp) = False Then ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            If txt.Contains("  ") Then ifValidationAppear(errpath, rec, fieldname, "Contains Multiple Spaces", txt)
            If Regex.IsMatch(txt, "^[0-9 ]+$") = True Then ifValidationAppear(errpath, rec, fieldname, "Pls Check Value is All numeric", txt)
            If txt.Length = 1 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
        End If
        alphaOrNum(errpath, rec, fieldname, txt, 1, , "-'")
        ''FIRSTNAME'' END''
    End Function
    Public Function lname(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''SURNAME''''
        If txt = "" Then
            ifValidationAppear(errpath, rec, fieldname, "Should Not Be Empty", txt)
        Else
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            If Char.IsUpper(tmp) = False Then ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            If txt.Contains("  ") = True Then ifValidationAppear(errpath, rec, fieldname, "Contains Double Spaces", txt)
            If txt.Length = 1 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
            If txt.Length > 30 Then ifValidationAppear(errpath, rec, fieldname, "  ", txt)
        End If
        ''SURNAME'' END ''
    End Function
    Public Function suffix(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''SURNAME''''
        If txt = "" Then
            ' ifValidationAppear(errpath, rec, fieldname, "Should Not Be Empty", txt)
        Else
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            If Char.IsUpper(tmp) = False Then ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            If txt.Contains("  ") = True Then ifValidationAppear(errpath, rec, fieldname, "Contains Double Spaces", txt)
            If txt.Length = 1 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
          End If
        ''SURNAME'' END ''
    End Function
    Public Function mname(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''FIRSTNAME''''
        If Not txt = "" Then
            If txt.Contains("  ") Then
                ifValidationAppear(errpath, rec, fieldname, "Contains Multiple Spaces", txt)
            End If
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            If Char.IsUpper(tmp) = False Then
                ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            End If
            alphaOrNum(errpath, rec, fieldname, txt, 1, , "-'")
            If txt.Length > 30 Then
                ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
            End If
        End If
        ''FIRSTNAME'' END''
    End Function
    Public Function state(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''STATE''''
        If Not txt = "" Then
            If txt.Contains("  ") = True Then
                ifValidationAppear(errpath, rec, fieldname, "Contains Multiple Spaces", txt)
            End If
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            If Char.IsUpper(tmp) = False Then
                ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            End If
            If txt.Length > 70 Then
                ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
            End If
        End If
        ''STATE'' END''
    End Function
    Public Function zipCode(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        If Not txt = "" Then
            If txt.Length > 15 Then
                ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
            End If
        End If
    End Function
    Public Function phone(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''CONTACTNUMBER''''
        If Not txt = "" Then
            If Not txt.Length <> 10 Then   ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)

            'Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            'If Not tmp = "" Then
            '    ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            'End If
        End If
        ''CONTACTNUMBER'' END ''
    End Function
    Public Function title(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''TITLE''
        If txt = "*" Then
            ifValidationAppear(errpath, rec, fieldname, "Pullout", txt)
        Else
            If Not txt = "" Then
                Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
                If Char.IsUpper(tmp) = False Then  If Char.IsLetter(txt) = True Then ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
                If txt.Length > 10 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
                Select Case txt
                    Case "MR", "MRS", "MISS", "MS", "MASTER", "LADY", "REV", "DR", "MAJOR", "SIR", "CAPT", "PROF", "BRIG", "CPL", ""
                        Return True
                    Case Else
                        ifValidationAppear(errpath, rec, fieldname, "Invalid Title", txt)
                End Select
            End If
        End If
        ''TITLE'' END''
    End Function
    Public Function qasflag(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        If txt.Replace(Chr(13), "") = "" Then
            ifValidationAppear(errpath, rec, fieldname, "Should Not Be Empty", txt)
        Else
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            alphaOrNum(errpath, rec, fieldname, txt, 2)
            If Char.IsUpper(tmp) = False Then ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            If txt.Length > 1 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
        End If
    End Function
    Public Function address(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''STREET''''
        If txt = "" Then
            ifValidationAppear(errpath, rec, fieldname, "Should Not Be Empty", txt)
        Else
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            If txt.Contains("  ") = True Then ifValidationAppear(errpath, rec, fieldname, "Contains Multiple Spaces", txt)
            If Char.IsUpper(tmp) = False Then ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            If txt.Length > 70 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
        End If
        ''STREET'' END''
    End Function
    Public Function city(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        ''''STREET''''
        If txt = "" Then
            ifValidationAppear(errpath, rec, fieldname, "Should Not Be Empty", txt)
        Else
            Dim tmp As String = Regex.Replace(txt, "[^A-Za-z]", "")
            If txt.Contains("  ") = True Then ifValidationAppear(errpath, rec, fieldname, "Contains Multiple Spaces", txt)
            If Char.IsUpper(tmp) = False Then ifValidationAppear(errpath, rec, fieldname, "Invalid Case", txt)
            If txt.Length > 30 Then ifValidationAppear(errpath, rec, fieldname, "Invalid Length", txt)
        End If
        ''STREET'' END''
    End Function
    Public Function commonReason(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByVal txt As String) As Boolean
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        Dim d As Integer
        Dim e As Integer
        Dim f As Integer
        Dim g As Integer
        Dim h As Integer
        Dim j As Integer
        For i As Integer = 0 To txt.Length - 1
            Select Case txt.Substring(i, 1)
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
        If a > 1 Or b > 1 Or c > 1 Or d > 1 Or e > 1 Or f > 1 Or g > 1 Or h > 1 Or j > 1 Then ifValidationAppear(errpath, rec, fieldname, "Contains Duplucate", txt)
        If txt.Length > 5 Then ifValidationAppear(errpath, rec, fieldname, "Invalid length", txt)
    End Function
    Public Sub CheckDate(ByVal errpath As String, ByVal rec As String, ByVal fieldname As String, ByRef value As String, ByVal formats As String, _
    Optional ByVal dayRng As Integer = 0, Optional ByVal monthRng As Integer = 0, Optional ByVal yearRng As Integer = 0, _
    Optional ByVal dayRng2 As Integer = 0, Optional ByVal monthRng2 As Integer = 0, Optional ByVal yearRng2 As Integer = 0, _
    Optional ByVal dayDD As Integer = 0, Optional ByVal monthDD As Integer = 0, Optional ByVal yearDD As Integer = 0)

        If value = "" Then
            Exit Sub
        End If
        Dim isValOk As String = ""
        Dim yyyy As String = ""
        Dim dd As String = ""
        Dim mm As String = ""
        Dim yy As String = ""
        Dim date1, date2, date3 As Date
        If value.Length = 6 Then
            mm = value.Substring(0, 2)
            dd = value.Substring(2, 2)
            yy = value.Substring(4, 2)

            If (dd.Trim.Equals("") And mm.Trim.Equals("") And yy.Trim.Equals("")) Or _
              (Val(dd) = 0 Or Val(mm) = 0) Then
                ifValidationAppear(errpath, rec, fieldname, "Date is invalid", value)
                Exit Sub
            Else
                'Month
                If Val(mm) > 12 Then
                    ifValidationAppear(errpath, rec, fieldname, "Invalid Month Detected", value)
                    Exit Sub
                Else
                    'Day
                    Select Case Val(mm)
                        Case 1, 3, 5, 7, 8, 10, 12
                            If Val(dd) > 31 Then
                                ifValidationAppear(errpath, rec, fieldname, "Invalid Day Detected", value)
                                Exit Sub
                            End If
                        Case 2
                            Dim leaf = Val(mm) Mod 4
                            If leaf = 0 Then
                                If Val(dd) > 29 Then
                                    ifValidationAppear(errpath, rec, fieldname, "Invalid Day Detected", value)
                                    Exit Sub
                                End If
                            Else
                                If Val(dd) > 28 Then
                                    ifValidationAppear(errpath, rec, fieldname, "Invalid Day Detected", value)
                                    Exit Sub
                                End If
                            End If
                        Case Else
                            If Val(dd) > 30 Then
                                ifValidationAppear(errpath, rec, fieldname, "Invalid Day Detected", value)
                                Exit Sub
                            End If
                    End Select
                    'Year
                    Dim dd2 As String
                    Dim mm2 As String
                    Dim yy2 As String
                    If yy > Val(Now.ToString("yy")) Then yyyy = "19" & yy Else yyyy = "20" & yy
                    'For Date Ranges
                    If Not dayRng = 0 Or Not monthRng = 0 Or Not yearRng = 0 Then
                        date1 = Date.ParseExact(dd & mm & yyyy, "ddMMyyyy", Nothing)
                        date2 = getExactDate(dayRng, monthRng, yearRng, "ddMMyyyy")
                        If date1 < date2 Then
                            ifValidationAppear(errpath, rec, fieldname, "Date is Invalid", value)
                        ElseIf mm.Trim <> "" And Val(mm) <> 0 And dd.Trim <> "" And Val(dd) <> 0 Then
                            date1 = Date.ParseExact(Format(Now, "ddMMyyyy"), "ddMMyyyy", Nothing)
                            date2 = Date.ParseExact(dd & mm & yyyy, "ddMMyyyy", Nothing)
                            date3 = getExactDate(dayRng2, monthRng2, yearRng2, "ddMMyyyy")
                            If date2.Date > date1.Date Then
                                ifValidationAppear(errpath, rec, fieldname, "pls check future dates", value)
                                Exit Sub
                            ElseIf date2 > date3 Then
                                ifValidationAppear(errpath, rec, fieldname, "Date is Invalid", value)
                            End If
                        End If
                        'For Due Dates
                    ElseIf Not dayDD = 0 And Not monthDD = 0 And Not yearDD = 0 Then
                        dd2 = Now.ToString("dd")
                        mm2 = Now.ToString("MM")
                        yy2 = Now.ToString("yy")
                        If mm2 + monthDD > 12 Then
                            mm2 = Val(mm2 + monthDD - 12)
                            yy2 += 1
                        Else
                            mm2 += monthDD
                        End If
                        If dd2 + dayDD > 12 Then
                            dd2 = Val(dd2 + dayDD - 30)
                            mm2 += 1
                        Else
                            dd2 += dayDD
                        End If
                        If mm2.Length = 1 Then mm2 = "0" & mm2
                        date2 = Date.ParseExact(mm & dd & yy, "MMddyy", Nothing)
                        date1 = Date.ParseExact(mm2 & dd2 & yy2 + yearDD, "MMddyy", Nothing)
                        If date2.Date > date1.Date Or date2.Date < Date.ParseExact(Format(Now, "MMddyy"), "MMddyy", Nothing) Then
                            ifValidationAppear(errpath, rec, fieldname, "Date is Invalid", value)
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Public Function getExactDate(ByVal ddRng As String, ByVal mmRng As String, ByVal yyRng As String, ByVal format As String) As Date
        Dim dd, mm, yy As String
        yy = Now.ToString("yyyy") - yyRng
        If Now.ToString("MM") - mmRng < 0 Then
            yy -= 1
            mm = (Now.ToString("MM") + 12) - mmRng
        Else
            mm = Now.ToString("MM") - mmRng
        End If
        If Now.ToString("dd") - ddRng < 0 Then
            mm -= 1
            dd = (Now.ToString("dd") + 30) - ddRng
        Else
            dd = Now.ToString("dd") - ddRng
        End If
        If dd = "0" Then dd = "30"
        If mm = "0" Then mm = "12"
        If dd.Length = 1 Then dd = "0" & dd
        If mm.Length = 1 Then mm = "0" & mm
        getExactDate = Date.ParseExact(dd & mm & yy, format, Nothing)
        Return getExactDate
    End Function


End Class
