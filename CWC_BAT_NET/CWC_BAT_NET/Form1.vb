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

Public Class Form1
    Private mouseDownLoc As Point
    Private currentPath As String = ""
    Private clientname As String = "CWC"
    Private jobName As String = "YYF_RSG"
    Private folders() As String
    Private CWCConn As OleDb.OleDbConnection
    Private cbatchConn As OleDb.OleDbConnection
    Private dblog As String
    Private batchCount As String
    Private batchSize As String
    Private batchLocation As String
    Private opID As String

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click
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
    Private Sub Label16_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label16.MouseHover
        Label16.ForeColor = Color.Transparent
        Label16.BackColor = Color.White
    End Sub
    Private Sub Label16_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label16.MouseLeave
        Label16.BackColor = Color.Transparent
        Label16.ForeColor = Color.White
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
        folders = IO.Directory.GetDirectories(DirListBox1.Path, "**", SearchOption.TopDirectoryOnly)
        For i As Integer = 0 To folders.Length - 1
            Dim imgCount() As String = IO.Directory.GetFiles(folders(i), "*.tif")
            Dim cntFiles As Integer = imgCount.Length
            If cntFiles > 0 Then
                Dim li As ListViewItem = ListView1.Items.Add(folders(i))
                li.SubItems.Add(cntFiles)
            End If
        Next
    End Function
    Private Sub DriveListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DriveListBox1.SelectedIndexChanged
        DirListBox1.Path = DriveListBox1.Drive
    End Sub
    Public Function createCWCConnection(ByVal DatabaseFullPath As String)
        CWCConn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & DatabaseFullPath & ";User Id=admin;Password=;")
        Return CWCConn
    End Function
    Public Function createCbatchConnection(ByVal DatabaseFullPath As String)
        cbatchConn = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "  Data Source=" & DatabaseFullPath & ";User Id=admin;Password=;")
        Return cbatchConn
    End Function
    Public Function CreateAccessDatabase(ByVal DatabaseFullPath As String) As Boolean
        Dim bAns As Boolean
        Dim cat As New ADOX.Catalog
        Try
            Dim sCreateString As String
            sCreateString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & DatabaseFullPath & ";" & _
                    "Jet OLEDB:Engine Type=5"
            cat.Create(sCreateString)
            bAns = True
        Catch Excep As System.Runtime.InteropServices.COMException
            bAns = False
        Finally
            cat = Nothing
        End Try
        Return bAns
    End Function
    Private Sub cmdAddAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddAll.Click
        For i As Integer = 0 To ListView1.Items.Count - 1
            ListView1.Items(i).Selected = True
            ListView1.Select()
        Next
        cmdAdd.PerformClick()
    End Sub
    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        If Not ListView2.SelectedItems.Count = 0 Then
            If ListView2.SelectedItems.Count > 0 Then
                For i As Integer = 0 To ListView2.SelectedItems.Count - 1
                    Dim li1 As ListViewItem.ListViewSubItem = ListView2.SelectedItems(i).SubItems(0)
                    Dim li2 As ListViewItem.ListViewSubItem = ListView2.SelectedItems(i).SubItems(1)
                    Dim li As ListViewItem = ListView1.Items.Add(li1.Text)
                    li.SubItems.Add(li2.Text)
                Next
                For Each i As ListViewItem In ListView2.SelectedItems
                    ListView2.Items.Remove(i)
                Next
            Else
                MsgBox("No Item selected")
            End If
        Else
            MsgBox("No Items to be Removed")
        End If
    End Sub
    Private Sub cmdClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearAll.Click
        If Not ListView2.Items.Count = 0 Then
            For i As Integer = 0 To ListView2.Items.Count - 1
                Dim li1 As ListViewItem.ListViewSubItem = ListView2.Items(i).SubItems(0)
                Dim li2 As ListViewItem.ListViewSubItem = ListView2.Items(i).SubItems(1)
                Dim li As ListViewItem = ListView1.Items.Add(li1.Text)
                li.SubItems.Add(li2.Text)
            Next
            ListView2.Items.Clear()
        Else
            MsgBox("No Items to be Cleared")
        End If
    End Sub
    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        For j As Integer = 0 To ListView1.SelectedItems.Count - 1
            For i As Integer = 0 To ListView2.Items.Count - 1
                If (ListView1.SelectedItems.Item(j).SubItems(0)).Text = (ListView2.Items(i).SubItems(0)).Text Then
                    MsgBox(" Selected Item is alreday selected")
                    Exit Sub
                End If
            Next i
        Next j
        For i As Integer = 0 To ListView1.SelectedItems.Count - 1
            Dim li1 As ListViewItem.ListViewSubItem = ListView1.SelectedItems(i).SubItems(0)
            Dim li2 As ListViewItem.ListViewSubItem = ListView1.SelectedItems(i).SubItems(1)

            Dim li As ListViewItem = ListView2.Items.Add(li1.Text)
            li.SubItems.Add(li2.Text)
        Next
        For Each i As ListViewItem In ListView1.SelectedItems
            ListView1.Items.Remove(i)
        Next
    End Sub
    Private Sub ListView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView2.DoubleClick
        For i As Integer = 0 To ListView2.SelectedItems.Count - 1
            Dim li1 As ListViewItem.ListViewSubItem = ListView2.SelectedItems(i).SubItems(0)
            Dim li2 As ListViewItem.ListViewSubItem = ListView2.SelectedItems(i).SubItems(1)
            Dim li As ListViewItem = ListView1.Items.Add(li1.Text)
            li.SubItems.Add(li2.Text)
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
        If batchSize >= imgCount Then
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
    Public Sub CreateCBatchCSV()
        'Open a connection to the database
        Dim dbCommand As OleDb.OleDbCommand
        Dim dbAdapter As OleDb.OleDbDataAdapter
        'Open the connection
        createCbatchConnection(Application.StartupPath & "\CBATCH" & "\" & jobName & ".mdb")
        cbatchConn.Open()
        'Retrieve all data from the job table.
        dbCommand = New OleDb.OleDbCommand("SELECT * FROM [" & jobName & "]", cbatchConn)
        dbAdapter = New OleDb.OleDbDataAdapter
        dbAdapter.SelectCommand = dbCommand

        'Populate a DataTable with the contents of the Job table.
        Dim dbTable As New DataTable
        dbAdapter.Fill(dbTable)

        'Create the CSV file with column headers first.
        Dim sw As New System.IO.StreamWriter(Application.StartupPath & "\CBATCH\" & jobName & ".csv")
        Dim columnHeaders As New System.Text.StringBuilder

        For Each col As DataColumn In dbTable.Columns
            If columnHeaders.ToString.Length > 0 Then columnHeaders.Append(",")
            columnHeaders.Append("""" & col.ColumnName & """")
        Next

        'Write out the column headers
        sw.WriteLine(columnHeaders.ToString)

        'Add the data
        Dim dataValues As New System.Text.StringBuilder

        For Each row As DataRow In dbTable.Rows
            For Each col As DataColumn In dbTable.Columns
                If dataValues.ToString.Length > 0 Then dataValues.Append(",")
                dataValues.Append("""" & row(col.ColumnName).ToString & """")
            Next
            sw.WriteLine(dataValues.ToString)
            dataValues = New System.Text.StringBuilder
        Next
        sw.Close()
        cbatchConn.Close()
        cbatchConn.Dispose()
    End Sub
    Public Function createLog(ByVal cbatchpath As String, ByVal datetime As String, ByVal batchloc As String, ByVal dbrec As String, ByVal shipment As String)
        createLog = ""
        Dim txt1 As String
        txt1 = "Date:              " & datetime & vbNewLine & _
               "Job              : " & "(" & jobName & ")" & clientname & " " & vbNewLine & _
               "Shipment/ZipFile : " & shipment & vbNewLine & _
               "Batch By (OperId): " & opID & vbNewLine & _
               "Directory Entry1 : " & batchloc & "\Entry1" & vbNewLine & _
               "Directory Entry2 : " & batchloc & "\Entry2" & vbNewLine & _
               "Directory Compare: " & batchloc & "\Compare" & vbNewLine & _
               "Date Received    : __________________________________________" & vbNewLine & _
               "Output Date      : __________________________________________" & vbNewLine & _
               "DE Target Date   : __________________________________________" & vbNewLine & _
               "QC Target Date   : __________________________________________" & vbNewLine & _
               "Val Target Date  : __________________________________________" & vbNewLine & _
               "                                                                 " & vbNewLine & _
               "┌─────────┬────────────────┬───────────┬───────────┬───────────┐" & vbNewLine & _
               "│ Number  │ MDB Name       │ Start Rec │ End Rec   │ Rec Count │" & vbNewLine & _
               "├─────────┼────────────────┼───────────┼───────────┼───────────┤" & dbrec & vbNewLine & _
               "└─────────┴────────────────┴───────────┴───────────┴───────────┘"
        My.Computer.FileSystem.WriteAllText(cbatchpath & ".log", txt1, False)
    End Function
    Function addLength(ByVal str As String, ByVal len As Integer) As String
        For i As Integer = str.Length To len - 1
            str &= " "
        Next
        Return str
    End Function
    Private Sub txtTargetDir_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTargetDir.Click
        Dim askBatchLoc As DialogResult = FolderBrowserDialog1.ShowDialog()
        If askBatchLoc = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        ElseIf askBatchLoc = Windows.Forms.DialogResult.OK Then
            txtTargetDir.Text = FolderBrowserDialog1.SelectedPath
            batchLocation = txtTargetDir.Text
        End If
    End Sub
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        For j As Integer = 0 To ListView1.SelectedItems.Count - 1
            For i As Integer = 0 To ListView2.Items.Count - 1
                If (ListView1.SelectedItems.Item(j).SubItems(0)).Text = (ListView2.Items(i).SubItems(0)).Text Then
                    MsgBox(" Selected Item is already selected")
                    Exit Sub
                End If
            Next i
        Next j
        If Not ListView1.SelectedItems.Count = 0 Then
            If ListView1.SelectedItems.Count > 0 Then
                For i As Integer = 0 To ListView1.SelectedItems.Count - 1
                    Dim li1 As ListViewItem.ListViewSubItem = ListView1.SelectedItems(i).SubItems(0)
                    Dim li2 As ListViewItem.ListViewSubItem = ListView1.SelectedItems(i).SubItems(1)
                    Dim li As ListViewItem = ListView2.Items.Add(li1.Text)
                    li.SubItems.Add(li2.Text)
                Next
                For Each i As ListViewItem In ListView1.SelectedItems
                    ListView1.Items.Remove(i)
                Next
            Else
                MsgBox("No Item selected")
            End If
        Else
            MsgBox("No Available Items")
        End If
     
    End Sub
    Private Sub cmdBatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBatch.Click
        If Not ListView2.Items.Count = 0 Then


            If txtTargetDir.Text = "" Then
                Dim askBatchLoc As DialogResult = FolderBrowserDialog1.ShowDialog()
                If askBatchLoc = Windows.Forms.DialogResult.Cancel Then
                    Exit Sub
                ElseIf askBatchLoc = Windows.Forms.DialogResult.OK Then
                    txtTargetDir.Text = FolderBrowserDialog1.SelectedPath
                    batchLocation = txtTargetDir.Text
                End If
            End If
            opID = ""
            Dialog1.ShowDialog()
            opID = Dialog1.opID
            If opID = "" Then
                Exit Sub
            End If
            ProgressBar1.Maximum = countAllImages()
            ProgressBar1.Value = 0
            ''Start Process
            For batches As Integer = 0 To ListView2.Items.Count - 1

                Dim eRec As Integer = 0
                Dim sRec As Integer = 1
                Dim RecCnt As Integer = 0
                ''Get the Database Count
                Dim batchSize As Integer = TextBox1.Text
                Dim imgcount As Integer = ListView2.Items(batches).SubItems(1).Text

                ''Get the ZipFile, DvdName, BatchDate, and SeqNo1 
                Dim dateTime As String = Now.ToString("G")
                Dim dateTime2 As String = Now.ToString("yyyyddMM")
                Dim imgLoc As New DirectoryInfo(ListView2.Items(batches).SubItems(0).Text)
                If imgLoc.Name.Length < 17 Then
                    valBox2("Invalid Batch Name")
                    Exit Sub
                End If
                Dim dvdNameParent As String = imgLoc.Parent.Name
                Dim zipFile As String = imgLoc.Name
                Dim dvdName As String = imgLoc.Name.Substring(imgLoc.Name.Length - 8)

                lblProcFolder.Text = ListView2.Items(batches).SubItems(0).Text
                Dim cbatchconnpath As String = Application.StartupPath & "\CBATCH" + "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\" & imgLoc.Name & ".mdb"
                Application.DoEvents()

                IO.Directory.CreateDirectory(Application.StartupPath & "\CBATCH" + "\" & jobName & "\" & dateTime2 & "\" & dvdName)
                Dim bool As Boolean
                bool = createD2112nOrigImage(cbatchconnpath, ListView2.Items(batches).SubItems(0).Text, Application.StartupPath & "\CBATCH" + "\" & jobName & "\" & dateTime2 & "\" & dvdName)
                If bool = False Then
                    GoTo nextLoop
                End If

                IO.Directory.CreateDirectory(batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Compare")
                IO.Directory.CreateDirectory(batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Entry1")
                IO.Directory.CreateDirectory(batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Entry2")
                IO.Directory.CreateDirectory(batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Images")

                Dim dbcount As Integer = countDB(imgcount, batchSize)
                Dim imgNames() As String = IO.Directory.GetFiles(ListView2.Items(batches).SubItems(0).Text, "*.tif")
                Dim mb As Integer
                Dim rec As Integer
                sRec = 1
                For DB As Integer = 0 To dbcount - 1
                    Dim dbname As Integer = DB
                    dbname += 1
                    createData001nBD2112(batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Entry1\" & dbname.ToString("00000000.##") & ".mdb", batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName)
                    createCWCConnection(batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Entry1\" & dbname.ToString("00000000.##") & ".mdb")
                    createCbatchConnection(cbatchconnpath)
                    Dim OImage001 As String = ""
                    Dim Image001 As String = ""
                    Dim OImage002 As String = ""
                    Dim Image002 As String = ""
                    Try
                        For m As Integer = 1 To batchSize
                            mb = mb + 1
                            Dim source As String = imgNames(mb - 1)
                            Dim dest As String = batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Images\" + mb.ToString("00000000.##") & ".tif"
                            System.IO.File.Copy(source, dest, True)
                            lblProcImage.Text = mb.ToString & " to " & imgNames.Length & " - " & imgNames(mb - 1).ToString
                            lblProcess.Text = (Integer.Parse(lblProcess.Text) + 100 \ ProgressBar1.Maximum)
                            ProgressBar1.Value += 1
                            Application.DoEvents()
                            Dim imgname As New DirectoryInfo(imgNames(mb - 1).ToString)
                            OImage001 = imgname.Name
                            Image001 = mb.ToString("00000000.##") & ".tif"
                            rec += 1
                            eRec += 1
                            RecCnt += 1
                            cbatchConn.Open()
                            CWCConn.Open()
                            Dim cmdG1 As New OleDb.OleDbCommand("INSERT INTO [" & jobName & "] ([RecNum],[ZipFile],[DvdName],[BatchNo],[RefInd],[VerFlag],[ComFlag],[UpdFlag],[QCFlag],[QASFlag],[ValFlag],[RefFlag],[LKFlag1],[LKFlag2],[LKFlag3],[LKFlag4],[LKFlag5],[OImage001],[Image001]) VALUES (@RecNum, @ZipFile, @DvdName, @BatchNo, @RefInd, @VerFlag, @ComFlag, @UpdFlag, @QCFlag, @QASFlag, @ValFlag, @RefFlag, @LKFlag1, @LKFlag2, @LKFlag3, @LKFlag4, @LKFlag5, @OImage001, @Image001);", CWCConn)
                            cmdG1.Parameters.AddWithValue("@RecNum", rec)
                            cmdG1.Parameters.AddWithValue("@ZipFile", zipFile)
                            cmdG1.Parameters.AddWithValue("@DvdName", dvdNameParent)
                            cmdG1.Parameters.AddWithValue("@BatchNo", dbname.ToString("00000000.##"))
                            cmdG1.Parameters.AddWithValue("@ReFInd", "0")
                            cmdG1.Parameters.AddWithValue("@VerFlag", "0")
                            cmdG1.Parameters.AddWithValue("@ComFlag", "0")
                            cmdG1.Parameters.AddWithValue("@UpdFlag", "0")
                            cmdG1.Parameters.AddWithValue("@QCFlag", "0")
                            cmdG1.Parameters.AddWithValue("@QASFlag", "0")
                            cmdG1.Parameters.AddWithValue("@ValFlag", "0")
                            cmdG1.Parameters.AddWithValue("@RefFlag", "0")
                            cmdG1.Parameters.AddWithValue("@LKFlag1", "0")
                            cmdG1.Parameters.AddWithValue("@LKFlag2", "0")
                            cmdG1.Parameters.AddWithValue("@LKFlag3", "0")
                            cmdG1.Parameters.AddWithValue("@LKFlag4", "0")
                            cmdG1.Parameters.AddWithValue("@LKFlag5", "0")
                            cmdG1.Parameters.AddWithValue("@OImage001", OImage001)
                            cmdG1.Parameters.AddWithValue("@Image001", Image001)
                            cmdG1.ExecuteNonQuery()

                            Dim cmd As New OleDb.OleDbCommand("INSERT INTO [" & jobName & "] ([recNum],[ZipFile],[DvdName],[BatchNo],[OImage001],[Image001]) VALUES (@RecNum, @ZipFile, @DvdName, @BatchNo, @OImage001, @Image001)", cbatchConn)
                            cmd.Parameters.AddWithValue("@RecNum", rec)
                            cmd.Parameters.AddWithValue("@ZipFile", zipFile)
                            cmd.Parameters.AddWithValue("@DvdName", dvdNameParent)
                            cmd.Parameters.AddWithValue("@BatchNo", dbname.ToString("00000000.##"))
                            cmd.Parameters.AddWithValue("@OImage001", OImage001)
                            cmd.Parameters.AddWithValue("@Image001", Image001)
                            cmd.ExecuteNonQuery()

                            Dim cmd11 As New OleDb.OleDbCommand("INSERT INTO [origImage] ([OImage001],[Image001]) VALUES (@OImage001, @Image001)", cbatchConn)
                            cmd11.Parameters.AddWithValue("@OImage001", OImage001)
                            cmd11.Parameters.AddWithValue("@Image001", Image001)
                            cmd11.ExecuteNonQuery()

                            Dim cmdg2 As New OleDb.OleDbCommand("INSERT INTO [data001] ([RecNum]) VALUES (@RecNum)", CWCConn)
                            cmdg2.Parameters.AddWithValue("@RecNum", rec)
                            cmdg2.ExecuteNonQuery()

                            cbatchConn.Close()
                            CWCConn.Close()
                            GC.Collect()
                            GC.WaitForPendingFinalizers()

                            OImage001 = ""
                            Image001 = ""
                        Next m

                        Dim sRecL As String
                        Dim eRecL As String
                        Dim batchSL As String
                        sRecL = addLength(sRec.ToString, 10)
                        eRecL = addLength(eRec.ToString, 10)
                        batchSL = addLength(RecCnt, 10)
                        dblog = dblog & vbNewLine & "│ " & dbname.ToString & "       │ " & dbname.ToString("00000000.##") & ".mdb   │ " & sRecL & "│ " & eRecL & "│ " & batchSL & "│"
                        sRec = eRec + 1
                        eRec = eRec
                        RecCnt = 0
                    Catch ex As Exception
                        ' MsgBox(ex.Message)
                        Dim sRecL As String
                        Dim eRecL As String
                        Dim batchSL As String
                        sRecL = addLength(sRec.ToString, 10)
                        eRecL = addLength(eRec.ToString, 10)
                        batchSL = addLength(RecCnt, 10)
                        dblog = dblog & vbNewLine & "│ " & dbname.ToString & "       │ " & dbname.ToString("00000000.##") & ".mdb   │ " & sRecL & "│ " & eRecL & "│ " & batchSL & "│"
                        sRec = eRec + 1
                        eRec = eRec
                        RecCnt = 0
                        GoTo skip
                    End Try

                Next DB

skip:
                cbatchConn.Dispose()
                CWCConn.Dispose()
                Dim DTSource = batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Entry1"
                Dim cmLoc As String = batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Compare"
                My.Computer.FileSystem.CopyDirectory(DTSource, cmLoc, True)
                Dim e2loc As String = batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\Entry2"
                My.Computer.FileSystem.CopyDirectory(DTSource, e2loc, True)


                createD2112(Application.StartupPath & "\CBATCH" & "\" & jobName & ".mdb")
                createCbatchConnection(Application.StartupPath & "\CBATCH" & "\" & jobName & ".mdb")
                Dim cmdK1 As String = "INSERT INTO [" & jobName & "] ([ZipFile],[DvdName],[SeqNo1],[SeqNo2],[JobCode],[ClientCode],[TotalImages],[TotalRecs],[BatchCnt],[BatchDate],[BatchSize]) VALUES (@ZipFile, @DvdName, @Seq1No, @SeqNo2, @JobCode, @ClientCode, @TotalImages, @TotalRecs, @BatchCnt, @BatchDate, @BathcSize);"
                Dim cmdK2 As New OleDb.OleDbCommand(cmdK1, cbatchConn)
                cmdK2.Parameters.AddWithValue("@ZipFile", zipFile)
                cmdK2.Parameters.AddWithValue("@DvdName", dvdName)
                cmdK2.Parameters.AddWithValue("@SeqNo1", dateTime2)
                cmdK2.Parameters.AddWithValue("@SeqNo2", zipFile)
                cmdK2.Parameters.AddWithValue("@JobCode", jobName)
                cmdK2.Parameters.AddWithValue("@ClientCode", clientname)
                cmdK2.Parameters.AddWithValue("@TotalImages", imgcount)
                cmdK2.Parameters.AddWithValue("@TotalRecs", rec)
                cmdK2.Parameters.AddWithValue("@BatchCnt", dbcount)
                cmdK2.Parameters.AddWithValue("@BatchDate", dateTime)
                cmdK2.Parameters.AddWithValue("@BatchSize", batchSize)
                cbatchConn.Open()
                cmdK2.ExecuteNonQuery()
                cbatchConn.Close()
                GC.Collect()
                GC.WaitForPendingFinalizers()
                cbatchConn.Dispose()
                createLog(Application.StartupPath & "\CBATCH" + "\" & jobName & "\" & dateTime2 & "\" & dvdName & "\" & imgLoc.Name, dateTime, batchLocation & "\" & clientname & "\" & jobName & "\" & dateTime2 & "\" & dvdName, dblog, dvdNameParent & "\" & zipFile)
                dblog = ""
                CreateCBatchCSV()


                rec = 0
                mb = 0
nextloop:       Application.DoEvents()
            Next batches
            lblProcess.Text = "100"
            valBox2("Done" & vbNewLine & "File Processed: " & ProgressBar1.Maximum & " image(s)")
        Else
            valBox2("No Items To Be Batch")
        End If
    End Sub
    Public Function createData001nBD2112(ByVal bountyPath As String, ByVal rebatchpath As String)
        createData001nBD2112 = ""


        CreateAccessDatabase(bountyPath)
        createCWCConnection(bountyPath)
        Try
back:       Dim cmd As New OleDb.OleDbCommand("CREATE TABLE Data001 (recNum int , [code] text(70) NULL, [title] text(70) NULL, [fname] text(70) NULL, " & _
"[mname] text(70) NULL, [lname] text(70) NULL, [suffix] text(70) NULL, [address] text(70) NULL, [city] text(70) NULL, [state] text(70) NULL, [zipcode] text(70) NULL," & _
" [qasFlag] text(70) NULL, [phone] text(70) NULL, [emailadd] text(70) NULL, [confirmReciept] text(70) NULL, [1stRecieved] text(70) NULL, [addBox] text(70) NULL," & _
" [dob] text(70) NULL, [edd] text(70) NULL, [1stBaby] text(70) NULL, [healthCare] text(70) NULL, [commonReason] text(70) NULL, [visit] text(70) NULL," & _
" [gender1] text(70) NULL, [dob1] text(70) NULL, [gender2] text(70) NULL, [dob2] text(70) NULL, [gender3] text(70) NULL, [dob3] text(70) NULL,[include1] text(70) NULL, [include2] text(70) NULL, [include3] text(70) NULL, [include4] text(70) NULL, [include5] text(70) NULL, [other] text(70) NULL, " & _
" [pet] text(70) NULL,  [content] text(70) NULL, [editorial] text(70) NULL,[issue] text(70) NULL, [pcAtHome] text(70) NULL, [usefulCoupons] text(70) NULL," & _
" [savingPlans] text(70) NULL, [couponBooklet] text(70) NULL, [receiveSamples] text(70) NULL, [receiveEmail] text(70) NULL,    [diet] text(70) NULL, " & _
"[prenatal] text(70) NULL, [infant] text(70) NULL,     [workOutside] text(70) NULL, [income] text(70) NULL, [education] text(70) NULL, [ethnicity] text(70) NULL, " & _
"[share] text(70) NULL, [yourPatient] text(70) NULL, [whyNot] text(70) NULL, [babies] text(70) NULL, [patient] text(70) NULL, [recieveBooks] text(70) NULL," & _
"[boxes] text(70) NULL, [officeAdd] text(70) NULL, [wishToReceive] text(70) NULL,   [offer] text(70) NULL, [product] text(70) NULL, " & _
"[otherProduct] text(70) NULL,[receiveEmail1] text(70) NULL,[receiveEmail2] text(70) NULL,[rate] text(70) NULL,[receive] text(70) NULL," & _
" [continue] text(70) NULL, [officeAddress] text(70) NULL, [officeCity] text(70) NULL, [officeState] text(70) NULL, [officeZipCode] text(70) NULL," & _
" [officeqasFlag] text(70) NULL, [relevant] text(70) NULL, [base] text(70) NULL, [topic] text(70) NULL, [noticeEggo] text(70) NULL, [eggoBrand] text(70) NULL, " & _
"[voucherSection] text(70) NULL, [buyPledge] text(70) NULL, [fieldFurther] text(70) NULL,   [productOther] text(70) NULL, [purchaseOnline] text(70) NULL, " & _
"[connection] text(70) NULL,[axa] text(70) NULL, [homeBusiness] text(70) NULL, [toysrus] text(70) NULL, [viacord] text(70) NULL, [adt] text(70) NULL, [HeartSmart] text(70) NULL," & _
"[advertisement] text(70) NULL, [notice] text(70) NULL, [coupon] text(70) NULL, [consider] text(70) NULL, [review] text(70) NULL,[makeaPurchase] text(70) NULL ," & _
"[financialfuture] text(70) NULL, [BOBA] text(70) NULL, [influence] text(70) NULL, " & _
"  [bbproduct] text(70) NULL,[review2] text(70) NULL, [interested] text(70) NULL," & _
"   [healthBenefit] text(70) NULL, [recommend] text(70) NULL,[hibiclens] text(70) NULL,[dhaO3] text(70) NULL,  [milk] text(70) NULL, [factsUp] text(70) NULL,  " & _
"[hemange] text(70) NULL, [hemangeolC] text(70) NULL, [stonyfie] text(70) NULL, [advetori] text(70) NULL,[receiveSample] text(70) NULL,  [dataCaptu] text(70) NULL, " & _
"[image] text(70) NULL, [Extra1] text(70) NULL, [Extra2] text(70) NULL, [Extra3] text(70) NULL, [Extra4] text(70) NULL, [Extra5] text(70) NULL )", CWCConn)


            CWCConn.Open()
            cmd.ExecuteNonQuery()
            Dim cmd1 As New OleDb.OleDbCommand("CREATE TABLE " & jobName & " ( " & "RecNum int," & "[ZipFile]  text(40)," & "[DvdName] text(40)," & "[BatchNo] text(40)," & _
             "[PullOut] text(40)," & "[RefInd] text(40)," & "[KeyFlag] text(40)," & "[KeyId] text(40)," & "[KeyDate] text(40)," & "[VerFlag] text(40)," & _
             "[VerID] text(40)," & "[VerDate] text(40)," & "[ComFlag] text(40)," & "[ComID] text(40)," & "[ComDate] text(40)," & "[UpdFlag] text(40)," & _
             "[UpdID] text(40)," & "[UpdDate] text(40)," & "[QCFlag] text(40)," & "[QCID] text(40)," & "[QCDate] text(40)," & _
             "[QASFlag] text(40)," & "[QASID] text(40)," & "[QASDate] text(40)," & "[ValFlag] text(40)," & "[ValID] text(40)," & _
             "[ValDate] text(40)," & "[RefFlag] text(40)," & "[RefID] text(40)," & "[RefDate] text(40)," & "[Accuracy1] text(40)," & _
             "[Accuracy2] text(40)," & "[Accuracy3] text(40)," & "[Accuracy4] text(40)," & "[Accuracy5] text(40)," & "[KSCount] text(40)," & _
             "[QCID1] text(40)," & "[QCID2] text(40)," & "[QCID3] text(40)," & "[QCID4] text(40)," & "[QCID5] text(40)," & "[LKFlag1] text(40)," & _
             "[LKFlag2] text(40)," & "[LKFlag3] text(40)," & "[LKFlag4] text(40)," & "[LKFlag5] text(40)," & "[LocId1] text(40)," & "[LocId2] text(40)," & _
             "[OImage001] text(70) NULL," & "[Image001] text(40)," & _
             "CONSTRAINT [pk_RecNum] PRIMARY KEY (RecNum)) ", CWCConn)
            cmd1.ExecuteNonQuery()
            CWCConn.Close()
            CWCConn.Dispose()
        Catch
            Dim delcmd1 As New OleDb.OleDbCommand("DELETE * FROM " & jobName, CWCConn)
            Dim delcmd2 As New OleDb.OleDbCommand("DELETE * FROM Data001 ", CWCConn)
            If CWCConn.State = ConnectionState.Closed Then
                CWCConn.Open()
            End If
            delcmd1.ExecuteNonQuery()
            delcmd2.ExecuteNonQuery()
            If CWCConn.State = ConnectionState.Open Then
                CWCConn.Close()
            End If
            '  GoTo back
        End Try

    End Function
    Public Function createD2112(ByVal path As String)
        createD2112 = ""
        If Not IO.File.Exists(path) Then
            CreateAccessDatabase(path)
            createCbatchConnection(path)
            Dim cmd4 As New OleDb.OleDbCommand("CREATE TABLE " & jobName & " ( " & "[ZipFile] text(50) ," & "[DvdName] text(50)," & "[SeqNo1] text(50)," & _
             "[SeqNo2] text(50)," & "[JobCode] text(50)," & "[ClientCode] text(50)," & "[TotalImages] int," & _
             "[TotalRecs] int," & "[BatchCnt] int," & "[BatchDate] text(50)," & "[BatchSize] int) ", cbatchConn)
            cbatchConn.Open()
            cmd4.ExecuteNonQuery()
            cbatchConn.Close()
        End If
    End Function
    Public Function createD2112nOrigImage(ByVal cbatchPath As String, ByVal batchPath As String, ByVal rebatchpath As String) As Boolean

back:   CreateAccessDatabase(cbatchPath)
        createCbatchConnection(cbatchPath)
        Try
            Dim cmd2 As New OleDb.OleDbCommand("CREATE TABLE " & jobName & " (RecNum int, [ZipFile] text(50),[DvdName] text(50),[BatchNo] text(50),[OImage001] text(70) NULL,[Image001] text(50))", cbatchConn)
            cbatchConn.Open()
            cmd2.ExecuteNonQuery()
            Dim cmd3 As New OleDb.OleDbCommand("CREATE TABLE origImage (OImage001 text(60) ,[Image001] text(50))", cbatchConn)
            cmd3.ExecuteNonQuery()
            cbatchConn.Close()
        Catch e As Exception
            '        MsgBox(e.Message)
            cbatchConn.Close()
            Dim bool As Boolean
            bool = valBox("File: " & batchPath & " is Already Batched....." & vbNewLine & "   Do You Want To Rebatch??", "STOP!!")
            If bool = True Then
back2:          Try
                    For Each file As String In My.Computer.FileSystem.GetFiles(rebatchpath)
                        My.Computer.FileSystem.DeleteFile(file)
                    Next
                    IO.Directory.Delete(rebatchpath, True)
                    IO.Directory.CreateDirectory(rebatchpath)
                Catch
                    createCbatchConnection(cbatchPath)
                    cbatchConn.Open()
                    cbatchConn.Close()
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                    GoTo back2
                End Try
                GoTo back
            Else
                valBox2("Rebatching File: " & vbNewLine & batchPath & "Has Been Canceled")
                Return False
            End If
        End Try
        cbatchConn.Dispose()
        Return True
    End Function
End Class
