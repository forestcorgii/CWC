<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.cmdClearAll = New System.Windows.Forms.Button
        Me.cmdAddAll = New System.Windows.Forms.Button
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.Batch = New System.Windows.Forms.ColumnHeader
        Me.ListView2 = New System.Windows.Forms.ListView
        Me.selectedBatches = New System.Windows.Forms.ColumnHeader
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DirListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
        Me.DriveListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.cmdKS = New System.Windows.Forms.Button
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.txtTotalPullout = New System.Windows.Forms.TextBox
        Me.txtAveKeyStroke = New System.Windows.Forms.TextBox
        Me.txtTotalKeyStroke = New System.Windows.Forms.TextBox
        Me.txtTotalValRec = New System.Windows.Forms.TextBox
        Me.txtTotalRec = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.rbPerMDB = New System.Windows.Forms.RadioButton
        Me.rbPerBatch = New System.Windows.Forms.RadioButton
        Me.rbCom = New System.Windows.Forms.RadioButton
        Me.rbE2 = New System.Windows.Forms.RadioButton
        Me.rbE1 = New System.Windows.Forms.RadioButton
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.GroupBox3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ProgressBar2)
        Me.GroupBox3.Controls.Add(Me.ProgressBar1)
        Me.GroupBox3.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.Black
        Me.GroupBox3.Location = New System.Drawing.Point(222, 356)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(417, 59)
        Me.GroupBox3.TabIndex = 23
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Process:"
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Location = New System.Drawing.Point(7, 41)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(404, 12)
        Me.ProgressBar2.TabIndex = 3
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(7, 15)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(404, 19)
        Me.ProgressBar1.TabIndex = 2
        '
        'cmdClearAll
        '
        Me.cmdClearAll.BackColor = System.Drawing.Color.White
        Me.cmdClearAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdClearAll.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClearAll.Location = New System.Drawing.Point(56, 289)
        Me.cmdClearAll.Name = "cmdClearAll"
        Me.cmdClearAll.Size = New System.Drawing.Size(123, 24)
        Me.cmdClearAll.TabIndex = 20
        Me.cmdClearAll.Text = "Clear All"
        Me.cmdClearAll.UseVisualStyleBackColor = False
        '
        'cmdAddAll
        '
        Me.cmdAddAll.BackColor = System.Drawing.Color.White
        Me.cmdAddAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdAddAll.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddAll.Location = New System.Drawing.Point(56, 260)
        Me.cmdAddAll.Name = "cmdAddAll"
        Me.cmdAddAll.Size = New System.Drawing.Size(123, 23)
        Me.cmdAddAll.TabIndex = 18
        Me.cmdAddAll.Text = "Select &All"
        Me.cmdAddAll.UseVisualStyleBackColor = False
        '
        'cmdAdd
        '
        Me.cmdAdd.BackColor = System.Drawing.Color.White
        Me.cmdAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdAdd.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.Location = New System.Drawing.Point(56, 231)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(123, 23)
        Me.cmdAdd.TabIndex = 17
        Me.cmdAdd.Text = "Add For Process"
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Batch})
        Me.ListView1.Font = New System.Drawing.Font("Segoe UI Light", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListView1.ForeColor = System.Drawing.Color.Black
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView1.Location = New System.Drawing.Point(222, 45)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(417, 129)
        Me.ListView1.TabIndex = 16
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'Batch
        '
        Me.Batch.Text = "Batches"
        Me.Batch.Width = 413
        '
        'ListView2
        '
        Me.ListView2.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.selectedBatches})
        Me.ListView2.Font = New System.Drawing.Font("Segoe UI Light", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListView2.ForeColor = System.Drawing.Color.Black
        Me.ListView2.FullRowSelect = True
        Me.ListView2.GridLines = True
        Me.ListView2.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView2.Location = New System.Drawing.Point(222, 190)
        Me.ListView2.Name = "ListView2"
        Me.ListView2.Size = New System.Drawing.Size(417, 125)
        Me.ListView2.TabIndex = 15
        Me.ListView2.UseCompatibleStateImageBehavior = False
        Me.ListView2.View = System.Windows.Forms.View.Details
        '
        'selectedBatches
        '
        Me.selectedBatches.Text = "Selected Batches :"
        Me.selectedBatches.Width = 413
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkSlateBlue
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Font = New System.Drawing.Font("Segoe UI Semibold", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(-6, -2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(922, 27)
        Me.Panel1.TabIndex = 13
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Segoe UI Semibold", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(8, 4)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(105, 21)
        Me.Label13.TabIndex = 19
        Me.Label13.Text = "CWC-KS-NET"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(867, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(55, 27)
        Me.Label15.TabIndex = 17
        Me.Label15.Text = "X"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DirListBox1)
        Me.GroupBox1.Controls.Add(Me.DriveListBox1)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(6, 115)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(206, 300)
        Me.GroupBox1.TabIndex = 25
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Choose Directory:"
        '
        'DirListBox1
        '
        Me.DirListBox1.Font = New System.Drawing.Font("Segoe UI Light", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DirListBox1.FormattingEnabled = True
        Me.DirListBox1.IntegralHeight = False
        Me.DirListBox1.Location = New System.Drawing.Point(12, 50)
        Me.DirListBox1.Name = "DirListBox1"
        Me.DirListBox1.Size = New System.Drawing.Size(188, 243)
        Me.DirListBox1.TabIndex = 1
        '
        'DriveListBox1
        '
        Me.DriveListBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.DriveListBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DriveListBox1.FormattingEnabled = True
        Me.DriveListBox1.Location = New System.Drawing.Point(12, 21)
        Me.DriveListBox1.Name = "DriveListBox1"
        Me.DriveListBox1.Size = New System.Drawing.Size(164, 24)
        Me.DriveListBox1.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdAdd)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.cmdClearAll)
        Me.GroupBox2.Controls.Add(Me.cmdKS)
        Me.GroupBox2.Controls.Add(Me.cmdAddAll)
        Me.GroupBox2.Controls.Add(Me.Panel6)
        Me.GroupBox2.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(647, 37)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(258, 371)
        Me.GroupBox2.TabIndex = 26
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Actions :"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.White
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(56, 322)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(123, 33)
        Me.Button1.TabIndex = 22
        Me.Button1.Text = "EXIT"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'cmdKS
        '
        Me.cmdKS.BackColor = System.Drawing.Color.White
        Me.cmdKS.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdKS.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdKS.ForeColor = System.Drawing.Color.DarkSlateBlue
        Me.cmdKS.Location = New System.Drawing.Point(56, 192)
        Me.cmdKS.Name = "cmdKS"
        Me.cmdKS.Size = New System.Drawing.Size(123, 32)
        Me.cmdKS.TabIndex = 19
        Me.cmdKS.Text = "Count KeyStroke"
        Me.cmdKS.UseVisualStyleBackColor = False
        '
        'Panel6
        '
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Controls.Add(Me.txtTotalPullout)
        Me.Panel6.Controls.Add(Me.txtAveKeyStroke)
        Me.Panel6.Controls.Add(Me.txtTotalKeyStroke)
        Me.Panel6.Controls.Add(Me.txtTotalValRec)
        Me.Panel6.Controls.Add(Me.txtTotalRec)
        Me.Panel6.Controls.Add(Me.Label6)
        Me.Panel6.Controls.Add(Me.Label1)
        Me.Panel6.Controls.Add(Me.Label4)
        Me.Panel6.Controls.Add(Me.Label3)
        Me.Panel6.Controls.Add(Me.Label2)
        Me.Panel6.Location = New System.Drawing.Point(14, 19)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(232, 128)
        Me.Panel6.TabIndex = 23
        '
        'txtTotalPullout
        '
        Me.txtTotalPullout.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalPullout.Location = New System.Drawing.Point(150, 98)
        Me.txtTotalPullout.Name = "txtTotalPullout"
        Me.txtTotalPullout.Size = New System.Drawing.Size(73, 22)
        Me.txtTotalPullout.TabIndex = 9
        Me.txtTotalPullout.Text = "0"
        Me.txtTotalPullout.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtAveKeyStroke
        '
        Me.txtAveKeyStroke.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAveKeyStroke.Location = New System.Drawing.Point(150, 74)
        Me.txtAveKeyStroke.Name = "txtAveKeyStroke"
        Me.txtAveKeyStroke.Size = New System.Drawing.Size(73, 22)
        Me.txtAveKeyStroke.TabIndex = 8
        Me.txtAveKeyStroke.Text = "0"
        Me.txtAveKeyStroke.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtTotalKeyStroke
        '
        Me.txtTotalKeyStroke.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalKeyStroke.Location = New System.Drawing.Point(150, 50)
        Me.txtTotalKeyStroke.Name = "txtTotalKeyStroke"
        Me.txtTotalKeyStroke.Size = New System.Drawing.Size(73, 22)
        Me.txtTotalKeyStroke.TabIndex = 7
        Me.txtTotalKeyStroke.Text = "0"
        Me.txtTotalKeyStroke.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtTotalValRec
        '
        Me.txtTotalValRec.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalValRec.Location = New System.Drawing.Point(150, 26)
        Me.txtTotalValRec.Name = "txtTotalValRec"
        Me.txtTotalValRec.Size = New System.Drawing.Size(73, 22)
        Me.txtTotalValRec.TabIndex = 6
        Me.txtTotalValRec.Text = "0"
        Me.txtTotalValRec.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtTotalRec
        '
        Me.txtTotalRec.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalRec.Location = New System.Drawing.Point(150, 2)
        Me.txtTotalRec.Name = "txtTotalRec"
        Me.txtTotalRec.Size = New System.Drawing.Size(73, 22)
        Me.txtTotalRec.TabIndex = 5
        Me.txtTotalRec.Text = "0"
        Me.txtTotalRec.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(15, 28)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(111, 15)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Total Valid Records:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(43, 100)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 15)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Total Pullouts:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(29, 52)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(97, 15)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Total KeyStrokes:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(38, 76)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 15)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Ave. KeyStroke:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(44, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 15)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Total Records:"
        '
        'rbPerMDB
        '
        Me.rbPerMDB.AutoCheck = False
        Me.rbPerMDB.AutoSize = True
        Me.rbPerMDB.Enabled = False
        Me.rbPerMDB.Font = New System.Drawing.Font("Segoe UI Light", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbPerMDB.Location = New System.Drawing.Point(109, 30)
        Me.rbPerMDB.Name = "rbPerMDB"
        Me.rbPerMDB.Size = New System.Drawing.Size(70, 19)
        Me.rbPerMDB.TabIndex = 29
        Me.rbPerMDB.Text = "Per MDB"
        Me.rbPerMDB.UseVisualStyleBackColor = True
        '
        'rbPerBatch
        '
        Me.rbPerBatch.AutoCheck = False
        Me.rbPerBatch.AutoSize = True
        Me.rbPerBatch.Checked = True
        Me.rbPerBatch.Font = New System.Drawing.Font("Segoe UI Light", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbPerBatch.Location = New System.Drawing.Point(13, 30)
        Me.rbPerBatch.Name = "rbPerBatch"
        Me.rbPerBatch.Size = New System.Drawing.Size(73, 19)
        Me.rbPerBatch.TabIndex = 28
        Me.rbPerBatch.Text = "Per Batch"
        Me.rbPerBatch.UseVisualStyleBackColor = True
        '
        'rbCom
        '
        Me.rbCom.AutoSize = True
        Me.rbCom.Font = New System.Drawing.Font("Segoe UI Light", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbCom.Location = New System.Drawing.Point(528, 328)
        Me.rbCom.Name = "rbCom"
        Me.rbCom.Size = New System.Drawing.Size(72, 19)
        Me.rbCom.TabIndex = 32
        Me.rbCom.Text = "Compare"
        Me.rbCom.UseVisualStyleBackColor = True
        '
        'rbE2
        '
        Me.rbE2.AutoSize = True
        Me.rbE2.Font = New System.Drawing.Font("Segoe UI Light", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbE2.Location = New System.Drawing.Point(399, 328)
        Me.rbE2.Name = "rbE2"
        Me.rbE2.Size = New System.Drawing.Size(56, 19)
        Me.rbE2.TabIndex = 31
        Me.rbE2.Text = "Entry2"
        Me.rbE2.UseVisualStyleBackColor = True
        '
        'rbE1
        '
        Me.rbE1.AutoSize = True
        Me.rbE1.Checked = True
        Me.rbE1.Font = New System.Drawing.Font("Segoe UI Light", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbE1.Location = New System.Drawing.Point(277, 328)
        Me.rbE1.Name = "rbE1"
        Me.rbE1.Size = New System.Drawing.Size(56, 19)
        Me.rbE1.TabIndex = 30
        Me.rbE1.TabStop = True
        Me.rbE1.Text = "Entry1"
        Me.rbE1.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.rbPerBatch)
        Me.GroupBox4.Controls.Add(Me.rbPerMDB)
        Me.GroupBox4.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(6, 38)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(206, 71)
        Me.GroupBox4.TabIndex = 33
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Process Type:"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(915, 420)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.rbCom)
        Me.Controls.Add(Me.rbE2)
        Me.Controls.Add(Me.rbE1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.ListView2)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.GroupBox3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdClearAll As System.Windows.Forms.Button
    Friend WithEvents cmdAddAll As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents Batch As System.Windows.Forms.ColumnHeader
    Friend WithEvents ListView2 As System.Windows.Forms.ListView
    Friend WithEvents selectedBatches As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DirListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents DriveListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdKS As System.Windows.Forms.Button
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents txtTotalPullout As System.Windows.Forms.TextBox
    Friend WithEvents txtAveKeyStroke As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalKeyStroke As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalValRec As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalRec As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents rbPerMDB As System.Windows.Forms.RadioButton
    Friend WithEvents rbPerBatch As System.Windows.Forms.RadioButton
    Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents rbCom As System.Windows.Forms.RadioButton
    Friend WithEvents rbE2 As System.Windows.Forms.RadioButton
    Friend WithEvents rbE1 As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox

End Class
