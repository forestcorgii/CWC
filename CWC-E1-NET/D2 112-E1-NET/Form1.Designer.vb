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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.OPENToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SAVEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.STARTENTRYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EDITRECORDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BROWSEMODEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.GOTOToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EXITToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ChangeIDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.FileToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.ROTATEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ZOOMINToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ZOOMOUTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MOVEIMAGETOLEFTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MOVEIMAGETORIGHTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MOVEIMAGEUPToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MOVEIMAGEDOWNToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblMode = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.lblDesc = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtDataBy = New System.Windows.Forms.TextBox
        Me.lblFieldName = New System.Windows.Forms.Label
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.fieldName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.value = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.txtValue = New System.Windows.Forms.TextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtUserID = New System.Windows.Forms.TextBox
        Me.txtLocID = New System.Windows.Forms.TextBox
        Me.txtCrntRec = New System.Windows.Forms.TextBox
        Me.txtTotalRec = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtCrntImg = New System.Windows.Forms.TextBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.imgvwer = New ImageViewer.ImageControl
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1372, 27)
        Me.Panel1.TabIndex = 0
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(1309, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(62, 27)
        Me.Label15.TabIndex = 19
        Me.Label15.Text = "X"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.Transparent
        Me.Label16.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.White
        Me.Label16.Location = New System.Drawing.Point(1259, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(50, 27)
        Me.Label16.TabIndex = 20
        Me.Label16.Text = "──"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Light", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(1369, 27)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "CWC-E1-NET"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Panel2.Controls.Add(Me.MenuStrip1)
        Me.Panel2.Location = New System.Drawing.Point(0, 27)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1372, 27)
        Me.Panel2.TabIndex = 2
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.Transparent
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ChangeIDToolStripMenuItem, Me.FileToolStripMenuItem1})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1372, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.BackColor = System.Drawing.Color.Transparent
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OPENToolStripMenuItem, Me.SAVEToolStripMenuItem, Me.STARTENTRYToolStripMenuItem, Me.EDITRECORDToolStripMenuItem, Me.BROWSEMODEToolStripMenuItem, Me.GOTOToolStripMenuItem, Me.EXITToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.FileToolStripMenuItem.Text = "FILE"
        '
        'OPENToolStripMenuItem
        '
        Me.OPENToolStripMenuItem.Name = "OPENToolStripMenuItem"
        Me.OPENToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.OPENToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.OPENToolStripMenuItem.Text = "OPEN"
        '
        'SAVEToolStripMenuItem
        '
        Me.SAVEToolStripMenuItem.Name = "SAVEToolStripMenuItem"
        Me.SAVEToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F2
        Me.SAVEToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.SAVEToolStripMenuItem.Text = "SAVE"
        '
        'STARTENTRYToolStripMenuItem
        '
        Me.STARTENTRYToolStripMenuItem.Name = "STARTENTRYToolStripMenuItem"
        Me.STARTENTRYToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F3
        Me.STARTENTRYToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.STARTENTRYToolStripMenuItem.Text = "START ENTRY"
        '
        'EDITRECORDToolStripMenuItem
        '
        Me.EDITRECORDToolStripMenuItem.Name = "EDITRECORDToolStripMenuItem"
        Me.EDITRECORDToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F4
        Me.EDITRECORDToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.EDITRECORDToolStripMenuItem.Text = "EDIT RECORD"
        '
        'BROWSEMODEToolStripMenuItem
        '
        Me.BROWSEMODEToolStripMenuItem.Name = "BROWSEMODEToolStripMenuItem"
        Me.BROWSEMODEToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F9
        Me.BROWSEMODEToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.BROWSEMODEToolStripMenuItem.Text = "BROWSE MODE"
        '
        'GOTOToolStripMenuItem
        '
        Me.GOTOToolStripMenuItem.Name = "GOTOToolStripMenuItem"
        Me.GOTOToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F5
        Me.GOTOToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.GOTOToolStripMenuItem.Text = "GO TO"
        '
        'EXITToolStripMenuItem
        '
        Me.EXITToolStripMenuItem.Name = "EXITToolStripMenuItem"
        Me.EXITToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.EXITToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.EXITToolStripMenuItem.Text = "EXIT "
        '
        'ChangeIDToolStripMenuItem
        '
        Me.ChangeIDToolStripMenuItem.BackColor = System.Drawing.Color.Transparent
        Me.ChangeIDToolStripMenuItem.Name = "ChangeIDToolStripMenuItem"
        Me.ChangeIDToolStripMenuItem.Size = New System.Drawing.Size(81, 20)
        Me.ChangeIDToolStripMenuItem.Text = "CHANGE ID"
        '
        'FileToolStripMenuItem1
        '
        Me.FileToolStripMenuItem1.BackColor = System.Drawing.Color.Transparent
        Me.FileToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ROTATEToolStripMenuItem, Me.ZOOMINToolStripMenuItem, Me.ZOOMOUTToolStripMenuItem, Me.MOVEIMAGETOLEFTToolStripMenuItem, Me.MOVEIMAGETORIGHTToolStripMenuItem, Me.MOVEIMAGEUPToolStripMenuItem, Me.MOVEIMAGEDOWNToolStripMenuItem})
        Me.FileToolStripMenuItem1.Name = "FileToolStripMenuItem1"
        Me.FileToolStripMenuItem1.Size = New System.Drawing.Size(55, 20)
        Me.FileToolStripMenuItem1.Text = "IMAGE"
        '
        'ROTATEToolStripMenuItem
        '
        Me.ROTATEToolStripMenuItem.Name = "ROTATEToolStripMenuItem"
        Me.ROTATEToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.ROTATEToolStripMenuItem.Size = New System.Drawing.Size(240, 22)
        Me.ROTATEToolStripMenuItem.Text = "ROTATE"
        '
        'ZOOMINToolStripMenuItem
        '
        Me.ZOOMINToolStripMenuItem.Name = "ZOOMINToolStripMenuItem"
        Me.ZOOMINToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Y), System.Windows.Forms.Keys)
        Me.ZOOMINToolStripMenuItem.Size = New System.Drawing.Size(240, 22)
        Me.ZOOMINToolStripMenuItem.Text = "ZOOM IN "
        '
        'ZOOMOUTToolStripMenuItem
        '
        Me.ZOOMOUTToolStripMenuItem.Name = "ZOOMOUTToolStripMenuItem"
        Me.ZOOMOUTToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.ZOOMOUTToolStripMenuItem.Size = New System.Drawing.Size(240, 22)
        Me.ZOOMOUTToolStripMenuItem.Text = "ZOOM OUT"
        '
        'MOVEIMAGETOLEFTToolStripMenuItem
        '
        Me.MOVEIMAGETOLEFTToolStripMenuItem.Name = "MOVEIMAGETOLEFTToolStripMenuItem"
        Me.MOVEIMAGETOLEFTToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
        Me.MOVEIMAGETOLEFTToolStripMenuItem.Size = New System.Drawing.Size(240, 22)
        Me.MOVEIMAGETOLEFTToolStripMenuItem.Text = "MOVE IMAGE TO LEFT"
        '
        'MOVEIMAGETORIGHTToolStripMenuItem
        '
        Me.MOVEIMAGETORIGHTToolStripMenuItem.Name = "MOVEIMAGETORIGHTToolStripMenuItem"
        Me.MOVEIMAGETORIGHTToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.J), System.Windows.Forms.Keys)
        Me.MOVEIMAGETORIGHTToolStripMenuItem.Size = New System.Drawing.Size(240, 22)
        Me.MOVEIMAGETORIGHTToolStripMenuItem.Text = "MOVE IMAGE TO RIGHT"
        '
        'MOVEIMAGEUPToolStripMenuItem
        '
        Me.MOVEIMAGEUPToolStripMenuItem.Name = "MOVEIMAGEUPToolStripMenuItem"
        Me.MOVEIMAGEUPToolStripMenuItem.ShortcutKeyDisplayString = "Ctrl+,"
        Me.MOVEIMAGEUPToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Oemcomma), System.Windows.Forms.Keys)
        Me.MOVEIMAGEUPToolStripMenuItem.Size = New System.Drawing.Size(240, 22)
        Me.MOVEIMAGEUPToolStripMenuItem.Text = "MOVE IMAGE UP"
        '
        'MOVEIMAGEDOWNToolStripMenuItem
        '
        Me.MOVEIMAGEDOWNToolStripMenuItem.Name = "MOVEIMAGEDOWNToolStripMenuItem"
        Me.MOVEIMAGEDOWNToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.I), System.Windows.Forms.Keys)
        Me.MOVEIMAGEDOWNToolStripMenuItem.Size = New System.Drawing.Size(240, 22)
        Me.MOVEIMAGEDOWNToolStripMenuItem.Text = "MOVE IMAGE DOWN"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblMode)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(868, 55)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(488, 41)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "MODE :"
        '
        'lblMode
        '
        Me.lblMode.Font = New System.Drawing.Font("Segoe UI Semibold", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMode.ForeColor = System.Drawing.Color.Black
        Me.lblMode.Location = New System.Drawing.Point(78, 10)
        Me.lblMode.Name = "lblMode"
        Me.lblMode.Size = New System.Drawing.Size(354, 29)
        Me.lblMode.TabIndex = 0
        Me.lblMode.Text = "BROWSE MODE"
        Me.lblMode.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblDesc)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.txtDataBy)
        Me.GroupBox2.Controls.Add(Me.lblFieldName)
        Me.GroupBox2.Controls.Add(Me.DataGridView1)
        Me.GroupBox2.Controls.Add(Me.txtValue)
        Me.GroupBox2.Location = New System.Drawing.Point(868, 102)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(488, 548)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        '
        'lblDesc
        '
        Me.lblDesc.BackColor = System.Drawing.Color.Black
        Me.lblDesc.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDesc.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.lblDesc.Location = New System.Drawing.Point(5, 39)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(476, 163)
        Me.lblDesc.TabIndex = 38
        Me.lblDesc.Text = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(7, 235)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 13)
        Me.Label4.TabIndex = 36
        Me.Label4.Text = "Data By:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDataBy
        '
        Me.txtDataBy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDataBy.Enabled = False
        Me.txtDataBy.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataBy.Location = New System.Drawing.Point(63, 232)
        Me.txtDataBy.Name = "txtDataBy"
        Me.txtDataBy.ReadOnly = True
        Me.txtDataBy.Size = New System.Drawing.Size(418, 22)
        Me.txtDataBy.TabIndex = 35
        '
        'lblFieldName
        '
        Me.lblFieldName.BackColor = System.Drawing.Color.SteelBlue
        Me.lblFieldName.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFieldName.ForeColor = System.Drawing.Color.White
        Me.lblFieldName.Location = New System.Drawing.Point(6, 11)
        Me.lblFieldName.Name = "lblFieldName"
        Me.lblFieldName.Size = New System.Drawing.Size(477, 28)
        Me.lblFieldName.TabIndex = 3
        Me.lblFieldName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.fieldName, Me.value})
        Me.DataGridView1.EnableHeadersVisualStyles = False
        Me.DataGridView1.Location = New System.Drawing.Point(6, 256)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.RoyalBlue
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.White
        Me.DataGridView1.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(476, 285)
        Me.DataGridView1.TabIndex = 34
        '
        'fieldName
        '
        Me.fieldName.HeaderText = "Fields"
        Me.fieldName.Name = "fieldName"
        Me.fieldName.ReadOnly = True
        Me.fieldName.Width = 200
        '
        'value
        '
        Me.value.HeaderText = "Value"
        Me.value.Name = "value"
        Me.value.ReadOnly = True
        Me.value.Width = 275
        '
        'txtValue
        '
        Me.txtValue.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtValue.Location = New System.Drawing.Point(6, 205)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(476, 25)
        Me.txtValue.TabIndex = 1
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.txtUserID)
        Me.GroupBox3.Controls.Add(Me.txtLocID)
        Me.GroupBox3.Controls.Add(Me.txtCrntRec)
        Me.GroupBox3.Controls.Add(Me.txtTotalRec)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.txtCrntImg)
        Me.GroupBox3.Location = New System.Drawing.Point(868, 654)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(488, 112)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.SteelBlue
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(284, 36)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 24)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "OF"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.SteelBlue
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(14, 82)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(101, 24)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "USER ID"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.SteelBlue
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(14, 60)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(101, 22)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "LOCATION ID"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.SteelBlue
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(14, 36)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(101, 24)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "RECORD"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtUserID
        '
        Me.txtUserID.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtUserID.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserID.Location = New System.Drawing.Point(115, 84)
        Me.txtUserID.Multiline = True
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.ReadOnly = True
        Me.txtUserID.Size = New System.Drawing.Size(367, 22)
        Me.txtUserID.TabIndex = 7
        Me.txtUserID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtLocID
        '
        Me.txtLocID.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtLocID.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocID.Location = New System.Drawing.Point(115, 61)
        Me.txtLocID.Multiline = True
        Me.txtLocID.Name = "txtLocID"
        Me.txtLocID.ReadOnly = True
        Me.txtLocID.Size = New System.Drawing.Size(367, 21)
        Me.txtLocID.TabIndex = 8
        Me.txtLocID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCrntRec
        '
        Me.txtCrntRec.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtCrntRec.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCrntRec.Location = New System.Drawing.Point(115, 38)
        Me.txtCrntRec.Multiline = True
        Me.txtCrntRec.Name = "txtCrntRec"
        Me.txtCrntRec.ReadOnly = True
        Me.txtCrntRec.Size = New System.Drawing.Size(169, 21)
        Me.txtCrntRec.TabIndex = 8
        Me.txtCrntRec.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtTotalRec
        '
        Me.txtTotalRec.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtTotalRec.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalRec.Location = New System.Drawing.Point(322, 38)
        Me.txtTotalRec.Multiline = True
        Me.txtTotalRec.Name = "txtTotalRec"
        Me.txtTotalRec.ReadOnly = True
        Me.txtTotalRec.Size = New System.Drawing.Size(160, 21)
        Me.txtTotalRec.TabIndex = 9
        Me.txtTotalRec.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.SteelBlue
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(14, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 24)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "IMAGE"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCrntImg
        '
        Me.txtCrntImg.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtCrntImg.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCrntImg.Location = New System.Drawing.Point(115, 12)
        Me.txtCrntImg.Multiline = True
        Me.txtCrntImg.Name = "txtCrntImg"
        Me.txtCrntImg.ReadOnly = True
        Me.txtCrntImg.Size = New System.Drawing.Size(367, 24)
        Me.txtCrntImg.TabIndex = 6
        Me.txtCrntImg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Filter = "MDB files (*.mdb)|*.mdb|All files (*.*)|*.*"
        '
        'imgvwer
        '
        Me.imgvwer.BackColor = System.Drawing.SystemColors.Control
        Me.imgvwer.Image = Nothing
        Me.imgvwer.ImagePath = ""
        Me.imgvwer.initialimage = Nothing
        Me.imgvwer.Location = New System.Drawing.Point(12, 60)
        Me.imgvwer.Name = "imgvwer"
        Me.imgvwer.Origin = New System.Drawing.Point(0, 0)
        Me.imgvwer.PageCount = 0
        Me.imgvwer.PageNumber = 1
        Me.imgvwer.PanButton = System.Windows.Forms.MouseButtons.Left
        Me.imgvwer.PanMode = True
        Me.imgvwer.ScrollbarsVisible = True
        Me.imgvwer.Size = New System.Drawing.Size(850, 714)
        Me.imgvwer.StretchImageToFit = False
        Me.imgvwer.TabIndex = 6
        Me.imgvwer.ZoomFactor = 1
        Me.imgvwer.ZoomOnMouseWheel = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1370, 750)
        Me.Controls.Add(Me.imgvwer)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Segoe UI Light", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChangeIDToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblMode As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents txtLocID As System.Windows.Forms.TextBox
    Friend WithEvents txtCrntRec As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalRec As System.Windows.Forms.TextBox
    Friend WithEvents txtCrntImg As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents OPENToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SAVEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents STARTENTRYToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EDITRECORDToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BROWSEMODEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GOTOToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EXITToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ROTATEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ZOOMINToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ZOOMOUTToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MOVEIMAGETOLEFTToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MOVEIMAGETORIGHTToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MOVEIMAGEUPToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MOVEIMAGEDOWNToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblFieldName As System.Windows.Forms.Label
    Friend WithEvents imgvwer As ImageViewer.ImageControl
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtDataBy As System.Windows.Forms.TextBox
    Friend WithEvents fieldName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents value As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents lblDesc As System.Windows.Forms.Label

End Class
