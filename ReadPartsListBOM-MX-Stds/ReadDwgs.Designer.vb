<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ReadDwgs
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReadDwgs))
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.BtnStart2 = New System.Windows.Forms.Button()
        Me.LblProgress = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.DwgsNotFoundList = New System.Windows.Forms.ListBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.DwgList = New System.Windows.Forms.ListBox()
        Me.SelectList = New System.Windows.Forms.ListBox()
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.BtnAddAll = New System.Windows.Forms.Button()
        Me.BtnRemove = New System.Windows.Forms.Button()
        Me.BtnClear = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.ComboBxRev = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.BtnGetMWInfo = New System.Windows.Forms.Button()
        Me.PathBox = New System.Windows.Forms.TextBox()
        Me.BtnVideo = New System.Windows.Forms.Button()
        Me.CancelButton_Renamed = New System.Windows.Forms.Button()
        Me.MatrixLogo = New System.Windows.Forms.PictureBox()
        Me.HandleErrorsToFabMat = New System.Windows.Forms.BindingSource(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.TxtBoxBOMItemsToProcess = New System.Windows.Forms.TextBox()
        Me.TxtBoxDwgsToProcess = New System.Windows.Forms.TextBox()
        Me.LblDwgsToProcess = New System.Windows.Forms.Label()
        Me.LblBOMItemsToProess = New System.Windows.Forms.Label()
        Me.LblRead_Dwgs = New System.Windows.Forms.Label()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.MatrixLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.HandleErrorsToFabMat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(13, 549)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(4)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(887, 27)
        Me.ProgressBar1.TabIndex = 48
        '
        'BtnStart2
        '
        Me.BtnStart2.BackColor = System.Drawing.SystemColors.Control
        Me.BtnStart2.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnStart2.Enabled = False
        Me.BtnStart2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnStart2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnStart2.Location = New System.Drawing.Point(815, 164)
        Me.BtnStart2.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnStart2.Name = "BtnStart2"
        Me.BtnStart2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnStart2.Size = New System.Drawing.Size(85, 39)
        Me.BtnStart2.TabIndex = 47
        Me.BtnStart2.Text = "Start"
        Me.BtnStart2.UseVisualStyleBackColor = False
        '
        'LblProgress
        '
        Me.LblProgress.BackColor = System.Drawing.SystemColors.Control
        Me.LblProgress.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblProgress.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblProgress.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblProgress.Location = New System.Drawing.Point(16, 517)
        Me.LblProgress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblProgress.Size = New System.Drawing.Size(792, 28)
        Me.LblProgress.TabIndex = 45
        Me.LblProgress.Text = "Progress........"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.DwgsNotFoundList)
        Me.GroupBox5.Location = New System.Drawing.Point(20, 120)
        Me.GroupBox5.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox5.Size = New System.Drawing.Size(216, 58)
        Me.GroupBox5.TabIndex = 44
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Comparison"
        Me.GroupBox5.Visible = False
        '
        'DwgsNotFoundList
        '
        Me.DwgsNotFoundList.BackColor = System.Drawing.SystemColors.Window
        Me.DwgsNotFoundList.Cursor = System.Windows.Forms.Cursors.Default
        Me.DwgsNotFoundList.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DwgsNotFoundList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.DwgsNotFoundList.ItemHeight = 17
        Me.DwgsNotFoundList.Location = New System.Drawing.Point(9, 23)
        Me.DwgsNotFoundList.Margin = New System.Windows.Forms.Padding(4)
        Me.DwgsNotFoundList.Name = "DwgsNotFoundList"
        Me.DwgsNotFoundList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DwgsNotFoundList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.DwgsNotFoundList.Size = New System.Drawing.Size(189, 21)
        Me.DwgsNotFoundList.TabIndex = 60
        Me.DwgsNotFoundList.Visible = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.DwgList)
        Me.GroupBox4.Controls.Add(Me.SelectList)
        Me.GroupBox4.Controls.Add(Me.BtnAdd)
        Me.GroupBox4.Controls.Add(Me.BtnAddAll)
        Me.GroupBox4.Controls.Add(Me.BtnRemove)
        Me.GroupBox4.Controls.Add(Me.BtnClear)
        Me.GroupBox4.Location = New System.Drawing.Point(16, 251)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox4.Size = New System.Drawing.Size(884, 262)
        Me.GroupBox4.TabIndex = 43
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Drawings Found in Directory above.                        Selected Drawings"
        '
        'DwgList
        '
        Me.DwgList.BackColor = System.Drawing.SystemColors.Window
        Me.DwgList.Cursor = System.Windows.Forms.Cursors.Default
        Me.DwgList.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DwgList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.DwgList.ItemHeight = 17
        Me.DwgList.Location = New System.Drawing.Point(8, 26)
        Me.DwgList.Margin = New System.Windows.Forms.Padding(4)
        Me.DwgList.Name = "DwgList"
        Me.DwgList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DwgList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.DwgList.Size = New System.Drawing.Size(379, 191)
        Me.DwgList.TabIndex = 16
        '
        'SelectList
        '
        Me.SelectList.BackColor = System.Drawing.SystemColors.Window
        Me.SelectList.Cursor = System.Windows.Forms.Cursors.Default
        Me.SelectList.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.SelectList.ItemHeight = 17
        Me.SelectList.Location = New System.Drawing.Point(493, 23)
        Me.SelectList.Margin = New System.Windows.Forms.Padding(4)
        Me.SelectList.Name = "SelectList"
        Me.SelectList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SelectList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.SelectList.Size = New System.Drawing.Size(379, 191)
        Me.SelectList.TabIndex = 15
        '
        'BtnAdd
        '
        Me.BtnAdd.BackColor = System.Drawing.SystemColors.Control
        Me.BtnAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnAdd.Enabled = False
        Me.BtnAdd.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnAdd.Location = New System.Drawing.Point(396, 26)
        Me.BtnAdd.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnAdd.Size = New System.Drawing.Size(89, 39)
        Me.BtnAdd.TabIndex = 8
        Me.BtnAdd.Text = "&Add ->"
        Me.BtnAdd.UseVisualStyleBackColor = False
        '
        'BtnAddAll
        '
        Me.BtnAddAll.BackColor = System.Drawing.SystemColors.Control
        Me.BtnAddAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnAddAll.Enabled = False
        Me.BtnAddAll.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnAddAll.Location = New System.Drawing.Point(396, 73)
        Me.BtnAddAll.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnAddAll.Name = "BtnAddAll"
        Me.BtnAddAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnAddAll.Size = New System.Drawing.Size(89, 39)
        Me.BtnAddAll.TabIndex = 9
        Me.BtnAddAll.Text = "A&dd All ->"
        Me.BtnAddAll.UseVisualStyleBackColor = False
        '
        'BtnRemove
        '
        Me.BtnRemove.BackColor = System.Drawing.SystemColors.Control
        Me.BtnRemove.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnRemove.Enabled = False
        Me.BtnRemove.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnRemove.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnRemove.Location = New System.Drawing.Point(396, 119)
        Me.BtnRemove.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnRemove.Name = "BtnRemove"
        Me.BtnRemove.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnRemove.Size = New System.Drawing.Size(89, 39)
        Me.BtnRemove.TabIndex = 10
        Me.BtnRemove.Text = "<-Remove"
        Me.BtnRemove.UseVisualStyleBackColor = False
        '
        'BtnClear
        '
        Me.BtnClear.BackColor = System.Drawing.SystemColors.Control
        Me.BtnClear.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnClear.Enabled = False
        Me.BtnClear.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnClear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnClear.Location = New System.Drawing.Point(396, 166)
        Me.BtnClear.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnClear.Size = New System.Drawing.Size(89, 39)
        Me.BtnClear.TabIndex = 11
        Me.BtnClear.Text = "&Clear"
        Me.BtnClear.UseVisualStyleBackColor = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ComboBxRev)
        Me.GroupBox3.Location = New System.Drawing.Point(630, 118)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Size = New System.Drawing.Size(175, 60)
        Me.GroupBox3.TabIndex = 42
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Revision Number"
        '
        'ComboBxRev
        '
        Me.ComboBxRev.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBxRev.FormattingEnabled = True
        Me.ComboBxRev.Location = New System.Drawing.Point(8, 21)
        Me.ComboBxRev.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBxRev.Name = "ComboBxRev"
        Me.ComboBxRev.Size = New System.Drawing.Size(115, 25)
        Me.ComboBxRev.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BtnGetMWInfo)
        Me.GroupBox2.Controls.Add(Me.PathBox)
        Me.GroupBox2.Location = New System.Drawing.Point(17, 186)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Size = New System.Drawing.Size(757, 58)
        Me.GroupBox2.TabIndex = 41
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Select Path"
        '
        'BtnGetMWInfo
        '
        Me.BtnGetMWInfo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnGetMWInfo.Location = New System.Drawing.Point(699, 17)
        Me.BtnGetMWInfo.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnGetMWInfo.Name = "BtnGetMWInfo"
        Me.BtnGetMWInfo.Size = New System.Drawing.Size(51, 33)
        Me.BtnGetMWInfo.TabIndex = 10
        Me.BtnGetMWInfo.Text = "..."
        Me.BtnGetMWInfo.UseVisualStyleBackColor = True
        '
        'PathBox
        '
        Me.PathBox.Location = New System.Drawing.Point(12, 23)
        Me.PathBox.Margin = New System.Windows.Forms.Padding(4)
        Me.PathBox.Name = "PathBox"
        Me.PathBox.Size = New System.Drawing.Size(677, 22)
        Me.PathBox.TabIndex = 9
        Me.PathBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'BtnVideo
        '
        Me.BtnVideo.BackColor = System.Drawing.SystemColors.Control
        Me.BtnVideo.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnVideo.Enabled = False
        Me.BtnVideo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnVideo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnVideo.Location = New System.Drawing.Point(815, 116)
        Me.BtnVideo.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnVideo.Name = "BtnVideo"
        Me.BtnVideo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnVideo.Size = New System.Drawing.Size(85, 39)
        Me.BtnVideo.TabIndex = 37
        Me.BtnVideo.Text = "Video"
        Me.BtnVideo.UseVisualStyleBackColor = False
        '
        'CancelButton_Renamed
        '
        Me.CancelButton_Renamed.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton_Renamed.Cursor = System.Windows.Forms.Cursors.Default
        Me.CancelButton_Renamed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CancelButton_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CancelButton_Renamed.Location = New System.Drawing.Point(815, 209)
        Me.CancelButton_Renamed.Margin = New System.Windows.Forms.Padding(4)
        Me.CancelButton_Renamed.Name = "CancelButton_Renamed"
        Me.CancelButton_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CancelButton_Renamed.Size = New System.Drawing.Size(85, 39)
        Me.CancelButton_Renamed.TabIndex = 38
        Me.CancelButton_Renamed.Text = "&Cancel"
        Me.CancelButton_Renamed.UseVisualStyleBackColor = False
        '
        'MatrixLogo
        '
        Me.MatrixLogo.BackgroundImage = CType(resources.GetObject("MatrixLogo.BackgroundImage"), System.Drawing.Image)
        Me.MatrixLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.MatrixLogo.Location = New System.Drawing.Point(17, 15)
        Me.MatrixLogo.Margin = New System.Windows.Forms.Padding(4)
        Me.MatrixLogo.Name = "MatrixLogo"
        Me.MatrixLogo.Size = New System.Drawing.Size(345, 97)
        Me.MatrixLogo.TabIndex = 53
        Me.MatrixLogo.TabStop = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'TxtBoxBOMItemsToProcess
        '
        Me.TxtBoxBOMItemsToProcess.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtBoxBOMItemsToProcess.ForeColor = System.Drawing.Color.SkyBlue
        Me.TxtBoxBOMItemsToProcess.Location = New System.Drawing.Point(832, 78)
        Me.TxtBoxBOMItemsToProcess.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtBoxBOMItemsToProcess.Name = "TxtBoxBOMItemsToProcess"
        Me.TxtBoxBOMItemsToProcess.Size = New System.Drawing.Size(68, 30)
        Me.TxtBoxBOMItemsToProcess.TabIndex = 299
        Me.TxtBoxBOMItemsToProcess.Visible = False
        '
        'TxtBoxDwgsToProcess
        '
        Me.TxtBoxDwgsToProcess.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtBoxDwgsToProcess.ForeColor = System.Drawing.Color.SkyBlue
        Me.TxtBoxDwgsToProcess.Location = New System.Drawing.Point(832, 38)
        Me.TxtBoxDwgsToProcess.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtBoxDwgsToProcess.Name = "TxtBoxDwgsToProcess"
        Me.TxtBoxDwgsToProcess.Size = New System.Drawing.Size(68, 30)
        Me.TxtBoxDwgsToProcess.TabIndex = 300
        Me.TxtBoxDwgsToProcess.Visible = False
        '
        'LblDwgsToProcess
        '
        Me.LblDwgsToProcess.AutoSize = True
        Me.LblDwgsToProcess.Location = New System.Drawing.Point(688, 47)
        Me.LblDwgsToProcess.Name = "LblDwgsToProcess"
        Me.LblDwgsToProcess.Size = New System.Drawing.Size(129, 16)
        Me.LblDwgsToProcess.TabIndex = 301
        Me.LblDwgsToProcess.Text = "Drawings to process"
        Me.LblDwgsToProcess.Visible = False
        '
        'LblBOMItemsToProess
        '
        Me.LblBOMItemsToProess.AutoSize = True
        Me.LblBOMItemsToProess.Location = New System.Drawing.Point(671, 86)
        Me.LblBOMItemsToProess.Name = "LblBOMItemsToProess"
        Me.LblBOMItemsToProess.Size = New System.Drawing.Size(142, 16)
        Me.LblBOMItemsToProess.TabIndex = 302
        Me.LblBOMItemsToProess.Text = "BOM Items on Drawing"
        Me.LblBOMItemsToProess.Visible = False
        '
        'LblRead_Dwgs
        '
        Me.LblRead_Dwgs.AutoSize = True
        Me.LblRead_Dwgs.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblRead_Dwgs.Location = New System.Drawing.Point(376, 11)
        Me.LblRead_Dwgs.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblRead_Dwgs.Name = "LblRead_Dwgs"
        Me.LblRead_Dwgs.Size = New System.Drawing.Size(429, 29)
        Me.LblRead_Dwgs.TabIndex = 303
        Me.LblRead_Dwgs.Text = "Read Drawings Win11 Version 2024"
        '
        'ReadDwgs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(915, 587)
        Me.Controls.Add(Me.LblRead_Dwgs)
        Me.Controls.Add(Me.LblBOMItemsToProess)
        Me.Controls.Add(Me.LblDwgsToProcess)
        Me.Controls.Add(Me.TxtBoxDwgsToProcess)
        Me.Controls.Add(Me.TxtBoxBOMItemsToProcess)
        Me.Controls.Add(Me.BtnVideo)
        Me.Controls.Add(Me.BtnStart2)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.MatrixLogo)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.CancelButton_Renamed)
        Me.Controls.Add(Me.GroupBox5)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ReadDwgs"
        Me.Text = "Menu---AutoCAD Read Drawings produce Raw Data Spreadsheet"
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.MatrixLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.HandleErrorsToFabMat, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents BtnStart2 As System.Windows.Forms.Button
    Public WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents DwgList As System.Windows.Forms.ListBox
    Public WithEvents SelectList As System.Windows.Forms.ListBox
    Public WithEvents BtnAdd As System.Windows.Forms.Button
    Public WithEvents BtnAddAll As System.Windows.Forms.Button
    Public WithEvents BtnRemove As System.Windows.Forms.Button
    Public WithEvents BtnClear As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBxRev As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents BtnVideo As System.Windows.Forms.Button
    Public WithEvents CancelButton_Renamed As System.Windows.Forms.Button
    Public WithEvents DwgsNotFoundList As System.Windows.Forms.ListBox
    Friend WithEvents MatrixLogo As System.Windows.Forms.PictureBox
    Friend WithEvents HandleErrorsToFabMat As System.Windows.Forms.BindingSource
    Friend WithEvents PathBox As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents BtnGetMWInfo As Button
    Friend WithEvents TxtBoxBOMItemsToProcess As TextBox
    Friend WithEvents TxtBoxDwgsToProcess As TextBox
    Friend WithEvents LblDwgsToProcess As Label
    Friend WithEvents LblBOMItemsToProess As Label
    Friend WithEvents LblRead_Dwgs As Label
End Class
