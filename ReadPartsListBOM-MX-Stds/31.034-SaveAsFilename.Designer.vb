<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class SaveAsFilename
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SaveAsFilename))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.LblFileName = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SaveButton = New System.Windows.Forms.Button()
        Me.CancelButton_Renamed = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LblFileName
        '
        Me.LblFileName.BackColor = System.Drawing.SystemColors.Control
        Me.LblFileName.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblFileName.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblFileName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblFileName.Location = New System.Drawing.Point(12, 9)
        Me.LblFileName.Name = "LblFileName"
        Me.LblFileName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblFileName.Size = New System.Drawing.Size(602, 56)
        Me.LblFileName.TabIndex = 9
        Me.LblFileName.Text = "BOM File Name:"
        '
        'TextBox1
        '
        Me.TextBox1.AcceptsReturn = True
        Me.TextBox1.AllowDrop = True
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Window
        Me.TextBox1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TextBox1.Location = New System.Drawing.Point(17, 127)
        Me.TextBox1.MaxLength = 0
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TextBox1.Size = New System.Drawing.Size(496, 65)
        Me.TextBox1.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(13, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(601, 45)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Below you can type in required Info, or Copy-Clip Text from Windows Explorier."
        '
        'SaveButton
        '
        Me.SaveButton.BackColor = System.Drawing.SystemColors.Control
        Me.SaveButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.SaveButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SaveButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SaveButton.Location = New System.Drawing.Point(519, 122)
        Me.SaveButton.Name = "SaveButton"
        Me.SaveButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SaveButton.Size = New System.Drawing.Size(96, 32)
        Me.SaveButton.TabIndex = 7
        Me.SaveButton.Text = "&Save"
        Me.SaveButton.UseVisualStyleBackColor = False
        '
        'CancelButton_Renamed
        '
        Me.CancelButton_Renamed.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton_Renamed.Cursor = System.Windows.Forms.Cursors.Default
        Me.CancelButton_Renamed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CancelButton_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CancelButton_Renamed.Location = New System.Drawing.Point(519, 160)
        Me.CancelButton_Renamed.Name = "CancelButton_Renamed"
        Me.CancelButton_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CancelButton_Renamed.Size = New System.Drawing.Size(96, 32)
        Me.CancelButton_Renamed.TabIndex = 8
        Me.CancelButton_Renamed.Text = "&Cancel"
        Me.CancelButton_Renamed.UseVisualStyleBackColor = False
        '
        'SaveAsFilename
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(630, 208)
        Me.Controls.Add(Me.LblFileName)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SaveButton)
        Me.Controls.Add(Me.CancelButton_Renamed)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 15)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SaveAsFilename"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public WithEvents LblFileName As Label
    Public WithEvents TextBox1 As TextBox
    Public WithEvents Label1 As Label
    Public WithEvents SaveButton As Button
    Public WithEvents CancelButton_Renamed As Button
#End Region
End Class