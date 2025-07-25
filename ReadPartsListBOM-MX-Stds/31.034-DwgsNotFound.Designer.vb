<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DwgsNotFound
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.LblDwgsNotFound = New System.Windows.Forms.Label
        Me.DwgsNotFoundList = New System.Windows.Forms.ListBox
        Me.BtnNo = New System.Windows.Forms.Button
        Me.BtnYes = New System.Windows.Forms.Button
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(41, 36)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(332, 21)
        Me.ComboBox1.TabIndex = 1
        '
        'LblDwgsNotFound
        '
        Me.LblDwgsNotFound.AutoSize = True
        Me.LblDwgsNotFound.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblDwgsNotFound.Location = New System.Drawing.Point(38, 9)
        Me.LblDwgsNotFound.Name = "LblDwgsNotFound"
        Me.LblDwgsNotFound.Size = New System.Drawing.Size(335, 19)
        Me.LblDwgsNotFound.TabIndex = 2
        Me.LblDwgsNotFound.Text = "The following Standard Drawings were not found."
        '
        'DwgsNotFoundList
        '
        Me.DwgsNotFoundList.BackColor = System.Drawing.SystemColors.Window
        Me.DwgsNotFoundList.Cursor = System.Windows.Forms.Cursors.Default
        Me.DwgsNotFoundList.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DwgsNotFoundList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.DwgsNotFoundList.Location = New System.Drawing.Point(41, 63)
        Me.DwgsNotFoundList.Name = "DwgsNotFoundList"
        Me.DwgsNotFoundList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DwgsNotFoundList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.DwgsNotFoundList.Size = New System.Drawing.Size(332, 160)
        Me.DwgsNotFoundList.TabIndex = 16
        '
        'BtnNo
        '
        Me.BtnNo.BackColor = System.Drawing.SystemColors.Control
        Me.BtnNo.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnNo.Enabled = False
        Me.BtnNo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnNo.Location = New System.Drawing.Point(521, 81)
        Me.BtnNo.Name = "BtnNo"
        Me.BtnNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnNo.Size = New System.Drawing.Size(64, 32)
        Me.BtnNo.TabIndex = 61
        Me.BtnNo.Text = "No"
        Me.BtnNo.UseVisualStyleBackColor = False
        '
        'BtnYes
        '
        Me.BtnYes.BackColor = System.Drawing.SystemColors.Control
        Me.BtnYes.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnYes.Enabled = False
        Me.BtnYes.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnYes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnYes.Location = New System.Drawing.Point(404, 81)
        Me.BtnYes.Name = "BtnYes"
        Me.BtnYes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnYes.Size = New System.Drawing.Size(64, 32)
        Me.BtnYes.TabIndex = 60
        Me.BtnYes.Text = "Yes"
        Me.BtnYes.UseVisualStyleBackColor = False
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.Window
        Me.ListBox1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ListBox1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ListBox1.Location = New System.Drawing.Point(404, 125)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBox1.Size = New System.Drawing.Size(181, 498)
        Me.ListBox1.TabIndex = 59
        Me.ListBox1.Visible = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(401, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(184, 69)
        Me.Label1.TabIndex = 58
        Me.Label1.Text = "The following Standard Drawings were not found, Do you want to continue ?"
        Me.Label1.Visible = False
        '
        'DwgsNotFound
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(607, 632)
        Me.Controls.Add(Me.BtnNo)
        Me.Controls.Add(Me.BtnYes)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DwgsNotFoundList)
        Me.Controls.Add(Me.LblDwgsNotFound)
        Me.Controls.Add(Me.ComboBox1)
        Me.Name = "DwgsNotFound"
        Me.Text = "DwgsNotFound"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents LblDwgsNotFound As System.Windows.Forms.Label
    Public WithEvents DwgsNotFoundList As System.Windows.Forms.ListBox
    Public WithEvents BtnNo As System.Windows.Forms.Button
    Public WithEvents BtnYes As System.Windows.Forms.Button
    Public WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
