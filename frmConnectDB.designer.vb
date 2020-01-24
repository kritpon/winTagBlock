<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConnectDB
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
        Me.components = New System.ComponentModel.Container()
        Me.cmdConnect = New System.Windows.Forms.Button()
        Me.cboDBlist = New System.Windows.Forms.ComboBox()
        Me.lbDBconnect = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.lbTimer1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdConnect
        '
        Me.cmdConnect.BackColor = System.Drawing.Color.DarkOliveGreen
        Me.cmdConnect.Font = New System.Drawing.Font("Tahoma", 10.875!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cmdConnect.ForeColor = System.Drawing.Color.White
        Me.cmdConnect.Location = New System.Drawing.Point(324, 22)
        Me.cmdConnect.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmdConnect.Name = "cmdConnect"
        Me.cmdConnect.Size = New System.Drawing.Size(90, 60)
        Me.cmdConnect.TabIndex = 0
        Me.cmdConnect.Text = "RUN"
        Me.cmdConnect.UseVisualStyleBackColor = False
        '
        'cboDBlist
        '
        Me.cboDBlist.Font = New System.Drawing.Font("Tahoma", 10.875!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cboDBlist.FormattingEnabled = True
        Me.cboDBlist.Location = New System.Drawing.Point(86, 34)
        Me.cboDBlist.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboDBlist.Name = "cboDBlist"
        Me.cboDBlist.Size = New System.Drawing.Size(162, 25)
        Me.cboDBlist.TabIndex = 1
        '
        'lbDBconnect
        '
        Me.lbDBconnect.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbDBconnect.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbDBconnect.ForeColor = System.Drawing.Color.Black
        Me.lbDBconnect.Location = New System.Drawing.Point(18, 98)
        Me.lbDBconnect.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbDBconnect.Name = "lbDBconnect"
        Me.lbDBconnect.Size = New System.Drawing.Size(404, 46)
        Me.lbDBconnect.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.Location = New System.Drawing.Point(18, 37)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "ติดต่อทาง"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'lbTimer1
        '
        Me.lbTimer1.AutoSize = True
        Me.lbTimer1.Location = New System.Drawing.Point(21, 165)
        Me.lbTimer1.Name = "lbTimer1"
        Me.lbTimer1.Size = New System.Drawing.Size(39, 13)
        Me.lbTimer1.TabIndex = 4
        Me.lbTimer1.Text = "Label2"
        '
        'frmConnectDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(476, 212)
        Me.Controls.Add(Me.lbTimer1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lbDBconnect)
        Me.Controls.Add(Me.cboDBlist)
        Me.Controls.Add(Me.cmdConnect)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "frmConnectDB"
        Me.Text = "เลือกฐานข้อมูล"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdConnect As Button
    Friend WithEvents cboDBlist As ComboBox
    Friend WithEvents lbDBconnect As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Timer1 As Timer
    Friend WithEvents lbTimer1 As Label
End Class
