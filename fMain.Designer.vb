<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fMain
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
        Me.txtConsole = New System.Windows.Forms.TextBox()
        Me.btnRegisterLibraries = New System.Windows.Forms.Button()
        Me.btnShowLibraries = New System.Windows.Forms.Button()
        Me.btnRunMapper = New System.Windows.Forms.Button()
        Me.btnAbout = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtConsole
        '
        Me.txtConsole.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtConsole.BackColor = System.Drawing.Color.White
        Me.txtConsole.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtConsole.Location = New System.Drawing.Point(12, 63)
        Me.txtConsole.Multiline = True
        Me.txtConsole.Name = "txtConsole"
        Me.txtConsole.ReadOnly = True
        Me.txtConsole.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtConsole.Size = New System.Drawing.Size(714, 363)
        Me.txtConsole.TabIndex = 4
        Me.txtConsole.WordWrap = False
        '
        'btnRegisterLibraries
        '
        Me.btnRegisterLibraries.Location = New System.Drawing.Point(12, 12)
        Me.btnRegisterLibraries.Name = "btnRegisterLibraries"
        Me.btnRegisterLibraries.Size = New System.Drawing.Size(120, 26)
        Me.btnRegisterLibraries.TabIndex = 0
        Me.btnRegisterLibraries.Text = "Register Libraries"
        Me.btnRegisterLibraries.UseVisualStyleBackColor = True
        '
        'btnShowLibraries
        '
        Me.btnShowLibraries.Location = New System.Drawing.Point(162, 12)
        Me.btnShowLibraries.Name = "btnShowLibraries"
        Me.btnShowLibraries.Size = New System.Drawing.Size(120, 26)
        Me.btnShowLibraries.TabIndex = 1
        Me.btnShowLibraries.Text = "Show Libraries"
        Me.btnShowLibraries.UseVisualStyleBackColor = True
        '
        'btnRunMapper
        '
        Me.btnRunMapper.Location = New System.Drawing.Point(312, 12)
        Me.btnRunMapper.Name = "btnRunMapper"
        Me.btnRunMapper.Size = New System.Drawing.Size(120, 26)
        Me.btnRunMapper.TabIndex = 2
        Me.btnRunMapper.Text = "Run Mapper"
        Me.btnRunMapper.UseVisualStyleBackColor = True
        '
        'btnAbout
        '
        Me.btnAbout.Location = New System.Drawing.Point(624, 12)
        Me.btnAbout.Name = "btnAbout"
        Me.btnAbout.Size = New System.Drawing.Size(102, 26)
        Me.btnAbout.TabIndex = 3
        Me.btnAbout.Text = "About"
        Me.btnAbout.UseVisualStyleBackColor = True
        '
        'fMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 438)
        Me.Controls.Add(Me.btnAbout)
        Me.Controls.Add(Me.btnRunMapper)
        Me.Controls.Add(Me.btnShowLibraries)
        Me.Controls.Add(Me.btnRegisterLibraries)
        Me.Controls.Add(Me.txtConsole)
        Me.MaximizeBox = False
        Me.Name = "fMain"
        Me.Text = "OP2Mapper2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtConsole As TextBox
    Friend WithEvents btnRegisterLibraries As Button
    Friend WithEvents btnShowLibraries As Button
    Friend WithEvents btnRunMapper As Button
    Friend WithEvents btnAbout As Button
End Class
