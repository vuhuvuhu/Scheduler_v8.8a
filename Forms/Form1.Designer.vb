<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.mainMenu = New System.Windows.Forms.MenuStrip()
        Me.BtnLogin = New System.Windows.Forms.Button()
        Me.LUser = New System.Windows.Forms.Label()
        Me.pnlMain = New System.Windows.Forms.Panel()
        Me.SuspendLayout()
        '
        'mainMenu
        '
        Me.mainMenu.Location = New System.Drawing.Point(0, 0)
        Me.mainMenu.Name = "mainMenu"
        Me.mainMenu.Size = New System.Drawing.Size(994, 24)
        Me.mainMenu.TabIndex = 0
        Me.mainMenu.Text = "MenuStrip1"
        '
        'BtnLogin
        '
        Me.BtnLogin.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnLogin.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.BtnLogin.ForeColor = System.Drawing.SystemColors.Window
        Me.BtnLogin.Location = New System.Drawing.Point(890, 1)
        Me.BtnLogin.Name = "BtnLogin"
        Me.BtnLogin.Size = New System.Drawing.Size(92, 23)
        Me.BtnLogin.TabIndex = 1
        Me.BtnLogin.Text = "ავტორიზაცია"
        Me.BtnLogin.UseVisualStyleBackColor = False
        '
        'LUser
        '
        Me.LUser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LUser.AutoSize = True
        Me.LUser.Location = New System.Drawing.Point(701, 6)
        Me.LUser.Name = "LUser"
        Me.LUser.Size = New System.Drawing.Size(183, 13)
        Me.LUser.TabIndex = 2
        Me.LUser.Text = "გთხოვთ გაიაროთ ავტორიზაცია"
        '
        'pnlMain
        '
        Me.pnlMain.BackColor = System.Drawing.Color.Transparent
        Me.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMain.Location = New System.Drawing.Point(0, 24)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Size = New System.Drawing.Size(994, 426)
        Me.pnlMain.TabIndex = 3
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(994, 450)
        Me.Controls.Add(Me.pnlMain)
        Me.Controls.Add(Me.LUser)
        Me.Controls.Add(Me.BtnLogin)
        Me.Controls.Add(Me.mainMenu)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.mainMenu
        Me.Name = "Form1"
        Me.Text = "Prospero Scheduler v8.8a"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents mainMenu As MenuStrip
    Friend WithEvents BtnLogin As Button
    Friend WithEvents LUser As Label
    Friend WithEvents pnlMain As Panel
End Class
