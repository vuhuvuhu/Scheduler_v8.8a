<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UC_Home
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UC_Home))
        Me.GBGreeting = New System.Windows.Forms.GroupBox()
        Me.LWish = New System.Windows.Forms.Label()
        Me.LUserName = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GBNow = New System.Windows.Forms.GroupBox()
        Me.LWeekDay = New System.Windows.Forms.Label()
        Me.LDate = New System.Windows.Forms.Label()
        Me.LTime = New System.Windows.Forms.Label()
        Me.GBBD = New System.Windows.Forms.GroupBox()
        Me.GB_Today = New System.Windows.Forms.GroupBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GBRedTasks = New System.Windows.Forms.GroupBox()
        Me.GBTools = New System.Windows.Forms.GroupBox()
        Me.BtnRefresh = New System.Windows.Forms.Button()
        Me.BtnAddAray = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GBGreeting.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GBNow.SuspendLayout()
        Me.GBTools.SuspendLayout()
        Me.SuspendLayout()
        '
        'GBGreeting
        '
        Me.GBGreeting.BackColor = System.Drawing.SystemColors.Control
        Me.GBGreeting.Controls.Add(Me.LWish)
        Me.GBGreeting.Controls.Add(Me.LUserName)
        Me.GBGreeting.Controls.Add(Me.PictureBox1)
        Me.GBGreeting.Location = New System.Drawing.Point(3, 3)
        Me.GBGreeting.Name = "GBGreeting"
        Me.GBGreeting.Size = New System.Drawing.Size(253, 77)
        Me.GBGreeting.TabIndex = 0
        Me.GBGreeting.TabStop = False
        Me.GBGreeting.Text = "მოგესალმებით"
        '
        'LWish
        '
        Me.LWish.AutoSize = True
        Me.LWish.BackColor = System.Drawing.Color.Transparent
        Me.LWish.Location = New System.Drawing.Point(62, 56)
        Me.LWish.Name = "LWish"
        Me.LWish.Size = New System.Drawing.Size(177, 13)
        Me.LWish.TabIndex = 2
        Me.LWish.Text = "გისურვებთ წარმატებულ დღეს"
        '
        'LUserName
        '
        Me.LUserName.AutoSize = True
        Me.LUserName.BackColor = System.Drawing.Color.Transparent
        Me.LUserName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LUserName.Location = New System.Drawing.Point(62, 19)
        Me.LUserName.Name = "LUserName"
        Me.LUserName.Size = New System.Drawing.Size(143, 20)
        Me.LUserName.TabIndex = 1
        Me.LUserName.Text = "მომხმარებელო"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(6, 19)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(50, 50)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'GBNow
        '
        Me.GBNow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBNow.Controls.Add(Me.LWeekDay)
        Me.GBNow.Controls.Add(Me.LDate)
        Me.GBNow.Controls.Add(Me.LTime)
        Me.GBNow.Location = New System.Drawing.Point(776, 3)
        Me.GBNow.Name = "GBNow"
        Me.GBNow.Size = New System.Drawing.Size(238, 77)
        Me.GBNow.TabIndex = 1
        Me.GBNow.TabStop = False
        Me.GBNow.Text = "მიმდინარე დრო:"
        '
        'LWeekDay
        '
        Me.LWeekDay.AutoSize = True
        Me.LWeekDay.BackColor = System.Drawing.Color.Transparent
        Me.LWeekDay.Location = New System.Drawing.Point(91, 56)
        Me.LWeekDay.Name = "LWeekDay"
        Me.LWeekDay.Size = New System.Drawing.Size(63, 13)
        Me.LWeekDay.TabIndex = 2
        Me.LWeekDay.Text = "ორშაბათი"
        '
        'LDate
        '
        Me.LDate.AutoSize = True
        Me.LDate.BackColor = System.Drawing.Color.Transparent
        Me.LDate.Location = New System.Drawing.Point(57, 39)
        Me.LDate.Name = "LDate"
        Me.LDate.Size = New System.Drawing.Size(144, 13)
        Me.LDate.TabIndex = 1
        Me.LDate.Text = "25 სექტემბერი 2025 წელი"
        '
        'LTime
        '
        Me.LTime.AutoSize = True
        Me.LTime.BackColor = System.Drawing.Color.Transparent
        Me.LTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LTime.Location = New System.Drawing.Point(83, 16)
        Me.LTime.Name = "LTime"
        Me.LTime.Size = New System.Drawing.Size(79, 20)
        Me.LTime.TabIndex = 0
        Me.LTime.Text = "12:12:12"
        '
        'GBBD
        '
        Me.GBBD.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBBD.Location = New System.Drawing.Point(776, 298)
        Me.GBBD.Name = "GBBD"
        Me.GBBD.Size = New System.Drawing.Size(238, 100)
        Me.GBBD.TabIndex = 2
        Me.GBBD.TabStop = False
        Me.GBBD.Text = "მოახლოებული დაბადების დღეები:"
        '
        'GB_Today
        '
        Me.GB_Today.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GB_Today.Location = New System.Drawing.Point(776, 86)
        Me.GB_Today.Name = "GB_Today"
        Me.GB_Today.Size = New System.Drawing.Size(238, 100)
        Me.GB_Today.TabIndex = 3
        Me.GB_Today.TabStop = False
        Me.GB_Today.Text = "დღეს"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Location = New System.Drawing.Point(776, 192)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(238, 100)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "აქტიური საქმეები"
        '
        'GBRedTasks
        '
        Me.GBRedTasks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBRedTasks.Location = New System.Drawing.Point(4, 86)
        Me.GBRedTasks.Name = "GBRedTasks"
        Me.GBRedTasks.Size = New System.Drawing.Size(766, 355)
        Me.GBRedTasks.TabIndex = 3
        Me.GBRedTasks.TabStop = False
        Me.GBRedTasks.Text = "მოითხოვს რეაგირებას"
        '
        'GBTools
        '
        Me.GBTools.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBTools.Controls.Add(Me.BtnRefresh)
        Me.GBTools.Controls.Add(Me.BtnAddAray)
        Me.GBTools.Location = New System.Drawing.Point(262, 3)
        Me.GBTools.Name = "GBTools"
        Me.GBTools.Size = New System.Drawing.Size(508, 77)
        Me.GBTools.TabIndex = 4
        Me.GBTools.TabStop = False
        Me.GBTools.Text = "ინსტრუმენტები"
        '
        'BtnRefresh
        '
        Me.BtnRefresh.AutoEllipsis = True
        Me.BtnRefresh.BackColor = System.Drawing.Color.Transparent
        Me.BtnRefresh.BackgroundImage = CType(resources.GetObject("BtnRefresh.BackgroundImage"), System.Drawing.Image)
        Me.BtnRefresh.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnRefresh.FlatAppearance.BorderSize = 0
        Me.BtnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnRefresh.Location = New System.Drawing.Point(68, 13)
        Me.BtnRefresh.Name = "BtnRefresh"
        Me.BtnRefresh.Size = New System.Drawing.Size(56, 56)
        Me.BtnRefresh.TabIndex = 1
        Me.BtnRefresh.UseVisualStyleBackColor = False
        '
        'BtnAddAray
        '
        Me.BtnAddAray.AutoEllipsis = True
        Me.BtnAddAray.BackColor = System.Drawing.Color.Transparent
        Me.BtnAddAray.BackgroundImage = CType(resources.GetObject("BtnAddAray.BackgroundImage"), System.Drawing.Image)
        Me.BtnAddAray.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnAddAray.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnAddAray.FlatAppearance.BorderSize = 0
        Me.BtnAddAray.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnAddAray.Location = New System.Drawing.Point(6, 14)
        Me.BtnAddAray.Name = "BtnAddAray"
        Me.BtnAddAray.Size = New System.Drawing.Size(56, 56)
        Me.BtnAddAray.TabIndex = 0
        Me.BtnAddAray.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button1.AutoEllipsis = True
        Me.Button1.BackColor = System.Drawing.Color.Transparent
        Me.Button1.BackgroundImage = CType(resources.GetObject("Button1.BackgroundImage"), System.Drawing.Image)
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(4, 447)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 56)
        Me.Button1.TabIndex = 5
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button2.AutoEllipsis = True
        Me.Button2.BackColor = System.Drawing.Color.Transparent
        Me.Button2.BackgroundImage = CType(resources.GetObject("Button2.BackgroundImage"), System.Drawing.Image)
        Me.Button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button2.FlatAppearance.BorderSize = 0
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Location = New System.Drawing.Point(129, 447)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(56, 56)
        Me.Button2.TabIndex = 1
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(66, 469)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 20)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Label1"
        '
        'UC_Home
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GBTools)
        Me.Controls.Add(Me.GBRedTasks)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GB_Today)
        Me.Controls.Add(Me.GBBD)
        Me.Controls.Add(Me.GBNow)
        Me.Controls.Add(Me.GBGreeting)
        Me.Name = "UC_Home"
        Me.Size = New System.Drawing.Size(1017, 506)
        Me.GBGreeting.ResumeLayout(False)
        Me.GBGreeting.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GBNow.ResumeLayout(False)
        Me.GBNow.PerformLayout()
        Me.GBTools.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GBGreeting As GroupBox
    Friend WithEvents GBNow As GroupBox
    Friend WithEvents GBBD As GroupBox
    Friend WithEvents GB_Today As GroupBox
    Friend WithEvents LWish As Label
    Friend WithEvents LUserName As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents LWeekDay As Label
    Friend WithEvents LDate As Label
    Friend WithEvents LTime As Label
    Friend WithEvents Timer1 As Timer
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GBRedTasks As GroupBox
    Friend WithEvents GBTools As GroupBox
    Friend WithEvents BtnRefresh As Button
    Friend WithEvents BtnAddAray As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
End Class
