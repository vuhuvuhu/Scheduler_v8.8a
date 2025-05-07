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
        Me.GBBD = New System.Windows.Forms.GroupBox()
        Me.GB_Today = New System.Windows.Forms.GroupBox()
        Me.LTime = New System.Windows.Forms.Label()
        Me.LDate = New System.Windows.Forms.Label()
        Me.LWeekDay = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GBGreeting.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GBNow.SuspendLayout()
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
        Me.GBNow.Controls.Add(Me.LWeekDay)
        Me.GBNow.Controls.Add(Me.LDate)
        Me.GBNow.Controls.Add(Me.LTime)
        Me.GBNow.Location = New System.Drawing.Point(262, 3)
        Me.GBNow.Name = "GBNow"
        Me.GBNow.Size = New System.Drawing.Size(155, 77)
        Me.GBNow.TabIndex = 1
        Me.GBNow.TabStop = False
        Me.GBNow.Text = "მიმდინარე დრო:"
        '
        'GBBD
        '
        Me.GBBD.Location = New System.Drawing.Point(73, 167)
        Me.GBBD.Name = "GBBD"
        Me.GBBD.Size = New System.Drawing.Size(238, 100)
        Me.GBBD.TabIndex = 2
        Me.GBBD.TabStop = False
        Me.GBBD.Text = "მოახლოებული დაბადების დღეები:"
        '
        'GB_Today
        '
        Me.GB_Today.Location = New System.Drawing.Point(366, 186)
        Me.GB_Today.Name = "GB_Today"
        Me.GB_Today.Size = New System.Drawing.Size(200, 100)
        Me.GB_Today.TabIndex = 3
        Me.GB_Today.TabStop = False
        Me.GB_Today.Text = "დღეს"
        '
        'LTime
        '
        Me.LTime.AutoSize = True
        Me.LTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LTime.Location = New System.Drawing.Point(37, 16)
        Me.LTime.Name = "LTime"
        Me.LTime.Size = New System.Drawing.Size(79, 20)
        Me.LTime.TabIndex = 0
        Me.LTime.Text = "12:12:12"
        '
        'LDate
        '
        Me.LDate.AutoSize = True
        Me.LDate.Location = New System.Drawing.Point(6, 43)
        Me.LDate.Name = "LDate"
        Me.LDate.Size = New System.Drawing.Size(144, 13)
        Me.LDate.TabIndex = 1
        Me.LDate.Text = "25 სექტემბერი 2025 წელი"
        '
        'LWeekDay
        '
        Me.LWeekDay.AutoSize = True
        Me.LWeekDay.Location = New System.Drawing.Point(44, 56)
        Me.LWeekDay.Name = "LWeekDay"
        Me.LWeekDay.Size = New System.Drawing.Size(63, 13)
        Me.LWeekDay.TabIndex = 2
        Me.LWeekDay.Text = "ორშაბათი"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'UC_Home
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
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
        Me.ResumeLayout(False)

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
End Class
