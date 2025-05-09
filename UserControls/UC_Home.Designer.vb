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
        Me.GBActiveTasks = New System.Windows.Forms.GroupBox()
        Me.GBRedTasks = New System.Windows.Forms.GroupBox()
        Me.GBTools = New System.Windows.Forms.GroupBox()
        Me.BtnRefresh = New System.Windows.Forms.Button()
        Me.BtnAddAray = New System.Windows.Forms.Button()
        Me.BtnPrev = New System.Windows.Forms.Button()
        Me.BtnNext = New System.Windows.Forms.Button()
        Me.LPage = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.LSe = New System.Windows.Forms.Label()
        Me.LDone = New System.Windows.Forms.Label()
        Me.LNDone = New System.Windows.Forms.Label()
        Me.LBenes = New System.Windows.Forms.Label()
        Me.LPers = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.LNeedReactionToday = New System.Windows.Forms.Label()
        Me.LNeerReaction = New System.Windows.Forms.Label()
        Me.GBGreeting.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GBNow.SuspendLayout()
        Me.GB_Today.SuspendLayout()
        Me.GBActiveTasks.SuspendLayout()
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
        Me.GBBD.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBBD.Location = New System.Drawing.Point(776, 367)
        Me.GBBD.Name = "GBBD"
        Me.GBBD.Size = New System.Drawing.Size(238, 100)
        Me.GBBD.TabIndex = 2
        Me.GBBD.TabStop = False
        Me.GBBD.Text = "მოახლოებული დაბადების დღეები:"
        '
        'GB_Today
        '
        Me.GB_Today.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GB_Today.Controls.Add(Me.LPers)
        Me.GB_Today.Controls.Add(Me.LBenes)
        Me.GB_Today.Controls.Add(Me.LNDone)
        Me.GB_Today.Controls.Add(Me.LDone)
        Me.GB_Today.Controls.Add(Me.LSe)
        Me.GB_Today.Controls.Add(Me.Label5)
        Me.GB_Today.Controls.Add(Me.Label4)
        Me.GB_Today.Controls.Add(Me.Label3)
        Me.GB_Today.Controls.Add(Me.Label2)
        Me.GB_Today.Controls.Add(Me.Label1)
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
        'GBActiveTasks
        '
        Me.GBActiveTasks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBActiveTasks.Controls.Add(Me.LNeedReactionToday)
        Me.GBActiveTasks.Controls.Add(Me.Label7)
        Me.GBActiveTasks.Controls.Add(Me.Label6)
        Me.GBActiveTasks.Controls.Add(Me.LNeerReaction)
        Me.GBActiveTasks.Location = New System.Drawing.Point(776, 192)
        Me.GBActiveTasks.Name = "GBActiveTasks"
        Me.GBActiveTasks.Size = New System.Drawing.Size(238, 169)
        Me.GBActiveTasks.TabIndex = 3
        Me.GBActiveTasks.TabStop = False
        Me.GBActiveTasks.Text = "აქტიური საქმეები"
        '
        'GBRedTasks
        '
        Me.GBRedTasks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBRedTasks.Location = New System.Drawing.Point(4, 86)
        Me.GBRedTasks.Name = "GBRedTasks"
        Me.GBRedTasks.Size = New System.Drawing.Size(766, 381)
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
        'BtnPrev
        '
        Me.BtnPrev.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnPrev.AutoEllipsis = True
        Me.BtnPrev.BackColor = System.Drawing.Color.Transparent
        Me.BtnPrev.BackgroundImage = CType(resources.GetObject("BtnPrev.BackgroundImage"), System.Drawing.Image)
        Me.BtnPrev.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnPrev.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnPrev.FlatAppearance.BorderSize = 0
        Me.BtnPrev.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnPrev.Location = New System.Drawing.Point(659, 473)
        Me.BtnPrev.Name = "BtnPrev"
        Me.BtnPrev.Size = New System.Drawing.Size(30, 30)
        Me.BtnPrev.TabIndex = 5
        Me.BtnPrev.UseVisualStyleBackColor = False
        '
        'BtnNext
        '
        Me.BtnNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnNext.AutoEllipsis = True
        Me.BtnNext.BackColor = System.Drawing.Color.Transparent
        Me.BtnNext.BackgroundImage = CType(resources.GetObject("BtnNext.BackgroundImage"), System.Drawing.Image)
        Me.BtnNext.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnNext.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnNext.FlatAppearance.BorderSize = 0
        Me.BtnNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnNext.Location = New System.Drawing.Point(740, 473)
        Me.BtnNext.Name = "BtnNext"
        Me.BtnNext.Size = New System.Drawing.Size(30, 30)
        Me.BtnNext.TabIndex = 1
        Me.BtnNext.UseVisualStyleBackColor = False
        '
        'LPage
        '
        Me.LPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LPage.AutoSize = True
        Me.LPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LPage.Location = New System.Drawing.Point(695, 482)
        Me.LPage.Name = "LPage"
        Me.LPage.Size = New System.Drawing.Size(39, 13)
        Me.LPage.TabIndex = 6
        Me.LPage.Text = "Label1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(196, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ჩანიშნული სეანსების რაოდენობა:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(6, 29)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "შესრულებული:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Location = New System.Drawing.Point(6, 42)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "შესასრულებელი:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Location = New System.Drawing.Point(6, 55)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "ბენეფიციარები:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Location = New System.Drawing.Point(6, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "თერაპევტები:"
        '
        'LSe
        '
        Me.LSe.AutoSize = True
        Me.LSe.BackColor = System.Drawing.Color.Transparent
        Me.LSe.Location = New System.Drawing.Point(199, 16)
        Me.LSe.Name = "LSe"
        Me.LSe.Size = New System.Drawing.Size(25, 13)
        Me.LSe.TabIndex = 5
        Me.LSe.Text = "123"
        '
        'LDone
        '
        Me.LDone.AutoSize = True
        Me.LDone.BackColor = System.Drawing.Color.Transparent
        Me.LDone.Location = New System.Drawing.Point(199, 29)
        Me.LDone.Name = "LDone"
        Me.LDone.Size = New System.Drawing.Size(25, 13)
        Me.LDone.TabIndex = 6
        Me.LDone.Text = "123"
        '
        'LNDone
        '
        Me.LNDone.AutoSize = True
        Me.LNDone.BackColor = System.Drawing.Color.Transparent
        Me.LNDone.Location = New System.Drawing.Point(199, 42)
        Me.LNDone.Name = "LNDone"
        Me.LNDone.Size = New System.Drawing.Size(25, 13)
        Me.LNDone.TabIndex = 7
        Me.LNDone.Text = "123"
        '
        'LBenes
        '
        Me.LBenes.AutoSize = True
        Me.LBenes.BackColor = System.Drawing.Color.Transparent
        Me.LBenes.Location = New System.Drawing.Point(199, 55)
        Me.LBenes.Name = "LBenes"
        Me.LBenes.Size = New System.Drawing.Size(25, 13)
        Me.LBenes.TabIndex = 8
        Me.LBenes.Text = "123"
        '
        'LPers
        '
        Me.LPers.AutoSize = True
        Me.LPers.BackColor = System.Drawing.Color.Transparent
        Me.LPers.Location = New System.Drawing.Point(199, 68)
        Me.LPers.Name = "LPers"
        Me.LPers.Size = New System.Drawing.Size(25, 13)
        Me.LPers.TabIndex = 9
        Me.LPers.Text = "123"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Location = New System.Drawing.Point(6, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(133, 13)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "რეაგირებას მოითხოვს:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Location = New System.Drawing.Point(6, 29)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(148, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "მათ შორის დღევანდელი:"
        '
        'LNeedReactionToday
        '
        Me.LNeedReactionToday.AutoSize = True
        Me.LNeedReactionToday.BackColor = System.Drawing.Color.Transparent
        Me.LNeedReactionToday.Location = New System.Drawing.Point(199, 29)
        Me.LNeedReactionToday.Name = "LNeedReactionToday"
        Me.LNeedReactionToday.Size = New System.Drawing.Size(25, 13)
        Me.LNeedReactionToday.TabIndex = 11
        Me.LNeedReactionToday.Text = "123"
        '
        'LNeerReaction
        '
        Me.LNeerReaction.AutoSize = True
        Me.LNeerReaction.BackColor = System.Drawing.Color.Transparent
        Me.LNeerReaction.Location = New System.Drawing.Point(199, 16)
        Me.LNeerReaction.Name = "LNeerReaction"
        Me.LNeerReaction.Size = New System.Drawing.Size(25, 13)
        Me.LNeerReaction.TabIndex = 10
        Me.LNeerReaction.Text = "123"
        '
        'UC_Home
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.LPage)
        Me.Controls.Add(Me.BtnNext)
        Me.Controls.Add(Me.BtnPrev)
        Me.Controls.Add(Me.GBTools)
        Me.Controls.Add(Me.GBRedTasks)
        Me.Controls.Add(Me.GBActiveTasks)
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
        Me.GB_Today.ResumeLayout(False)
        Me.GB_Today.PerformLayout()
        Me.GBActiveTasks.ResumeLayout(False)
        Me.GBActiveTasks.PerformLayout()
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
    Friend WithEvents GBActiveTasks As GroupBox
    Friend WithEvents GBRedTasks As GroupBox
    Friend WithEvents GBTools As GroupBox
    Friend WithEvents BtnRefresh As Button
    Friend WithEvents BtnAddAray As Button
    Friend WithEvents BtnPrev As Button
    Friend WithEvents BtnNext As Button
    Friend WithEvents LPage As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents LPers As Label
    Friend WithEvents LBenes As Label
    Friend WithEvents LNDone As Label
    Friend WithEvents LDone As Label
    Friend WithEvents LSe As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents LNeedReactionToday As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents LNeerReaction As Label
    Friend WithEvents Label6 As Label
End Class
