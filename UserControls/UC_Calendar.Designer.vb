<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class UC_Calendar
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UC_Calendar))
        Me.pnlFIlter = New System.Windows.Forms.Panel()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RBSpace = New System.Windows.Forms.RadioButton()
        Me.RBPer = New System.Windows.Forms.RadioButton()
        Me.RBBene = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.DTPCalendar = New System.Windows.Forms.DateTimePicker()
        Me.cbFinish = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbStart = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbMonth = New System.Windows.Forms.RadioButton()
        Me.rbWeek = New System.Windows.Forms.RadioButton()
        Me.rbDay = New System.Windows.Forms.RadioButton()
        Me.BtnAddSchedule = New System.Windows.Forms.Button()
        Me.BtnRef = New System.Windows.Forms.Button()
        Me.BtnHDown = New System.Windows.Forms.Button()
        Me.BtnHUp = New System.Windows.Forms.Button()
        Me.BtnVDown = New System.Windows.Forms.Button()
        Me.BtnVUp = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pnlCalendarGrid = New System.Windows.Forms.Panel()
        Me.CheckBox6 = New System.Windows.Forms.CheckBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox7 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.pnlFIlter.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlFIlter
        '
        Me.pnlFIlter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlFIlter.Controls.Add(Me.GroupBox2)
        Me.pnlFIlter.Controls.Add(Me.CheckBox1)
        Me.pnlFIlter.Controls.Add(Me.Label3)
        Me.pnlFIlter.Controls.Add(Me.CheckBox6)
        Me.pnlFIlter.Controls.Add(Me.CheckBox4)
        Me.pnlFIlter.Controls.Add(Me.CheckBox2)
        Me.pnlFIlter.Controls.Add(Me.Label4)
        Me.pnlFIlter.Controls.Add(Me.CheckBox7)
        Me.pnlFIlter.Controls.Add(Me.DTPCalendar)
        Me.pnlFIlter.Controls.Add(Me.cbFinish)
        Me.pnlFIlter.Controls.Add(Me.Label1)
        Me.pnlFIlter.Controls.Add(Me.cbStart)
        Me.pnlFIlter.Controls.Add(Me.GroupBox1)
        Me.pnlFIlter.Controls.Add(Me.BtnAddSchedule)
        Me.pnlFIlter.Controls.Add(Me.BtnRef)
        Me.pnlFIlter.Controls.Add(Me.BtnHDown)
        Me.pnlFIlter.Controls.Add(Me.CheckBox3)
        Me.pnlFIlter.Controls.Add(Me.BtnHUp)
        Me.pnlFIlter.Controls.Add(Me.BtnVDown)
        Me.pnlFIlter.Controls.Add(Me.BtnVUp)
        Me.pnlFIlter.Controls.Add(Me.CheckBox5)
        Me.pnlFIlter.Controls.Add(Me.Label2)
        Me.pnlFIlter.Location = New System.Drawing.Point(0, 0)
        Me.pnlFIlter.Name = "pnlFIlter"
        Me.pnlFIlter.Size = New System.Drawing.Size(1302, 80)
        Me.pnlFIlter.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RBSpace)
        Me.GroupBox2.Controls.Add(Me.RBPer)
        Me.GroupBox2.Controls.Add(Me.RBBene)
        Me.GroupBox2.Location = New System.Drawing.Point(257, -2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(101, 77)
        Me.GroupBox2.TabIndex = 36
        Me.GroupBox2.TabStop = False
        '
        'RBSpace
        '
        Me.RBSpace.AutoSize = True
        Me.RBSpace.Location = New System.Drawing.Point(6, 10)
        Me.RBSpace.Name = "RBSpace"
        Me.RBSpace.Size = New System.Drawing.Size(63, 17)
        Me.RBSpace.TabIndex = 28
        Me.RBSpace.TabStop = True
        Me.RBSpace.Text = "სივრცე"
        Me.RBSpace.UseVisualStyleBackColor = True
        '
        'RBPer
        '
        Me.RBPer.AutoSize = True
        Me.RBPer.Location = New System.Drawing.Point(6, 33)
        Me.RBPer.Name = "RBPer"
        Me.RBPer.Size = New System.Drawing.Size(86, 17)
        Me.RBPer.TabIndex = 29
        Me.RBPer.TabStop = True
        Me.RBPer.Text = "თერაპევტი"
        Me.RBPer.UseVisualStyleBackColor = True
        '
        'RBBene
        '
        Me.RBBene.AutoSize = True
        Me.RBBene.Location = New System.Drawing.Point(6, 56)
        Me.RBBene.Name = "RBBene"
        Me.RBBene.Size = New System.Drawing.Size(97, 17)
        Me.RBBene.TabIndex = 30
        Me.RBBene.TabStop = True
        Me.RBBene.Text = "ბენეფიციარი"
        Me.RBBene.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(364, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 13)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "შესრულება:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(127, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 13)
        Me.Label4.TabIndex = 36
        Me.Label4.Text = "-მდე"
        '
        'DTPCalendar
        '
        Me.DTPCalendar.Location = New System.Drawing.Point(71, 55)
        Me.DTPCalendar.Name = "DTPCalendar"
        Me.DTPCalendar.Size = New System.Drawing.Size(181, 20)
        Me.DTPCalendar.TabIndex = 26
        '
        'cbFinish
        '
        Me.cbFinish.FormattingEnabled = True
        Me.cbFinish.Location = New System.Drawing.Point(71, 32)
        Me.cbFinish.Name = "cbFinish"
        Me.cbFinish.Size = New System.Drawing.Size(50, 21)
        Me.cbFinish.TabIndex = 48
        Me.cbFinish.Text = "20:00"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(127, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "-დან"
        '
        'cbStart
        '
        Me.cbStart.FormattingEnabled = True
        Me.cbStart.Location = New System.Drawing.Point(71, 8)
        Me.cbStart.Name = "cbStart"
        Me.cbStart.Size = New System.Drawing.Size(50, 21)
        Me.cbStart.TabIndex = 47
        Me.cbStart.Text = "09:00"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbMonth)
        Me.GroupBox1.Controls.Add(Me.rbWeek)
        Me.GroupBox1.Controls.Add(Me.rbDay)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(62, 75)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'rbMonth
        '
        Me.rbMonth.AutoSize = True
        Me.rbMonth.Location = New System.Drawing.Point(6, 55)
        Me.rbMonth.Name = "rbMonth"
        Me.rbMonth.Size = New System.Drawing.Size(47, 17)
        Me.rbMonth.TabIndex = 2
        Me.rbMonth.TabStop = True
        Me.rbMonth.Text = "თვე"
        Me.rbMonth.UseVisualStyleBackColor = True
        '
        'rbWeek
        '
        Me.rbWeek.AutoSize = True
        Me.rbWeek.Location = New System.Drawing.Point(6, 32)
        Me.rbWeek.Name = "rbWeek"
        Me.rbWeek.Size = New System.Drawing.Size(55, 17)
        Me.rbWeek.TabIndex = 1
        Me.rbWeek.TabStop = True
        Me.rbWeek.Text = "კვირა"
        Me.rbWeek.UseVisualStyleBackColor = True
        '
        'rbDay
        '
        Me.rbDay.AutoSize = True
        Me.rbDay.Location = New System.Drawing.Point(6, 9)
        Me.rbDay.Name = "rbDay"
        Me.rbDay.Size = New System.Drawing.Size(49, 17)
        Me.rbDay.TabIndex = 0
        Me.rbDay.TabStop = True
        Me.rbDay.Text = "დღე"
        Me.rbDay.UseVisualStyleBackColor = True
        '
        'BtnAddSchedule
        '
        Me.BtnAddSchedule.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnAddSchedule.BackgroundImage = CType(resources.GetObject("BtnAddSchedule.BackgroundImage"), System.Drawing.Image)
        Me.BtnAddSchedule.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnAddSchedule.Location = New System.Drawing.Point(1193, 13)
        Me.BtnAddSchedule.Name = "BtnAddSchedule"
        Me.BtnAddSchedule.Size = New System.Drawing.Size(50, 50)
        Me.BtnAddSchedule.TabIndex = 45
        Me.BtnAddSchedule.UseVisualStyleBackColor = True
        '
        'BtnRef
        '
        Me.BtnRef.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnRef.BackgroundImage = CType(resources.GetObject("BtnRef.BackgroundImage"), System.Drawing.Image)
        Me.BtnRef.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnRef.Location = New System.Drawing.Point(1249, 13)
        Me.BtnRef.Name = "BtnRef"
        Me.BtnRef.Size = New System.Drawing.Size(50, 50)
        Me.BtnRef.TabIndex = 46
        Me.BtnRef.UseVisualStyleBackColor = True
        '
        'BtnHDown
        '
        Me.BtnHDown.BackgroundImage = CType(resources.GetObject("BtnHDown.BackgroundImage"), System.Drawing.Image)
        Me.BtnHDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnHDown.Location = New System.Drawing.Point(864, 26)
        Me.BtnHDown.Name = "BtnHDown"
        Me.BtnHDown.Size = New System.Drawing.Size(30, 30)
        Me.BtnHDown.TabIndex = 44
        Me.BtnHDown.UseVisualStyleBackColor = True
        '
        'BtnHUp
        '
        Me.BtnHUp.BackgroundImage = CType(resources.GetObject("BtnHUp.BackgroundImage"), System.Drawing.Image)
        Me.BtnHUp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnHUp.Location = New System.Drawing.Point(936, 26)
        Me.BtnHUp.Name = "BtnHUp"
        Me.BtnHUp.Size = New System.Drawing.Size(30, 30)
        Me.BtnHUp.TabIndex = 43
        Me.BtnHUp.UseVisualStyleBackColor = True
        '
        'BtnVDown
        '
        Me.BtnVDown.BackgroundImage = CType(resources.GetObject("BtnVDown.BackgroundImage"), System.Drawing.Image)
        Me.BtnVDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnVDown.Location = New System.Drawing.Point(900, 41)
        Me.BtnVDown.Name = "BtnVDown"
        Me.BtnVDown.Size = New System.Drawing.Size(30, 30)
        Me.BtnVDown.TabIndex = 42
        Me.BtnVDown.UseVisualStyleBackColor = True
        '
        'BtnVUp
        '
        Me.BtnVUp.BackgroundImage = CType(resources.GetObject("BtnVUp.BackgroundImage"), System.Drawing.Image)
        Me.BtnVUp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnVUp.Location = New System.Drawing.Point(900, 8)
        Me.BtnVUp.Name = "BtnVUp"
        Me.BtnVUp.Size = New System.Drawing.Size(30, 30)
        Me.BtnVUp.TabIndex = 41
        Me.BtnVUp.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(200, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "ჩვენება:"
        '
        'pnlCalendarGrid
        '
        Me.pnlCalendarGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlCalendarGrid.AutoScroll = True
        Me.pnlCalendarGrid.BackColor = System.Drawing.Color.White
        Me.pnlCalendarGrid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCalendarGrid.Location = New System.Drawing.Point(3, 81)
        Me.pnlCalendarGrid.Name = "pnlCalendarGrid"
        Me.pnlCalendarGrid.Size = New System.Drawing.Size(1296, 496)
        Me.pnlCalendarGrid.TabIndex = 3
        '
        'CheckBox6
        '
        Me.CheckBox6.AutoSize = True
        Me.CheckBox6.Location = New System.Drawing.Point(444, 52)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(144, 17)
        Me.CheckBox6.TabIndex = 39
        Me.CheckBox6.Text = "პროგრამით გატარება"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Location = New System.Drawing.Point(594, 29)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(140, 17)
        Me.CheckBox4.TabIndex = 37
        Me.CheckBox4.Text = "გაცდენა არასაპატიო"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(594, 8)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(110, 17)
        Me.CheckBox2.TabIndex = 35
        Me.CheckBox2.Text = "შესრულებული"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox7
        '
        Me.CheckBox7.AutoSize = True
        Me.CheckBox7.Location = New System.Drawing.Point(594, 52)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.Size = New System.Drawing.Size(94, 17)
        Me.CheckBox7.TabIndex = 40
        Me.CheckBox7.Text = "გაუქმებული"
        Me.CheckBox7.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(444, 29)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(120, 17)
        Me.CheckBox3.TabIndex = 36
        Me.CheckBox3.Text = "გაცდენა საპატიო"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Location = New System.Drawing.Point(694, 52)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(74, 17)
        Me.CheckBox5.TabIndex = 38
        Me.CheckBox5.Text = "აღდგენა"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(444, 8)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(90, 17)
        Me.CheckBox1.TabIndex = 33
        Me.CheckBox1.Text = "დაგეგმილი"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'UC_Calendar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.pnlCalendarGrid)
        Me.Controls.Add(Me.pnlFIlter)
        Me.Name = "UC_Calendar"
        Me.Size = New System.Drawing.Size(1302, 580)
        Me.pnlFIlter.ResumeLayout(False)
        Me.pnlFIlter.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlFIlter As Panel
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents rbMonth As RadioButton
    Friend WithEvents rbWeek As RadioButton
    Friend WithEvents rbDay As RadioButton
    Friend WithEvents BtnAddSchedule As Button
    Friend WithEvents DTPCalendar As DateTimePicker
    Friend WithEvents BtnRef As Button
    Friend WithEvents BtnHDown As Button
    Friend WithEvents BtnHUp As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents BtnVDown As Button
    Friend WithEvents BtnVUp As Button
    Friend WithEvents RBSpace As RadioButton
    Friend WithEvents RBPer As RadioButton
    Friend WithEvents RBBene As RadioButton
    Friend WithEvents Label2 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label4 As Label
    Friend WithEvents cbFinish As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cbStart As ComboBox
    Friend WithEvents pnlCalendarGrid As Panel
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents CheckBox6 As CheckBox
    Friend WithEvents CheckBox4 As CheckBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox7 As CheckBox
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents CheckBox5 As CheckBox
End Class
