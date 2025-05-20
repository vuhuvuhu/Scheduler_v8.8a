<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UC_Schedule
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UC_Schedule))
        Me.Label9 = New System.Windows.Forms.Label()
        Me.CBSpace = New System.Windows.Forms.ComboBox()
        Me.CBDaf = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.LName = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ChBGroup = New System.Windows.Forms.CheckBox()
        Me.CBTer = New System.Windows.Forms.ComboBox()
        Me.CBPer = New System.Windows.Forms.ComboBox()
        Me.CBBeneSurname = New System.Windows.Forms.ComboBox()
        Me.CBBeneName = New System.Windows.Forms.ComboBox()
        Me.DtpMde = New System.Windows.Forms.DateTimePicker()
        Me.DtpDan = New System.Windows.Forms.DateTimePicker()
        Me.DgvSchedule = New System.Windows.Forms.DataGridView()
        Me.pnlFilter = New System.Windows.Forms.Panel()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.CheckBox7 = New System.Windows.Forms.CheckBox()
        Me.CheckBox6 = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.BtnAddSchedule = New System.Windows.Forms.Button()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.BtnRef = New System.Windows.Forms.Button()
        CType(Me.DgvSchedule, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 54)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(71, 13)
        Me.Label9.TabIndex = 55
        Me.Label9.Text = "შესრულება"
        '
        'CBSpace
        '
        Me.CBSpace.FormattingEnabled = True
        Me.CBSpace.Location = New System.Drawing.Point(793, 27)
        Me.CBSpace.Name = "CBSpace"
        Me.CBSpace.Size = New System.Drawing.Size(182, 21)
        Me.CBSpace.TabIndex = 54
        '
        'CBDaf
        '
        Me.CBDaf.FormattingEnabled = True
        Me.CBDaf.Location = New System.Drawing.Point(793, 0)
        Me.CBDaf.Name = "CBDaf"
        Me.CBDaf.Size = New System.Drawing.Size(182, 21)
        Me.CBDaf.TabIndex = 53
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(706, 29)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(45, 13)
        Me.Label8.TabIndex = 52
        Me.Label8.Text = "სივრცე"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(706, 4)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(81, 13)
        Me.Label7.TabIndex = 51
        Me.Label7.Text = "დაფინანსება"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(444, 30)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(54, 13)
        Me.Label6.TabIndex = 50
        Me.Label6.Text = "თერაპია"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(444, 3)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 13)
        Me.Label5.TabIndex = 49
        Me.Label5.Text = "თერაპევტი"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(262, 30)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 13)
        Me.Label4.TabIndex = 48
        Me.Label4.Text = "გვარი"
        '
        'LName
        '
        Me.LName.AutoSize = True
        Me.LName.Location = New System.Drawing.Point(262, 3)
        Me.LName.Name = "LName"
        Me.LName.Size = New System.Drawing.Size(49, 13)
        Me.LName.TabIndex = 47
        Me.LName.Text = "სახელი"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(209, 30)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "-მდე"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(209, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 13)
        Me.Label2.TabIndex = 45
        Me.Label2.Text = "-დან"
        '
        'ChBGroup
        '
        Me.ChBGroup.AutoSize = True
        Me.ChBGroup.Location = New System.Drawing.Point(981, 3)
        Me.ChBGroup.Name = "ChBGroup"
        Me.ChBGroup.Size = New System.Drawing.Size(80, 17)
        Me.ChBGroup.TabIndex = 44
        Me.ChBGroup.Text = "ჯგუფური"
        Me.ChBGroup.UseVisualStyleBackColor = True
        '
        'CBTer
        '
        Me.CBTer.FormattingEnabled = True
        Me.CBTer.Location = New System.Drawing.Point(518, 27)
        Me.CBTer.Name = "CBTer"
        Me.CBTer.Size = New System.Drawing.Size(182, 21)
        Me.CBTer.TabIndex = 42
        '
        'CBPer
        '
        Me.CBPer.FormattingEnabled = True
        Me.CBPer.Location = New System.Drawing.Point(518, 0)
        Me.CBPer.Name = "CBPer"
        Me.CBPer.Size = New System.Drawing.Size(182, 21)
        Me.CBPer.TabIndex = 41
        '
        'CBBeneSurname
        '
        Me.CBBeneSurname.FormattingEnabled = True
        Me.CBBeneSurname.Location = New System.Drawing.Point(317, 27)
        Me.CBBeneSurname.Name = "CBBeneSurname"
        Me.CBBeneSurname.Size = New System.Drawing.Size(121, 21)
        Me.CBBeneSurname.TabIndex = 40
        '
        'CBBeneName
        '
        Me.CBBeneName.FormattingEnabled = True
        Me.CBBeneName.Location = New System.Drawing.Point(317, 0)
        Me.CBBeneName.Name = "CBBeneName"
        Me.CBBeneName.Size = New System.Drawing.Size(121, 21)
        Me.CBBeneName.TabIndex = 39
        '
        'DtpMde
        '
        Me.DtpMde.Location = New System.Drawing.Point(3, 27)
        Me.DtpMde.Name = "DtpMde"
        Me.DtpMde.Size = New System.Drawing.Size(200, 20)
        Me.DtpMde.TabIndex = 38
        '
        'DtpDan
        '
        Me.DtpDan.Location = New System.Drawing.Point(3, 3)
        Me.DtpDan.Name = "DtpDan"
        Me.DtpDan.Size = New System.Drawing.Size(200, 20)
        Me.DtpDan.TabIndex = 37
        '
        'DgvSchedule
        '
        Me.DgvSchedule.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvSchedule.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvSchedule.Location = New System.Drawing.Point(0, 346)
        Me.DgvSchedule.Name = "DgvSchedule"
        Me.DgvSchedule.Size = New System.Drawing.Size(1161, 151)
        Me.DgvSchedule.TabIndex = 34
        '
        'pnlFilter
        '
        Me.pnlFilter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlFilter.Controls.Add(Me.CheckBox5)
        Me.pnlFilter.Controls.Add(Me.CheckBox7)
        Me.pnlFilter.Controls.Add(Me.CheckBox6)
        Me.pnlFilter.Controls.Add(Me.CheckBox1)
        Me.pnlFilter.Controls.Add(Me.CheckBox4)
        Me.pnlFilter.Controls.Add(Me.BtnAddSchedule)
        Me.pnlFilter.Controls.Add(Me.CheckBox2)
        Me.pnlFilter.Controls.Add(Me.CheckBox3)
        Me.pnlFilter.Controls.Add(Me.BtnRef)
        Me.pnlFilter.Controls.Add(Me.DtpDan)
        Me.pnlFilter.Controls.Add(Me.Label9)
        Me.pnlFilter.Controls.Add(Me.CBSpace)
        Me.pnlFilter.Controls.Add(Me.CBDaf)
        Me.pnlFilter.Controls.Add(Me.DtpMde)
        Me.pnlFilter.Controls.Add(Me.Label8)
        Me.pnlFilter.Controls.Add(Me.CBBeneName)
        Me.pnlFilter.Controls.Add(Me.Label7)
        Me.pnlFilter.Controls.Add(Me.CBBeneSurname)
        Me.pnlFilter.Controls.Add(Me.Label6)
        Me.pnlFilter.Controls.Add(Me.CBPer)
        Me.pnlFilter.Controls.Add(Me.Label5)
        Me.pnlFilter.Controls.Add(Me.CBTer)
        Me.pnlFilter.Controls.Add(Me.Label4)
        Me.pnlFilter.Controls.Add(Me.LName)
        Me.pnlFilter.Controls.Add(Me.ChBGroup)
        Me.pnlFilter.Controls.Add(Me.Label3)
        Me.pnlFilter.Controls.Add(Me.Label2)
        Me.pnlFilter.Location = New System.Drawing.Point(3, 3)
        Me.pnlFilter.Name = "pnlFilter"
        Me.pnlFilter.Size = New System.Drawing.Size(1473, 79)
        Me.pnlFilter.TabIndex = 56
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Location = New System.Drawing.Point(814, 53)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(74, 17)
        Me.CheckBox5.TabIndex = 61
        Me.CheckBox5.Text = "აღდგენა"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'CheckBox7
        '
        Me.CheckBox7.AutoSize = True
        Me.CheckBox7.Location = New System.Drawing.Point(714, 53)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.Size = New System.Drawing.Size(94, 17)
        Me.CheckBox7.TabIndex = 63
        Me.CheckBox7.Text = "გაუქმებული"
        Me.CheckBox7.UseVisualStyleBackColor = True
        '
        'CheckBox6
        '
        Me.CheckBox6.AutoSize = True
        Me.CheckBox6.Location = New System.Drawing.Point(564, 53)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(144, 17)
        Me.CheckBox6.TabIndex = 62
        Me.CheckBox6.Text = "პროგრამით გატარება"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(80, 53)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(90, 17)
        Me.CheckBox1.TabIndex = 57
        Me.CheckBox1.Text = "დაგეგმილი"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Location = New System.Drawing.Point(418, 53)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(140, 17)
        Me.CheckBox4.TabIndex = 60
        Me.CheckBox4.Text = "გაცდენა არასაპატიო"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'BtnAddSchedule
        '
        Me.BtnAddSchedule.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnAddSchedule.BackgroundImage = CType(resources.GetObject("BtnAddSchedule.BackgroundImage"), System.Drawing.Image)
        Me.BtnAddSchedule.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnAddSchedule.Location = New System.Drawing.Point(1363, 4)
        Me.BtnAddSchedule.Name = "BtnAddSchedule"
        Me.BtnAddSchedule.Size = New System.Drawing.Size(50, 50)
        Me.BtnAddSchedule.TabIndex = 64
        Me.BtnAddSchedule.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(176, 53)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(110, 17)
        Me.CheckBox2.TabIndex = 58
        Me.CheckBox2.Text = "შესრულებული"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(292, 53)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(120, 17)
        Me.CheckBox3.TabIndex = 59
        Me.CheckBox3.Text = "გაცდენა საპატიო"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'BtnRef
        '
        Me.BtnRef.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnRef.BackgroundImage = CType(resources.GetObject("BtnRef.BackgroundImage"), System.Drawing.Image)
        Me.BtnRef.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnRef.Location = New System.Drawing.Point(1419, 4)
        Me.BtnRef.Name = "BtnRef"
        Me.BtnRef.Size = New System.Drawing.Size(50, 50)
        Me.BtnRef.TabIndex = 65
        Me.BtnRef.UseVisualStyleBackColor = True
        '
        'UC_Schedule
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.pnlFilter)
        Me.Controls.Add(Me.DgvSchedule)
        Me.Name = "UC_Schedule"
        Me.Size = New System.Drawing.Size(1479, 621)
        CType(Me.DgvSchedule, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFilter.ResumeLayout(False)
        Me.pnlFilter.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label9 As Label
    Friend WithEvents CBSpace As ComboBox
    Friend WithEvents CBDaf As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents LName As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents ChBGroup As CheckBox
    Friend WithEvents CBTer As ComboBox
    Friend WithEvents CBPer As ComboBox
    Friend WithEvents CBBeneSurname As ComboBox
    Friend WithEvents CBBeneName As ComboBox
    Friend WithEvents DtpMde As DateTimePicker
    Friend WithEvents DtpDan As DateTimePicker
    Friend WithEvents DgvSchedule As DataGridView
    Friend WithEvents pnlFilter As Panel
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents CheckBox6 As CheckBox
    Friend WithEvents CheckBox4 As CheckBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox7 As CheckBox
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents CheckBox5 As CheckBox
    Friend WithEvents BtnAddSchedule As Button
    Friend WithEvents BtnRef As Button
End Class
