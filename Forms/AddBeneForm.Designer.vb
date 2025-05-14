<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddBeneForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AddBeneForm))
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.BtnClear = New System.Windows.Forms.Button()
        Me.BtnPhoto = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TMail = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TTel = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TCarmPN = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TCarm = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TPN = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.RBF = New System.Windows.Forms.RadioButton()
        Me.RBM = New System.Windows.Forms.RadioButton()
        Me.TSurname = New System.Windows.Forms.TextBox()
        Me.TName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.DTPBene = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LN = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblStatus = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnAdd
        '
        Me.BtnAdd.BackColor = System.Drawing.Color.Transparent
        Me.BtnAdd.BackgroundImage = CType(resources.GetObject("BtnAdd.BackgroundImage"), System.Drawing.Image)
        Me.BtnAdd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnAdd.Location = New System.Drawing.Point(387, 241)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(40, 40)
        Me.BtnAdd.TabIndex = 288
        Me.BtnAdd.UseVisualStyleBackColor = False
        Me.BtnAdd.Visible = False
        '
        'BtnClear
        '
        Me.BtnClear.BackColor = System.Drawing.Color.Transparent
        Me.BtnClear.BackgroundImage = CType(resources.GetObject("BtnClear.BackgroundImage"), System.Drawing.Image)
        Me.BtnClear.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnClear.Location = New System.Drawing.Point(397, 210)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(30, 30)
        Me.BtnClear.TabIndex = 287
        Me.BtnClear.UseVisualStyleBackColor = False
        '
        'BtnPhoto
        '
        Me.BtnPhoto.Location = New System.Drawing.Point(327, 135)
        Me.BtnPhoto.Name = "BtnPhoto"
        Me.BtnPhoto.Size = New System.Drawing.Size(100, 23)
        Me.BtnPhoto.TabIndex = 286
        Me.BtnPhoto.Text = "ფოტო"
        Me.BtnPhoto.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(327, 28)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(100, 100)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 285
        Me.PictureBox1.TabStop = False
        '
        'TMail
        '
        Me.TMail.Location = New System.Drawing.Point(145, 233)
        Me.TMail.Name = "TMail"
        Me.TMail.Size = New System.Drawing.Size(176, 20)
        Me.TMail.TabIndex = 284
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(12, 236)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 13)
        Me.Label10.TabIndex = 283
        Me.Label10.Text = "ელ. ფოსტა:"
        '
        'TTel
        '
        Me.TTel.Location = New System.Drawing.Point(145, 207)
        Me.TTel.Name = "TTel"
        Me.TTel.Size = New System.Drawing.Size(176, 20)
        Me.TTel.TabIndex = 282
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(12, 210)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(36, 13)
        Me.Label9.TabIndex = 281
        Me.Label9.Text = "ტელ:"
        '
        'TCarmPN
        '
        Me.TCarmPN.Location = New System.Drawing.Point(145, 181)
        Me.TCarmPN.Name = "TCarmPN"
        Me.TCarmPN.Size = New System.Drawing.Size(176, 20)
        Me.TCarmPN.TabIndex = 280
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 184)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(127, 13)
        Me.Label8.TabIndex = 279
        Me.Label8.Text = "წარმომადგენლის პ/ნ:"
        '
        'TCarm
        '
        Me.TCarm.Location = New System.Drawing.Point(145, 155)
        Me.TCarm.Name = "TCarm"
        Me.TCarm.Size = New System.Drawing.Size(176, 20)
        Me.TCarm.TabIndex = 278
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 158)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(108, 13)
        Me.Label6.TabIndex = 277
        Me.Label6.Text = "წარმომადგენელი:"
        '
        'TPN
        '
        Me.TPN.Location = New System.Drawing.Point(145, 80)
        Me.TPN.Name = "TPN"
        Me.TPN.Size = New System.Drawing.Size(176, 20)
        Me.TPN.TabIndex = 276
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(26, 13)
        Me.Label5.TabIndex = 275
        Me.Label5.Text = "პ/ნ:"
        '
        'RBF
        '
        Me.RBF.AutoSize = True
        Me.RBF.Location = New System.Drawing.Point(199, 132)
        Me.RBF.Name = "RBF"
        Me.RBF.Size = New System.Drawing.Size(59, 17)
        Me.RBF.TabIndex = 274
        Me.RBF.TabStop = True
        Me.RBF.Text = "Female"
        Me.RBF.UseVisualStyleBackColor = True
        '
        'RBM
        '
        Me.RBM.AutoSize = True
        Me.RBM.Location = New System.Drawing.Point(145, 132)
        Me.RBM.Name = "RBM"
        Me.RBM.Size = New System.Drawing.Size(48, 17)
        Me.RBM.TabIndex = 273
        Me.RBM.TabStop = True
        Me.RBM.Text = "Male"
        Me.RBM.UseVisualStyleBackColor = True
        '
        'TSurname
        '
        Me.TSurname.Location = New System.Drawing.Point(145, 54)
        Me.TSurname.Name = "TSurname"
        Me.TSurname.Size = New System.Drawing.Size(176, 20)
        Me.TSurname.TabIndex = 272
        '
        'TName
        '
        Me.TName.Location = New System.Drawing.Point(145, 28)
        Me.TName.Name = "TName"
        Me.TName.Size = New System.Drawing.Size(176, 20)
        Me.TName.TabIndex = 271
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(10, 111)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(119, 13)
        Me.Label7.TabIndex = 270
        Me.Label7.Text = "დაბადების თარიღი:"
        '
        'DTPBene
        '
        Me.DTPBene.Location = New System.Drawing.Point(145, 106)
        Me.DTPBene.Name = "DTPBene"
        Me.DTPBene.Size = New System.Drawing.Size(176, 20)
        Me.DTPBene.TabIndex = 269
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 136)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(38, 13)
        Me.Label4.TabIndex = 268
        Me.Label4.Text = "სქესი:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 13)
        Me.Label3.TabIndex = 267
        Me.Label3.Text = "გვარი:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 266
        Me.Label2.Text = "სახელი:"
        '
        'LN
        '
        Me.LN.AutoSize = True
        Me.LN.Location = New System.Drawing.Point(90, 8)
        Me.LN.Name = "LN"
        Me.LN.Size = New System.Drawing.Size(39, 13)
        Me.LN.TabIndex = 265
        Me.LN.Text = "Label2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 264
        Me.Label1.Text = "ჩანაწერი N:"
        '
        'LblStatus
        '
        Me.LblStatus.AutoSize = True
        Me.LblStatus.Location = New System.Drawing.Point(15, 268)
        Me.LblStatus.Name = "LblStatus"
        Me.LblStatus.Size = New System.Drawing.Size(45, 13)
        Me.LblStatus.TabIndex = 289
        Me.LblStatus.Text = "Label11"
        '
        'AddBeneForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(431, 293)
        Me.Controls.Add(Me.LblStatus)
        Me.Controls.Add(Me.BtnAdd)
        Me.Controls.Add(Me.BtnClear)
        Me.Controls.Add(Me.BtnPhoto)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.TMail)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TTel)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.TCarmPN)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TCarm)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TPN)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.RBF)
        Me.Controls.Add(Me.RBM)
        Me.Controls.Add(Me.TSurname)
        Me.Controls.Add(Me.TName)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.DTPBene)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LN)
        Me.Controls.Add(Me.Label1)
        Me.Name = "AddBeneForm"
        Me.Text = "AddBeneForm"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BtnAdd As Button
    Friend WithEvents BtnClear As Button
    Friend WithEvents BtnPhoto As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents TMail As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents TTel As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents TCarmPN As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents TCarm As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents TPN As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents RBF As RadioButton
    Friend WithEvents RBM As RadioButton
    Friend WithEvents TSurname As TextBox
    Friend WithEvents TName As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents DTPBene As DateTimePicker
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents LN As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents LblStatus As Label
End Class
