<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddPerForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AddPerForm))
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.BtnClear = New System.Windows.Forms.Button()
        Me.BtnPhoto = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TMail = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TTel = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TPN = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TSurname = New System.Windows.Forms.TextBox()
        Me.TName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.DTPPer = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LN = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RBActive = New System.Windows.Forms.RadioButton()
        Me.RBPasive = New System.Windows.Forms.RadioButton()
        Me.LblStatus = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnAdd
        '
        Me.BtnAdd.BackColor = System.Drawing.Color.Transparent
        Me.BtnAdd.BackgroundImage = CType(resources.GetObject("BtnAdd.BackgroundImage"), System.Drawing.Image)
        Me.BtnAdd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnAdd.Location = New System.Drawing.Point(387, 196)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(40, 40)
        Me.BtnAdd.TabIndex = 324
        Me.BtnAdd.UseVisualStyleBackColor = False
        Me.BtnAdd.Visible = False
        '
        'BtnClear
        '
        Me.BtnClear.BackColor = System.Drawing.Color.Transparent
        Me.BtnClear.BackgroundImage = CType(resources.GetObject("BtnClear.BackgroundImage"), System.Drawing.Image)
        Me.BtnClear.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnClear.Location = New System.Drawing.Point(397, 164)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(30, 30)
        Me.BtnClear.TabIndex = 323
        Me.BtnClear.UseVisualStyleBackColor = False
        '
        'BtnPhoto
        '
        Me.BtnPhoto.Location = New System.Drawing.Point(327, 135)
        Me.BtnPhoto.Name = "BtnPhoto"
        Me.BtnPhoto.Size = New System.Drawing.Size(100, 23)
        Me.BtnPhoto.TabIndex = 322
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
        Me.PictureBox1.TabIndex = 321
        Me.PictureBox1.TabStop = False
        '
        'TMail
        '
        Me.TMail.Location = New System.Drawing.Point(145, 164)
        Me.TMail.Name = "TMail"
        Me.TMail.Size = New System.Drawing.Size(176, 20)
        Me.TMail.TabIndex = 320
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(12, 167)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 13)
        Me.Label10.TabIndex = 319
        Me.Label10.Text = "ელ. ფოსტა:"
        '
        'TTel
        '
        Me.TTel.Location = New System.Drawing.Point(145, 138)
        Me.TTel.Name = "TTel"
        Me.TTel.Size = New System.Drawing.Size(176, 20)
        Me.TTel.TabIndex = 318
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(12, 141)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(36, 13)
        Me.Label9.TabIndex = 317
        Me.Label9.Text = "ტელ:"
        '
        'TPN
        '
        Me.TPN.Location = New System.Drawing.Point(145, 80)
        Me.TPN.Name = "TPN"
        Me.TPN.Size = New System.Drawing.Size(176, 20)
        Me.TPN.TabIndex = 316
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(26, 13)
        Me.Label5.TabIndex = 315
        Me.Label5.Text = "პ/ნ:"
        '
        'TSurname
        '
        Me.TSurname.Location = New System.Drawing.Point(145, 54)
        Me.TSurname.Name = "TSurname"
        Me.TSurname.Size = New System.Drawing.Size(176, 20)
        Me.TSurname.TabIndex = 314
        '
        'TName
        '
        Me.TName.Location = New System.Drawing.Point(145, 28)
        Me.TName.Name = "TName"
        Me.TName.Size = New System.Drawing.Size(176, 20)
        Me.TName.TabIndex = 313
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(10, 111)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(119, 13)
        Me.Label7.TabIndex = 312
        Me.Label7.Text = "დაბადების თარიღი:"
        '
        'DTPPer
        '
        Me.DTPPer.Location = New System.Drawing.Point(145, 106)
        Me.DTPPer.Name = "DTPPer"
        Me.DTPPer.Size = New System.Drawing.Size(176, 20)
        Me.DTPPer.TabIndex = 311
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 13)
        Me.Label3.TabIndex = 310
        Me.Label3.Text = "გვარი:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 309
        Me.Label2.Text = "სახელი:"
        '
        'LN
        '
        Me.LN.AutoSize = True
        Me.LN.Location = New System.Drawing.Point(90, 8)
        Me.LN.Name = "LN"
        Me.LN.Size = New System.Drawing.Size(39, 13)
        Me.LN.TabIndex = 308
        Me.LN.Text = "Label2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 307
        Me.Label1.Text = "ჩანაწერი N:"
        '
        'RBActive
        '
        Me.RBActive.AutoSize = True
        Me.RBActive.Location = New System.Drawing.Point(145, 190)
        Me.RBActive.Name = "RBActive"
        Me.RBActive.Size = New System.Drawing.Size(71, 17)
        Me.RBActive.TabIndex = 325
        Me.RBActive.TabStop = True
        Me.RBActive.Text = "აქტიური"
        Me.RBActive.UseVisualStyleBackColor = True
        '
        'RBPasive
        '
        Me.RBPasive.AutoSize = True
        Me.RBPasive.Location = New System.Drawing.Point(222, 190)
        Me.RBPasive.Name = "RBPasive"
        Me.RBPasive.Size = New System.Drawing.Size(70, 17)
        Me.RBPasive.TabIndex = 326
        Me.RBPasive.TabStop = True
        Me.RBPasive.Text = "პასიური"
        Me.RBPasive.UseVisualStyleBackColor = True
        '
        'LblStatus
        '
        Me.LblStatus.AutoSize = True
        Me.LblStatus.Location = New System.Drawing.Point(12, 223)
        Me.LblStatus.Name = "LblStatus"
        Me.LblStatus.Size = New System.Drawing.Size(45, 13)
        Me.LblStatus.TabIndex = 327
        Me.LblStatus.Text = "Label11"
        '
        'AddPerForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LemonChiffon
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(435, 243)
        Me.Controls.Add(Me.LblStatus)
        Me.Controls.Add(Me.RBPasive)
        Me.Controls.Add(Me.RBActive)
        Me.Controls.Add(Me.BtnAdd)
        Me.Controls.Add(Me.BtnClear)
        Me.Controls.Add(Me.BtnPhoto)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.TMail)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TTel)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.TPN)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TSurname)
        Me.Controls.Add(Me.TName)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.DTPPer)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LN)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AddPerForm"
        Me.Text = "პერსონალის დამატება"
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
    Friend WithEvents TPN As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents TSurname As TextBox
    Friend WithEvents TName As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents DTPPer As DateTimePicker
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents LN As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents RBActive As RadioButton
    Friend WithEvents RBPasive As RadioButton
    Friend WithEvents LblStatus As Label
End Class
