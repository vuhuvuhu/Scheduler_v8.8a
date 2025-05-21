' ===========================================
' 📄 Services/SessionCardFactory.vb
' -------------------------------------------
' სესიის ბარათების შექმნის სერვისი - შესწორებული რედაქტირების ღილაკის ფუნქციონალით
' ===========================================
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services
    ''' <summary>
    ''' სესიის ბარათების შექმნის კლასი - პასუხისმგებელია კალენდარზე ბარათების შექმნაზე
    ''' </summary>
    Public Class SessionCardFactory
        ' კონსტანტები
        Private Const HEADER_HEIGHT As Integer = 24 ' ბარათის ჰედერის სიმაღლე

        ' დელეგატი რედაქტირების ღილაკის დააჭერისათვის
        Public Delegate Sub SessionCardActionDelegate(sessionId As Integer)

        ''' <summary>
        ''' სესიის ბარათის შექმნა
        ''' </summary>
        Public Shared Function CreateSessionCard(session As SessionModel,
                                              cardWidth As Integer,
                                              cardHeight As Integer,
                                              cardPosition As Point,
                                              Optional editDelegate As SessionCardActionDelegate = Nothing) As Panel
            Try
                Debug.WriteLine($"CreateSessionCard: ID={session.Id}, სტატუსი='{session.Status}', ზომა: {cardWidth}x{cardHeight}")

                ' ფერების განსაზღვრა
                Dim cardColor As Color = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                Dim borderColor As Color = SessionStatusColors.GetStatusBorderColor(session.Status, session.DateTime)
                Dim headerColor As Color = Color.FromArgb(
                    Math.Max(0, cardColor.R - 50),
                    Math.Max(0, cardColor.G - 50),
                    Math.Max(0, cardColor.B - 50)
                )

                ' ბარათის შექმნა
                Dim sessionCard As New Panel()
                sessionCard.Size = New Size(cardWidth, cardHeight)
                sessionCard.Location = cardPosition
                sessionCard.BackColor = cardColor
                sessionCard.BorderStyle = BorderStyle.None
                sessionCard.Tag = session.Id
                sessionCard.Cursor = Cursors.Hand

                ' ზედა მუქი ზოლი (header)
                Dim headerPanel As New Panel()
                headerPanel.Size = New Size(sessionCard.Width, HEADER_HEIGHT)
                headerPanel.Location = New Point(0, 0)
                headerPanel.BackColor = headerColor
                sessionCard.Controls.Add(headerPanel)

                ' ბენეფიციარის სახელი (ჯერ გვარი, მერე სახელი)
                Dim mtavruliFont As String = "Sylfaen" ' ნაგულისხმევი ფონტი
                Try
                    If FontFamily.Families.Any(Function(f) f.Name = "BPG_Nino_Mtavruli") Then
                        mtavruliFont = "BPG_Nino_Mtavruli"
                    ElseIf FontFamily.Families.Any(Function(f) f.Name = "ALK_Tall_Mtavruli") Then
                        mtavruliFont = "ALK_Tall_Mtavruli"
                    End If
                Catch ex As Exception
                    ' შეცდომის შემთხვევაში ნაგულისხმევ ფონტს ვიყენებთ
                End Try

                Dim lblBeneficiary As New Label()
                Dim beneficiaryText As String = $"{session.BeneficiarySurname.ToUpper()} {session.BeneficiaryName.ToUpper()}"
                lblBeneficiary.Text = beneficiaryText
                lblBeneficiary.Location = New Point(8, 2)
                lblBeneficiary.Size = New Size(headerPanel.Width - 16, 20)
                lblBeneficiary.Font = New Font(mtavruliFont, 8, FontStyle.Bold)
                lblBeneficiary.ForeColor = Color.White
                lblBeneficiary.TextAlign = ContentAlignment.MiddleLeft
                headerPanel.Controls.Add(lblBeneficiary)

                ' კასტომ ჩარჩოს დამატება
                AddHandler sessionCard.Paint, Sub(sender, e)
                                                  ' ძირითადი ფონი
                                                  Using brush As New SolidBrush(cardColor)
                                                      e.Graphics.FillRectangle(brush, sessionCard.ClientRectangle)
                                                  End Using

                                                  ' ჩარჩო
                                                  Using pen As New Pen(borderColor, 2)
                                                      e.Graphics.DrawRectangle(pen, 1, 1, sessionCard.Width - 2, sessionCard.Height - 2)
                                                  End Using
                                              End Sub

                ' დინამიური ლეიბლების განთავსება
                Dim currentY As Integer = HEADER_HEIGHT + 5
                Dim labelSpacing As Integer = 12

                ' სტატუსი
                If cardHeight > HEADER_HEIGHT + 15 Then
                    Dim lblStatus As New Label()
                    lblStatus.Text = session.Status
                    lblStatus.Size = New Size(sessionCard.Width - 10, 10)
                    lblStatus.Location = New Point(5, currentY)
                    lblStatus.Font = New Font("Sylfaen", 6, FontStyle.Bold)
                    lblStatus.ForeColor = Color.DarkSlateGray
                    lblStatus.BackColor = Color.Transparent
                    lblStatus.TextAlign = ContentAlignment.TopCenter
                    sessionCard.Controls.Add(lblStatus)
                    currentY += labelSpacing
                End If

                ' თერაპევტი
                If cardHeight > HEADER_HEIGHT + 30 AndAlso Not String.IsNullOrEmpty(session.TherapistName) Then
                    Dim lblTherapist As New Label()
                    lblTherapist.Text = session.TherapistName
                    lblTherapist.Size = New Size(sessionCard.Width - 10, 11)
                    lblTherapist.Location = New Point(5, currentY)
                    lblTherapist.Font = New Font("Sylfaen", 7, FontStyle.Regular)
                    lblTherapist.ForeColor = Color.FromArgb(40, 40, 40)
                    lblTherapist.BackColor = Color.Transparent
                    lblTherapist.TextAlign = ContentAlignment.TopLeft
                    sessionCard.Controls.Add(lblTherapist)
                    currentY += labelSpacing
                End If

                ' სივრცე ან თერაპიის ტიპი
                If cardHeight > HEADER_HEIGHT + 50 Then
                    If Not String.IsNullOrEmpty(session.Space) Then
                        Dim lblSpace As New Label()
                        lblSpace.Text = $"სივრცე: {session.Space}"
                        lblSpace.Size = New Size(sessionCard.Width - 10, 10)
                        lblSpace.Location = New Point(5, currentY)
                        lblSpace.Font = New Font("Sylfaen", 6, FontStyle.Italic)
                        lblSpace.ForeColor = Color.FromArgb(60, 60, 60)
                        lblSpace.BackColor = Color.Transparent
                        lblSpace.TextAlign = ContentAlignment.TopLeft
                        sessionCard.Controls.Add(lblSpace)
                        currentY += labelSpacing
                    ElseIf Not String.IsNullOrEmpty(session.TherapyType) Then
                        Dim lblTherapyType As New Label()
                        lblTherapyType.Text = session.TherapyType
                        lblTherapyType.Size = New Size(sessionCard.Width - 10, 10)
                        lblTherapyType.Location = New Point(5, currentY)
                        lblTherapyType.Font = New Font("Sylfaen", 6, FontStyle.Italic)
                        lblTherapyType.ForeColor = Color.FromArgb(60, 60, 60)
                        lblTherapyType.BackColor = Color.Transparent
                        lblTherapyType.TextAlign = ContentAlignment.TopLeft
                        sessionCard.Controls.Add(lblTherapyType)
                        currentY += labelSpacing
                    End If
                End If

                ' რედაქტირების ღილაკი
                Dim btnEdit As New Button()
                btnEdit.Text = "✎" ' ფანქრის სიმბოლო
                btnEdit.Font = New Font("Segoe UI Symbol", 10, FontStyle.Bold)
                btnEdit.ForeColor = Color.White
                btnEdit.BackColor = Color.FromArgb(80, 80, 80)
                btnEdit.Size = New Size(24, 24)
                btnEdit.Location = New Point(sessionCard.Width - 30, sessionCard.Height - 30)
                btnEdit.FlatStyle = FlatStyle.Flat
                btnEdit.FlatAppearance.BorderSize = 0
                btnEdit.Tag = session.Id
                btnEdit.Cursor = Cursors.Hand

                ' რედაქტირების ღილაკზე დაჭერის ივენთი
                AddHandler btnEdit.Click, Sub(sender, e)
                                              ' გამოვიძახოთ სესიის რედაქტირების ფორმა პირდაპირ
                                              Try
                                                  Dim sessionId As Integer = CInt(btnEdit.Tag)
                                                  Debug.WriteLine($"რედაქტირების ღილაკზე დაჭერა: სესია ID={sessionId}")

                                                  ' თუ დელეგატი მითითებულია, გამოვიძახოთ
                                                  If editDelegate IsNot Nothing Then
                                                      editDelegate(sessionId)
                                                  End If
                                              Catch ex As Exception
                                                  Debug.WriteLine($"შეცდომა რედაქტირების ღილაკზე: {ex.Message}")
                                              End Try
                                          End Sub

                ' მოვუმრგვალოთ ღილაკი
                Dim btnPath As New GraphicsPath()
                btnPath.AddEllipse(0, 0, btnEdit.Width, btnEdit.Height)
                btnEdit.Region = New Region(btnPath)

                ' დავამატოთ ღილაკი ბარათზე
                sessionCard.Controls.Add(btnEdit)
                btnEdit.BringToFront()

                ' მომრგვალებული კუთხეები
                Try
                    Dim path As New GraphicsPath()
                    Dim cornerRadius As Integer = 6
                    If sessionCard.Width > cornerRadius * 2 AndAlso sessionCard.Height > cornerRadius * 2 Then
                        path.AddArc(0, 0, cornerRadius * 2, cornerRadius * 2, 180, 90)
                        path.AddArc(sessionCard.Width - cornerRadius * 2, 0, cornerRadius * 2, cornerRadius * 2, 270, 90)
                        path.AddArc(sessionCard.Width - cornerRadius * 2, sessionCard.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 0, 90)
                        path.AddArc(0, sessionCard.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 90, 90)
                        path.CloseFigure()
                        sessionCard.Region = New Region(path)
                    End If
                Catch ex As Exception
                    ' Region შექმნის პრობლემისას გაგრძელება
                End Try

                Return sessionCard
            Catch ex As Exception
                Debug.WriteLine($"CreateSessionCard: შეცდომა - {ex.Message}")
                Return New Panel() ' ცარიელი პანელი შეცდომის შემთხვევაში
            End Try
        End Function

        ''' <summary>
        ''' კონტექსტური მენიუს შექმნა სესიის ბარათისთვის
        ''' </summary>
        Public Shared Sub CreateContextMenu(sessionCard As Panel, session As SessionModel, statusChangeHandler As EventHandler)
            Try
                ' შევქმნათ კონტექსტური მენიუ
                Dim contextMenu As New ContextMenuStrip()

                ' სესიის ინფორმაციის მენიუს პუნქტი
                Dim infoMenuItem As New ToolStripMenuItem($"სესია #{session.Id} - {session.BeneficiaryName} {session.BeneficiarySurname}")
                infoMenuItem.Font = New Font(infoMenuItem.Font, FontStyle.Bold)
                infoMenuItem.Enabled = False
                contextMenu.Items.Add(infoMenuItem)

                ' გამყოფი
                contextMenu.Items.Add(New ToolStripSeparator())

                ' რედაქტირების პუნქტი - ივენთ ჰენდლერის გარეშე, ეს უნდა იმართებოდეს UC_Calendar-დან
                Dim editMenuItem As New ToolStripMenuItem("რედაქტირება")
                editMenuItem.Tag = session.Id
                contextMenu.Items.Add(editMenuItem)

                ' სტატუსის შეცვლის ქვემენიუ
                Dim statusMenuItem As New ToolStripMenuItem("სტატუსის შეცვლა")

                ' ყველა შესაძლო სტატუსი
                Dim statuses As String() = {"დაგეგმილი", "შესრულებული", "გაუქმებული", "გაცდენა საპატიო", "გაცდენა არასაპატიო", "აღდგენა", "პროგრამით გატარება"}

                ' დავამატოთ ყველა სტატუსი ქვემენიუში
                For Each status In statuses
                    ' ამჟამინდელი სტატუსი არ გვინდა მენიუში
                    If Not String.Equals(status, session.Status, StringComparison.OrdinalIgnoreCase) Then
                        Dim statusItem As New ToolStripMenuItem(status)

                        ' ვუყენებთ შესაბამის ფერს
                        statusItem.BackColor = SessionStatusColors.GetStatusColor(status, DateTime.Now)

                        ' ვინახავთ სესიის ID-ს და სტატუსს Tag-ში
                        statusItem.Tag = New Tuple(Of Integer, String)(session.Id, status)

                        ' ვამატებთ ქმედებას
                        If statusChangeHandler IsNot Nothing Then
                            AddHandler statusItem.Click, statusChangeHandler
                        End If

                        statusMenuItem.DropDownItems.Add(statusItem)
                    End If
                Next

                contextMenu.Items.Add(statusMenuItem)

                ' გამყოფი
                contextMenu.Items.Add(New ToolStripSeparator())

                ' ინფორმაციის ჩვენების პუნქტი - ივენთ ჰენდლერის გარეშე, ეს უნდა იმართებოდეს UC_Calendar-დან
                Dim showInfoMenuItem As New ToolStripMenuItem("ინფორმაციის ნახვა")
                showInfoMenuItem.Tag = session.Id
                contextMenu.Items.Add(showInfoMenuItem)

                ' კონტექსტური მენიუს მიბმა პანელზე
                sessionCard.ContextMenuStrip = contextMenu

            Catch ex As Exception
                Debug.WriteLine($"CreateContextMenu: შეცდომა - {ex.Message}")
            End Try
        End Sub
    End Class
End Namespace