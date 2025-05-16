' ===========================================
' 📄 UserControls/UC_Calendar.vb
' -------------------------------------------
' კალენდრის UserControl - ასახავს სესიების კალენდარს
' ===========================================
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_Calendar
    Inherits UserControl

    ' ViewModel კალენდრის მონაცემებისთვის
    Private ReadOnly viewModel As CalendarViewModel

    ' მონაცემთა სერვისი
    Private dataService As IDataService = Nothing

    ' სესიები მეხსიერებაში
    Private allSessions As List(Of SessionModel) = New List(Of SessionModel)()

    ' კალენდრის დღეების ღილაკები
    Private calendarButtons(41) As Button

    ' მომხმარებლის ელფოსტა
    Private userEmail As String = String.Empty

    ' კონტროლები კალენდრის UI-სთვის
    Private btnPrevMonth As Button
    Private btnNextMonth As Button
    Private btnToday As Button
    Private lblMonthYear As Label
    Private pnlCalendar As Panel
    Private pnlDayDetails As Panel
    Private lblSelectedDate As Label
    Private pnlDaySessions As Panel
    Private btnAddSession As Button


    ' ამ ცვლადებს დაამატეთ კლასის დასაწყისში (ვარაუდობთ რომ ფაილი უკვე გაქვთ შექმნილი)

    Private spaces As List(Of String) = New List(Of String)()
    Private timeIntervals As List(Of DateTime) = New List(Of DateTime)()
    Private gridCells(,) As Panel ' უჯრედები სესიების განთავსებისთვის
    Private calendarContentPanel As Panel = Nothing

    ' მაშტაბირების კოეფიციენტები
    Private hScale As Double = 1.0 ' ჰორიზონტალური მასშტაბი
    Private vScale As Double = 1.0 ' ვერტიკალური მასშტაბი

    ' მასშტაბირების ნაბიჯები
    Private Const SCALE_STEP As Double = 0.2 ' მასშტაბის ცვლილების ნაბიჯი
    Private Const MIN_SCALE As Double = 0.5 ' მინიმალური მასშტაბი
    Private Const MAX_SCALE As Double = 2.5 ' მაქსიმალური მასშტაბი

    ' მაშტაბირებამდე ზომები
    Private Const BASE_HEADER_HEIGHT As Integer = 40 ' სათაურების საბაზისო სიმაღლე - არ იცვლება
    Private Const BASE_TIME_COLUMN_WIDTH As Integer = 60 ' დროის სვეტის საბაზისო სიგანე - არ იცვლება
    Private Const BASE_ROW_HEIGHT As Integer = 40 ' მწკრივის საბაზისო სიმაღლე - იცვლება vScale-ით
    Private Const BASE_SPACE_COLUMN_WIDTH As Integer = 150 ' სივრცის სვეტის საბაზისო სიგანე - იცვლება hScale-ით

    ''' <summary>
    ''' მონაცემთა სერვისის მითითება
    ''' </summary>
    Public Sub SetDataService(service As IDataService)
        dataService = service
        Debug.WriteLine("UC_Calendar_Advanced.SetDataService: მითითებულია მონაცემთა სერვისი")

        ' სივრცეებისა და სესიების ჩატვირთვა
        LoadSpaces()
        LoadSessions()

        ' კალენდრის ხედის განახლება
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' მომხმარებლის ელფოსტის მითითება
    ''' </summary>
    Public Sub SetUserEmail(email As String)
        userEmail = email
        Debug.WriteLine($"UC_Calendar_Advanced.SetUserEmail: მითითებულია მომხმარებლის ელფოსტა: {email}")
    End Sub

    ''' <summary>
    ''' დროის ინტერვალების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeTimeIntervals()
        ' გავასუფთავოთ სია
        timeIntervals.Clear()

        ' საწყისი და საბოლოო დრო
        Dim startTime As DateTime
        Dim endTime As DateTime

        ' ვამოწმებთ, არჩეულია თუ არა კომბობოქსები
        If cbStart.SelectedItem IsNot Nothing AndAlso cbFinish.SelectedItem IsNot Nothing Then
            Dim startTimeString As String = cbStart.SelectedItem.ToString()
            Dim endTimeString As String = cbFinish.SelectedItem.ToString()

            ' ვქმნით DateTime ობიექტებს
            Dim baseDate As DateTime = DateTime.Today
            Dim startHour As Integer = Integer.Parse(startTimeString.Split(":"c)(0))
            Dim startMinute As Integer = Integer.Parse(startTimeString.Split(":"c)(1))
            Dim endHour As Integer = Integer.Parse(endTimeString.Split(":"c)(0))
            Dim endMinute As Integer = Integer.Parse(endTimeString.Split(":"c)(1))

            startTime = baseDate.AddHours(startHour).AddMinutes(startMinute)
            endTime = baseDate.AddHours(endHour).AddMinutes(endMinute)
        Else
            ' საწყისი მნიშვნელობები
            startTime = DateTime.Today.AddHours(9) ' 09:00
            endTime = DateTime.Today.AddHours(20) ' 20:00
        End If

        ' ინტერვალების შექმნა ნახევარი საათით
        Dim currentTime As DateTime = startTime
        While currentTime <= endTime
            timeIntervals.Add(currentTime)
            currentTime = currentTime.AddMinutes(30)
        End While

        Debug.WriteLine($"InitializeTimeIntervals: შექმნილია {timeIntervals.Count} დროის ინტერვალი")
    End Sub

    ''' <summary>
    ''' კალენდრის ხედის განახლება
    ''' </summary>
    Private Sub UpdateCalendarView()
        Try
            ' პირველ ჯერზე წამოვიღოთ მონაცემები თუ ცარიელია
            If spaces.Count = 0 Then
                LoadSpaces()
            End If

            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                LoadSessions()
            End If

            ' განვაახლოთ მასშტაბის ლეიბლები
            UpdateScaleLabels()

            ' განვაახლოთ დროის ინტერვალები
            InitializeTimeIntervals()

            ' შევამოწმოთ, რომელი ხედია არჩეული
            If rbDay.Checked Then
                ' დღის ხედი
                If RBSpace.Checked Then
                    ' სივრცეების მიხედვით
                    ShowDayViewBySpace()
                ElseIf RBPer.Checked Then
                    ' თერაპევტების მიხედვით
                    ShowDayViewByTherapist()
                ElseIf RBBene.Checked Then
                    ' ბენეფიციარების მიხედვით
                    ShowDayViewByBeneficiary()
                End If
            ElseIf rbWeek.Checked Then
                ' კვირის ხედი
                ShowWeekView()
            ElseIf rbMonth.Checked Then
                ' თვის ხედი
                ShowMonthView()
            End If
        Catch ex As Exception
            Debug.WriteLine($"UpdateCalendarView: შეცდომა - {ex.Message}")
            MessageBox.Show($"კალენდრის ხედის განახლების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სივრცეების ჩატვირთვა მონაცემთა წყაროდან განსაზღვრული წესრიგით
    ''' </summary>
    Private Sub LoadSpaces()
        Try
            ' გავასუფთავოთ სივრცეების სია
            spaces.Clear()

            ' განსაზღვრული წესრიგის სია - დავამატოთ საკონფერენციო, მშობლები და გარე
            Dim orderedSpaces As New List(Of String) From {
            "მწვანე აბა", "ლურჯი აბა", "სენსორი", "ფიზიკური", "მეტყველება",
            "მუსიკა", "თერაპევტი", "არტი", "სხვა", "საკონფერენციო",
            "მშობლები", "ონლაინ", "სახლი", "გარე"
        }

            ' შევამოწმოთ მონაცემთა სერვისი
            If dataService Is Nothing Then
                Debug.WriteLine("LoadSpaces: მონაცემთა სერვისი არ არის ინიციალიზებული")

                ' თუ მონაცემთა სერვისი არ არის, გამოვიყენოთ განსაზღვრული სია
                spaces.AddRange(orderedSpaces)
                Return
            End If

            ' წამოვიღოთ უნიკალური სივრცეები ყველა სესიიდან
            Dim spaceSet As New HashSet(Of String)()

            ' წამოვიღოთ ყველა სესია ან გამოვიყენოთ ლოკალური ასლი
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                LoadSessions()
            End If

            ' ამოვკრიბოთ უნიკალური სივრცეები
            For Each session In allSessions
                If Not String.IsNullOrEmpty(session.Space?.Trim()) Then
                    spaceSet.Add(session.Space.Trim())
                End If
            Next

            ' თუ ვერ ვიპოვეთ სივრცეები სესიებიდან, ცხრილიდან წამოვიღოთ უშუალოდ
            If spaceSet.Count = 0 AndAlso dataService IsNot Nothing Then
                ' ცხრილიდან სივრცეების წამოღება
                Dim spacesData = dataService.GetData("DB-Space!B2:B")

                If spacesData IsNot Nothing Then
                    For Each row In spacesData
                        If row.Count > 0 AndAlso row(0) IsNot Nothing AndAlso Not String.IsNullOrEmpty(row(0).ToString().Trim()) Then
                            spaceSet.Add(row(0).ToString().Trim())
                        End If
                    Next
                End If
            End If

            ' დავამატოთ სივრცეები განსაზღვრული წესრიგით
            For Each orderedSpace In orderedSpaces
                If spaceSet.Contains(orderedSpace) Then
                    spaces.Add(orderedSpace)
                    spaceSet.Remove(orderedSpace) ' წავშალოთ სეტიდან, რომ აღარ დავამატოთ ხელახლა
                Else
                    ' დავამატოთ სივრცე თუ ის არ არის წყაროში, მაგრამ არის ჩვენს განსაზღვრულ სიაში
                    spaces.Add(orderedSpace)
                    Debug.WriteLine($"LoadSpaces: დაემატა სივრცე სიაში, რომელიც არ იყო წყაროში: {orderedSpace}")
                End If
            Next

            ' დავამატოთ ყველა დარჩენილი სივრცე, რომელიც არ იყო განსაზღვრულ სიაში
            spaces.AddRange(spaceSet.OrderBy(Function(s) s))

            ' თუ საერთოდ ვერ ვიპოვეთ სივრცეები, გამოვიყენოთ საწყისი სია
            If spaces.Count = 0 Then
                spaces.AddRange(orderedSpaces)
            End If

            Debug.WriteLine($"LoadSpaces: ჩატვირთულია {spaces.Count} სივრცე")
        Catch ex As Exception
            Debug.WriteLine($"LoadSpaces: შეცდომა - {ex.Message}")
            ' შეცდომის შემთხვევაში გამოვიყენოთ საწყისი სია - დავამატოთ საკონფერენციო, მშობლები და გარე
            spaces.AddRange(New List(Of String) From {
            "მწვანე აბა", "ლურჯი აბა", "სენსორი", "ფიზიკური", "მეტყველება",
            "მუსიკა", "თერაპევტი", "არტი", "სხვა", "საკონფერენციო",
            "მშობლები", "ონლაინ", "სახლი", "გარე"
        })
        End Try
    End Sub

    ''' <summary>
    ''' ვერტიკალური თარიღის ლეიბლის დახატვა
    ''' </summary>
    Private Sub DateLabel_Paint(sender As Object, e As PaintEventArgs)
        ' ლეიბლის მიღება
        Dim dateLabel As Label = DirectCast(sender, Label)

        ' ტექსტის შებრუნება ვერტიკალურად
        e.Graphics.TranslateTransform(5, dateLabel.Height)
        e.Graphics.RotateTransform(-90)
        Using brush As New SolidBrush(dateLabel.ForeColor)
            e.Graphics.DrawString(dateLabel.Text, dateLabel.Font, brush, 0, 0)
        End Using
    End Sub
    ''' <summary>
    ''' უჯრედის ჩარჩოს დახატვა
    ''' </summary>
    Private Sub Cell_Paint(sender As Object, e As PaintEventArgs)
        Dim cell As Panel = DirectCast(sender, Panel)

        Using pen As New Pen(Color.FromArgb(200, 200, 200), 1)
            ' მხოლოდ მარჯვენა ვერტიკალური ხაზი
            e.Graphics.DrawLine(pen, cell.Width - 1, 0, cell.Width - 1, cell.Height)
        End Using
    End Sub
    ''' <summary>
    ''' სესიის პანელის მომრგვალებული ჩარჩოს დახატვა
    ''' </summary>
    Private Sub SessionPanel_Paint(sender As Object, e As PaintEventArgs)
        Dim sessionPanel As Panel = DirectCast(sender, Panel)

        Dim path As New Drawing2D.GraphicsPath()
        Dim radius As Integer = 5
        Dim rect As New Rectangle(0, 0, sessionPanel.Width, sessionPanel.Height)

        ' მომრგვალებული კუთხეები
        path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90) ' ზედა მარცხენა
        path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90) ' ზედა მარჯვენა
        path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90) ' ქვედა მარჯვენა
        path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90) ' ქვედა მარცხენა
        path.CloseFigure()

        sessionPanel.Region = New Region(path)

        ' დახატოს ჩარჩო
        Using pen As New Pen(Color.FromArgb(150, 150, 150), 1)
            e.Graphics.DrawPath(pen, path)
        End Using
    End Sub
    ''' <summary>
    ''' ჩვენება დღის ხედის სივრცეების მიხედვით - გამოსწორებული ვერსია
    ''' </summary>
    Private Sub ShowDayViewBySpace()
        Try
            ' გავასუფთავოთ კალენდრის პანელი
            pnlCalendarGrid.Controls.Clear()

            ' დავაყენოთ pnlFilter-ის ფერი - ნახევრად გამჭვირვალე თეთრი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            Debug.WriteLine("ShowDayViewBySpace: დაიწყო პანელის ჩვენება")
            Debug.WriteLine($"ShowDayViewBySpace: მასშტაბები - hScale={hScale:F2}, vScale={vScale:F2}")

            ' შევამოწმოთ არის თუ არა სივრცეები ჩატვირთული
            If spaces.Count = 0 Then
                LoadSpaces()
            End If

            ' დროითი ინტერვალები თუ არ არის, შევქმნათ
            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' პარამეტრები გრიდისთვის
            Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT ' უცვლელი - არ იმასშტაბირდება ვერტიკალურად
            Dim TIME_COLUMN_WIDTH As Integer = BASE_TIME_COLUMN_WIDTH ' უცვლელი - არ იმასშტაბირდება ჰორიზონტალურად
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale) ' მხოლოდ მწკრივები იმასშტაბირდება ვერტიკალურად
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale) ' სივრცის სვეტები იმასშტაბირდება ჰორიზონტალურად

            ' დავამატოთ დეტალური დებაგ ინფორმაცია
            Debug.WriteLine($"ShowDayViewBySpace: ზომები - HEADER_HEIGHT={HEADER_HEIGHT}, " &
                     $"TIME_COLUMN_WIDTH={TIME_COLUMN_WIDTH}, " &
                     $"ROW_HEIGHT={ROW_HEIGHT}, " &
                     $"SPACE_COLUMN_WIDTH={SPACE_COLUMN_WIDTH}")

            ' გრიდის მთლიანი სიგანე და სიმაღლე - მთელ ხილულ ფართობზე
            Dim totalWidth As Integer = pnlCalendarGrid.ClientSize.Width - 5 ' 5 პიქსელი დაშორებისთვის
            Dim totalHeight As Integer = pnlCalendarGrid.ClientSize.Height - 5

            Debug.WriteLine($"ShowDayViewBySpace: კონტეინერის ზომები - Width={totalWidth}, Height={totalHeight}")

            ' ======= 1. სივრცეების სათაურების რიგი გრიდის ზემოთ =======
            Dim headerRow As New Panel()
            headerRow.Size = New Size(totalWidth, HEADER_HEIGHT)
            headerRow.Location = New Point(0, 0)
            headerRow.BackColor = Color.FromArgb(220, 220, 240)
            pnlCalendarGrid.Controls.Add(headerRow)

            ' დროის სვეტის სათაური - უცვლელი ზომით
            Dim timeHeaderLabel As New Label()
            timeHeaderLabel.Size = New Size(TIME_COLUMN_WIDTH, HEADER_HEIGHT)
            timeHeaderLabel.Location = New Point(0, 0)
            timeHeaderLabel.BackColor = Color.FromArgb(180, 180, 220)
            timeHeaderLabel.TextAlign = ContentAlignment.MiddleCenter
            timeHeaderLabel.Text = "დრო"
            timeHeaderLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            headerRow.Controls.Add(timeHeaderLabel)

            ' მთავრული ფონტი
            Dim mtavruliFont As New Font("Sylfaen", 10, FontStyle.Bold)
            Try
                If FontFamily.Families.Any(Function(f) f.Name = "BPG_Nino_Mtavruli") Then
                    mtavruliFont = New Font("BPG_Nino_Mtavruli", 10, FontStyle.Bold)
                ElseIf FontFamily.Families.Any(Function(f) f.Name = "ALK_Tall_Mtavruli") Then
                    mtavruliFont = New Font("ALK_Tall_Mtavruli", 10, FontStyle.Bold)
                End If
            Catch ex As Exception
                ' ფონტის შეცდომის შემთხვევაში ვიყენებთ Sylfaen
            End Try

            ' სივრცეების სვეტების სიგანე - მნიშვნელოვანია მასშტაბირებისთვის
            Debug.WriteLine($"ShowDayViewBySpace: გამოთვლილი SPACE_COLUMN_WIDTH={SPACE_COLUMN_WIDTH}")

            ' გადავამოწმოთ რომ ყველა სვეტი ეტევა
            Dim availableWidth As Integer = totalWidth - TIME_COLUMN_WIDTH
            Dim totalColumnsWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count

            Debug.WriteLine($"ShowDayViewBySpace: availableWidth={availableWidth}, totalColumnsWidth={totalColumnsWidth}")

            ' თუ საჭიროა, დავაკორექტიროთ სვეტის სიგანე
            Dim adjustedSpaceColumnWidth As Integer = SPACE_COLUMN_WIDTH
            If totalColumnsWidth > availableWidth And spaces.Count > 0 Then
                adjustedSpaceColumnWidth = availableWidth / spaces.Count
                Debug.WriteLine($"ShowDayViewBySpace: სვეტის სიგანე დაკორექტირდა {SPACE_COLUMN_WIDTH}-დან {adjustedSpaceColumnWidth}-მდე")
            End If

            For i As Integer = 0 To spaces.Count - 1
                Dim spaceHeader As New Label()
                spaceHeader.Size = New Size(adjustedSpaceColumnWidth, HEADER_HEIGHT)
                spaceHeader.Location = New Point(TIME_COLUMN_WIDTH + (i * adjustedSpaceColumnWidth), 0)
                spaceHeader.BackColor = Color.FromArgb(60, 80, 150)
                spaceHeader.ForeColor = Color.White
                spaceHeader.TextAlign = ContentAlignment.MiddleCenter
                spaceHeader.Text = spaces(i).ToUpper()
                spaceHeader.Font = mtavruliFont
                headerRow.Controls.Add(spaceHeader)
            Next

            ' ======= 2. გრიდის კონტეინერი =======
            Dim gridContainer As New Panel()
            gridContainer.Size = New Size(totalWidth, totalHeight - HEADER_HEIGHT - 5)
            gridContainer.Location = New Point(0, HEADER_HEIGHT)
            gridContainer.BackColor = Color.White
            gridContainer.AutoScroll = True
            pnlCalendarGrid.Controls.Add(gridContainer)

            ' ======= 3. დროის სვეტი მარცხნივ =======
            Dim timeColumn As New Panel()
            timeColumn.Size = New Size(TIME_COLUMN_WIDTH, timeIntervals.Count * ROW_HEIGHT)
            timeColumn.Location = New Point(0, 0)
            timeColumn.BackColor = Color.FromArgb(240, 240, 245)
            gridContainer.Controls.Add(timeColumn)

            ' დროის ეტიკეტები - მასშტაბირებული სიმაღლით
            For i As Integer = 0 To timeIntervals.Count - 1
                Dim timeLabel As New Label()
                timeLabel.Size = New Size(TIME_COLUMN_WIDTH - 5, ROW_HEIGHT)
                timeLabel.Location = New Point(0, i * ROW_HEIGHT)
                timeLabel.Text = timeIntervals(i).ToString("HH:mm")
                timeLabel.TextAlign = ContentAlignment.MiddleRight
                timeLabel.Padding = New Padding(0, 0, 5, 0)

                ' ალტერნატიული ფერები მწკრივებისთვის
                If i Mod 2 = 0 Then
                    timeLabel.BackColor = Color.FromArgb(240, 240, 240)
                Else
                    timeLabel.BackColor = Color.FromArgb(245, 245, 245)
                End If

                timeColumn.Controls.Add(timeLabel)
            Next

            ' ======= 4. სივრცეების გრიდი =======
            Dim grid As New Panel()
            grid.Size = New Size(adjustedSpaceColumnWidth * spaces.Count, timeIntervals.Count * ROW_HEIGHT)
            grid.Location = New Point(TIME_COLUMN_WIDTH, 0)
            gridContainer.Controls.Add(grid)

            ' უჯრედების მასივი
            ReDim gridCells(spaces.Count - 1, timeIntervals.Count - 1)

            ' უჯრედების შექმნა
            For col As Integer = 0 To spaces.Count - 1
                For row As Integer = 0 To timeIntervals.Count - 1
                    Dim cell As New Panel()
                    cell.Size = New Size(adjustedSpaceColumnWidth, ROW_HEIGHT)
                    cell.Location = New Point(col * adjustedSpaceColumnWidth, row * ROW_HEIGHT)

                    ' ალტერნატიული ფერები
                    If row Mod 2 = 0 Then
                        cell.BackColor = Color.FromArgb(250, 250, 250)
                    Else
                        cell.BackColor = Color.FromArgb(245, 245, 245)
                    End If

                    ' უჯრედის ჩარჩო
                    AddHandler cell.Paint, AddressOf Cell_Paint

                    gridCells(col, row) = cell
                    grid.Controls.Add(cell)
                Next
            Next

            ' ======= 5. სესიების ბარათების განთავსება =======
            If allSessions IsNot Nothing AndAlso allSessions.Count > 0 Then
                Dim selectedDate As DateTime = DTPCalendar.Value.Date
                Dim daySessions = allSessions.Where(Function(s) s.DateTime.Date = selectedDate).ToList()

                ' თუ სესიები ცოტაა, დავამატოთ საცდელი სესიები
                If daySessions.Count < 5 Then
                    daySessions.RemoveAll(Function(s) s.Id >= 990 AndAlso s.Id <= 999)
                    AddTestSessions(daySessions, selectedDate)
                End If

                PlaceSessionsOnGrid(daySessions)
            End If

            Debug.WriteLine("ShowDayViewBySpace: დასრულდა პანელის ჩვენება")
        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewBySpace: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewBySpace: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"დღის ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' საცდელი სესიების დამატება სხვადასხვა დროსა და სივრცეში
    ''' </summary>
    Private Sub AddTestSessions(daySessions As List(Of SessionModel), selectedDate As DateTime)
        Try
            ' პირველი საცდელი სესია - 10:00, მწვანე აბა
            Dim session1 As New SessionModel()
            session1.Id = 999
            session1.BeneficiaryName = "გიორგი"
            session1.BeneficiarySurname = "ბერიძე"
            session1.DateTime = selectedDate.AddHours(10)
            session1.TherapistName = "ნინო თერაპევტი"
            session1.TherapyType = "მეტყველება"
            session1.Space = "მწვანე აბა"
            session1.Duration = 60
            session1.Status = "დაგეგმილი"
            daySessions.Add(session1)

            ' მეორე საცდელი სესია - 12:00, ლურჯი აბა
            Dim session2 As New SessionModel()
            session2.Id = 998
            session2.BeneficiaryName = "მარიამი"
            session2.BeneficiarySurname = "სამხარაძე"
            session2.DateTime = selectedDate.AddHours(12)
            session2.TherapistName = "ვახო კვირკველია"
            session2.TherapyType = "ფიზიკური"
            session2.Space = "ლურჯი აბა"
            session2.Duration = 60
            session2.Status = "დაგეგმილი"
            daySessions.Add(session2)

            ' მესამე საცდელი სესია - 13:00, სენსორი
            Dim session3 As New SessionModel()
            session3.Id = 997
            session3.BeneficiaryName = "ნინო"
            session3.BeneficiarySurname = "ქავთარაძე"
            session3.DateTime = selectedDate.AddHours(13)
            session3.TherapistName = "მაია მენაბდიშვილი"
            session3.TherapyType = "სენსორული"
            session3.Space = "სენსორი"
            session3.Duration = 60
            session3.Status = "დაგეგმილი"
            daySessions.Add(session3)

            ' მეოთხე საცდელი სესია - 12:00, მეტყველება
            Dim session4 As New SessionModel()
            session4.Id = 996
            session4.BeneficiaryName = "გვანცა"
            session4.BeneficiarySurname = "კაპანაძე"
            session4.DateTime = selectedDate.AddHours(12)
            session4.TherapistName = "ლელა ჯიქია"
            session4.TherapyType = "მეტყველება"
            session4.Space = "მეტყველება"
            session4.Duration = 60
            session4.Status = "დაგეგმილი"
            daySessions.Add(session4)

            ' მეხუთე საცდელი სესია - 13:00, ფიზიკური
            Dim session5 As New SessionModel()
            session5.Id = 995
            session5.BeneficiaryName = "ზურა"
            session5.BeneficiarySurname = "ბითაძე"
            session5.DateTime = selectedDate.AddHours(13)
            session5.TherapistName = "მარიამ ბარათაშვილი"
            session5.TherapyType = "ფიზიკური"
            session5.Space = "ფიზიკური"
            session5.Duration = 60
            session5.Status = "დაგეგმილი"
            daySessions.Add(session5)

            Debug.WriteLine($"AddTestSessions: დაემატა 5 საცდელი სესია")
        Catch ex As Exception
            Debug.WriteLine($"AddTestSessions: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების განთავსება გრიდზე - გამოსწორებული ვერსია
    ''' </summary>
    Private Sub PlaceSessionsOnGrid(sessions As List(Of SessionModel))
        Try
            ' შევამოწმოთ უჯრედები და სესიები
            If gridCells Is Nothing OrElse sessions Is Nothing OrElse sessions.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnGrid: gridCells ან sessions არის Nothing/ცარიელი")
                Return
            End If

            Debug.WriteLine($"PlaceSessionsOnGrid: დაიწყო {sessions.Count} სესიის განთავსება")
            Debug.WriteLine($"PlaceSessionsOnGrid: gridCells ზომები: {gridCells.GetLength(0)}x{gridCells.GetLength(1)}")

            For Each session In sessions
                ' ვიპოვოთ სივრცის ინდექსი
                Dim spaceIndex As Integer = spaces.IndexOf(session.Space)
                Debug.WriteLine($"PlaceSessionsOnGrid: სესია ID={session.Id}, სივრცე={session.Space}, spaceIndex={spaceIndex}")

                If spaceIndex < 0 AndAlso spaces.Count > 0 Then
                    Debug.WriteLine($"PlaceSessionsOnGrid: სივრცე '{session.Space}' ვერ მოიძებნა სივრცეების სიაში. ვიყენებთ პირველ სივრცეს.")
                    spaceIndex = 0
                End If

                ' ვიპოვოთ დროის ინდექსი
                Dim timeIndex As Integer = -1

                ' ვცადოთ ზუსტი დროის შესაბამისობა
                For i As Integer = 0 To timeIntervals.Count - 1
                    If timeIntervals(i).Hour = session.DateTime.Hour AndAlso
                   timeIntervals(i).Minute = session.DateTime.Minute Then
                        timeIndex = i
                        Debug.WriteLine($"PlaceSessionsOnGrid: ზუსტი დროის დამთხვევა - {session.DateTime:HH:mm}, timeIndex={i}")
                        Exit For
                    End If
                Next

                ' თუ ზუსტი დამთხვევა არ გვაქვს, ვიპოვოთ უახლოესი
                If timeIndex < 0 Then
                    Dim minDiff As Integer = Integer.MaxValue
                    For i As Integer = 0 To timeIntervals.Count - 1
                        Dim diff As Integer = Math.Abs((timeIntervals(i) - session.DateTime).TotalMinutes)
                        If diff < minDiff Then
                            minDiff = diff
                            timeIndex = i
                        End If
                    Next
                    Debug.WriteLine($"PlaceSessionsOnGrid: უახლოესი დრო - {session.DateTime:HH:mm}, timeIndex={timeIndex}, სხვაობა={minDiff}წთ")
                End If

                ' თუ ინდექსები ვალიდურია
                If spaceIndex >= 0 AndAlso timeIndex >= 0 AndAlso
               spaceIndex < gridCells.GetLength(0) AndAlso timeIndex < gridCells.GetLength(1) Then

                    ' სესიის ხანგრძლივობა (ნახევარსაათიანი ინტერვალების რაოდენობა)
                    Dim durationInCells As Integer = Math.Max(1, Math.Ceiling(session.Duration / 30.0))
                    Debug.WriteLine($"PlaceSessionsOnGrid: ხანგრძლივობა={session.Duration}წთ, durationInCells={durationInCells}")

                    ' ბოლო მწკრივის ინდექსი (შეზღუდული გრიდის ზომით)
                    Dim lastRowIndex As Integer = Math.Min(timeIndex + durationInCells - 1, gridCells.GetLength(1) - 1)

                    ' სესიის სრული სიმაღლე
                    Dim sessionHeight As Integer = (lastRowIndex - timeIndex + 1) * gridCells(spaceIndex, timeIndex).Height
                    Debug.WriteLine($"PlaceSessionsOnGrid: სესიის სიმაღლე={sessionHeight}px")

                    ' უჯრედის ზომები
                    Dim cellWidth As Integer = gridCells(spaceIndex, timeIndex).Width
                    Debug.WriteLine($"PlaceSessionsOnGrid: უჯრედის სიგანე={cellWidth}px")

                    ' შევქმნათ სესიის პანელი
                    Dim sessionPanel As New Panel()
                    sessionPanel.Size = New Size(cellWidth - 4, sessionHeight - 2)
                    sessionPanel.Location = New Point(2, 1)
                    sessionPanel.BackColor = GetSessionColor(session)

                    ' მომრგვალებული კუთხეები
                    AddHandler sessionPanel.Paint, AddressOf SessionPanel_Paint

                    ' სესიის ინფორმაციის დამატება
                    AddSessionInfo(sessionPanel, session)

                    ' სესიის დამატება უჯრედზე
                    gridCells(spaceIndex, timeIndex).Controls.Add(sessionPanel)

                    ' დამატება Tag-ის და Click ივენთის
                    sessionPanel.Tag = session.Id
                    AddHandler sessionPanel.Click, AddressOf SessionPanel_Click

                    Debug.WriteLine($"PlaceSessionsOnGrid: სესია ID={session.Id} განთავსდა უჯრედზე ({spaceIndex},{timeIndex})")
                Else
                    Debug.WriteLine($"PlaceSessionsOnGrid: გადაცილებულია უჯრედების საზღვრებს - spaceIndex={spaceIndex}, timeIndex={timeIndex}, " &
                             $"gridCells ზომა={gridCells.GetLength(0)}x{gridCells.GetLength(1)}")
                End If
            Next

            Debug.WriteLine("PlaceSessionsOnGrid: დასრულდა სესიების განთავსება")
        Catch ex As Exception
            Debug.WriteLine($"PlaceSessionsOnGrid: შეცდომა - {ex.Message}")
            Debug.WriteLine($"PlaceSessionsOnGrid: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის ფერის განსაზღვრა სტატუსის მიხედვით
    ''' </summary>
    Private Function GetSessionColor(session As SessionModel) As Color
        ' სტატუსის მიხედვით ფერის შერჩევა
        Dim status As String = session.Status.Trim().ToLower()

        If status = "შესრულებული" Then
            Return Color.FromArgb(225, 255, 225)      ' ღია მწვანე
        ElseIf status = "გაუქმებული" OrElse status = "გაუქმება" Then
            Return Color.FromArgb(230, 230, 230)      ' ღია ნაცრისფერი
        ElseIf session.IsOverdue Then
            Return Color.FromArgb(255, 225, 225)      ' ღია წითელი
        Else
            Return Color.FromArgb(255, 255, 225)      ' ღია ყვითელი
        End If
    End Function

    ''' <summary>
    ''' სესიის ინფორმაციის დამატება პანელზე - მინიმალისტური და მკაფიო
    ''' </summary>
    Private Sub AddSessionInfo(panel As Panel, session As SessionModel)
        Try
            ' ფიქსირებული ფონტის ზომები - არ იცვლება მასშტაბირებისას
            Dim fontSize1 As Single = 9 ' ძირითადი ტექსტი
            Dim fontSize2 As Single = 8 ' მეორადი ტექსტი

            ' დრო
            Dim lblTime As New Label()
            lblTime.Text = session.DateTime.ToString("HH:mm")
            lblTime.AutoSize = False
            lblTime.Size = New Size(Math.Min(45, panel.Width - 10), 18)
            lblTime.Location = New Point(5, 2)
            lblTime.Font = New Font("Segoe UI", fontSize1, FontStyle.Bold)
            lblTime.TextAlign = ContentAlignment.TopLeft
            panel.Controls.Add(lblTime)

            ' ბენეფიციარი
            Dim lblBeneficiary As New Label()
            lblBeneficiary.Text = $"{session.BeneficiaryName} {session.BeneficiarySurname}"
            lblBeneficiary.AutoSize = False
            lblBeneficiary.Size = New Size(panel.Width - 10, 18)
            lblBeneficiary.Location = New Point(5, 20)
            lblBeneficiary.Font = New Font("Sylfaen", fontSize1, FontStyle.Bold)
            panel.Controls.Add(lblBeneficiary)

            ' თერაპევტი - თუ საკმარისი სიმაღლეა
            If panel.Height > 45 Then
                Dim lblTherapist As New Label()
                lblTherapist.Text = session.TherapistName
                lblTherapist.AutoSize = False
                lblTherapist.Size = New Size(panel.Width - 10, 16)
                lblTherapist.Location = New Point(5, 39)
                lblTherapist.Font = New Font("Sylfaen", fontSize2)
                panel.Controls.Add(lblTherapist)
            End If
        Catch ex As Exception
            Debug.WriteLine($"AddSessionInfo: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ყველა სესიის ჩატვირთვა
    ''' </summary>
    Private Sub LoadSessions()
        Try
            ' შევამოწმოთ მონაცემთა სერვისი
            If dataService Is Nothing Then
                Debug.WriteLine("LoadSessions: მონაცემთა სერვისი არ არის ინიციალიზებული")
                Return
            End If

            ' ჩავტვირთოთ სესიები
            allSessions = dataService.GetAllSessions()
            Debug.WriteLine($"LoadSessions: ჩატვირთულია {allSessions.Count} სესია")

        Catch ex As Exception
            Debug.WriteLine($"LoadSessions: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის უჯრედზე დაჭერის დამუშავება
    ''' </summary>
    Private Sub SessionPanel_Click(sender As Object, e As EventArgs)
        Try
            Dim panel = DirectCast(sender, Panel)
            Dim sessionId As Integer = CInt(panel.Tag)

            ' სესიის რედაქტირების ფორმის გახსნა...
            Debug.WriteLine($"SessionPanel_Click: გამოძახება სესიისთვის ID={sessionId}")

            ' შევამოწმოთ გვაქვს თუ არა dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' შევამოწმოთ უკვე გახსნილია თუ არა NewRecordForm
            For Each frm As Form In Application.OpenForms
                If TypeOf frm Is NewRecordForm Then
                    ' თუ უკვე გახსნილია, მოვიტანოთ წინ და გამოვიდეთ მეთოდიდან
                    frm.Focus()
                    Return
                End If
            Next

            ' შევქმნათ რედაქტირების ფორმა
            Dim editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_Calendar_Advanced")

            ' გავხსნათ ფორმა
            If editForm.ShowDialog() = DialogResult.OK Then
                ' თუ რედაქტირება წარმატებით დასრულდა, განვაახლოთ მონაცემები
                LoadSessions()
                UpdateCalendarView()
            End If
        Catch ex As Exception
            Debug.WriteLine($"SessionPanel_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის დროის ინდექსის პოვნა - გამოსწორებული ვერსია
    ''' </summary>
    Private Function GetTimeIndexForSession(session As SessionModel) As Integer
        ' გავიტანოთ დებაგ ინფორმაცია
        Debug.WriteLine($"GetTimeIndexForSession: ვეძებთ ინდექსს დროისთვის {session.DateTime:HH:mm}")

        ' დროის ინტერვალების რაოდენობა
        If timeIntervals.Count = 0 Then
            Debug.WriteLine("GetTimeIndexForSession: timeIntervals ცარიელია")
            Return -1
        End If

        ' ნაგულისხმევი მნიშვნელობა
        Dim result As Integer = -1

        ' ვამოწმებთ სესიის თარიღს
        Dim sessionTime As DateTime = session.DateTime

        ' ვამოწმებთ ყველა ინტერვალს
        For i As Integer = 0 To timeIntervals.Count - 1
            ' თუ ზუსტად ემთხვევა
            If sessionTime.Hour = timeIntervals(i).Hour AndAlso
           sessionTime.Minute = timeIntervals(i).Minute Then
                result = i
                Debug.WriteLine($"GetTimeIndexForSession: ზუსტი დამთხვევა ინდექსზე {i}")
                Exit For
            End If

            ' თუ სესიის დრო ორ ინტერვალს შორისაა, ვირჩევთ უახლოეს წინა ინტერვალს
            If i < timeIntervals.Count - 1 AndAlso
           sessionTime >= timeIntervals(i) AndAlso
           sessionTime < timeIntervals(i + 1) Then
                result = i
                Debug.WriteLine($"GetTimeIndexForSession: ინტერვალებს შორის დრო, ვიღებთ ინდექსს {i}")
                Exit For
            End If
        Next

        ' თუ ვერ ვიპოვეთ, მოდით შევამოწმოთ ახლოსაა თუ არა პირველ ან ბოლო ინტერვალთან
        If result = -1 Then
            If sessionTime < timeIntervals(0) Then
                ' თუ სესიის დრო პირველ ინტერვალამდეა, ვიღებთ პირველს
                result = 0
                Debug.WriteLine("GetTimeIndexForSession: დრო პირველ ინტერვალამდეა, ვიღებთ ინდექსს 0")
            ElseIf sessionTime >= timeIntervals(timeIntervals.Count - 1) Then
                ' თუ სესიის დრო ბოლო ინტერვალის შემდეგაა, ვიღებთ ბოლოს
                result = timeIntervals.Count - 1
                Debug.WriteLine($"GetTimeIndexForSession: დრო ბოლო ინტერვალის შემდეგაა, ვიღებთ ინდექსს {result}")
            End If
        End If

        Debug.WriteLine($"GetTimeIndexForSession: საბოლოო ინდექსი არის {result}")
        Return result
    End Function
    ''' <summary>
    ''' სესიის ხანგრძლივობის გათვლა უჯრედებში
    ''' </summary>
    Private Function CalculateSessionHeight(session As SessionModel) As Integer
        ' სესიის ხანგრძლივობა წუთებში
        Dim durationMinutes As Integer = session.Duration

        ' მინიმუმ 1 უჯრედი
        If durationMinutes <= 30 Then
            Return 1
        End If

        ' ვთვლით, რამდენი ნახევარსაათიანი ინტერვალი დასჭირდება
        Return CInt(Math.Ceiling(durationMinutes / 30.0))
    End Function

    ''' <summary>
    ''' დროებითი მეთოდი - ჯერ არ არის იმპლემენტირებული
    ''' </summary>
    Private Sub ShowDayViewByTherapist()
        ' თერაპევტების ხედით ჩვენება
        Dim lblNotImplemented As New Label()
        lblNotImplemented.Text = "თერაპევტების ხედი ჯერ არ არის იმპლემენტირებული"
        lblNotImplemented.AutoSize = True
        lblNotImplemented.Location = New Point(20, 20)
        pnlCalendarGrid.Controls.Clear()
        pnlCalendarGrid.Controls.Add(lblNotImplemented)
    End Sub

    ''' <summary>
    ''' დროებითი მეთოდი - ჯერ არ არის იმპლემენტირებული
    ''' </summary>
    Private Sub ShowDayViewByBeneficiary()
        ' ბენეფიციარების ხედით ჩვენება
        Dim lblNotImplemented As New Label()
        lblNotImplemented.Text = "ბენეფიციარების ხედი ჯერ არ არის იმპლემენტირებული"
        lblNotImplemented.AutoSize = True
        lblNotImplemented.Location = New Point(20, 20)
        pnlCalendarGrid.Controls.Clear()
        pnlCalendarGrid.Controls.Add(lblNotImplemented)
    End Sub

    ''' <summary>
    ''' დროებითი მეთოდი - ჯერ არ არის იმპლემენტირებული
    ''' </summary>
    Private Sub ShowWeekView()
        ' კვირის ხედი
        Dim lblNotImplemented As New Label()
        lblNotImplemented.Text = "კვირის ხედი ჯერ არ არის იმპლემენტირებული"
        lblNotImplemented.AutoSize = True
        lblNotImplemented.Location = New Point(20, 20)
        pnlCalendarGrid.Controls.Clear()
        pnlCalendarGrid.Controls.Add(lblNotImplemented)
    End Sub

    ''' <summary>
    ''' დროებითი მეთოდი - ჯერ არ არის იმპლემენტირებული
    ''' </summary>
    Private Sub ShowMonthView()
        ' თვის ხედი
        Dim lblNotImplemented As New Label()
        lblNotImplemented.Text = "თვის ხედი ჯერ არ არის იმპლემენტირებული"
        lblNotImplemented.AutoSize = True
        lblNotImplemented.Location = New Point(20, 20)
        pnlCalendarGrid.Controls.Clear()
        pnlCalendarGrid.Controls.Add(lblNotImplemented)
    End Sub

    ' ივენთ ჰენდლერები რომლებიც უნდა დაუკავშიროთ კონტროლებს Form-ის ინიციალიზაციაში

    ''' <summary>
    ''' კალენდრის თარიღის ცვლილების დამუშავება
    ''' </summary>
    Private Sub DTPCalendar_ValueChanged(sender As Object, e As EventArgs)
        ' განვაახლოთ თარიღის ლეიბლი
        'lblDate.Text = DTPCalendar.Value.ToString("d MMMM yyyy, dddd", New Globalization.CultureInfo("ka-GE"))

        ' განვაახლოთ კალენდრის ხედი
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' კალენდრის ხედის ცვლილება (დღე/კვირა/თვე)
    ''' </summary>
    Private Sub ViewType_CheckedChanged(sender As Object, e As EventArgs)
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' ფილტრის ცვლილება (სივრცე/თერაპევტი/ბენეფიციარი)
    ''' </summary>
    Private Sub FilterType_CheckedChanged(sender As Object, e As EventArgs)
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' საწყისი/საბოლოო დროის ცვლილება
    ''' </summary>
    Private Sub TimeRange_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' ვამოწმებთ ორივე დრო არის თუ არა არჩეული და ვალიდურია თუ არა
        If cbStart.SelectedIndex < 0 OrElse cbFinish.SelectedIndex < 0 Then
            Return
        End If

        ' განვაახლოთ კალენდარის ხედი
        UpdateCalendarView()
    End Sub
    ''' <summary>
    ''' UserControl-ის Resize ივენთის დამმუშავებელი - ზომების ცვლილებისას
    ''' კალენდრის მთელს ეკრანზე გასაშლელად
    ''' </summary>
    Private Sub UC_Calendar_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Try
            ' თუ კონტროლი არ არის ინიციალიზებული, გამოვიდეთ
            If Not Me.IsHandleCreated Then
                Return
            End If

            ' განვაახლოთ კალენდრის გრიდის ზომები კონტეინერის მიხედვით
            pnlCalendarGrid.Width = Me.Width - pnlCalendarGrid.Left - 20
            pnlCalendarGrid.Height = Me.Height - pnlCalendarGrid.Top - 20

            ' ფილტრის პანელის ზომის განახლება
            pnlFIlter.Width = Me.Width - pnlFIlter.Left - 20

            ' თუ კალენდრის ხედი უკვე გაშლილია, განვაახლოთ შემცველობა
            UpdateCalendarView()

            Debug.WriteLine($"UC_Calendar_Resize: კალენდრის ზომები განახლდა, ახალი ზომები: {Me.Width}x{Me.Height}")
        Catch ex As Exception
            Debug.WriteLine($"UC_Calendar_Resize: შეცდომა - {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' ივენთი, რომელიც გაეშვება UserControl-ის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_Calendar_Advanced_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            ' საწყისი პარამეტრების დაყენება (თუ ისინი არ არის დაყენებული)
            InitializeTimeIntervals()

            ' კალენდრის განახლება
            UpdateCalendarView()

            Debug.WriteLine("UC_Calendar_Advanced_Load: კალენდრის კონტროლი წარმატებით ჩაიტვირთა")
        Catch ex As Exception
            Debug.WriteLine($"UC_Calendar_Advanced_Load: შეცდომა - {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' კონსტრუქტორი კალენდრის ViewModel-ით
    ''' </summary>
    ''' <param name="calendarVm">კალენდრის ViewModel</param>
    Public Sub New(calendarVm As CalendarViewModel)
        ' UI ელემენტების ინიციალიზაცია
        InitializeComponent()

        Me.Dock = DockStyle.Fill
        ' ViewModel-ის მინიჭება
        If calendarVm IsNot Nothing Then
            viewModel = calendarVm
        Else
            viewModel = New CalendarViewModel()
        End If

        ' კომბობოქსების შევსება დროებით
        FillTimeComboBoxes()

        ' ივენთების მიბმა კონტროლებთან
        AddHandler DTPCalendar.ValueChanged, AddressOf DTPCalendar_ValueChanged
        AddHandler rbDay.CheckedChanged, AddressOf ViewType_CheckedChanged
        AddHandler rbWeek.CheckedChanged, AddressOf ViewType_CheckedChanged
        AddHandler rbMonth.CheckedChanged, AddressOf ViewType_CheckedChanged
        AddHandler RBSpace.CheckedChanged, AddressOf FilterType_CheckedChanged
        AddHandler RBPer.CheckedChanged, AddressOf FilterType_CheckedChanged
        AddHandler RBBene.CheckedChanged, AddressOf FilterType_CheckedChanged
        AddHandler cbStart.SelectedIndexChanged, AddressOf TimeRange_SelectedIndexChanged
        AddHandler cbFinish.SelectedIndexChanged, AddressOf TimeRange_SelectedIndexChanged

        ' საწყისი მნიშვნელობების დაყენება
        'lblDate.Text = DTPCalendar.Value.ToString("d MMMM yyyy, dddd", New Globalization.CultureInfo("ka-GE"))

        ' რადიობუტონების საწყისი მნიშვნელობები
        rbDay.Checked = True
        RBSpace.Checked = True
    End Sub

    ''' <summary>
    ''' დროის კომბობოქსების შევსება
    ''' </summary>
    Private Sub FillTimeComboBoxes()
        ' გავასუფთავოთ კომბობოქსები
        cbStart.Items.Clear()
        cbFinish.Items.Clear()

        ' დავამატოთ დროის ინტერვალები 30 წუთიანი ბიჯით 08:00-დან 21:00-მდე
        For hour As Integer = 8 To 21
            For minute As Integer = 0 To 30 Step 30
                Dim timeString As String = $"{hour:00}:{minute:00}"
                cbStart.Items.Add(timeString)
                cbFinish.Items.Add(timeString)
            Next
        Next

        ' საწყისი მნიშვნელობების დაყენება
        cbStart.SelectedItem = "09:00"
        cbFinish.SelectedItem = "20:00"
    End Sub
    ''' <summary>
    ''' მასშტაბის ლეიბლების განახლება
    ''' </summary>
    ''' <remarks>ამ მეთოდში ვქმნით ან განვაახლებთ ლეიბლებს, რომლებიც აჩვენებენ კალენდრის ჰორიზონტალურ და ვერტიკალურ მასშტაბებს</remarks>
    Private Sub UpdateScaleLabels()
        Try
            ' შევქმნათ ან განვაახლოთ ლეიბლი ჰორიზონტალური მასშტაბისთვის
            Dim lblHScale As Label = DirectCast(Controls.Find("lblHScale", True).FirstOrDefault(), Label)
            If lblHScale Is Nothing Then
                lblHScale = New Label()
                lblHScale.Name = "lblHScale"
                lblHScale.AutoSize = True
                lblHScale.Location = New Point(BtnHUp.Left + BtnHUp.Width + 5, BtnHUp.Top + 3)
                lblHScale.Font = New Font("Segoe UI", 8)
                Me.Controls.Add(lblHScale)
            End If
            lblHScale.Text = $"×{hScale:F2}"

            ' შევქმნათ ან განვაახლოთ ლეიბლი ვერტიკალური მასშტაბისთვის
            Dim lblVScale As Label = DirectCast(Controls.Find("lblVScale", True).FirstOrDefault(), Label)
            If lblVScale Is Nothing Then
                lblVScale = New Label()
                lblVScale.Name = "lblVScale"
                lblVScale.AutoSize = True
                lblVScale.Location = New Point(BtnVUp.Left + BtnVUp.Width + 5, BtnVUp.Top + 3)
                lblVScale.Font = New Font("Segoe UI", 8)
                Me.Controls.Add(lblVScale)
            End If
            lblVScale.Text = $"×{vScale:F2}"
        Catch ex As Exception
            Debug.WriteLine($"UpdateScaleLabels: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ვერტიკალური მასშტაბის გაზრდა - მწკრივების სიმაღლის გაზრდა
    ''' </summary>
    Private Sub BtnVUp_Click(sender As Object, e As EventArgs) Handles BtnVUp.Click
        ' მასშტაბის გაზრდა დადგენილი ნაბიჯით
        vScale += SCALE_STEP

        ' მაქსიმალური მასშტაბის შეზღუდვა
        If vScale > MAX_SCALE Then vScale = MAX_SCALE

        Debug.WriteLine($"BtnVUp_Click: ვერტიკალური მასშტაბი გაიზარდა: {vScale:F2}")

        ' კალენდრის განახლება ახალი მასშტაბით
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' ვერტიკალური მასშტაბის შემცირება - მწკრივების სიმაღლის შემცირება
    ''' </summary>
    Private Sub BtnVDown_Click(sender As Object, e As EventArgs) Handles BtnVDown.Click
        ' მასშტაბის შემცირება დადგენილი ნაბიჯით
        vScale -= SCALE_STEP

        ' მინიმალური მასშტაბის შეზღუდვა
        If vScale < MIN_SCALE Then vScale = MIN_SCALE

        Debug.WriteLine($"BtnVDown_Click: ვერტიკალური მასშტაბი შემცირდა: {vScale:F2}")

        ' კალენდრის განახლება ახალი მასშტაბით
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' ჰორიზონტალური მასშტაბის გაზრდა - სივრცეების სვეტების სიგანის გაზრდა
    ''' </summary>
    Private Sub BtnHUp_Click(sender As Object, e As EventArgs) Handles BtnHUp.Click
        ' მასშტაბის გაზრდა დადგენილი ნაბიჯით
        hScale += SCALE_STEP

        ' მაქსიმალური მასშტაბის შეზღუდვა
        If hScale > MAX_SCALE Then hScale = MAX_SCALE

        Debug.WriteLine($"BtnHUp_Click: ჰორიზონტალური მასშტაბი გაიზარდა: {hScale:F2}")

        ' კალენდრის განახლება ახალი მასშტაბით
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' ჰორიზონტალური მასშტაბის შემცირება - სივრცეების სვეტების სიგანის შემცირება
    ''' </summary>
    Private Sub BtnHDown_Click(sender As Object, e As EventArgs) Handles BtnHDown.Click
        ' მასშტაბის შემცირება დადგენილი ნაბიჯით
        hScale -= SCALE_STEP

        ' მინიმალური მასშტაბის შეზღუდვა
        If hScale < MIN_SCALE Then hScale = MIN_SCALE

        Debug.WriteLine($"BtnHDown_Click: ჰორიზონტალური მასშტაბი შემცირდა: {hScale:F2}")

        ' კალენდრის განახლება ახალი მასშტაბით
        UpdateCalendarView()
    End Sub

End Class