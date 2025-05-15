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

            ' განვაახლოთ დროის ინტერვალები
            InitializeTimeIntervals()

            ' შევამოწმოთ, რომელი ხედია არჩეული
            If rbDay.Checked Then
                ' დღის ხედი
                If RBSpace.Checked Then
                    ' სივრცეების მიხედვით
                    lblInfo.Text = "დღის ხედი, სივრცეებით"
                    ShowDayViewBySpace()
                ElseIf RBPer.Checked Then
                    ' თერაპევტების მიხედვით
                    lblInfo.Text = "დღის ხედი, თერაპევტებით"
                    ShowDayViewByTherapist()
                ElseIf RBBene.Checked Then
                    ' ბენეფიციარების მიხედვით
                    lblInfo.Text = "დღის ხედი, ბენეფიციარებით"
                    ShowDayViewByBeneficiary()
                End If
            ElseIf rbWeek.Checked Then
                ' კვირის ხედი
                lblInfo.Text = "კვირის ხედი - მუშავდება"
                ShowWeekView()
            ElseIf rbMonth.Checked Then
                ' თვის ხედი
                lblInfo.Text = "თვის ხედი - მუშავდება"
                ShowMonthView()
            End If
        Catch ex As Exception
            Debug.WriteLine($"UpdateCalendarView: შეცდომა - {ex.Message}")
            MessageBox.Show($"კალენდრის ხედის განახლების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სივრცეების ჩატვირთვა მონაცემთა წყაროდან
    ''' </summary>
    Private Sub LoadSpaces()
        Try
            ' გავასუფთავოთ სივრცეების სია
            spaces.Clear()

            ' შევამოწმოთ მონაცემთა სერვისი
            If dataService Is Nothing Then
                Debug.WriteLine("LoadSpaces: მონაცემთა სერვისი არ არის ინიციალიზებული")
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

            ' დავამატოთ სივრცეები სიაში
            spaces.AddRange(spaceSet.OrderBy(Function(s) s))

            ' თუ ვერ ვიპოვეთ სივრცეები სესიებიდან, ცხრილიდან წამოვიღოთ უშუალოდ
            If spaces.Count = 0 AndAlso dataService IsNot Nothing Then
                ' ცხრილიდან სივრცეების წამოღება
                Dim spacesData = dataService.GetData("DB-Space!B2:B")

                If spacesData IsNot Nothing Then
                    For Each row In spacesData
                        If row.Count > 0 AndAlso row(0) IsNot Nothing AndAlso Not String.IsNullOrEmpty(row(0).ToString().Trim()) Then
                            spaceSet.Add(row(0).ToString().Trim())
                        End If
                    Next

                    ' დავამატოთ სივრცეები სიაში
                    spaces.AddRange(spaceSet.OrderBy(Function(s) s))
                End If
            End If

            Debug.WriteLine($"LoadSpaces: ჩატვირთულია {spaces.Count} სივრცე")
        Catch ex As Exception
            Debug.WriteLine($"LoadSpaces: შეცდომა - {ex.Message}")
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
    ''' ჩვენება დღის ხედის სივრცეების მიხედვით
    ''' </summary>
    Private Sub ShowDayViewBySpace()
        Try
            ' გავასუფთავოთ კალენდრის პანელი
            pnlCalendarGrid.Controls.Clear()

            ' შევამოწმოთ არის თუ არა სივრცეები ჩატვირთული
            If spaces.Count = 0 Then
                ' თუ არა, ჩავტვირთოთ სივრცეები
                LoadSpaces()

                ' თუ მაინც ცარიელია, გამოვიდეთ
                If spaces.Count = 0 Then
                    Dim lblNoSpaces As New Label()
                    lblNoSpaces.Text = "სივრცეები არ არის ხელმისაწვდომი"
                    lblNoSpaces.AutoSize = True
                    lblNoSpaces.Location = New Point(20, 20)
                    pnlCalendarGrid.Controls.Add(lblNoSpaces)
                    Return
                End If
            End If

            ' სვეტების რაოდენობა: იმდენი უჯრედი, რამდენი სივრცეც გვაქვს
            Dim cellsPerRow As Integer = spaces.Count

            ' დროის ინტერვალების რაოდენობა
            Dim rowCount As Integer = timeIntervals.Count

            ' უჯრედის ზომები
            Const COLUMN_WIDTH As Integer = 150
            Const ROW_HEIGHT As Integer = 60
            Const TIME_COLUMN_WIDTH As Integer = 80

            ' ხილული სივრცე გრიდისთვის
            Dim contentWidth As Integer = TIME_COLUMN_WIDTH + (cellsPerRow * COLUMN_WIDTH)
            Dim contentHeight As Integer = (rowCount + 1) * ROW_HEIGHT ' +1 სათაურის მწკრივისთვის

            ' შევქმნათ კონტენტის პანელი
            calendarContentPanel = New Panel()
            calendarContentPanel.Size = New Size(contentWidth, contentHeight)
            calendarContentPanel.Location = New Point(0, 0)
            calendarContentPanel.BackColor = Color.White
            pnlCalendarGrid.Controls.Add(calendarContentPanel)

            ' შევქმნათ სივრცეების სათაურების მწკრივი
            ' დროის ცარიელი უჯრედი
            Dim emptyCell As New Label()
            emptyCell.Size = New Size(TIME_COLUMN_WIDTH, ROW_HEIGHT)
            emptyCell.Location = New Point(0, 0)
            emptyCell.BackColor = Color.FromArgb(230, 230, 250)
            emptyCell.BorderStyle = BorderStyle.FixedSingle
            calendarContentPanel.Controls.Add(emptyCell)

            For i As Integer = 0 To cellsPerRow - 1
                ' სივრცის სათაური
                Dim headerLabel As New Label()
                headerLabel.Text = spaces(i)
                headerLabel.TextAlign = ContentAlignment.MiddleCenter
                headerLabel.Size = New Size(COLUMN_WIDTH, ROW_HEIGHT)
                headerLabel.Location = New Point(TIME_COLUMN_WIDTH + (i * COLUMN_WIDTH), 0)
                headerLabel.BackColor = Color.FromArgb(230, 230, 250)
                headerLabel.BorderStyle = BorderStyle.FixedSingle
                calendarContentPanel.Controls.Add(headerLabel)
            Next

            ' შევქმნათ დროის ინტერვალების სვეტი
            For i As Integer = 0 To rowCount - 1
                ' დროის ლეიბლი
                Dim timeLabel As New Label()
                timeLabel.Text = timeIntervals(i).ToString("HH:mm")
                timeLabel.TextAlign = ContentAlignment.MiddleRight
                timeLabel.Size = New Size(TIME_COLUMN_WIDTH, ROW_HEIGHT)
                timeLabel.Location = New Point(0, ROW_HEIGHT + (i * ROW_HEIGHT))
                timeLabel.BackColor = Color.FromArgb(240, 240, 240)
                timeLabel.BorderStyle = BorderStyle.FixedSingle
                calendarContentPanel.Controls.Add(timeLabel)
            Next

            ' შევქმნათ უჯრედები გრიდისთვის
            ReDim gridCells(cellsPerRow - 1, rowCount - 1)

            For col As Integer = 0 To cellsPerRow - 1
                For row As Integer = 0 To rowCount - 1
                    ' ახალი უჯრედი
                    Dim cell As New Panel()
                    cell.Size = New Size(COLUMN_WIDTH, ROW_HEIGHT)
                    cell.Location = New Point(TIME_COLUMN_WIDTH + (col * COLUMN_WIDTH), ROW_HEIGHT + (row * ROW_HEIGHT))
                    cell.BackColor = Color.White
                    cell.BorderStyle = BorderStyle.FixedSingle

                    ' შევინახოთ უჯრედი მასივში
                    gridCells(col, row) = cell

                    ' დავამატოთ უჯრედი კონტენტის პანელზე
                    calendarContentPanel.Controls.Add(cell)
                Next
            Next

            ' თუ სესიები გვაქვს, განვათავსოთ გრიდზე
            If allSessions IsNot Nothing AndAlso allSessions.Count > 0 Then
                ' ფილტრაცია არჩეული დღის მიხედვით
                Dim selectedDate As DateTime = DTPCalendar.Value.Date
                Dim daySessions = allSessions.Where(Function(s) s.DateTime.Date = selectedDate).ToList()

                ' განვათავსოთ სესიები გრიდზე
                PlaceSessionsOnGrid(daySessions)
            End If

            Debug.WriteLine("ShowDayViewBySpace: დღის ხედი სივრცეების მიხედვით წარმატებით გამოჩნდა")
        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewBySpace: შეცდომა - {ex.Message}")
            MessageBox.Show($"დღის ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სესიების განთავსება გრიდზე
    ''' </summary>
    Private Sub PlaceSessionsOnGrid(sessions As List(Of SessionModel))
        Try
            ' შევამოწმოთ უჯრედები და სესიები
            If gridCells Is Nothing OrElse sessions Is Nothing OrElse sessions.Count = 0 Then
                Return
            End If

            ' განვათავსოთ თითოეული სესია
            For Each session In sessions
                ' ვიპოვოთ სივრცის ინდექსი
                Dim spaceIndex As Integer = spaces.IndexOf(session.Space)

                ' თუ სივრცე ნაპოვნია
                If spaceIndex >= 0 Then
                    ' ვიპოვოთ დროის ინდექსი
                    Dim timeIndex As Integer = GetTimeIndexForSession(session)

                    ' თუ დრო ჩვენი ინტერვალების ფარგლებშია
                    If timeIndex >= 0 AndAlso
                   spaceIndex < gridCells.GetLength(0) AndAlso
                   timeIndex < gridCells.GetLength(1) Then

                        ' სესიის ხანგრძლივობის გათვალისწინება (რამდენ უჯრედს დაიკავებს)
                        Dim durationInCells As Integer = CalculateSessionHeight(session)

                        ' შეზღუდვა მაქსიმალური რაოდენობით
                        Dim lastRowIndex As Integer = Math.Min(timeIndex + durationInCells - 1, gridCells.GetLength(1) - 1)

                        ' სესიის სრული სიმაღლე
                        Dim sessionHeight As Integer = (lastRowIndex - timeIndex + 1) * gridCells(spaceIndex, timeIndex).Height

                        ' შევქმნათ სესიის პანელი
                        Dim sessionPanel As New Panel()
                        sessionPanel.Size = New Size(gridCells(spaceIndex, timeIndex).Width - 2, sessionHeight)
                        sessionPanel.Location = New Point(1, 1)
                        sessionPanel.BackColor = GetSessionColor(session)

                        ' სესიის ინფორმაციის დამატება
                        AddSessionInfo(sessionPanel, session)

                        ' სესიის დამატება უჯრედზე
                        gridCells(spaceIndex, timeIndex).Controls.Add(sessionPanel)

                        ' სესიის ID-ის შენახვა Tag-ში დაჭერაზე რეაგირებისთვის
                        sessionPanel.Tag = session.Id
                        AddHandler sessionPanel.Click, AddressOf SessionPanel_Click
                    End If
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"PlaceSessionsOnGrid: შეცდომა - {ex.Message}")
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
    ''' სესიის დროის ინდექსის პოვნა
    ''' </summary>
    Private Function GetTimeIndexForSession(session As SessionModel) As Integer
        ' ნაგულისხმევი მნიშვნელობა
        Dim result As Integer = -1

        ' ვამოწმებთ სესიის თარიღს
        Dim sessionTime As DateTime = session.DateTime

        ' ვამოწმებთ ყველა ინტერვალს
        For i As Integer = 0 To timeIntervals.Count - 1
            ' ვნახავთ, რომელ ინტერვალს ემთხვევა სესიის დრო
            If sessionTime.Hour = timeIntervals(i).Hour AndAlso
           sessionTime.Minute = timeIntervals(i).Minute Then
                result = i
                Exit For
            End If

            ' თუ სესიის დრო ორ ინტერვალს შორისაა, ვირჩევთ უახლოეს
            If i < timeIntervals.Count - 1 AndAlso
           sessionTime >= timeIntervals(i) AndAlso
           sessionTime < timeIntervals(i + 1) Then
                result = i
                Exit For
            End If
        Next

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
    ''' სესიის ფერის განსაზღვრა სტატუსის მიხედვით
    ''' </summary>
    Private Function GetSessionColor(session As SessionModel) As Color
        ' სტატუსის მიხედვით ფერის შერჩევა
        Dim status As String = session.Status.Trim().ToLower()

        If status = "შესრულებული" Then
            Return Color.FromArgb(220, 255, 220) ' მწვანე
        ElseIf status = "გაუქმებული" OrElse status = "გაუქმება" Then
            Return Color.FromArgb(230, 230, 230) ' ნაცრისფერი
        ElseIf session.IsOverdue Then
            Return Color.FromArgb(255, 220, 220) ' წითელი
        Else
            Return Color.FromArgb(255, 255, 220) ' ყვითელი
        End If
    End Function

    ''' <summary>
    ''' სესიის ინფორმაციის დამატება პანელზე
    ''' </summary>
    Private Sub AddSessionInfo(panel As Panel, session As SessionModel)
        Try
            ' დრო
            Dim lblTime As New Label()
            lblTime.Text = session.DateTime.ToString("HH:mm")
            lblTime.AutoSize = True
            lblTime.Location = New Point(5, 3)
            lblTime.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            panel.Controls.Add(lblTime)

            ' ბენეფიციარი
            Dim lblBeneficiary As New Label()
            lblBeneficiary.Text = session.FullName
            lblBeneficiary.AutoSize = True
            lblBeneficiary.Location = New Point(5, 20)
            lblBeneficiary.Font = New Font("Sylfaen", 9, FontStyle.Bold)
            panel.Controls.Add(lblBeneficiary)

            ' თერაპევტი
            Dim lblTherapist As New Label()
            lblTherapist.Text = session.TherapistName
            lblTherapist.AutoSize = True
            lblTherapist.Location = New Point(5, 37)
            lblTherapist.Font = New Font("Sylfaen", 8)
            panel.Controls.Add(lblTherapist)

        Catch ex As Exception
            Debug.WriteLine($"AddSessionInfo: შეცდომა - {ex.Message}")
        End Try
    End Sub

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
        lblDate.Text = DTPCalendar.Value.ToString("d MMMM yyyy, dddd", New Globalization.CultureInfo("ka-GE"))

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
End Class