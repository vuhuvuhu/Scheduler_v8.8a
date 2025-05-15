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

    ''' <summary>
    ''' კონსტრუქტორი: შექმნის კალენდრის კონტროლს
    ''' </summary>
    ''' <param name="calendarViewModel">CalendarViewModel ობიექტი</param>
    Public Sub New(calendarViewModel As CalendarViewModel)
        ' კონტროლების ინიციალიზაცია
        InitializeComponent()

        ' ViewModel-ის შენახვა
        viewModel = calendarViewModel

        ' კალენდრის UI კონტროლების შექმნა
        CreateCalendarControls()

        ' ViewModel-ის PropertyChanged ივენთის მიბმა
        AddHandler viewModel.PropertyChanged, AddressOf ViewModel_PropertyChanged

        ' საწყისი მონაცემების განახლება
        UpdateCalendarUI()
    End Sub

    ''' <summary>
    ''' კალენდრის UI კონტროლების შექმნა
    ''' </summary>
    Private Sub CreateCalendarControls()
        ' კალენდრის ნავიგაციის პანელი
        Dim pnlNavigation As New Panel()
        pnlNavigation.Dock = DockStyle.Top
        pnlNavigation.Height = 40
        pnlNavigation.BackColor = Color.FromArgb(200, 200, 255)
        Controls.Add(pnlNavigation)

        ' წინა თვის ღილაკი
        btnPrevMonth = New Button()
        btnPrevMonth.Text = "◀"
        btnPrevMonth.Width = 40
        btnPrevMonth.Height = 30
        btnPrevMonth.Location = New Point(10, 5)
        btnPrevMonth.BackColor = Color.White
        AddHandler btnPrevMonth.Click, AddressOf BtnPrevMonth_Click
        pnlNavigation.Controls.Add(btnPrevMonth)

        ' თვისა და წლის ლეიბლი
        lblMonthYear = New Label()
        lblMonthYear.AutoSize = True
        lblMonthYear.Font = New Font("Sylfaen", 14, FontStyle.Bold)
        lblMonthYear.Location = New Point(100, 8)
        pnlNavigation.Controls.Add(lblMonthYear)

        ' შემდეგი თვის ღილაკი
        btnNextMonth = New Button()
        btnNextMonth.Text = "▶"
        btnNextMonth.Width = 40
        btnNextMonth.Height = 30
        btnNextMonth.Location = New Point(300, 5)
        btnNextMonth.BackColor = Color.White
        AddHandler btnNextMonth.Click, AddressOf BtnNextMonth_Click
        pnlNavigation.Controls.Add(btnNextMonth)

        ' დღევანდელ დღეზე დაბრუნების ღილაკი
        btnToday = New Button()
        btnToday.Text = "დღეს"
        btnToday.Width = 80
        btnToday.Height = 30
        btnToday.Location = New Point(350, 5)
        btnToday.BackColor = Color.White
        AddHandler btnToday.Click, AddressOf BtnToday_Click
        pnlNavigation.Controls.Add(btnToday)

        ' კალენდრის პანელი
        pnlCalendar = New Panel()
        pnlCalendar.Dock = DockStyle.Left
        pnlCalendar.Width = 500
        pnlCalendar.BackColor = Color.White
        Controls.Add(pnlCalendar)

        ' კვირის დღეების სათაურები
        Dim dayNames() As String = {"ორშ", "სამ", "ოთხ", "ხუთ", "პარ", "შაბ", "კვი"}
        For i As Integer = 0 To 6
            Dim lblDay As New Label()
            lblDay.Text = dayNames(i)
            lblDay.Width = 70
            lblDay.Height = 30
            lblDay.Location = New Point(i * 70 + 10, 10)
            lblDay.TextAlign = ContentAlignment.MiddleCenter
            lblDay.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            lblDay.BackColor = Color.FromArgb(230, 230, 250)
            pnlCalendar.Controls.Add(lblDay)
        Next

        ' კალენდრის დღეების ღილაკების შექმნა
        For i As Integer = 0 To 41
            calendarButtons(i) = New Button()
            calendarButtons(i).Width = 70
            calendarButtons(i).Height = 70
            calendarButtons(i).BackColor = Color.White
            calendarButtons(i).FlatStyle = FlatStyle.Flat
            calendarButtons(i).Tag = i ' შევინახოთ ინდექსი Tag-ში

            ' ღილაკზე დაჭერის ივენთი
            AddHandler calendarButtons(i).Click, AddressOf CalendarDay_Click

            ' ღილაკის პოზიციის დაყენება (6x7 ბადე)
            Dim row As Integer = i \ 7
            Dim col As Integer = i Mod 7
            calendarButtons(i).Location = New Point(col * 70 + 10, row * 70 + 40)

            pnlCalendar.Controls.Add(calendarButtons(i))
        Next

        ' დღის დეტალების პანელი
        pnlDayDetails = New Panel()
        pnlDayDetails.Dock = DockStyle.Fill
        pnlDayDetails.BackColor = Color.FromArgb(245, 245, 250)
        Controls.Add(pnlDayDetails)

        ' შერჩეული თარიღის ლეიბლი
        lblSelectedDate = New Label()
        lblSelectedDate.AutoSize = True
        lblSelectedDate.Font = New Font("Sylfaen", 14, FontStyle.Bold)
        lblSelectedDate.Location = New Point(10, 10)
        pnlDayDetails.Controls.Add(lblSelectedDate)

        ' სესიის დამატების ღილაკი
        btnAddSession = New Button()
        btnAddSession.Text = "სესიის დამატება"
        btnAddSession.Width = 150
        btnAddSession.Height = 30
        btnAddSession.Location = New Point(300, 10)
        btnAddSession.BackColor = Color.FromArgb(200, 255, 200)
        AddHandler btnAddSession.Click, AddressOf BtnAddSession_Click
        pnlDayDetails.Controls.Add(btnAddSession)

        ' დღის სესიების პანელი
        pnlDaySessions = New Panel()
        pnlDaySessions.Location = New Point(10, 50)
        pnlDaySessions.Width = pnlDayDetails.Width - 20
        pnlDaySessions.Height = pnlDayDetails.Height - 60
        pnlDaySessions.AutoScroll = True
        pnlDayDetails.Controls.Add(pnlDaySessions)

        ' დავაკავშიროთ პანელის Resize ივენთი
        AddHandler pnlDayDetails.Resize, AddressOf PnlDayDetails_Resize
    End Sub

    ''' <summary>
    ''' დღის დეტალების პანელის ზომის ცვლილებაზე რეაგირება
    ''' </summary>
    Private Sub PnlDayDetails_Resize(sender As Object, e As EventArgs)
        ' დავარეგულიროთ დღის სესიების პანელის ზომა
        pnlDaySessions.Width = pnlDayDetails.Width - 20
        pnlDaySessions.Height = pnlDayDetails.Height - 60
    End Sub

    ''' <summary>
    ''' ViewModel-ის PropertyChanged ივენთის დამმუშავებელი
    ''' </summary>
    Private Sub ViewModel_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
        ' UI ნაკადის შემოწმება
        If Me.InvokeRequired Then
            Me.Invoke(Sub() ViewModel_PropertyChanged(sender, e))
            Return
        End If

        ' შესაბამისი UI კომპონენტების განახლება
        Select Case e.PropertyName
            Case NameOf(viewModel.SelectedMonth)
                UpdateCalendarUI()
            Case NameOf(viewModel.SelectedDate)
                UpdateSelectedDayUI()
            Case NameOf(viewModel.SelectedDaySessions)
                UpdateDaySessionsUI()
        End Select
    End Sub

    ''' <summary>
    ''' კალენდრის UI-ის განახლება
    ''' </summary>
    Private Sub UpdateCalendarUI()
        ' თვისა და წლის განახლება
        lblMonthYear.Text = $"{viewModel.MonthName} {viewModel.Year}"

        ' გამოვთვალოთ პირველი დღის პოზიცია
        Dim firstDayOfMonth As DateTime = viewModel.SelectedMonth
        Dim firstDayOfWeek As Integer = CInt(firstDayOfMonth.DayOfWeek)

        ' შევცვალოთ 0 (კვირა) 7-ზე ქართული კალენდრისთვის
        If firstDayOfWeek = 0 Then firstDayOfWeek = 7

        ' პოზიციის კორექტირება (ორშაბათი პირველი დღე)
        Dim startPos As Integer = firstDayOfWeek - 1

        ' დღეების რაოდენობა თვეში
        Dim daysInMonth As Integer = DateTime.DaysInMonth(viewModel.SelectedMonth.Year, viewModel.SelectedMonth.Month)

        ' წინა თვის თარიღები
        Dim prevMonth As DateTime = firstDayOfMonth.AddMonths(-1)
        Dim daysInPrevMonth As Integer = DateTime.DaysInMonth(prevMonth.Year, prevMonth.Month)

        ' გავასუფთავოთ ყველა ღილაკი
        For i As Integer = 0 To 41
            calendarButtons(i).Text = ""
            calendarButtons(i).ForeColor = Color.Black
            calendarButtons(i).BackColor = Color.White
            calendarButtons(i).Tag = Nothing
        Next

        ' წინა თვის დღეები
        For i As Integer = 0 To startPos - 1
            Dim day As Integer = daysInPrevMonth - startPos + i + 1
            calendarButtons(i).Text = day.ToString()
            calendarButtons(i).ForeColor = Color.Gray
            calendarButtons(i).Tag = New DateTime(prevMonth.Year, prevMonth.Month, day)
        Next

        ' მიმდინარე თვის დღეები
        For i As Integer = 1 To daysInMonth
            Dim pos As Integer = startPos + i - 1
            calendarButtons(pos).Text = i.ToString()
            calendarButtons(pos).Tag = New DateTime(viewModel.SelectedMonth.Year, viewModel.SelectedMonth.Month, i)

            ' დღევანდელი დღის გამოყოფა
            If viewModel.SelectedMonth.Year = DateTime.Today.Year AndAlso
               viewModel.SelectedMonth.Month = DateTime.Today.Month AndAlso
               i = DateTime.Today.Day Then
                calendarButtons(pos).BackColor = Color.LightBlue
            End If

            ' შერჩეული დღის გამოყოფა
            If viewModel.SelectedDate.Year = viewModel.SelectedMonth.Year AndAlso
               viewModel.SelectedDate.Month = viewModel.SelectedMonth.Month AndAlso
               i = viewModel.SelectedDate.Day Then
                calendarButtons(pos).BackColor = Color.FromArgb(200, 200, 255)
            End If
        Next

        ' შემდეგი თვის დღეები
        Dim nextMonth As DateTime = firstDayOfMonth.AddMonths(1)
        Dim nextMonthDay As Integer = 1
        For i As Integer = startPos + daysInMonth To 41
            calendarButtons(i).Text = nextMonthDay.ToString()
            calendarButtons(i).ForeColor = Color.Gray
            calendarButtons(i).Tag = New DateTime(nextMonth.Year, nextMonth.Month, nextMonthDay)
            nextMonthDay += 1
        Next

        ' სესიების ჩვენება კალენდარში
        ShowSessionsInCalendar()
    End Sub

    ''' <summary>
    ''' სესიების ჩვენება კალენდარში
    ''' </summary>
    Private Sub ShowSessionsInCalendar()
        ' სესიების რაოდენობა თითოეული დღისთვის
        Dim sessionsPerDay As New Dictionary(Of DateTime, Integer)

        ' სესიების სტატუსების საკონტროლოდ
        Dim hasCompletedSessions As New Dictionary(Of DateTime, Boolean)
        Dim hasCancelledSessions As New Dictionary(Of DateTime, Boolean)
        Dim hasOverdueSessions As New Dictionary(Of DateTime, Boolean)

        ' დავთვალოთ სესიები დღეების მიხედვით
        For Each session In viewModel.Sessions
            Dim sessionDate As DateTime = session.DateTime.Date

            ' დავთვალოთ სესიების რაოდენობა დღეში
            If sessionsPerDay.ContainsKey(sessionDate) Then
                sessionsPerDay(sessionDate) += 1
            Else
                sessionsPerDay(sessionDate) = 1
            End If

            ' შევამოწმოთ სესიის სტატუსი
            Dim status As String = session.Status.Trim().ToLower()
            If status = "შესრულებული" Then
                hasCompletedSessions(sessionDate) = True
            ElseIf status = "გაუქმებული" OrElse status = "გაუქმება" Then
                hasCancelledSessions(sessionDate) = True
            ElseIf session.IsOverdue Then
                hasOverdueSessions(sessionDate) = True
            End If
        Next

        ' განვაახლოთ კალენდრის ღილაკები
        For i As Integer = 0 To 41
            If calendarButtons(i).Tag IsNot Nothing AndAlso TypeOf calendarButtons(i).Tag Is DateTime Then
                Dim buttonDate As DateTime = CType(calendarButtons(i).Tag, DateTime)

                ' სესიების ინდიკატორები
                If sessionsPerDay.ContainsKey(buttonDate) Then
                    Dim sessionsCount As Integer = sessionsPerDay(buttonDate)

                    ' დავამატოთ სესიების რაოდენობა ტექსტში
                    calendarButtons(i).Text = $"{calendarButtons(i).Text} ({sessionsCount})"

                    ' ფერის ცვლილება სტატუსის მიხედვით
                    If hasOverdueSessions.ContainsKey(buttonDate) AndAlso hasOverdueSessions(buttonDate) Then
                        ' ვადაგადაცილებული სესიები - წითელი ფონი
                        calendarButtons(i).BackColor = Color.FromArgb(255, 200, 200)
                    ElseIf hasCancelledSessions.ContainsKey(buttonDate) AndAlso hasCancelledSessions(buttonDate) Then
                        ' გაუქმებული სესიები - ნაცრისფერი ფონი
                        calendarButtons(i).BackColor = Color.FromArgb(200, 200, 200)
                    ElseIf hasCompletedSessions.ContainsKey(buttonDate) AndAlso hasCompletedSessions(buttonDate) Then
                        ' შესრულებული სესიები - მწვანე ფონი
                        calendarButtons(i).BackColor = Color.FromArgb(200, 255, 200)
                    Else
                        ' დაგეგმილი სესიები - ყვითელი ფონი
                        calendarButtons(i).BackColor = Color.FromArgb(255, 255, 200)
                    End If
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' შერჩეული დღის UI-ის განახლება
    ''' </summary>
    Private Sub UpdateSelectedDayUI()
        ' შერჩეული თარიღის ლეიბლის განახლება
        lblSelectedDate.Text = viewModel.SelectedDate.ToString("d MMMM yyyy, dddd", New Globalization.CultureInfo("ka-GE"))

        ' განვაახლოთ კალენდრის ღილაკების არჩევა
        UpdateCalendarUI()

        ' განვაახლოთ დღის სესიების ჩვენება
        UpdateDaySessionsUI()
    End Sub

    ''' <summary>
    ''' დღის სესიების UI-ის განახლება
    ''' </summary>
    Private Sub UpdateDaySessionsUI()
        ' გავასუფთავოთ დღის სესიების პანელი
        pnlDaySessions.Controls.Clear()

        ' თუ არ არის სესიები, ვაჩვენოთ შეტყობინება
        If viewModel.SelectedDaySessions.Count = 0 Then
            Dim lblNoSessions As New Label()
            lblNoSessions.Text = "არ არის სესიები ამ დღეს"
            lblNoSessions.AutoSize = True
            lblNoSessions.Location = New Point(10, 10)
            lblNoSessions.Font = New Font("Sylfaen", 12, FontStyle.Regular)
            pnlDaySessions.Controls.Add(lblNoSessions)
            Return
        End If

        ' ვაჩვენოთ ყველა სესია
        Dim yPos As Integer = 10
        For Each session In viewModel.SelectedDaySessions
            ' სესიის პანელი
            Dim pnlSession As New Panel()
            pnlSession.Width = pnlDaySessions.Width - 30
            pnlSession.Height = 100
            pnlSession.Location = New Point(10, yPos)
            pnlSession.BorderStyle = BorderStyle.FixedSingle

            ' სესიის ფონის ფერი სტატუსის მიხედვით
            Dim status As String = session.Status.Trim().ToLower()
            If status = "შესრულებული" Then
                pnlSession.BackColor = Color.FromArgb(220, 255, 220) ' მწვანე
            ElseIf status = "გაუქმებული" OrElse status = "გაუქმება" Then
                pnlSession.BackColor = Color.FromArgb(230, 230, 230) ' ნაცრისფერი
            ElseIf session.IsOverdue Then
                pnlSession.BackColor = Color.FromArgb(255, 220, 220) ' წითელი
            Else
                pnlSession.BackColor = Color.FromArgb(255, 255, 220) ' ყვითელი
            End If

            ' სესიის დრო
            Dim lblTime As New Label()
            lblTime.Text = session.DateTime.ToString("HH:mm")
            lblTime.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            lblTime.Location = New Point(10, 10)
            lblTime.AutoSize = True
            pnlSession.Controls.Add(lblTime)

            ' სესიის ხანგრძლივობა
            Dim lblDuration As New Label()
            lblDuration.Text = $"{session.Duration} წთ"
            lblDuration.Font = New Font("Segoe UI", 10, FontStyle.Regular)
            lblDuration.Location = New Point(10, 35)
            lblDuration.AutoSize = True
            pnlSession.Controls.Add(lblDuration)

            ' ბენეფიციარი
            Dim lblBeneficiary As New Label()
            lblBeneficiary.Text = session.FullName
            lblBeneficiary.Font = New Font("Sylfaen", 12, FontStyle.Bold)
            lblBeneficiary.Location = New Point(100, 10)
            lblBeneficiary.AutoSize = True
            pnlSession.Controls.Add(lblBeneficiary)

            ' თერაპევტი
            Dim lblTherapist As New Label()
            lblTherapist.Text = $"თერაპევტი: {session.TherapistName}"
            lblTherapist.Font = New Font("Sylfaen", 10, FontStyle.Regular)
            lblTherapist.Location = New Point(100, 35)
            lblTherapist.AutoSize = True
            pnlSession.Controls.Add(lblTherapist)

            ' თერაპიის ტიპი
            Dim lblTherapyType As New Label()
            lblTherapyType.Text = $"თერაპია: {session.TherapyType}"
            lblTherapyType.Font = New Font("Sylfaen", 10, FontStyle.Regular)
            lblTherapyType.Location = New Point(100, 55)
            lblTherapyType.AutoSize = True
            pnlSession.Controls.Add(lblTherapyType)

            ' სტატუსი
            Dim lblStatus As New Label()
            lblStatus.Text = $"სტატუსი: {session.Status}"
            lblStatus.Font = New Font("Sylfaen", 10, FontStyle.Regular)
            lblStatus.Location = New Point(100, 75)
            lblStatus.AutoSize = True
            pnlSession.Controls.Add(lblStatus)

            ' რედაქტირების ღილაკი
            Dim btnEdit As New Button()
            btnEdit.Text = "✎"
            btnEdit.Font = New Font("Segoe UI Symbol", 10, FontStyle.Bold)
            btnEdit.Width = 30
            btnEdit.Height = 30
            btnEdit.Location = New Point(pnlSession.Width - 40, 10)
            btnEdit.FlatStyle = FlatStyle.Flat
            btnEdit.Tag = session.Id ' შევინახოთ სესიის ID
            AddHandler btnEdit.Click, AddressOf BtnEditSession_Click
            pnlSession.Controls.Add(btnEdit)

            ' დავამატოთ სესიის პანელი
            pnlDaySessions.Controls.Add(pnlSession)

            ' გადავიდეთ შემდეგ პოზიციაზე
            yPos += 110
        Next
    End Sub

    ''' <summary>
    ''' წინა თვის ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnPrevMonth_Click(sender As Object, e As EventArgs)
        viewModel.PreviousMonth()
    End Sub

    ''' <summary>
    ''' შემდეგი თვის ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnNextMonth_Click(sender As Object, e As EventArgs)
        viewModel.NextMonth()
    End Sub

    ''' <summary>
    ''' დღევანდელ დღეზე დაბრუნების ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnToday_Click(sender As Object, e As EventArgs)
        ' დავაყენოთ მიმდინარე თვე და დღევანდელი თარიღი
        viewModel.CurrentMonth()
        viewModel.SelectedDate = DateTime.Today
    End Sub

    ''' <summary>
    ''' კალენდრის დღეზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub CalendarDay_Click(sender As Object, e As EventArgs)
        Dim btn = TryCast(sender, Button)
        If btn IsNot Nothing AndAlso btn.Tag IsNot Nothing AndAlso TypeOf btn.Tag Is DateTime Then
            ' დავაყენოთ შერჩეული თარიღი
            viewModel.SelectedDate = CType(btn.Tag, DateTime)

            ' თუ შერჩეული თარიღი სხვა თვეშია, შევცვალოთ თვეც
            If viewModel.SelectedDate.Month <> viewModel.SelectedMonth.Month OrElse
               viewModel.SelectedDate.Year <> viewModel.SelectedMonth.Year Then
                viewModel.SelectedMonth = New DateTime(viewModel.SelectedDate.Year, viewModel.SelectedDate.Month, 1)
            End If
        End If
    End Sub

    ''' <summary>
    ''' სესიის დამატების ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnAddSession_Click(sender As Object, e As EventArgs)
        Try
            ' შევამოწმოთ გვაქვს თუ არა მონაცემთა სერვისი
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' მომხმარებლის ავტორიზაციის შემოწმება
            If String.IsNullOrEmpty(userEmail) Then
                MessageBox.Show("სესიის დამატებისთვის საჭიროა ავტორიზაცია", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' შევამოწმოთ უკვე გახსნილია თუ არა NewRecordForm
            For Each frm As Form In Application.OpenForms
                If TypeOf frm Is NewRecordForm Then
                    ' თუ უკვე გახსნილია, მოვიტანოთ წინ და გამოვიდეთ მეთოდიდან
                    Debug.WriteLine("BtnAddSession_Click: NewRecordForm უკვე გახსნილია, ფოკუსის გადატანა")
                    frm.Focus()
                    Return
                End If
            Next

            Debug.WriteLine("BtnAddSession_Click: იხსნება ახალი სესიის დამატების ფორმა")

            ' შევქმნათ ახალი სესიის ფორმა
            Dim newSessionForm As New NewRecordForm(dataService, "სესია", userEmail, "UC_Calendar")

            ' დავაყენოთ არჩეული თარიღი
            ' რეფლექსიით მოვძებნოთ DateTimePicker კონტროლი სახელით DTP1
            For Each ctrl As Control In newSessionForm.Controls
                If ctrl.Name = "DTP1" AndAlso TypeOf ctrl Is DateTimePicker Then
                    DirectCast(ctrl, DateTimePicker).Value = viewModel.SelectedDate
                    Debug.WriteLine($"BtnAddSession_Click: DTP1 თარიღი დაყენებულია: {viewModel.SelectedDate:dd.MM.yyyy}")
                    Exit For
                End If
            Next

            ' გავხსნათ ფორმა
            Debug.WriteLine("BtnAddSession_Click: ფორმა იხსნება ShowDialog-ით")
            Dim result = newSessionForm.ShowDialog()

            ' თუ დავამატეთ სესია, განვაახლოთ მონაცემები
            If result = DialogResult.OK Then
                Debug.WriteLine("BtnAddSession_Click: ფორმა დაიხურა DialogResult.OK-ით, ვაახლებთ მონაცემებს")
                LoadMonthSessions()
            Else
                Debug.WriteLine($"BtnAddSession_Click: ფორმა დაიხურა შედეგით: {result}")
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnAddSession_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnAddSession_Click: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"სესიის დამატების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სესიის რედაქტირების ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnEditSession_Click(sender As Object, e As EventArgs)
        Try
            Dim btn = TryCast(sender, Button)
            If btn Is Nothing OrElse btn.Tag Is Nothing Then Return

            Dim sessionId As Integer = Convert.ToInt32(btn.Tag)

            ' შევამოწმოთ გვაქვს თუ არა მონაცემთა სერვისი
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
            Dim editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_Calendar")

            ' გავხსნათ ფორმა
            If editForm.ShowDialog() = DialogResult.OK Then
                ' თუ რედაქტირება წარმატებით დასრულდა, განვაახლოთ მონაცემები
                LoadMonthSessions()
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnEditSession_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიის რედაქტირების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' არჩეული თვის სესიების ჩატვირთვა
    ''' </summary>
    Public Sub LoadMonthSessions()
        Try
            ' შევამოწმოთ გვაქვს თუ არა dataService
            If dataService Is Nothing Then
                Debug.WriteLine("LoadMonthSessions: dataService არ არის ინიციალიზებული")
                Return
            End If

            ' მოვიპოვოთ არჩეული თვის პირველი და ბოლო დღის თარიღები
            Dim firstDayOfMonth As New DateTime(viewModel.SelectedMonth.Year, viewModel.SelectedMonth.Month, 1)
            Dim lastDayOfMonth As New DateTime(viewModel.SelectedMonth.Year, viewModel.SelectedMonth.Month, DateTime.DaysInMonth(viewModel.SelectedMonth.Year, viewModel.SelectedMonth.Month))

            ' 1 კვირა გამოვაჩინოთ წინა და შემდეგი თვიდანაც (კალენდრის სრულად შესავსებად)
            Dim startDate As DateTime = firstDayOfMonth.AddDays(-7)
            Dim endDate As DateTime = lastDayOfMonth.AddDays(7)

            Debug.WriteLine($"LoadMonthSessions: პერიოდი {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}")

            ' წამოვიღოთ ყველა სესია მონაცემთა ბაზიდან
            Dim allSessions = dataService.GetAllSessions()
            Debug.WriteLine($"LoadMonthSessions: მიღებულია {allSessions.Count} სესია")

            ' გავფილტროთ არჩეული პერიოდისთვის
            Dim monthSessions As New List(Of SessionModel)
            For Each session In allSessions
                Dim sessionDate As DateTime = session.DateTime.Date
                If sessionDate >= startDate.Date AndAlso sessionDate <= endDate.Date Then
                    monthSessions.Add(session)
                End If
            Next

            Debug.WriteLine($"LoadMonthSessions: გაფილტრულია {monthSessions.Count} სესია")

            ' გადავცეთ გაფილტრული სესიები ViewModel-ს
            viewModel.LoadMonthSessions(monthSessions)

            ' განვაახლოთ UI
            UpdateCalendarUI()

        Catch ex As Exception
            Debug.WriteLine($"LoadMonthSessions: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონაცემთა სერვისის მითითება
    ''' </summary>
    Public Sub SetDataService(service As IDataService)
        dataService = service
        Debug.WriteLine("UC_Calendar.SetDataService: მითითებულია მონაცემთა სერვისი")

        ' ჩავტვირთოთ საწყისი მონაცემები
        LoadMonthSessions()
    End Sub

    ''' <summary>
    ''' მომხმარებლის ელფოსტის მითითება
    ''' </summary>
    Public Sub SetUserEmail(email As String)
        userEmail = email
        viewModel.UserEmail = email
        Debug.WriteLine($"UC_Calendar.SetUserEmail: მითითებულია მომხმარებლის ელფოსტა: {email}")
    End Sub

    ''' <summary>
    ''' ივენთი, რომელიც გაეშვება UserControl-ის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_Calendar_Load(sender As Object, e As EventArgs) Handles Me.Load
        Debug.WriteLine("UC_Calendar_Load: კალენდრის კონტროლი ჩაიტვირთა")

        ' განვაახლოთ UI
        UpdateCalendarUI()
    End Sub

    ''' <summary>
    ''' ივენთი, რომელიც გაეშვება UserControl-ის ზომის ცვლილებისას
    ''' </summary>
    Private Sub UC_Calendar_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ' დავარეგულიროთ ზოგიერთი კონტროლის ზომა მშობელი კონტროლის ზომაზე დამოკიდებულებით
        pnlCalendar.Height = Me.Height
        pnlDaySessions.Width = pnlDayDetails.Width - 20
        pnlDaySessions.Height = pnlDayDetails.Height - 60

        ' განვაახლოთ კალენდრის კონტროლები
        UpdateCalendarUI()
    End Sub
End Class