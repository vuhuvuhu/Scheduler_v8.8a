Imports System.ComponentModel
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock
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

    ' სივრცეები და დროის ინტერვალები
    Private spaces As List(Of String) = New List(Of String)()
    Private timeIntervals As List(Of DateTime) = New List(Of DateTime)()
    Private gridCells(,) As Panel ' უჯრედები სესიების განთავსებისთვის
    Private calendarContentPanel As Panel = Nothing

    ' მასშტაბირების კოეფიციენტები
    Private hScale As Double = 1.0 ' ჰორიზონტალური მასშტაბი
    Private vScale As Double = 1.0 ' ვერტიკალური მასშტაბი

    ' მასშტაბირების ნაბიჯები
    Private Const SCALE_STEP As Double = 0.1 ' მასშტაბის ცვლილების ნაბიჯი
    Private Const MIN_SCALE As Double = 0.5 ' მინიმალური მასშტაბი
    Private Const MAX_SCALE As Double = 4 ' მაქსიმალური მასშტაბი

    ' მასშტაბირებამდე ზომები
    Private Const BASE_HEADER_HEIGHT As Integer = 40 ' სათაურების საბაზისო სიმაღლე - არ იცვლება
    Private Const BASE_TIME_COLUMN_WIDTH As Integer = 60 ' დროის სვეტის საბაზისო სიგანე - არ იცვლება
    Private Const BASE_ROW_HEIGHT As Integer = 40 ' მწკრივის საბაზისო სიმაღლე - იცვლება vScale-ით
    Private Const BASE_SPACE_COLUMN_WIDTH As Integer = 150 ' სივრცის სვეტის საბაზისო სიგანე - იცვლება hScale-ით

    ' კლასის დონეზე კონსტანტები
    Private Const BASE_DATE_COLUMN_WIDTH As Integer = 40 ' თარიღის ზოლის საბაზისო სიგანე - არ იცვლება მასშტაბით

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

        ' საწყისი მასშტაბი
        hScale = 1.0
        vScale = 1.0

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

        ' რადიობუტონების საწყისი მნიშვნელობები
        rbDay.Checked = True
        RBSpace.Checked = True
    End Sub

    ''' <summary>
    ''' ცარიელი კონსტრუქტორი - DesignerMode-სთვის
    ''' </summary>
    Public Sub New()
        ' UI ინიციალიზაცია
        InitializeComponent()

        ' ViewModel-ის მინიჭება
        viewModel = New CalendarViewModel()

        ' საწყისი მასშტაბი
        hScale = 1.0
        vScale = 1.0
    End Sub

    ''' <summary>
    ''' მონაცემთა სერვისის მითითება
    ''' </summary>
    Public Sub SetDataService(service As IDataService)
        dataService = service
        Debug.WriteLine("UC_Calendar.SetDataService: მითითებულია მონაცემთა სერვისი")

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
        viewModel.UserEmail = email
        Debug.WriteLine($"UC_Calendar.SetUserEmail: მითითებულია მომხმარებლის ელფოსტა: {email}")
    End Sub

    ''' <summary>
    ''' დროის კომბობოქსების შევსება
    ''' </summary>
    Private Sub FillTimeComboBoxes()
        Try
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
            Dim startIndex As Integer = cbStart.FindStringExact("09:00")
            If startIndex >= 0 Then
                cbStart.SelectedIndex = startIndex
            Else
                cbStart.SelectedIndex = 0
            End If

            Dim finishIndex As Integer = cbFinish.FindStringExact("20:00")
            If finishIndex >= 0 Then
                cbFinish.SelectedIndex = finishIndex
            Else
                cbFinish.SelectedIndex = cbFinish.Items.Count - 1
            End If
        Catch ex As Exception
            Debug.WriteLine($"FillTimeComboBoxes: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მასშტაბის ჩვენება ინტერფეისში
    ''' </summary>
    Private Sub UpdateScaleLabels()
        Try
            ' შევქმნათ ან განვაახლოთ ლეიბლი ჰორიზონტალური მასშტაბისთვის
            Dim lblHScale As Label = DirectCast(Controls.Find("lblHScale", True).FirstOrDefault(), Label)
            If lblHScale Is Nothing Then
                ' თუ ლეიბლი არ არსებობს, შევქმნათ ახალი
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
                ' თუ ლეიბლი არ არსებობს, შევქმნათ ახალი
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
    ''' დროის ინტერვალების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeTimeIntervals()
        Try
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
        Catch ex As Exception
            Debug.WriteLine($"InitializeTimeIntervals: შეცდომა - {ex.Message}")
            timeIntervals.Clear()

            ' შეცდომის შემთხვევაში დავამატოთ რამდენიმე საათი
            Dim baseDate As DateTime = DateTime.Today
            For hour As Integer = 9 To 20
                timeIntervals.Add(baseDate.AddHours(hour))
                timeIntervals.Add(baseDate.AddHours(hour).AddMinutes(30))
            Next
        End Try
    End Sub

    ''' <summary>
    ''' სივრცეების ჩატვირთვა მონაცემთა წყაროდან
    ''' </summary>
    Private Sub LoadSpaces()
        Try
            ' გავასუფთავოთ სივრცეების სია
            spaces.Clear()

            ' განსაზღვრული წესრიგის სია
            Dim orderedSpaces As New List(Of String) From {
                "მწვანე აბა", "ლურჯი აბა", "სენსორი", "მეტყველება",
                "მუსიკა", "თერაპევტი", "არტი", "სხვა", "ფიზიკური", "საკონფერენციო",
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
                    spaceSet.Remove(orderedSpace) ' წავშალოთ სეტიდან
                Else
                    ' დავამატოთ მაინც სტანდარტული სივრცე
                    spaces.Add(orderedSpace)
                End If
            Next

            ' დავამატოთ ყველა დარჩენილი სივრცე
            spaces.AddRange(spaceSet.OrderBy(Function(s) s))

            ' თუ საერთოდ ვერ ვიპოვეთ სივრცეები, გამოვიყენოთ საწყისი სია
            If spaces.Count = 0 Then
                spaces.AddRange(orderedSpaces)
            End If

            Debug.WriteLine($"LoadSpaces: ჩატვირთულია {spaces.Count} სივრცე")
        Catch ex As Exception
            Debug.WriteLine($"LoadSpaces: შეცდომა - {ex.Message}")

            ' შეცდომის შემთხვევაში გამოვიყენოთ საწყისი სია
            spaces.Clear()
            spaces.AddRange(New List(Of String) From {
                "მწვანე აბა", "ლურჯი აბა", "სენსორი", "მეტყველება",
                "მუსიკა", "თერაპევტი", "არტი", "სხვა", "ფიზიკური", "საკონფერენციო",
                "მშობლები", "ონლაინ", "სახლი", "გარე"
            })
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
            ' შეცდომის შემთხვევაში შევქმნათ ცარიელი სია
            allSessions = New List(Of SessionModel)()
        End Try
    End Sub

    ''' <summary>
    ''' კალენდრის დღის ხედის პანელების ინიციალიზაცია 
    ''' 🔧 ძირითადი შესწორება: თარიღის სვეტი და სივრცეების სათაურები შენარჩუნდებიან მასშტაბირებისას
    ''' </summary>
    Private Sub InitializeDayViewPanels()
        Try
            Debug.WriteLine("InitializeDayViewPanels: დაიწყო პანელების ინიციალიზაცია")

            ' გავასუფთავოთ კალენდრის პანელი
            pnlCalendarGrid.Controls.Clear()

            ' ძალიან მნიშვნელოვანი: pnlCalendarGrid-ს არ ჰქონდეს AutoScroll!
            pnlCalendarGrid.AutoScroll = False

            ' თუ დროის ინტერვალები არ არის ინიციალიზებული, დავამატოთ
            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' წავკითხოთ მშობელი პანელის ზომები
            Dim totalWidth As Integer = pnlCalendarGrid.ClientSize.Width
            Dim totalHeight As Integer = pnlCalendarGrid.ClientSize.Height

            Debug.WriteLine($"InitializeDayViewPanels: კონტეინერის ზომები - Width={totalWidth}, Height={totalHeight}")

            ' კონსტანტები, რომლებიც არ იცვლება მასშტაბირებისას
            Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT
            Dim TIME_COLUMN_WIDTH As Integer = BASE_TIME_COLUMN_WIDTH
            Dim DATE_COLUMN_WIDTH As Integer = BASE_DATE_COLUMN_WIDTH ' ეს არ იმასშტაბირდება!

            ' გამოვთვალოთ დროისა და თარიღის პანელების ზომები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim totalRowsHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            Debug.WriteLine($"InitializeDayViewPanels: ROW_HEIGHT={ROW_HEIGHT}, totalRowsHeight={totalRowsHeight}, vScale={vScale}")

            ' ======= 1. შევქმნათ თარიღის ზოლის პანელი (ყველაზე მარცხნივ) =======
            Dim dateColumnPanel As New Panel()
            dateColumnPanel.Name = "dateColumnPanel"
            dateColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left ' 🔧 მნიშვნელოვანი დამატება!
            dateColumnPanel.AutoScroll = False
            dateColumnPanel.Size = New Size(DATE_COLUMN_WIDTH, totalRowsHeight)
            dateColumnPanel.Location = New Point(0, HEADER_HEIGHT)
            dateColumnPanel.BackColor = Color.FromArgb(60, 80, 150)
            dateColumnPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(dateColumnPanel)

            ' ======= 2. შევქმნათ საათების პანელი (თარიღის ზოლის მარჯვნივ) =======
            Dim timeColumnPanel As New Panel()
            timeColumnPanel.Name = "timeColumnPanel"
            timeColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left ' 🔧 მნიშვნელოვანი დამატება!
            timeColumnPanel.AutoScroll = False
            timeColumnPanel.Size = New Size(TIME_COLUMN_WIDTH, totalRowsHeight)
            timeColumnPanel.Location = New Point(DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            timeColumnPanel.BackColor = Color.FromArgb(240, 240, 245)
            timeColumnPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(timeColumnPanel)

            ' ======= 3. შევქმნათ თარიღის სათაურის პანელი (მარცხენა ზედა კუთხეში) =======
            Dim dateHeaderPanel As New Panel()
            dateHeaderPanel.Name = "dateHeaderPanel"
            dateHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left ' 🔧 მნიშვნელოვანი დამატება!
            dateHeaderPanel.Size = New Size(DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            dateHeaderPanel.Location = New Point(0, 0)
            dateHeaderPanel.BackColor = Color.FromArgb(40, 60, 120)
            dateHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(dateHeaderPanel)

            ' თარიღის სათაურის ლეიბლი
            Dim dateHeaderLabel As New Label()
            dateHeaderLabel.Size = New Size(DATE_COLUMN_WIDTH - 2, HEADER_HEIGHT - 2)
            dateHeaderLabel.Location = New Point(1, 1)
            dateHeaderLabel.TextAlign = ContentAlignment.MiddleCenter
            dateHeaderLabel.Text = "თარიღი"
            dateHeaderLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            dateHeaderLabel.ForeColor = Color.White
            dateHeaderPanel.Controls.Add(dateHeaderLabel)

            ' ======= 4. შევქმნათ დროის სათაურის პანელი (თარიღის სათაურის მარჯვნივ) =======
            Dim timeHeaderPanel As New Panel()
            timeHeaderPanel.Name = "timeHeaderPanel"
            timeHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left ' 🔧 მნიშვნელოვანი დამატება!
            timeHeaderPanel.Size = New Size(TIME_COLUMN_WIDTH, HEADER_HEIGHT)
            timeHeaderPanel.Location = New Point(DATE_COLUMN_WIDTH, 0)
            timeHeaderPanel.BackColor = Color.FromArgb(180, 180, 220)
            timeHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(timeHeaderPanel)

            ' დროის სათაურის ლეიბლი
            Dim timeHeaderLabel As New Label()
            timeHeaderLabel.Size = New Size(TIME_COLUMN_WIDTH - 2, HEADER_HEIGHT - 2)
            timeHeaderLabel.Location = New Point(1, 1)
            timeHeaderLabel.TextAlign = ContentAlignment.MiddleCenter
            timeHeaderLabel.Text = "დრო"
            timeHeaderLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            timeHeaderPanel.Controls.Add(timeHeaderLabel)

            ' ======= 5. შევქმნათ სივრცეების სათაურების პანელი (ზემოთ) =======
            Dim spacesHeaderPanel As New Panel()
            spacesHeaderPanel.Name = "spacesHeaderPanel"
            spacesHeaderPanel.AutoScroll = False
            ' 🔧 კრიტიკული შესწორება: ზომა დამოკიდებული იქნება მასშტაბზე!
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim totalSpacesWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count
            spacesHeaderPanel.Size = New Size(totalSpacesWidth, HEADER_HEIGHT)
            spacesHeaderPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, 0)
            spacesHeaderPanel.BackColor = Color.FromArgb(220, 220, 240)
            spacesHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(spacesHeaderPanel)

            ' ======= 6. შევქმნათ მთავარი გრიდის პანელი =======
            Dim mainGridPanel As New Panel()
            mainGridPanel.Name = "mainGridPanel"
            mainGridPanel.Size = New Size(totalWidth - TIME_COLUMN_WIDTH - DATE_COLUMN_WIDTH, totalHeight - HEADER_HEIGHT)
            mainGridPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            mainGridPanel.BackColor = Color.White
            mainGridPanel.BorderStyle = BorderStyle.FixedSingle

            ' მნიშვნელოვანი: მხოლოდ ეს პანელი არის სკროლებადი!
            mainGridPanel.AutoScroll = True
            pnlCalendarGrid.Controls.Add(mainGridPanel)

            Debug.WriteLine("InitializeDayViewPanels: პანელების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"InitializeDayViewPanels: შეცდომა - {ex.Message}")
            Debug.WriteLine($"InitializeDayViewPanels: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"პანელების ინიციალიზაციის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' თარიღის ზოლის პანელის შევსება არჩეული თარიღის ინფორმაციით
    ''' 🔧 გასწორება: RotatedLabel-ს ზომა და ფონტი ჩასწორდა მასშტაბირების შენარჩუნებისთვის
    ''' </summary>
    Private Sub FillDateColumnPanel()
        Try
            Debug.WriteLine("FillDateColumnPanel: დაიწყო თარიღის ზოლის შევსება")

            ' ვიპოვოთ თარიღის ზოლის პანელი
            Dim dateColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("dateColumnPanel", False).FirstOrDefault(), Panel)

            If dateColumnPanel Is Nothing Then
                Debug.WriteLine("FillDateColumnPanel: თარიღის ზოლის პანელი ვერ მოიძებნა")
                Return
            End If

            ' გავასუფთავოთ
            dateColumnPanel.Controls.Clear()

            ' არჩეული თარიღი კალენდრიდან
            Dim selectedDate As Date = DTPCalendar.Value.Date

            ' ქართული კულტურა თარიღის ფორმატირებისთვის
            Dim georgianCulture As New Globalization.CultureInfo("ka-GE")

            ' კვირის დღე, რიცხვი, თვე და წელი
            Dim weekDay As String = selectedDate.ToString("dddd", georgianCulture)
            Dim dayOfMonth As String = selectedDate.ToString("dd", georgianCulture)
            Dim month As String = selectedDate.ToString("MMMM", georgianCulture)
            Dim year As String = selectedDate.ToString("yyyy", georgianCulture)

            ' 🔧 კრიტიკული შესწორება: RotatedLabel ზომა დამოკიდებული უნდა იყოს მასშტაბზე!
            Dim rotatedDateLabel As New RotatedLabel()

            ' შევინახოთ სათანადო ზომები - მასშტაბის დამოუკიდებლად
            rotatedDateLabel.Size = New Size(dateColumnPanel.Width - 4, dateColumnPanel.Height - 4)
            rotatedDateLabel.Location = New Point(2, 2)
            rotatedDateLabel.BackColor = Color.FromArgb(60, 80, 150)
            rotatedDateLabel.ForeColor = Color.White

            ' 🔧 შრიფტის ზომა ვერტიკალური მასშტაბის დამოუკიდებლად (ევროპული სტანდარტი)
            Dim fontSize As Single = 10 ' ფიქსირებული ზომა, რომელიც არ იმასშტაბირდება
            rotatedDateLabel.Font = New Font("Sylfaen", fontSize, FontStyle.Bold)

            ' ვერტიკალური ტექსტის ფორმირება
            Dim dateText As New StringBuilder()
            dateText.Append(weekDay)
            dateText.Append("  ")
            dateText.Append(dayOfMonth)
            dateText.Append("  ")
            dateText.Append(month)
            dateText.Append("  ")
            dateText.Append(year)

            rotatedDateLabel.Text = dateText.ToString()
            rotatedDateLabel.RotationAngle = 270 ' შემობრუნება 270 გრადუსით

            ' დავამატოთ ლეიბლი პანელზე
            dateColumnPanel.Controls.Add(rotatedDateLabel)

            Debug.WriteLine($"FillDateColumnPanel: თარიღის ზოლი შევსებულია, პანელის ზომა: {dateColumnPanel.Size}, fontSize: {fontSize}")

        Catch ex As Exception
            Debug.WriteLine($"FillDateColumnPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' დროის პანელის შევსება ინტერვალებით
    ''' გაუმჯობესებული ვერსია გრძელი მწკრივების ჩვენებისთვის
    ''' </summary>
    Private Sub FillTimeColumnPanel()
        Try
            Debug.WriteLine("FillTimeColumnPanel: დაიწყო დროის პანელის შევსება")

            ' ვიპოვოთ დროის პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)

            If timeColumnPanel Is Nothing Then
                Debug.WriteLine("FillTimeColumnPanel: დროის პანელი ვერ მოიძებნა")
                Return
            End If

            ' გავასუფთავოთ
            timeColumnPanel.Controls.Clear()

            ' თუ დროის ინტერვალები არ არის ინიციალიზებული
            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' მწკრივის სიმაღლე (იმასშტაბირდება)
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)

            ' განვაახლოთ დროის პანელის სიმაღლე
            timeColumnPanel.Height = ROW_HEIGHT * timeIntervals.Count

            Debug.WriteLine($"FillTimeColumnPanel: დროის პანელის სიმაღლე: {timeColumnPanel.Height}, ROW_HEIGHT: {ROW_HEIGHT}, ინტერვალები: {timeIntervals.Count}")

            ' დროის ეტიკეტები
            For i As Integer = 0 To timeIntervals.Count - 1
                Dim timeLabel As New Label()
                timeLabel.Size = New Size(timeColumnPanel.Width - 5, ROW_HEIGHT)
                timeLabel.Location = New Point(0, i * ROW_HEIGHT)
                timeLabel.Text = timeIntervals(i).ToString("HH:mm")
                timeLabel.TextAlign = ContentAlignment.MiddleRight
                timeLabel.Padding = New Padding(0, 0, 5, 0)
                timeLabel.Font = New Font("Segoe UI", 8, FontStyle.Regular)

                ' შევინახოთ თავდაყრო Y პოზიცია Tag-ში სქროლის სინქრონიზაციისთვის
                timeLabel.Tag = i * ROW_HEIGHT

                ' ალტერნატიული ფერები მწკრივებისთვის
                If i Mod 2 = 0 Then
                    timeLabel.BackColor = Color.FromArgb(245, 245, 250)
                Else
                    timeLabel.BackColor = Color.FromArgb(250, 250, 255)
                End If

                ' ჩარჩოს დამატება
                timeLabel.BorderStyle = BorderStyle.FixedSingle

                timeColumnPanel.Controls.Add(timeLabel)

                Debug.WriteLine($"FillTimeColumnPanel: დაემატა დროის ლეიბლი [{i}]: {timeIntervals(i):HH:mm}, Tag={timeLabel.Tag}")
            Next

            Debug.WriteLine($"FillTimeColumnPanel: დასრულდა - დაემატა {timeIntervals.Count} დროის ეტიკეტი")

        Catch ex As Exception
            Debug.WriteLine($"FillTimeColumnPanel: შეცდომა - {ex.Message}")
            Debug.WriteLine($"FillTimeColumnPanel: StackTrace: {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' სივრცეების სათაურების პანელის შევსება და ზომის დაყენება
    ''' Tag-ების შენახვით სქროლის სინქრონიზაციისთვის
    ''' </summary>
    Private Sub FillSpacesHeaderPanel()
        Try
            Debug.WriteLine("FillSpacesHeaderPanel: დაიწყო სივრცეების სათაურების შევსება")

            ' ვიპოვოთ სივრცეების სათაურების პანელი
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)

            If spacesHeaderPanel Is Nothing Then
                Debug.WriteLine("FillSpacesHeaderPanel: სივრცეების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' გავასუფთავოთ
            spacesHeaderPanel.Controls.Clear()

            ' თუ სივრცეები არ არის ჩატვირთული
            If spaces.Count = 0 Then
                LoadSpaces()
            End If

            Debug.WriteLine($"FillSpacesHeaderPanel: ჩატვირთული სივრცეების რაოდენობა: {spaces.Count}")

            ' სივრცის სვეტის სიგანე (იმასშტაბირდება)
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' გამოვთვალოთ სათაურების პანელის სრული სიგანე
            Dim totalHeaderWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count

            ' შევცვალოთ პანელის სიგანე მასშტაბის მიხედვით
            spacesHeaderPanel.Width = totalHeaderWidth

            Debug.WriteLine($"FillSpacesHeaderPanel: პანელის სიგანე: {spacesHeaderPanel.Width}, SPACE_COLUMN_WIDTH: {SPACE_COLUMN_WIDTH}")

            ' მთავრული ფონტი
            Dim mtavruliFont As New Font("Sylfaen", 9, FontStyle.Bold)
            Try
                If FontFamily.Families.Any(Function(f) f.Name = "BPG_Nino_Mtavruli") Then
                    mtavruliFont = New Font("BPG_Nino_Mtavruli", 9, FontStyle.Bold)
                    Debug.WriteLine("FillSpacesHeaderPanel: გამოყენებული BPG_Nino_Mtavruli ფონტი")
                ElseIf FontFamily.Families.Any(Function(f) f.Name = "ALK_Tall_Mtavruli") Then
                    mtavruliFont = New Font("ALK_Tall_Mtavruli", 9, FontStyle.Bold)
                    Debug.WriteLine("FillSpacesHeaderPanel: გამოყენებული ALK_Tall_Mtavruli ფონტი")
                Else
                    Debug.WriteLine("FillSpacesHeaderPanel: გამოყენებული Sylfaen ფონტი")
                End If
            Catch fontEx As Exception
                Debug.WriteLine($"FillSpacesHeaderPanel: ფონტის შეცდომა - {fontEx.Message}")
            End Try

            ' სივრცეების სათაურების შექმნა
            For i As Integer = 0 To spaces.Count - 1
                Dim spaceHeader As New Label()
                spaceHeader.Size = New Size(SPACE_COLUMN_WIDTH - 1, spacesHeaderPanel.Height - 2)
                spaceHeader.Location = New Point(i * SPACE_COLUMN_WIDTH, 1)
                spaceHeader.BackColor = Color.FromArgb(60, 80, 150)
                spaceHeader.ForeColor = Color.White
                spaceHeader.TextAlign = ContentAlignment.MiddleCenter
                spaceHeader.Text = spaces(i).ToUpper()
                spaceHeader.Font = mtavruliFont
                spaceHeader.BorderStyle = BorderStyle.FixedSingle

                ' შევინახოთ თავდაყრო X პოზიცია Tag-ში სქროლის სინქრონიზაციისთვის
                spaceHeader.Tag = i * SPACE_COLUMN_WIDTH

                spacesHeaderPanel.Controls.Add(spaceHeader)

                Debug.WriteLine($"FillSpacesHeaderPanel: დაემატა სივრცის სათაური [{i}]: {spaces(i)}, Tag={spaceHeader.Tag}")
            Next

            Debug.WriteLine($"FillSpacesHeaderPanel: დაემატა {spaces.Count} სივრცის სათაური, მასშტაბი: {hScale}")
            Debug.WriteLine($"FillSpacesHeaderPanel: სათაურების პანელის სიგანე: {totalHeaderWidth}")

        Catch ex As Exception
            Debug.WriteLine($"FillSpacesHeaderPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის შევსება - ფუნდამენტურად გადამუშავებული ორმაგი სკროლის თავიდან აცილებისთვის
    ''' </summary>
    Private Sub FillMainGridPanel()
        Try
            Debug.WriteLine("FillMainGridPanel: დაიწყო მთავარი გრიდის შევსება")

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("FillMainGridPanel: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' გავასუფთავოთ მთავარი პანელი - ამჯერად არ ვქმნით ცალკე შიდა პანელს!
            mainGridPanel.Controls.Clear()

            ' შევამოწმოთ სივრცეები და დრო
            If spaces.Count = 0 Then
                LoadSpaces()
            End If

            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' პარამეტრები - მასშტაბირების გამოყენებით
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)

            ' გრიდის სრული სიგანე (გავითვალისწინოთ მასშტაბი)
            Dim gridWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count

            ' გრიდის სრული სიმაღლე
            Dim gridHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            ' მნიშვნელოვანი: დავაყენოთ mainGridPanel-ის AutoScrollMinSize
            ' ეს არის ერთადერთი ადგილი, სადაც ვაკონტროლებთ სკროლს!
            mainGridPanel.AutoScrollMinSize = New Size(gridWidth, gridHeight)

            Debug.WriteLine($"FillMainGridPanel: გრიდის ზომები - Width={gridWidth}, Height={gridHeight}, hScale={hScale}, vScale={vScale}")

            ' უჯრედების მასივი
            ReDim gridCells(spaces.Count - 1, timeIntervals.Count - 1)

            ' უჯრედების შექმნა - პირდაპირ mainGridPanel-ზე
            For col As Integer = 0 To spaces.Count - 1
                For row As Integer = 0 To timeIntervals.Count - 1
                    Dim cell As New Panel()
                    cell.Size = New Size(SPACE_COLUMN_WIDTH, ROW_HEIGHT)
                    cell.Location = New Point(col * SPACE_COLUMN_WIDTH, row * ROW_HEIGHT)

                    ' ალტერნატიული ფერები
                    If row Mod 2 = 0 Then
                        cell.BackColor = Color.FromArgb(250, 250, 250)
                    Else
                        cell.BackColor = Color.FromArgb(245, 245, 245)
                    End If

                    ' უჯრედის ჩარჩო
                    AddHandler cell.Paint, AddressOf Cell_Paint

                    ' ვინახავთ უჯრედის ობიექტს მასივში
                    gridCells(col, row) = cell

                    ' მნიშვნელოვანი: ვამატებთ უჯრედს პირდაპირ mainGridPanel-ზე
                    mainGridPanel.Controls.Add(cell)
                Next
            Next

            Debug.WriteLine($"FillMainGridPanel: შეიქმნა {spaces.Count * timeIntervals.Count} უჯრედი")

        Catch ex As Exception
            Debug.WriteLine($"FillMainGridPanel: შეცდომა - {ex.Message}")
            Debug.WriteLine($"FillMainGridPanel: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' დღის ხედის ჩვენება სივრცეების მიხედვით - 🔧 ბარათების გასუფთავების დამატებით
    ''' </summary>
    Private Sub ShowDayViewBySpace()
        Try
            Debug.WriteLine("ShowDayViewBySpace: დაიწყო დღის ხედის ჩვენება სივრცეების მიხედვით")

            ' დავაყენოთ pnlFilter-ის ფერი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' ======= 1. პანელების ინიციალიზაცია =======
            InitializeDayViewPanels()

            ' ======= 2. თარიღის ზოლის შევსება =======
            FillDateColumnPanel()

            ' ======= 3. დროის პანელის შევსება =======
            FillTimeColumnPanel()

            ' ======= 4. სივრცეების სათაურების პანელის შევსება =======
            FillSpacesHeaderPanel()

            ' ======= 5. მთავარი გრიდის პანელის შევსება =======
            FillMainGridPanel()

            ' ======= 6. მთავარი გრიდის პანელის ზომის დაყენება =======
            SetMainGridPanelSize()

            ' ======= 7. სინქრონიზაცია სქროლისთვის =======
            SetupScrollSynchronization()

            ' ======= 8. 🔧 ძველი ბარათების გასუფთავება =======
            ClearSessionCardsFromGrid()

            ' ======= 9. სესიების განთავსება გრიდში =======
            PlaceSessionsOnGrid()

            Debug.WriteLine("ShowDayViewBySpace: დღის ხედის ჩვენება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewBySpace: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewBySpace: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"დღის ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სქროლის სინქრონიზაცია - გასწორებული ვერსია
    ''' </summary>
    Private Sub SetupScrollSynchronization()
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("SetupScrollSynchronization: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' ვიპოვოთ დროის პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)

            If timeColumnPanel Is Nothing Then
                Debug.WriteLine("SetupScrollSynchronization: დროის პანელი ვერ მოიძებნა")
                Return
            End If

            ' ვიპოვოთ სივრცეების სათაურების პანელი
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)

            If spacesHeaderPanel Is Nothing Then
                Debug.WriteLine("SetupScrollSynchronization: სივრცეების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' ჯერ მოვხსნათ არსებული ივენთის ჰენდლერები, თუ ისინი არსებობენ
            RemoveHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' მივაბათ Scroll ივენთი ხელახლა
            AddHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' ასევე, დავამატოთ ივენთი გორგოლჭის სქროლისთვის
            AddHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            Debug.WriteLine("SetupScrollSynchronization: სქროლის სინქრონიზაცია დაყენებულია")

        Catch ex As Exception
            Debug.WriteLine($"SetupScrollSynchronization: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის სქროლის ივენთი 
    ''' 🔧 გასწორებული ვერსია: თარიღისა და დროის სვეტების სწორი სინქრონიზაცია
    ''' </summary>
    Private Sub MainGridPanel_Scroll(sender As Object, e As ScrollEventArgs)
        Try
            Dim mainGridPanel As Panel = DirectCast(sender, Panel)

            ' ვიპოვოთ ყველა საჭირო პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)
            Dim dateColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("dateColumnPanel", False).FirstOrDefault(), Panel)
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)

            ' ვერტიკალური სქროლი - 🔧 კრიტიკული შესწორება
            If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
                Dim scrollOffset As Integer = -mainGridPanel.VerticalScroll.Value

                ' დროის პანელის Y სინქრონიზაცია
                If timeColumnPanel IsNot Nothing Then
                    timeColumnPanel.Top = BASE_HEADER_HEIGHT + scrollOffset
                End If

                ' თარიღის პანელის Y სინქრონიზაცია
                If dateColumnPanel IsNot Nothing Then
                    dateColumnPanel.Top = BASE_HEADER_HEIGHT + scrollOffset
                End If

                Debug.WriteLine($"MainGridPanel_Scroll: ვერტიკალური სქროლი - Value = {mainGridPanel.VerticalScroll.Value}, Offset = {scrollOffset}")
            End If

            ' ჰორიზონტალური სქროლი - 🔧 განსაკუთრებით კრიტიკული შესწორება!
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                Dim scrollOffset As Integer = -mainGridPanel.HorizontalScroll.Value

                ' 🔧 მნიშვნელოვანი: სივრცეების სათაურები მოძრაობენ ჰორიზონტალური სქროლის მიხედვით
                If spacesHeaderPanel IsNot Nothing Then
                    ' ასე გამოვთვლით სწორ X პოზიციას
                    Dim fixedLeftPosition As Integer = BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH
                    spacesHeaderPanel.Left = fixedLeftPosition + scrollOffset
                    Debug.WriteLine($"MainGridPanel_Scroll: სივრცეების სათაურები გადაადგილდა Left={spacesHeaderPanel.Left}")
                End If

                ' 🔧 პრობლემა იყო აქ - თარიღისა და დროის სვეტები არ უნდა იმოძრაონ ჰორიზონტალურად!
                ' ეს ყოველთვის ფიქსირებული პოზიციებზე რჩებიან
                Debug.WriteLine($"MainGridPanel_Scroll: ჰორიზონტალური სქროლი - Value = {mainGridPanel.HorizontalScroll.Value}, Offset = {scrollOffset}")
            End If

        Catch ex As Exception
            Debug.WriteLine($"MainGridPanel_Scroll: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' შემობრუნებული ტექსტის ლეიბლი - საშუალებას გვაძლევს შევაბრუნოთ ტექსტი ნებისმიერი კუთხით
    ''' </summary>
    Public Class RotatedLabel
        Inherits Control

        Private _text As String = String.Empty
        Private _rotationAngle As Integer = 0

        ''' <summary>
        ''' შესაბრუნებელი კუთხე გრადუსებში (0-360)
        ''' </summary>
        Public Property RotationAngle As Integer
            Get
                Return _rotationAngle
            End Get
            Set(value As Integer)
                If _rotationAngle <> value Then
                    _rotationAngle = value
                    Me.Invalidate() ' გადავხატოთ კონტროლი
                End If
            End Set
        End Property

        ''' <summary>
        ''' ლეიბლის ტექსტი
        ''' </summary>
        Public Overrides Property Text As String
            Get
                Return _text
            End Get
            Set(value As String)
                If _text <> value Then
                    _text = value
                    Me.Invalidate() ' გადავხატოთ კონტროლი
                End If
            End Set
        End Property

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        Public Sub New()
            Me.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
            Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            Me.SetStyle(ControlStyles.UserPaint, True)
            Me.SetStyle(ControlStyles.ResizeRedraw, True)
        End Sub

        ''' <summary>
        ''' კონტროლის მოხატვის მეთოდი
        ''' </summary>
        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)

            If String.IsNullOrEmpty(_text) Then
                Return
            End If

            ' ფონის შევსება
            Using brush As New SolidBrush(Me.BackColor)
                e.Graphics.FillRectangle(brush, Me.ClientRectangle)
            End Using

            ' ანტიალიასინგის ჩართვა (უფრო გლუვი ტექსტისთვის)
            e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

            ' დავამატოთ ტრანსფორმაცია (ბრუნვა)
            e.Graphics.TranslateTransform(Me.Width / 2, Me.Height / 2)
            e.Graphics.RotateTransform(_rotationAngle)

            ' გავზომოთ ტექსტის ზომა
            Dim textSize = e.Graphics.MeasureString(_text, Me.Font)

            ' გამოვთვალოთ ტექსტის პოზიცია - ცენტრირებული
            Dim x = -textSize.Width / 2
            Dim y = -textSize.Height / 2

            ' დავხატოთ ტექსტი
            Using brush As New SolidBrush(Me.ForeColor)
                e.Graphics.DrawString(_text, Me.Font, brush, x, y)
            End Using

            ' დავაბრუნოთ ტრანსფორმაცია საწყის მდგომარეობაში
            e.Graphics.ResetTransform()
        End Sub
    End Class

    ''' <summary>
    ''' მაუსის გორგოლჭის ივენთი - 🔧 გასწორებული სქროლის სინქრონიზაციისთვის
    ''' </summary>
    Private Sub MainGridPanel_MouseWheel(sender As Object, e As MouseEventArgs)
        Try
            Dim mainGridPanel As Panel = DirectCast(sender, Panel)

            ' ვიპოვოთ საჭირო პანელები
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)
            Dim dateColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("dateColumnPanel", False).FirstOrDefault(), Panel)
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)

            ' 🔧 ვერტიკალური სქროლის სინქრონიზაცია (მუშაობს კარგად)
            Dim verticalScrollOffset As Integer = -mainGridPanel.VerticalScroll.Value

            If timeColumnPanel IsNot Nothing Then
                timeColumnPanel.Top = BASE_HEADER_HEIGHT + verticalScrollOffset
            End If

            If dateColumnPanel IsNot Nothing Then
                dateColumnPanel.Top = BASE_HEADER_HEIGHT + verticalScrollOffset
            End If

            ' 🔧 ჰორიზონტალური სქროლის სინქრონიზაცია (ახალი ლოგიკა)
            Dim horizontalScrollOffset As Integer = -mainGridPanel.HorizontalScroll.Value

            If spacesHeaderPanel IsNot Nothing Then
                Dim fixedLeftPosition As Integer = BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH
                spacesHeaderPanel.Left = fixedLeftPosition + horizontalScrollOffset
            End If

        Catch ex As Exception
            Debug.WriteLine($"MainGridPanel_MouseWheel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' უჯრედის ჩარჩოს დახატვა
    ''' </summary>
    Private Sub Cell_Paint(sender As Object, e As PaintEventArgs)
        Dim cell As Panel = DirectCast(sender, Panel)

        Using pen As New Pen(Color.FromArgb(200, 200, 200), 1)
            ' 🔧 მხოლოდ ვერტიკალური ხაზები - მარჯვენა ბორდერი
            e.Graphics.DrawLine(pen, cell.Width - 1, 0, cell.Width - 1, cell.Height - 1)

            ' 🔧 არასავალდებულო: მარცხენა ბორდერი მხოლოდ პირველი სვეტისთვის
            ' თუ გსურს სრული ვერტიკალური გამიჯვნა, ამ ხაზებს ნუ შეცვლი
            ' თუ გსურს მხოლოდ სივრცეებს შორის გამიჯვნა, დაუმატე პირობა
        End Using
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
        lblNotImplemented.Font = New Font("Sylfaen", 12, FontStyle.Bold)
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
        lblNotImplemented.Font = New Font("Sylfaen", 12, FontStyle.Bold)
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
        lblNotImplemented.Font = New Font("Sylfaen", 12, FontStyle.Bold)
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
        lblNotImplemented.Font = New Font("Sylfaen", 12, FontStyle.Bold)
        pnlCalendarGrid.Controls.Clear()
        pnlCalendarGrid.Controls.Add(lblNotImplemented)
    End Sub

    ''' <summary>
    ''' კალენდრის ხედის განახლება - ფუნდამენტურად გადამუშავებული ορმაგი სკროლის თავიდან აცილებისთვის
    ''' </summary>
    Public Sub UpdateCalendarView()
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
            Debug.WriteLine($"UpdateCalendarView: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"კალენდრის ხედის განახლების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

        ' განვაახლოთ კალენდრის ხედი
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' კალენდრის თარიღის ცვლილების დამუშავება
    ''' </summary>
    Private Sub DTPCalendar_ValueChanged(sender As Object, e As EventArgs)
        ' განვაახლოთ კალენდრის ხედი
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' ივენთი, რომელიც გაეშვება UserControl-ის ჩატვირთვისას
    ''' გასწორებული ორმაგი სკროლის პრობლემისთვის
    ''' </summary>
    Private Sub UC_Calendar_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            ' მნიშვნელოვანი: უზრუნველვყოფთ, რომ UserControl-ს და pnlCalendarGrid-ს არ ჰქონდეს AutoScroll
            Me.AutoScroll = False
            If pnlCalendarGrid IsNot Nothing Then
                pnlCalendarGrid.AutoScroll = False
            End If

            ' საწყისი პარამეტრების დაყენება (თუ ისინი არ არის დაყენებული)
            InitializeTimeIntervals()

            ' კალენდრის განახლება
            UpdateCalendarView()

            Debug.WriteLine("UC_Calendar_Load: კალენდრის კონტროლი წარმატებით ჩაიტვირთა, AutoScroll გამორთულია")
        Catch ex As Exception
            Debug.WriteLine($"UC_Calendar_Load: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' UserControl-ის Resize ივენთის დამმუშავებელი
    ''' გასწორებული ორმაგი სკროლისთვის
    ''' </summary>
    Private Sub UC_Calendar_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Try
            ' თუ კონტროლი არ არის ინიციალიზებული, გამოვიდეთ
            If Not Me.IsHandleCreated Then
                Return
            End If

            ' მნიშვნელოვანი: დავრწმუნდეთ რომ არ არის AutoScroll ჩართული
            Me.AutoScroll = False
            If pnlCalendarGrid IsNot Nothing Then
                pnlCalendarGrid.AutoScroll = False

                ' განვაახლოთ კალენდრის გრიდის ზომები კონტეინერის მიხედვით
                pnlCalendarGrid.Width = Me.Width - pnlCalendarGrid.Left - 20
                pnlCalendarGrid.Height = Me.Height - pnlCalendarGrid.Top - 20
            End If

            ' ფილტრის პანელის ზომის განახლება
            If pnlFIlter IsNot Nothing Then
                pnlFIlter.Width = Me.Width - pnlFIlter.Left - 20
            End If

            ' თუ კალენდრის ხედი უკვე გაშლილია, განვაახლოთ შემცველობა
            UpdateCalendarView()

            Debug.WriteLine($"UC_Calendar_Resize: კალენდრის ზომები განახლდა, ახალი ზომები: {Me.Width}x{Me.Height}, AutoScroll={Me.AutoScroll}")
        Catch ex As Exception
            Debug.WriteLine($"UC_Calendar_Resize: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' უჯრედების ძველი ბარათების გასუფთავება - 🔧 ახლა mainGridPanel-დანაც შლის
    ''' </summary>
    Private Sub ClearSessionCardsFromGrid()
        Try
            Debug.WriteLine("ClearSessionCardsFromGrid: ვშლით ყველა ძველ ბარათს")

            ' 1. ვშლით ბარათებს უჯრედებიდან (თუ რამე დაშენარჩუნებულა)
            If gridCells IsNot Nothing Then
                For col As Integer = 0 To gridCells.GetLength(0) - 1
                    For row As Integer = 0 To gridCells.GetLength(1) - 1
                        Dim cell = gridCells(col, row)
                        If cell IsNot Nothing Then
                            cell.Controls.Clear()
                        End If
                    Next
                Next
            End If

            ' 2. 🔧 ახალი: ვშლით ბარათებს mainGridPanel-დანაც
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            If mainGridPanel IsNot Nothing Then
                ' ვშლით ყველა ბარათს, რომელთაც Tag აქვთ (სესიის ID)
                Dim cardsToRemove As New List(Of Control)
                For Each ctrl As Control In mainGridPanel.Controls
                    If TypeOf ctrl Is Panel AndAlso ctrl.Tag IsNot Nothing Then
                        ' ეს სესიის ბარათია
                        cardsToRemove.Add(ctrl)
                    End If
                Next

                ' ვშლით ყველა ნაპოვნ ბარათს
                For Each card In cardsToRemove
                    mainGridPanel.Controls.Remove(card)
                    card.Dispose()
                Next

                Debug.WriteLine($"ClearSessionCardsFromGrid: mainGridPanel-დან წაიშალა {cardsToRemove.Count} ბარათი")
            End If

            Debug.WriteLine("ClearSessionCardsFromGrid: ყველა ბარათი წარმატებით მოხსნილია")
        Catch ex As Exception
            Debug.WriteLine($"ClearSessionCardsFromGrid: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ჰორიზონტალური მასშტაბის გაზრდა - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Sub BtnHUp_Click(sender As Object, e As EventArgs) Handles BtnHUp.Click
        Try
            Dim oldHScale As Double = hScale
            hScale += SCALE_STEP
            If hScale > MAX_SCALE Then hScale = MAX_SCALE

            Debug.WriteLine($"BtnHUp_Click: ჰორიზონტალური მასშტაბი შეიცვალა {oldHScale:F2} -> {hScale:F2}")

            ' თუ მასშტაბი შეიცვალა, განვაახლოთ კალენდარი
            If oldHScale <> hScale Then
                UpdateCalendarView() ' 🔧 სრული განახლება!
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnHUp_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ჰორიზონტალური მასშტაბის შემცირება - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Sub BtnHDown_Click(sender As Object, e As EventArgs) Handles BtnHDown.Click
        Try
            Dim oldHScale As Double = hScale
            hScale -= SCALE_STEP
            If hScale < MIN_SCALE Then hScale = MIN_SCALE

            Debug.WriteLine($"BtnHDown_Click: ჰორიზონტალური მასშტაბი შეიცვალა {oldHScale:F2} -> {hScale:F2}")

            ' თუ მასშტაბი შეიცვალა, განვაახლოთ კალენდარი
            If oldHScale <> hScale Then
                UpdateCalendarView() ' 🔧 სრული განახლება!
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnHDown_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ვერტიკალური მასშტაბის გაზრდა - გასწორებული ვერსია თარიღის სვეტთან
    ''' </summary>
    Private Sub BtnVUp_Click(sender As Object, e As EventArgs) Handles BtnVUp.Click
        Try
            ' მიმდინარე მასშტაბის შენახვა დებაგინგისთვის
            Dim oldVScale As Double = vScale

            ' მასშტაბის გაზრდა დადგენილი ნაბიჯით
            vScale += SCALE_STEP

            ' მაქსიმალური მასშტაბის შეზღუდვა
            If vScale > MAX_SCALE Then vScale = MAX_SCALE

            Debug.WriteLine($"BtnVUp_Click: ვერტიკალური მასშტაბი შეიცვალა {oldVScale:F2} -> {vScale:F2}")

            ' კალენდრის განახლება ახალი მასშტაბით მხოლოდ მაშინ, თუ მასშტაბი შეიცვალა
            If oldVScale <> vScale Then
                ' განაახლეთ მასშტაბის ლეიბლი
                UpdateScaleLabels()

                ' საჭიროა სრული განახლება ვერტიკალური სქეილის ცვლილებისას
                ' რადგან თარიღის სვეტი, დროის სვეტი და გრიდის სიმაღლე იცვლება
                UpdateCalendarView()
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnVUp_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnVUp_Click: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ვერტიკალური მასშტაბის შემცირება - გასწორებული ვერსია თარიღის სვეტთან
    ''' </summary>
    Private Sub BtnVDown_Click(sender As Object, e As EventArgs) Handles BtnVDown.Click
        Try
            ' მიმდინარე მასშტაბის შენახვა დებაგინგისთვის
            Dim oldVScale As Double = vScale

            ' მასშტაბის შემცირება დადგენილი ნაბიჯით
            vScale -= SCALE_STEP

            ' მინიმალური მასშტაბის შეზღუდვა
            If vScale < MIN_SCALE Then vScale = MIN_SCALE

            Debug.WriteLine($"BtnVDown_Click: ვერტიკალური მასშტაბი შეიცვალა {oldVScale:F2} -> {vScale:F2}")

            ' კალენდრის განახლება ახალი მასშტაბით მხოლოდ მაშინ, თუ მასშტაბი შეიცვალა
            If oldVScale <> vScale Then
                ' განაახლეთ მასშტაბის ლეიბლი
                UpdateScaleLabels()

                ' საჭიროა სრული განახლება ვერტიკალური სქეილის ცვლილებისას
                ' რადგან თარიღის სვეტი, დროის სვეტი და გრიდის სიმაღლე იცვლება
                UpdateCalendarView()
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnVDown_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnVDown_Click: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' კალენდარში სესიების ჩვენება
    ''' </summary>
    Private Sub ShowSessionsInCalendar()
        Try
            Debug.WriteLine("ShowSessionsInCalendar: დაიწყო სესიების ჩვენება კალენდარში")

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("ShowSessionsInCalendar: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' შევამოწმოთ არის თუ არა სესიები ჩატვირთული
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                Debug.WriteLine("ShowSessionsInCalendar: სესიები არ არის ჩატვირთული")
                Return
            End If

            ' არჩეული თარიღი (DTPCalendar-დან)
            Dim selectedDate As DateTime = DTPCalendar.Value.Date

            ' ფილტრაცია: მხოლოდ არჩეული დღის სესიები
            Dim daySessions As List(Of SessionModel) = allSessions.Where(Function(s) s.DateTime.Date = selectedDate).ToList()

            Debug.WriteLine($"ShowSessionsInCalendar: ნაპოვნია {daySessions.Count} სესია {selectedDate:dd.MM.yyyy} თარიღისთვის")

            ' სესიების პანელები
            For Each session In daySessions
                Try
                    ' ვიპოვოთ შესაბამისი სივრცე და დრო
                    Dim spaceIndex As Integer = spaces.IndexOf(session.Space)
                    If spaceIndex < 0 Then
                        ' თუ სივრცე ვერ მოიძებნა, გავაგრძელოთ
                        Continue For
                    End If

                    ' მოვძებნოთ უახლოესი დროის ინტერვალი
                    Dim sessionTime As DateTime = New DateTime(
                        selectedDate.Year,
                        selectedDate.Month,
                        selectedDate.Day,
                        session.DateTime.Hour,
                        session.DateTime.Minute,
                        0)

                    Dim timeIndex As Integer = -1
                    For i As Integer = 0 To timeIntervals.Count - 1
                        If timeIntervals(i) >= sessionTime Then
                            timeIndex = i
                            Exit For
                        End If
                    Next

                    ' თუ ვერ მოვძებნეთ დროის ინტერვალი, გავაგრძელოთ
                    If timeIndex < 0 Then
                        Continue For
                    End If

                    ' შევამოწმოთ არის თუ არა უჯრედი შექმნილი
                    If gridCells Is Nothing OrElse gridCells.GetLength(0) <= spaceIndex OrElse gridCells.GetLength(1) <= timeIndex Then
                        Continue For
                    End If

                    ' მოვიპოვოთ უჯრედი
                    Dim cell As Panel = gridCells(spaceIndex, timeIndex)

                    ' შევქმნათ სესიის ბარათი
                    Dim sessionCard As New Panel()
                    sessionCard.Size = New Size(cell.Width - 4, cell.Height - 4)
                    sessionCard.Location = New Point(2, 2)
                    sessionCard.BackColor = GetSessionColor(session.Status)
                    sessionCard.Tag = session.Id ' შევინახოთ სესიის ID ბარათის Tag-ში

                    ' ვიზუალური გაფორმება - მომრგვალებული კუთხეები
                    AddHandler sessionCard.Paint, AddressOf SessionCard_Paint

                    ' ბენეფიციარის სახელი
                    Dim lblBeneficiary As New Label()
                    lblBeneficiary.Location = New Point(3, 3)
                    lblBeneficiary.Size = New Size(sessionCard.Width - 6, 20)
                    lblBeneficiary.Font = New Font("Sylfaen", 8, FontStyle.Bold)
                    lblBeneficiary.Text = session.BeneficiaryName + " " + session.BeneficiarySurname
                    lblBeneficiary.TextAlign = ContentAlignment.MiddleCenter
                    sessionCard.Controls.Add(lblBeneficiary)

                    ' თერაპევტის სახელი
                    Dim lblTherapist As New Label()
                    lblTherapist.Location = New Point(3, 23)
                    lblTherapist.Size = New Size(sessionCard.Width - 6, 14)
                    lblTherapist.Font = New Font("Sylfaen", 7, FontStyle.Regular)
                    lblTherapist.Text = session.TherapistName
                    lblTherapist.TextAlign = ContentAlignment.MiddleLeft
                    sessionCard.Controls.Add(lblTherapist)

                    ' სესიის ID და დრო
                    Dim lblInfo As New Label()
                    lblInfo.Location = New Point(3, sessionCard.Height - 17)
                    lblInfo.Size = New Size(sessionCard.Width - 6, 14)
                    lblInfo.Font = New Font("Sylfaen", 7, FontStyle.Regular)
                    lblInfo.Text = $"#{session.Id} - {session.DateTime:HH:mm}"
                    lblInfo.TextAlign = ContentAlignment.MiddleRight
                    sessionCard.Controls.Add(lblInfo)

                    ' დავამატოთ ივენთის ჰენდლერი - ორჯერ დაწკაპუნებით რედაქტირება
                    AddHandler sessionCard.DoubleClick, AddressOf SessionCard_DoubleClick

                    ' დავამატოთ ბარათი უჯრედში
                    cell.Controls.Add(sessionCard)
                    sessionCard.BringToFront()

                    Debug.WriteLine($"ShowSessionsInCalendar: დაემატა სესია ID={session.Id}, სივრცე={session.Space}, დრო={session.DateTime:HH:mm}")

                Catch ex As Exception
                    Debug.WriteLine($"ShowSessionsInCalendar: შეცდომა სესიის დამატებისას - {ex.Message}")
                    ' გავაგრძელოთ შემდეგი სესიით შეცდომის შემთხვევაში
                    Continue For
                End Try
            Next

            Debug.WriteLine("ShowSessionsInCalendar: სესიების ჩვენება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowSessionsInCalendar: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის სტატუსის მიხედვით ფერის განსაზღვრა
    ''' </summary>
    ''' <param name="status">სესიის სტატუსი</param>
    ''' <returns>შესაბამისი ფერი</returns>
    Private Function GetSessionColor(status As String) As Color
        ' ფერების განსაზღვრა სტატუსის მიხედვით
        Select Case status.Trim().ToLower()
            Case "დაგეგმილი"
                Return Color.FromArgb(200, 230, 255) ' ღია ლურჯი
            Case "შესრულებული"
                Return Color.FromArgb(200, 255, 200) ' ღია მწვანე
            Case "გაუქმებული", "გაუქმება"
                Return Color.FromArgb(255, 200, 200) ' ღია წითელი
            Case Else
                Return Color.FromArgb(240, 240, 240) ' ღია ნაცრისფერი
        End Select
    End Function

    ''' <summary>
    ''' სესიის ბარათის დახატვის ივენთი - მომრგვალებული კუთხეები
    ''' </summary>
    Private Sub SessionCard_Paint(sender As Object, e As PaintEventArgs)
        Dim card As Panel = DirectCast(sender, Panel)
        Dim cornerRadius As Integer = 5

        ' მომრგვალებული კუთხეების გზის შექმნა
        Dim path As New Drawing2D.GraphicsPath()
        path.AddArc(0, 0, cornerRadius * 2, cornerRadius * 2, 180, 90)
        path.AddArc(card.Width - cornerRadius * 2, 0, cornerRadius * 2, cornerRadius * 2, 270, 90)
        path.AddArc(card.Width - cornerRadius * 2, card.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 0, 90)
        path.AddArc(0, card.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 90, 90)
        path.CloseFigure()

        ' რეგიონის დაყენება
        card.Region = New Region(path)

        ' ჩარჩოს დახატვა
        Using pen As New Pen(Color.FromArgb(120, 120, 150), 1)
            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            e.Graphics.DrawPath(pen, path)
        End Using
    End Sub

    ''' <summary>
    ''' სესიის ბარათზე ორჯერ დაწკაპუნების ივენთი - რედაქტირებისთვის
    ''' </summary>
    Private Sub SessionCard_DoubleClick(sender As Object, e As EventArgs)
        Try
            ' ვიპოვოთ სესიის ID
            Dim sessionCard As Panel = DirectCast(sender, Panel)
            Dim sessionId As Integer = CInt(sessionCard.Tag)

            Debug.WriteLine($"SessionCard_DoubleClick: სესიის რედაქტირების მოთხოვნა, ID={sessionId}")

            ' შევამოწმოთ არის თუ არა მომხმარებელი ავტორიზებული რედაქტირებისთვის
            If String.IsNullOrEmpty(userEmail) Then
                MessageBox.Show("ავტორიზაციის გარეშე ვერ მოხდება სესიის რედაქტირება", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ხელმისაწვდომი", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' შევქმნათ და გავხსნათ რედაქტირების ფორმა
            ' NewRecordForm იღებს dataService, "სესია" ტიპს, სესიის ID-ს და მომხმარებლის ელფოსტას
            Dim editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_Calendar")

            ' გავხსნათ ფორმა
            Dim result = editForm.ShowDialog()

            ' თუ ფორმა დაიხურა OK რეზულტატით, განვაახლოთ მონაცემები
            If result = DialogResult.OK Then
                Debug.WriteLine($"SessionCard_DoubleClick: სესია ID={sessionId} წარმატებით განახლდა")

                ' ჩავტვირთოთ სესიები თავიდან
                LoadSessions()

                ' განვაახლოთ კალენდრის ხედი
                UpdateCalendarView()
            End If

        Catch ex As Exception
            Debug.WriteLine($"SessionCard_DoubleClick: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიის რედაქტირების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' პანელების ზომების და პოზიციების განახლება მასშტაბირების შემდეგ
    ''' </summary>
    Private Sub UpdatePanelSizesAndPositions()
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            ' ვიპოვოთ სივრცეების სათაურების პანელი
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)

            ' ვიპოვოთ დროის პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing OrElse spacesHeaderPanel Is Nothing OrElse timeColumnPanel Is Nothing Then
                Debug.WriteLine("UpdatePanelSizesAndPositions: ერთ-ერთი პანელი ვერ მოიძებნა")
                Return
            End If

            ' გამოვთვალოთ სივრცეების სათაურების პანელის სრული სიგანე
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim totalHeaderWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count

            ' განვაახლოთ სივრცეების სათაურების პანელის სიგანე
            spacesHeaderPanel.Width = totalHeaderWidth

            ' დავაყენოთ სკროლის პოზიცია, როდესაც mainGridPanel-ის სკროლი იცვლება
            Dim scrollPosition As Point = mainGridPanel.AutoScrollPosition
            ' AutoScrollPosition აბრუნებს უარყოფით რიცხვებს, ამიტომ ვაბრუნებთ დადებითად
            Dim actualHScrollPos As Integer = Math.Abs(scrollPosition.X)

            ' ვაყენებთ spacesHeaderPanel-ის პოზიციას ჰორიზონტალური სკროლის გათვალისწინებით
            spacesHeaderPanel.Location = New Point(BASE_TIME_COLUMN_WIDTH - actualHScrollPos, 0)

            Debug.WriteLine($"UpdatePanelSizesAndPositions: განახლდა პანელების ზომები და პოზიციები, hScroll={actualHScrollPos}")

        Catch ex As Exception
            Debug.WriteLine($"UpdatePanelSizesAndPositions: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის ზომის დაყენება - განახლებული თარიღის ზოლით
    ''' </summary>
    Private Sub SetMainGridPanelSize()
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("SetMainGridPanelSize: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' ვიპოვოთ შიდა პანელი
            If calendarContentPanel Is Nothing Then
                Debug.WriteLine("SetMainGridPanelSize: შიდა პანელი არ არის")
                Return
            End If

            ' სივრცეების რაოდენობის შემოწმება
            If spaces.Count = 0 Then
                LoadSpaces()
            End If

            ' დროის ინტერვალების შემოწმება
            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' პარამეტრები მასშტაბის გათვალისწინებით
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)

            ' გრიდის სრული სიგანე და სიმაღლე
            Dim gridWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count
            Dim gridHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            ' დავაყენოთ შიდა პანელის ზომა
            calendarContentPanel.Size = New Size(gridWidth, gridHeight)

            ' დავაყენოთ მთავარი გრიდის პანელის AutoScrollMinSize
            mainGridPanel.AutoScrollMinSize = New Size(gridWidth, gridHeight)

            Debug.WriteLine($"SetMainGridPanelSize: გრიდის ზომა დაყენებულია - Width={gridWidth}, Height={gridHeight}")

        Catch ex As Exception
            Debug.WriteLine($"SetMainGridPanelSize: შეცდომა - {ex.Message}")
            Debug.WriteLine($"SetMainGridPanelSize: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების განთავსება გრიდში - 🎨 ახალი ფერების სისტემით
    ''' ✅ გამოყენებს SessionStatusColors კლასს ფერების გასაზღვრავად
    ''' </summary>
    Private Sub PlaceSessionsOnGrid()
        Try
            Debug.WriteLine("=== PlaceSessionsOnGrid: დაიწყო (ახალი ფერების სისტემით) ===")

            ' ძირითადი შემოწმებები
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnGrid: სესიები არ არის")
                Return
            End If

            If gridCells Is Nothing Then
                Debug.WriteLine("PlaceSessionsOnGrid: გრიდის უჯრედები არ არის")
                Return
            End If

            If timeIntervals.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnGrid: დროის ინტერვალები არ არის")
                Return
            End If

            ' მნიშვნელოვანი: ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            If mainGridPanel Is Nothing Then
                Debug.WriteLine("❌ მთავარი გრიდის პანელი ვერ მოიძებნა!")
                Return
            End If

            ' მასშტაბირებული პარამეტრები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' არჩეული თარიღი
            Dim selectedDate As DateTime = DTPCalendar.Value.Date
            Debug.WriteLine($"არჩეული თარიღი: {selectedDate:dd.MM.yyyy}")

            ' ფილტრი - მხოლოდ არჩეული დღის სესიები
            Dim daySessions = allSessions.Where(Function(s) s.DateTime.Date = selectedDate).ToList()
            Debug.WriteLine($"ნაპოვნია {daySessions.Count} სესია ამ თარიღისთვის")

            ' სესიების განთავსება
            For Each session In daySessions
                Try
                    Debug.WriteLine($"--- ვამუშავებთ სესიას ID={session.Id}, სტატუსი='{session.Status}', დრო={session.DateTime:HH:mm}, ხანგრძლივობა={session.Duration}წთ ---")

                    ' 1. სივრცის ინდექსის განსაზღვრა
                    Dim spaceIndex As Integer = -1
                    For i As Integer = 0 To spaces.Count - 1
                        If String.Equals(spaces(i).Trim(), session.Space.Trim(), StringComparison.OrdinalIgnoreCase) Then
                            spaceIndex = i
                            Exit For
                        End If
                    Next

                    If spaceIndex < 0 Then
                        Debug.WriteLine($"❌ სივრცე '{session.Space}' ვერ მოიძებნა!")
                        Continue For
                    End If

                    ' 2. დროის ინტერვალის განსაზღვრა
                    Dim sessionTime As TimeSpan = session.DateTime.TimeOfDay
                    Dim timeIndex As Integer = -1

                    ' ვიპოვოთ ყველაზე ახლო დროის ინტერვალი
                    Dim minDifference As TimeSpan = TimeSpan.MaxValue
                    For i As Integer = 0 To timeIntervals.Count - 1
                        Dim intervalTime As TimeSpan = timeIntervals(i).TimeOfDay
                        Dim difference As TimeSpan = If(sessionTime >= intervalTime,
                                                  sessionTime - intervalTime,
                                                  intervalTime - sessionTime)

                        If difference < minDifference Then
                            minDifference = difference
                            timeIndex = i
                        End If
                    Next

                    ' ალტერნატიული მეთოდი
                    If timeIndex < 0 OrElse minDifference.TotalMinutes > 15 Then
                        Dim startTime As TimeSpan = timeIntervals(0).TimeOfDay
                        Dim elapsedMinutes As Double = (sessionTime - startTime).TotalMinutes
                        timeIndex = CInt(Math.Round(elapsedMinutes / 30))
                        If timeIndex < 0 Then timeIndex = 0
                        If timeIndex >= timeIntervals.Count Then timeIndex = timeIntervals.Count - 1
                    End If

                    ' 3. საზღვრების შემოწმება
                    If spaceIndex < 0 OrElse timeIndex < 0 OrElse
                  spaceIndex >= spaces.Count OrElse timeIndex >= timeIntervals.Count Then
                        Debug.WriteLine($"❌ არასწორი ინდექსები: space={spaceIndex}, time={timeIndex}")
                        Continue For
                    End If

                    ' 4. ბარათის პოზიციის გამოთვლა
                    Dim cardX As Integer = spaceIndex * SPACE_COLUMN_WIDTH + 4
                    Dim cardY As Integer = timeIndex * ROW_HEIGHT + 2

                    ' 5. პროპორციული ბარათის სიმაღლის გამოთვლა (შემცირებული მაქსიმუმი)
                    Dim sessionDurationMinutes As Integer = session.Duration
                    Dim baseCardHeight As Double = ROW_HEIGHT * (sessionDurationMinutes / 30.0)
                    Dim minCardHeight As Integer = ROW_HEIGHT
                    Dim cardHeight As Integer = CInt(Math.Max(baseCardHeight, minCardHeight))
                    Dim maxCardHeight As Integer = ROW_HEIGHT * 4
                    If cardHeight > maxCardHeight Then cardHeight = maxCardHeight

                    ' 6. 🎨 ახალი ფერების განსაზღვრა SessionStatusColors კლასიდან
                    Dim cardColor As Color = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                    Dim borderColor As Color = SessionStatusColors.GetStatusBorderColor(session.Status, session.DateTime)
                    ' ახალი: header-ისთვის მუქი ფერი
                    Dim headerColor As Color = Color.FromArgb(
                   Math.Max(0, cardColor.R - 50),
                   Math.Max(0, cardColor.G - 50),
                   Math.Max(0, cardColor.B - 50)
               )

                    Debug.WriteLine($"🎨 ფერის განსაზღვრა: სტატუსი='{session.Status}' -> ფონი={cardColor}, ჩარჩო={borderColor}, header={headerColor}")

                    ' 7. ბარათის შექმნა
                    Dim sessionCard As New Panel()
                    sessionCard.Size = New Size(SPACE_COLUMN_WIDTH - 8, cardHeight)
                    sessionCard.Location = New Point(cardX, cardY)
                    sessionCard.BackColor = cardColor
                    sessionCard.BorderStyle = BorderStyle.None
                    sessionCard.Tag = session.Id
                    sessionCard.Cursor = Cursors.Hand

                    ' 8. ზედა მუქი ზოლი (header)
                    Dim HEADER_HEIGHT As Integer = 24
                    Dim headerPanel As New Panel()
                    headerPanel.Size = New Size(sessionCard.Width, HEADER_HEIGHT)
                    headerPanel.Location = New Point(0, 0)
                    headerPanel.BackColor = headerColor
                    sessionCard.Controls.Add(headerPanel)

                    ' 9. ბენეფიციარის სახელი header-ში მთავრული ფონტით (ჯერ გვარი, მერე სახელი)
                    Dim mtavruliFont As String = "Sylfaen" ' ნაგულისხმევი ფონტი
                    Try
                        Using testFont As New Font("KA_LITERATURULI_MT", 8)
                            mtavruliFont = "KA_LITERATURULI_MT"
                        End Using
                    Catch
                        ' ალტერნატიული ფონტები
                        Dim altFonts As String() = {"BPG_Nino_Mtavruli", "ALK_Tall_Mtavruli", "Sylfaen"}
                        For Each fontName In altFonts
                            Try
                                Using testFont As New Font(fontName, 8)
                                    mtavruliFont = fontName
                                    Exit For
                                End Using
                            Catch
                                Continue For
                            End Try
                        Next
                    End Try

                    Dim lblBeneficiary As New Label()
                    ' ცვლილება: ჯერ გვარი, მერე სახელი
                    Dim beneficiaryText As String = $"{session.BeneficiarySurname.ToUpper()} {session.BeneficiaryName.ToUpper()}"
                    lblBeneficiary.Text = beneficiaryText
                    lblBeneficiary.Location = New Point(8, 2)
                    lblBeneficiary.Size = New Size(headerPanel.Width - 16, 20)
                    lblBeneficiary.Font = New Font(mtavruliFont, 8, FontStyle.Bold)
                    lblBeneficiary.ForeColor = Color.White
                    lblBeneficiary.TextAlign = ContentAlignment.MiddleLeft
                    headerPanel.Controls.Add(lblBeneficiary)

                    ' 10. კასტომ ჩარჩოს დამატება
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

                    ' 11. ლეიბლების დინამიური განთავსება (ოდნავ გაზრდილი პადინგი)
                    Dim currentY As Integer = HEADER_HEIGHT + 5 ' 3-დან 5-ზე დავაბრუნე
                    Dim labelSpacing As Integer = 12 ' 10-დან 12-ზე გავაზრდე

                    ' სტატუსი (კომპაქტური, მაგრამ ხილვადი)
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

                    ' თერაპევტი (თუ ბარათი საკმარისად დიდია)
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

                    ' თერაპიის ტიპი (თუ ბარათი ძალიან დიდია)
                    If cardHeight > HEADER_HEIGHT + 50 AndAlso Not String.IsNullOrEmpty(session.TherapyType) Then
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

                    ' 12. რედაქტირების ღილაკი ბარათის ქვედა მარჯვენა კუთხეში
                    Dim btnEdit As New Button()
                    btnEdit.Text = "✎" ' ფანქრის სიმბოლო
                    btnEdit.Font = New Font("Segoe UI Symbol", 10, FontStyle.Bold)
                    btnEdit.ForeColor = Color.White
                    btnEdit.BackColor = Color.FromArgb(80, 80, 80) ' მუქი ფონი
                    btnEdit.Size = New Size(24, 24) ' მრგვალი ღილაკისთვის
                    btnEdit.Location = New Point(sessionCard.Width - 30, sessionCard.Height - 30)
                    btnEdit.FlatStyle = FlatStyle.Flat
                    btnEdit.FlatAppearance.BorderSize = 0
                    btnEdit.Tag = session.Id ' შევინახოთ სესიის ID ღილაკის Tag-ში
                    btnEdit.Cursor = Cursors.Hand

                    ' მოვუმრგვალოთ ღილაკი
                    Dim btnPath As New Drawing2D.GraphicsPath()
                    btnPath.AddEllipse(0, 0, btnEdit.Width, btnEdit.Height)
                    btnEdit.Region = New Region(btnPath)

                    ' ღილაკის მიბმა ფუნქციაზე
                    AddHandler btnEdit.Click, AddressOf BtnEditSession_Click

                    ' ღილაკის დამატება ბარათზე
                    sessionCard.Controls.Add(btnEdit)
                    btnEdit.BringToFront()

                    ' 13. მომრგვალებული კუთხეები
                    Try
                        Dim path As New Drawing2D.GraphicsPath()
                        Dim cornerRadius As Integer = 6
                        If sessionCard.Width > cornerRadius * 2 AndAlso sessionCard.Height > cornerRadius * 2 Then
                            path.AddArc(0, 0, cornerRadius * 2, cornerRadius * 2, 180, 90)
                            path.AddArc(sessionCard.Width - cornerRadius * 2, 0, cornerRadius * 2, cornerRadius * 2, 270, 90)
                            path.AddArc(sessionCard.Width - cornerRadius * 2, sessionCard.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 0, 90)
                            path.AddArc(0, sessionCard.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 90, 90)
                            path.CloseFigure()
                            sessionCard.Region = New Region(path)
                        End If
                    Catch
                        ' Region შექმნის პრობლემისას გაგრძელება
                    End Try

                    ' 14. ივენთების მიბმა
                    AddHandler sessionCard.Click, AddressOf SessionCard_Click
                    AddHandler sessionCard.DoubleClick, AddressOf SessionCard_DoubleClick

                    ' 15. ბარათის mainGridPanel-ზე დამატება
                    mainGridPanel.Controls.Add(sessionCard)
                    sessionCard.BringToFront()

                    Debug.WriteLine($"✅ ბარათი განთავსდა: ID={session.Id}, სტატუსი='{session.Status}', " &
                             $"ზომა={sessionCard.Width}x{sessionCard.Height}px, ფერი={cardColor}")

                Catch sessionEx As Exception
                    Debug.WriteLine($"❌ შეცდომა სესია ID={session.Id}: {sessionEx.Message}")
                    Continue For
                End Try
            Next

            Debug.WriteLine("=== PlaceSessionsOnGrid: დასრულება ===")

        Catch ex As Exception
            Debug.WriteLine($"❌ PlaceSessionsOnGrid: ზოგადი შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის ბარათზე დაჭერის ივენთი
    ''' </summary>
    Private Sub SessionCard_Click(sender As Object, e As EventArgs)
        Try
            ' მივიღოთ დაჭერილი ბარათი
            Dim sessionCard As Panel = DirectCast(sender, Panel)

            ' მივიღოთ სესიის ID
            Dim sessionId As Integer = CInt(sessionCard.Tag)

            ' ვიპოვოთ სესია ID-ით
            Dim session = allSessions.FirstOrDefault(Function(s) s.Id = sessionId)

            If session IsNot Nothing Then
                ' გამოვაჩინოთ სესიის დეტალები MessageBox-ში
                Dim sb As New StringBuilder()
                sb.AppendLine($"სესიის ინფორმაცია (ID: {session.Id})")
                sb.AppendLine("----------------------------")
                sb.AppendLine($"ბენეფიციარი: {session.BeneficiaryName} {session.BeneficiarySurname}")
                sb.AppendLine($"თარიღი: {session.FormattedDateTime}")
                sb.AppendLine($"ხანგრძლივობა: {session.Duration} წუთი")
                sb.AppendLine($"თერაპევტი: {session.TherapistName}")
                sb.AppendLine($"თერაპია: {session.TherapyType}")
                sb.AppendLine($"სივრცე: {session.Space}")
                sb.AppendLine($"სტატუსი: {session.Status}")

                MessageBox.Show(sb.ToString(), "სესიის დეტალები", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            Debug.WriteLine($"SessionCard_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' რედაქტირების ღილაკზე დაჭერის დამუშავება
    ''' </summary>
    Private Sub BtnEditSession_Click(sender As Object, e As EventArgs)
        Try
            Dim btn = DirectCast(sender, Button)
            Dim sessionId As Integer = CInt(btn.Tag)

            Debug.WriteLine($"BtnEditSession_Click: სესიის რედაქტირების მოთხოვნა, ID={sessionId}")

            ' შევამოწმოთ არის თუ არა მომხმარებელი ავტორიზებული რედაქტირებისთვის
            If String.IsNullOrEmpty(userEmail) Then
                MessageBox.Show("ავტორიზაციის გარეშე ვერ მოხდება სესიის რედაქტირება", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ხელმისაწვდომი", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' შევქმნათ და გავხსნათ რედაქტირების ფორმა
            Dim editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_Calendar")

            ' გავხსნათ ფორმა
            Dim result = editForm.ShowDialog()

            ' თუ ფორმა დაიხურა OK რეზულტატით, განვაახლოთ მონაცემები
            If result = DialogResult.OK Then
                Debug.WriteLine($"BtnEditSession_Click: სესია ID={sessionId} წარმატებით განახლდა")

                ' ჩავტვირთოთ სესიები თავიდან
                LoadSessions()

                ' განვაახლოთ კალენდრის ხედი
                UpdateCalendarView()
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnEditSession_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიის რედაქტირების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
''' <summary>
''' სტატუსების ფერების კონფიგურაცია - ცენტრალური ადგილი ყველა სტატუსის ფერისთვის
''' გამოიყენება როგორც კალენდარში, ასევე სხვა ადგილებში სადაც სესიის ბარათები ჩანს
''' </summary>
Public Class SessionStatusColors

    ''' <summary>
    ''' სტატუსის მიხედვით ფერის დაბრუნება
    ''' </summary>
    ''' <param name="status">სესიის სტატუსი</param>
    ''' <param name="sessionDateTime">სესიის თარიღი და დრო (ვადაგადაცილების შემოწმებისთვის)</param>
    ''' <returns>შესაბამისი ფერი</returns>
    Public Shared Function GetStatusColor(status As String, sessionDateTime As DateTime) As Color
        ' სტატუსის დანორმალიზება
        Dim normalizedStatus As String = status.Trim().ToLower()

        ' დაგეგმილი სესიებისთვის ვამოწმებთ ვადაგადაცილებას
        If normalizedStatus = "დაგეგმილი" Then
            ' თუ ვადა გავიდა - ვარდისფერი
            If sessionDateTime < DateTime.Now Then
                Return OverdueColor
            Else
                ' თუ ვადა ჯერ არ მოსულა - თეთრი
                Return PlannedColor
            End If
        End If

        ' დანარჩენი სტატუსები
        Select Case normalizedStatus
            Case "შესრულებული"
                Return CompletedColor
            Case "გაუქმებული", "გაუქმება"
                Return CancelledColor
            Case "პროგრამით გატარება"
                Return AutoProcessedColor
            Case "აღდგენა"
                Return RestoredColor
            Case "გაცდენა არასაპატიო"
                Return MissedUnexcusedColor
            Case "გაცდენა საპატიო"
                Return MissedExcusedColor
            Case "შესრულების პროცესში"
                Return InProgressColor
            Case Else
                ' უცნობი სტატუსისთვის ნაგულისხმევი ფერი
                Return DefaultColor
        End Select
    End Function

    ''' <summary>
    ''' ღია მწვანე - შესრულებული სესიებისთვის
    ''' </summary>
    Public Shared ReadOnly Property CompletedColor As Color
        Get
            Return Color.FromArgb(200, 255, 200) ' ღია მწვანე
        End Get
    End Property

    ''' <summary>
    ''' თეთრი - დაგეგმილი სესიებისთვის (ვადა ჯერ არ მოსულა)
    ''' </summary>
    Public Shared ReadOnly Property PlannedColor As Color
        Get
            Return Color.FromArgb(255, 255, 255) ' თეთრი
        End Get
    End Property

    ''' <summary>
    ''' ვარდისფერი - ვადაგადაცილებული დაგეგმილი სესიებისთვის
    ''' </summary>
    Public Shared ReadOnly Property OverdueColor As Color
        Get
            Return Color.FromArgb(255, 182, 193) ' ღია ვარდისფერი
        End Get
    End Property

    ''' <summary>
    ''' წითელი - გაუქმებული სესიებისთვის
    ''' </summary>
    Public Shared ReadOnly Property CancelledColor As Color
        Get
            Return Color.FromArgb(255, 150, 150) ' წითელი
        End Get
    End Property

    ''' <summary>
    ''' ნაცრისფერი - პროგრამით გატარებული სესიებისთვის
    ''' </summary>
    Public Shared ReadOnly Property AutoProcessedColor As Color
        Get
            Return Color.FromArgb(220, 220, 220) ' ნაცრისფერი
        End Get
    End Property

    ''' <summary>
    ''' უფრო ღია მწვანე - აღდგენილი სესიებისთვის
    ''' </summary>
    Public Shared ReadOnly Property RestoredColor As Color
        Get
            Return Color.FromArgb(170, 255, 170) ' უფრო ღია მწვანე
        End Get
    End Property

    ''' <summary>
    ''' ღია იასამნისფერი - არასაპატიო გაცდენისთვის
    ''' </summary>
    Public Shared ReadOnly Property MissedUnexcusedColor As Color
        Get
            Return Color.FromArgb(230, 230, 255) ' ღია იასამნისფერი
        End Get
    End Property

    ''' <summary>
    ''' ყვითელი - საპატიო გაცდენისთვის
    ''' </summary>
    Public Shared ReadOnly Property MissedExcusedColor As Color
        Get
            Return Color.FromArgb(255, 255, 200) ' ღია ყვითელი
        End Get
    End Property

    ''' <summary>
    ''' ღია ლურჯი - შესრულების პროცესშია
    ''' </summary>
    Public Shared ReadOnly Property InProgressColor As Color
        Get
            Return Color.FromArgb(200, 230, 255) ' ღია ლურჯი
        End Get
    End Property

    ''' <summary>
    ''' ნაგულისხმევი ფერი უცნობი სტატუსებისთვის
    ''' </summary>
    Public Shared ReadOnly Property DefaultColor As Color
        Get
            Return Color.FromArgb(240, 240, 240) ' ღია ნაცრისფერი
        End Get
    End Property

    ''' <summary>
    ''' ყველა შესაძლო სტატუსის სია მათი ფერებით
    ''' გამოიყენება ლეგენდისა და რეპორტებისთვის
    ''' </summary>
    Public Shared Function GetAllStatusColorsForLegend() As Dictionary(Of String, Color)
        Dim statusColors As New Dictionary(Of String, Color)

        statusColors.Add("შესრულებული", CompletedColor)
        statusColors.Add("დაგეგმილი (ვადაში)", PlannedColor)
        statusColors.Add("დაგეგმილი (ვადაგადაცილებული)", OverdueColor)
        statusColors.Add("გაუქმებული", CancelledColor)
        statusColors.Add("პროგრამით გატარება", AutoProcessedColor)
        statusColors.Add("აღდგენა", RestoredColor)
        statusColors.Add("გაცდენა არასაპატიო", MissedUnexcusedColor)
        statusColors.Add("გაცდენა საპატიო", MissedExcusedColor)
        statusColors.Add("შესრულების პროცესში", InProgressColor)

        Return statusColors
    End Function

    ''' <summary>
    ''' ბარათის ჩარჩოს ფერის დაბრუნება სტატუსის მიხედვით
    ''' ზოგიერთი ბარათისთვის შეიძლება სხვადასხვა ჩარჩო გვჭირდებოდეს
    ''' </summary>
    Public Shared Function GetStatusBorderColor(status As String, sessionDateTime As DateTime) As Color
        ' ძირითადი ფერი
        Dim mainColor As Color = GetStatusColor(status, sessionDateTime)

        ' ჩარჩო ოდნავ უფრო მუქია
        Return Color.FromArgb(
            Math.Max(0, mainColor.R - 30),
            Math.Max(0, mainColor.G - 30),
            Math.Max(0, mainColor.B - 30)
        )
    End Function
End Class