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

    ' კლასის დონეზე ვამატებთ თერაპევტების სიას
    Private therapists As List(Of String) = New List(Of String)()

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

        ' ✨ ახალი: ფილტრების ინიციალიზაცია
        InitializeStatusFilters()

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
    ''' ჩეკბოქსების ფილტრაციის ინიციალიზაცია
    ''' ჩეკბოქსები უკვე შექმნილია დიზაინერში, მხოლოდ ივენთების მიბმა ხდება
    ''' </summary>
    Private Sub InitializeStatusFilters()
        Try
            Debug.WriteLine("InitializeStatusFilters: ჩეკბოქსების ინიციალიზაცია")

            ' ყველა ჩეკბოქსი საწყისად ჩართული უნდა იყოს
            CheckBox1.Checked = True  ' დაგეგმილი
            CheckBox2.Checked = True  ' შესრულებული
            CheckBox3.Checked = True  ' გაცდენა საპატიო
            CheckBox4.Checked = True  ' გაცდენა არასაპატიო
            CheckBox5.Checked = True  ' აღდგენა
            CheckBox6.Checked = True  ' პროგრამით გატარება
            CheckBox7.Checked = True  ' გაუქმებული

            ' ივენთების მიბმა - CheckedChanged ივენთებზე
            AddHandler CheckBox1.CheckedChanged, AddressOf StatusFilter_CheckedChanged
            AddHandler CheckBox2.CheckedChanged, AddressOf StatusFilter_CheckedChanged
            AddHandler CheckBox3.CheckedChanged, AddressOf StatusFilter_CheckedChanged
            AddHandler CheckBox4.CheckedChanged, AddressOf StatusFilter_CheckedChanged
            AddHandler CheckBox5.CheckedChanged, AddressOf StatusFilter_CheckedChanged
            AddHandler CheckBox6.CheckedChanged, AddressOf StatusFilter_CheckedChanged
            AddHandler CheckBox7.CheckedChanged, AddressOf StatusFilter_CheckedChanged

            Debug.WriteLine("InitializeStatusFilters: ყველა ჩეკბოქსი ჩართულია და ივენთები მიბმულია")
        Catch ex As Exception
            Debug.WriteLine($"InitializeStatusFilters: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ჩეკბოქსის ცვლილების ივენთი - სტატუსის ფილტრაციისთვის
    ''' </summary>
    Private Sub StatusFilter_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Dim checkbox As CheckBox = DirectCast(sender, CheckBox)
            Debug.WriteLine($"StatusFilter_CheckedChanged: {checkbox.Name} = {checkbox.Checked}")

            ' კალენდრის ხედის განახლება ფილტრაციით
            UpdateCalendarView()
        Catch ex As Exception
            Debug.WriteLine($"StatusFilter_CheckedChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფილტრაციისთვის შესაბამისი სესიების მიღება სტატუსის მიხედვით
    ''' </summary>
    Private Function GetFilteredSessions() As List(Of SessionModel)
        Try
            Debug.WriteLine("GetFilteredSessions: სესიების ფილტრაცია")

            ' არჩეული თარიღი
            Dim selectedDate As DateTime = DTPCalendar.Value.Date

            ' მხოლოდ არჩეული დღის სესიები
            Dim daySessions = allSessions.Where(Function(s) s.DateTime.Date = selectedDate).ToList()

            ' ფილტრაცია სტატუსების მიხედვით
            Dim filteredSessions As New List(Of SessionModel)()

            For Each session In daySessions
                Dim shouldShow As Boolean = False
                Dim sessionStatus As String = session.Status.Trim().ToLower()

                ' CheckBox1 - დაგეგმილი
                If CheckBox1.Checked AndAlso sessionStatus = "დაგეგმილი" Then
                    shouldShow = True
                End If

                ' CheckBox2 - შესრულებული
                If CheckBox2.Checked AndAlso sessionStatus = "შესრულებული" Then
                    shouldShow = True
                End If

                ' CheckBox3 - გაცდენა საპატიო
                If CheckBox3.Checked AndAlso sessionStatus = "გაცდენა საპატიო" Then
                    shouldShow = True
                End If

                ' CheckBox4 - გაცდენა არასაპატიო
                If CheckBox4.Checked AndAlso sessionStatus = "გაცდენა არასაპატიო" Then
                    shouldShow = True
                End If

                ' CheckBox5 - აღდგენა
                If CheckBox5.Checked AndAlso sessionStatus = "აღდგენა" Then
                    shouldShow = True
                End If

                ' CheckBox6 - პროგრამით გატარება
                If CheckBox6.Checked AndAlso sessionStatus = "პროგრამით გატარება" Then
                    shouldShow = True
                End If

                ' CheckBox7 - გაუქმებული
                If CheckBox7.Checked AndAlso (sessionStatus = "გაუქმებული" OrElse sessionStatus = "გაუქმება") Then
                    shouldShow = True
                End If

                ' თუ ჩეკბოქსი ჩართულია, დავამატოთ სესია
                If shouldShow Then
                    filteredSessions.Add(session)
                End If
            Next

            Debug.WriteLine($"GetFilteredSessions: {daySessions.Count}-დან ფილტრაციის შემდეგ დარჩა {filteredSessions.Count} სესია")
            Return filteredSessions

        Catch ex As Exception
            Debug.WriteLine($"GetFilteredSessions: შეცდომა - {ex.Message}")
            Return New List(Of SessionModel)()
        End Try
    End Function

    ' ═══════════════════════════════════════
    ' 🔧 მოდიფიკაცია: კონსტრუქტორში ფილტრების ინიციალიზაცია
    ' ═══════════════════════════════════════

    ''' <summary>
    ''' კონსტრუქტორი კალენდრის ViewModel-ით (განახლებული ვერსია ფილტრებით)
    ''' ჩაამატე ეს ხაზი არსებულ კონსტრუქტორში, სადაც სხვა ივენთებია მიბმული
    ''' </summary>
    Private Sub InitializeConstructorWithFilters()
        ' ... შენი არსებული კოდი ...

        ' რადიობუტონების საწყისი მნიშვნელობები
        rbDay.Checked = True
        RBSpace.Checked = True

        ' ✨ ახალი: ფილტრების ინიციალიზაცია
        InitializeStatusFilters()
    End Sub

    ''' <summary>
    ''' თერაპევტების ჩატვირთვა მონაცემთა წყაროდან არჩეული თარიღისთვის
    ''' მხოლოდ ის თერაპევტები, რომლებიც მოცემულ თარიღში იქნებიან დაკავებულები
    ''' </summary>
    Private Sub LoadTherapistsForDate()
        Try
            Debug.WriteLine("LoadTherapistsForDate: დაიწყო თერაპევტების ჩატვირთვა")

            ' გავასუფთავოთ თერაპევტების სია
            therapists.Clear()

            ' არჩეული თარიღი
            Dim selectedDate As DateTime = DTPCalendar.Value.Date
            Debug.WriteLine($"LoadTherapistsForDate: არჩეული თარიღი: {selectedDate:dd.MM.yyyy}")

            ' שևამოწმოთ მონაცემთა სერვისი
            If dataService Is Nothing Then
                Debug.WriteLine("LoadTherapistsForDate: მონაცემთა სერვისი არ არის ინიციალიზებული")
                ' საცდელი თერაპევტების ჩამატება ტესტისთვის
                therapists.AddRange({"ჩანაწერი არ არის"})
                Return
            End If

            ' თუ სესიები არ არის ჩატვირთული, ჩავტვირთოთ
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                LoadSessions()
            End If

            ' ფილტრაცია: მხოლოდ არჩეული დღის სესიები
            Dim daySessions = allSessions.Where(Function(s) s.DateTime.Date = selectedDate).ToList()
            Debug.WriteLine($"LoadTherapistsForDate: ნაპოვნია {daySessions.Count} სესია {selectedDate:dd.MM.yyyy} თარიღისთვის")

            ' შევკრიბოთ უნიკალური თერაპევტები არჩეული დღისთვის
            Dim therapistSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each session In daySessions
                If Not String.IsNullOrWhiteSpace(session.TherapistName) Then
                    ' ვამატებთ მხოლოდ არაცარიელ თერაპევტებს
                    Dim therapistName As String = session.TherapistName.Trim()
                    therapistSet.Add(therapistName)
                    Debug.WriteLine($"LoadTherapistsForDate: დაემატა თერაპევტი: '{therapistName}'")
                End If
            Next

            ' ვარეთავთ თერაპევტებს სახელოვნად
            therapists.AddRange(therapistSet.OrderBy(Function(t) t))

            ' თუ არ მოიძებნა თერაპევტები, დავამატოთ საცდელი
            If therapists.Count = 0 Then
                Debug.WriteLine("LoadTherapistsForDate: თერაპევტები არ მოიძებნა, ვამატებთ საცდელ მონაცემებს")
                therapists.AddRange({"ჩანაწერი არ არის"})
            End If

            Debug.WriteLine($"LoadTherapistsForDate: ჩატვირთულია {therapists.Count} თერაპევტი")
            For i As Integer = 0 To therapists.Count - 1
                Debug.WriteLine($"  [{i}] {therapists(i)}")
            Next

        Catch ex As Exception
            Debug.WriteLine($"LoadTherapistsForDate: შეცდომა - {ex.Message}")

            ' შეცდომის შემთხვევაში გვაუყენოთ საცდელი თერაპევტები
            therapists.Clear()
            therapists.AddRange({"ჩანაწერი არ არის"})
        End Try
    End Sub

    ''' <summary>
    ''' სქროლის სინქრონიზაციის განახლება - გასწორებული ვერსია თერაპევტების მხარდაჭერით
    ''' (ეს ჩანაცვლდება ან დაემატება არსებულ SetupScrollSynchronization მეთოდს)
    ''' </summary>
    Private Sub SetupScrollSynchronizationForBothGrids()
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("SetupScrollSynchronizationForBothGrids: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' ჯერ მოვხსნათ არსებული ივენთის ჰენდლერები, თუ ისინი არსებობენ
            RemoveHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll
            RemoveHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            ' მივაბათ Scroll ივენთი ხელახლა (ახლა მუშაობს როგორც სივრცეებისთვის, ასევე თერაპევტებისთვის)
            AddHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' ასევე, დავამატოთ ივენთი გორგოლჭის სქროლისთვის
            AddHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            Debug.WriteLine("SetupScrollSynchronizationForBothGrids: სქროლის სინქრონიზაცია დაყენებულია სივრცეების და თერაპევტების გრიდებისთვის")

        Catch ex As Exception
            Debug.WriteLine($"SetupScrollSynchronizationForBothGrids: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კალენდრის პანელების ზომების განახლება მასშტაბის შეცვლისას - მხარდაჭერა თერაპევტებისთვის
    ''' </summary>
    Private Sub UpdatePanelSizesForCurrentGrid()
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            ' შევამოწმოთ, რომელი სათაურების პანელია გამოყენებული
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)
            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("UpdatePanelSizesForCurrentGrid: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' სივრცეების ხედისას
            If spacesHeaderPanel IsNot Nothing AndAlso spacesHeaderPanel.Visible Then
                ' გამოვთვალოთ სივრცეების სათაურების პანელის სრული სიგანე
                Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
                Dim totalSpacesWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count

                ' განვაახლოთ სივრცეების სათაურების პანელის სიგანე
                spacesHeaderPanel.Width = totalSpacesWidth

                ' სკროლის პოზიცია
                Dim scrollPosition As Point = mainGridPanel.AutoScrollPosition
                Dim actualHScrollPos As Integer = Math.Abs(scrollPosition.X)

                ' ვაყენებთ spacesHeaderPanel-ის პოზიციას ჰორიზონტალური სკროლის გათვალისწინებით
                spacesHeaderPanel.Location = New Point(BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH - actualHScrollPos, 0)

                Debug.WriteLine($"UpdatePanelSizesForCurrentGrid: განახლდა სივრცეების პანელის ზომები, Width={totalSpacesWidth}")
            End If

            ' თერაპევტების ხედისას
            If therapistsHeaderPanel IsNot Nothing AndAlso therapistsHeaderPanel.Visible Then
                ' გამოვთვალოთ თერაპევტების სათაურების პანელის სრული სიგანე
                Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
                Dim totalTherapistsWidth As Integer = THERAPIST_COLUMN_WIDTH * therapists.Count

                ' განვაახლოთ თერაპევტების სათაურების პანელის სიგანე
                therapistsHeaderPanel.Width = totalTherapistsWidth

                ' სკროლის პოზიცია
                Dim scrollPosition As Point = mainGridPanel.AutoScrollPosition
                Dim actualHScrollPos As Integer = Math.Abs(scrollPosition.X)

                ' ვაყენებთ therapistsHeaderPanel-ის პოზიციას ჰორიზონტალური სკროლის გათვალისწინებით
                therapistsHeaderPanel.Location = New Point(BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH - actualHScrollPos, 0)

                Debug.WriteLine($"UpdatePanelSizesForCurrentGrid: განახლდა თერაპევტების პანელის ზომები, Width={totalTherapistsWidth}")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UpdatePanelSizesForCurrentGrid: შეცდომა - {ex.Message}")
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
    ''' სქროლის სინქრონიზაცია - ᲛᲗᲚᲘᲐᲜᲐᲓ გასწორებული ვერსია
    ''' ეს მეთოდი ჩანაცვლდება არსებული SetupScrollSynchronization მეთოდი
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

            ' ვიპოვოთ თერაპევტების სათაურების პანელი
            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)

            ' ჯერ მოვხსნათ არსებული ივენთის ჰენდლერები, თუ ისინი არსებობენ
            RemoveHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' მივაბათ Scroll ივენთი ხელახლა
            AddHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' ასევე, დავამატოთ ივენთი გორგოლჭის სქროლისთვის
            AddHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            Debug.WriteLine("SetupScrollSynchronization: სქროლის სინქრონიზაცია დაყენებულია")
            Debug.WriteLine($"  spacesHeaderPanel: {If(spacesHeaderPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")
            Debug.WriteLine($"  therapistsHeaderPanel: {If(therapistsHeaderPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")

        Catch ex As Exception
            Debug.WriteLine($"SetupScrollSynchronization: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის სქროლის ივენთი 
    ''' 🔧 ᲛᲗᲚᲘᲐᲜᲐᲓ გასწორებული ვერსია - მუშაობს ყველა გრიდისთვის
    ''' ეს მეთოდი ჩანაცვლდება არსებული MainGridPanel_Scroll მეთოდი
    ''' </summary>
    Private Sub MainGridPanel_Scroll(sender As Object, e As ScrollEventArgs)
        Try
            Dim mainGridPanel As Panel = DirectCast(sender, Panel)

            ' ვიპოვოთ ყველა საჭირო პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)
            Dim dateColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("dateColumnPanel", False).FirstOrDefault(), Panel)

            ' 🔧 ახალი: შევამოწმოთ ყველა სათაურების პანელი
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)
            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)

            ' ვერტიკალური სქროლი
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

            ' ჰორიზონტალური სქროლი
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                Dim scrollOffset As Integer = -mainGridPanel.HorizontalScroll.Value

                ' 🔧 მნიშვნელოვანი: ვინახავთ ფიქსირებულ მარცხენა პოზიციას
                Dim fixedLeftPosition As Integer = BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH

                ' 🔧 სივრცეების სათაურები (თუ არსებობს და ხილვადია)
                If spacesHeaderPanel IsNot Nothing AndAlso spacesHeaderPanel.Visible Then
                    spacesHeaderPanel.Left = fixedLeftPosition + scrollOffset
                    Debug.WriteLine($"MainGridPanel_Scroll: სივრცეების სათაურები გადაადგილდა Left={spacesHeaderPanel.Left}")
                End If

                ' 🔧 ᲐᲮᲐᲚᲘ: თერაპევტების სათაურები (თუ არსებობს და ხილვადია)
                If therapistsHeaderPanel IsNot Nothing AndAlso therapistsHeaderPanel.Visible Then
                    therapistsHeaderPanel.Left = fixedLeftPosition + scrollOffset
                    Debug.WriteLine($"MainGridPanel_Scroll: თერაპევტების სათაურები გადაადგილდა Left={therapistsHeaderPanel.Left}")
                End If

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
    ''' მაუსის გორგოლჭის ივენთი - 🔧 ᲛᲗᲚᲘᲐᲜᲐᲓ გასწორებული ვერსია
    ''' ეს მეთოდი ჩანაცვლდება არსებული MainGridPanel_MouseWheel მეთოდი
    ''' </summary>
    Private Sub MainGridPanel_MouseWheel(sender As Object, e As MouseEventArgs)
        Try
            Dim mainGridPanel As Panel = DirectCast(sender, Panel)

            ' ვიპოვოთ საჭირო პანელები
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)
            Dim dateColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("dateColumnPanel", False).FirstOrDefault(), Panel)

            ' 🔧 ახალი: შევამოწმოთ ყველა სათაურების პანელი
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)
            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)

            ' 🔧 ვერტიკალური სქროლის სინქრონიზაცია (მუშაობს კარგად)
            Dim verticalScrollOffset As Integer = -mainGridPanel.VerticalScroll.Value

            If timeColumnPanel IsNot Nothing Then
                timeColumnPanel.Top = BASE_HEADER_HEIGHT + verticalScrollOffset
            End If

            If dateColumnPanel IsNot Nothing Then
                dateColumnPanel.Top = BASE_HEADER_HEIGHT + verticalScrollOffset
            End If

            ' 🔧 ჰორიზონტალური სქროლის სინქრონიზაცია
            Dim horizontalScrollOffset As Integer = -mainGridPanel.HorizontalScroll.Value
            Dim fixedLeftPosition As Integer = BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH

            ' 🔧 სივრცეების სათაურები (თუ არსებობს და ხილვადია)
            If spacesHeaderPanel IsNot Nothing AndAlso spacesHeaderPanel.Visible Then
                spacesHeaderPanel.Left = fixedLeftPosition + horizontalScrollOffset
            End If

            ' 🔧 ᲐᲮᲐᲚᲘ: თერაპევტების სათაურები (თუ არსებობს და ხილვადია)
            If therapistsHeaderPanel IsNot Nothing AndAlso therapistsHeaderPanel.Visible Then
                therapistsHeaderPanel.Left = fixedLeftPosition + horizontalScrollOffset
            End If

        Catch ex As Exception
            Debug.WriteLine($"MainGridPanel_MouseWheel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 დიაგნოსტიკური მეთოდი - შეამოწმეთ სქროლის მუშაობა
    ''' გამოიძახეთ ეს მეთოდი, როდესაც პრობლემაა სქროლთან
    ''' </summary>
    Public Sub DiagnoseScrollIssue()
        Try
            Debug.WriteLine("=== DiagnoseScrollIssue: დაიწყო ===")

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            Debug.WriteLine($"mainGridPanel: {If(mainGridPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")

            If mainGridPanel IsNot Nothing Then
                Debug.WriteLine($"  ზომა: {mainGridPanel.Size}")
                Debug.WriteLine($"  სქროლის მნიშვნელობები: H={mainGridPanel.HorizontalScroll.Value}, V={mainGridPanel.VerticalScroll.Value}")
                Debug.WriteLine($"  AutoScrollMinSize: {mainGridPanel.AutoScrollMinSize}")
            End If

            ' ვიპოვოთ სათაურების პანელები
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)
            Debug.WriteLine($"spacesHeaderPanel: {If(spacesHeaderPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")
            If spacesHeaderPanel IsNot Nothing Then
                Debug.WriteLine($"  ზომა: {spacesHeaderPanel.Size}, ადგილმდებარეობა: {spacesHeaderPanel.Location}, ხილვადი: {spacesHeaderPanel.Visible}")
            End If

            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)
            Debug.WriteLine($"therapistsHeaderPanel: {If(therapistsHeaderPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")
            If therapistsHeaderPanel IsNot Nothing Then
                Debug.WriteLine($"  ზომა: {therapistsHeaderPanel.Size}, ადგილმდებარეობა: {therapistsHeaderPanel.Location}, ხილვადი: {therapistsHeaderPanel.Visible}")
            End If

            ' ვიპოვოთ დროის პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)
            Debug.WriteLine($"timeColumnPanel: {If(timeColumnPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")

            Debug.WriteLine("=== DiagnoseScrollIssue: დასრულდა ===")

        Catch ex As Exception
            Debug.WriteLine($"DiagnoseScrollIssue: შეცდომა - {ex.Message}")
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
    ''' დღის ხედის ჩვენება თერაპევტების მიხედვით
    ''' იგივე ლოგიკა, რაც სივრცეებისთვის, მაგრამ therapists სიით
    ''' </summary>
    Private Sub ShowDayViewByTherapist()
        Try
            Debug.WriteLine("ShowDayViewByTherapist: დაიწყო დღის ხედის ჩვენება თერაპევტების მიხედვით")

            ' ======= 1. თერაპევტების ჩატვირთვა მოცემული თარიღისთვის =======
            LoadTherapistsForDate()

            ' დავაყენოთ pnlFilter-ის ფერი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' ======= 2. პანელების ინიციალიზაცია =======
            InitializeDayViewPanelsForTherapists()

            ' ======= 3. თარიღის ზოლის შევსება =======
            FillDateColumnPanel()

            ' ======= 4. დროის პანელის შევსება =======
            FillTimeColumnPanel()

            ' ======= 5. თერაპევტების სათაურების პანელის შევსება =======
            FillTherapistsHeaderPanel()

            ' ======= 6. მთავარი გრიდის პანელის შევსება თერაპევტების მიხედვით =======
            FillMainGridPanelForTherapists()

            ' ======= 7. სკროლის სინქრონიზაცია =======
            SetupScrollSynchronizationForBothGrids() ' ახალი მეთოდი, რომელიც მხარს უჭერს თერაპევტების გრიდსაც

            ' ======= 8. ძველი ბარათების გასუფთავება =======
            ClearSessionCardsFromGrid()

            ' ======= 9. სესიების განთავსება გრიდში თერაპევტების მიხედვით =======
            PlaceSessionsOnTherapistGrid()

            Debug.WriteLine("ShowDayViewByTherapist: დღის ხედის ჩვენება დასრულდა თერაპევტების მიხედვით")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewByTherapist: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewByTherapist: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"თერაპევტების ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დღის ხედის პანელების ინიციალიზაცია თერაპევტებისთვის
    ''' იგივე ლოგიკა, რაც სივრცეებისთვის, მაგრამ therapists-ის რაოდენობით
    ''' </summary>
    Private Sub InitializeDayViewPanelsForTherapists()
        Try
            Debug.WriteLine("InitializeDayViewPanelsForTherapists: დაიწყო პანელების ინიციალიზაცია תერაპევტებისთვის")

            ' გავასუფთავოთ კალენდრის პანელი
            pnlCalendarGrid.Controls.Clear()

            ' pnlCalendarGrid-ს არ ჰქონდეს AutoScroll!
            pnlCalendarGrid.AutoScroll = False

            ' როუ დროის ინტერვალები არ არის ინიციალიზებული, დავამატოთ
            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' თუ თერაპევტები არ არის ჩატვირთული, ჩავტვირთოთ
            If therapists.Count = 0 Then
                LoadTherapistsForDate()
            End If

            ' წავკითხოთ მშობელი პანელის ზომები
            Dim totalWidth As Integer = pnlCalendarGrid.ClientSize.Width
            Dim totalHeight As Integer = pnlCalendarGrid.ClientSize.Height

            Debug.WriteLine($"InitializeDayViewPanelsForTherapists: კონტეინერის ზომები - Width={totalWidth}, Height={totalHeight}")

            ' კონსტანტები, რომლებიც არ იცვლება მასშტაბირებისას
            Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT
            Dim TIME_COLUMN_WIDTH As Integer = BASE_TIME_COLUMN_WIDTH
            Dim DATE_COLUMN_WIDTH As Integer = BASE_DATE_COLUMN_WIDTH

            ' გამოვთვალოთ დროისა და თარიღის პანელების ზომები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim totalRowsHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            Debug.WriteLine($"InitializeDayViewPanelsForTherapists: ROW_HEIGHT={ROW_HEIGHT}, totalRowsHeight={totalRowsHeight}, vScale={vScale}")

            ' ======= 1. შევქმნათ თარიღის ზოლის პანელი =======
            Dim dateColumnPanel As New Panel()
            dateColumnPanel.Name = "dateColumnPanel"
            dateColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            dateColumnPanel.AutoScroll = False
            dateColumnPanel.Size = New Size(DATE_COLUMN_WIDTH, totalRowsHeight)
            dateColumnPanel.Location = New Point(0, HEADER_HEIGHT)
            dateColumnPanel.BackColor = Color.FromArgb(60, 80, 150)
            dateColumnPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(dateColumnPanel)

            ' ======= 2. שევქმნათ საათების პანელი =======
            Dim timeColumnPanel As New Panel()
            timeColumnPanel.Name = "timeColumnPanel"
            timeColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            timeColumnPanel.AutoScroll = False
            timeColumnPanel.Size = New Size(TIME_COLUMN_WIDTH, totalRowsHeight)
            timeColumnPanel.Location = New Point(DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            timeColumnPanel.BackColor = Color.FromArgb(240, 240, 245)
            timeColumnPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(timeColumnPanel)

            ' ======= 3. შევქმნათ თარიღის სათაურის პანელი =======
            Dim dateHeaderPanel As New Panel()
            dateHeaderPanel.Name = "dateHeaderPanel"
            dateHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            dateHeaderPanel.Size = New Size(DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            dateHeaderPanel.Location = New Point(0, 0)
            dateHeaderPanel.BackColor = Color.FromArgb(40, 60, 120)
            dateHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(dateHeaderPanel)

            ' თარიღის სათაურის ლეიבლი
            Dim dateHeaderLabel As New Label()
            dateHeaderLabel.Size = New Size(DATE_COLUMN_WIDTH - 2, HEADER_HEIGHT - 2)
            dateHeaderLabel.Location = New Point(1, 1)
            dateHeaderLabel.TextAlign = ContentAlignment.MiddleCenter
            dateHeaderLabel.Text = "თარიღი"
            dateHeaderLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            dateHeaderLabel.ForeColor = Color.White
            dateHeaderPanel.Controls.Add(dateHeaderLabel)

            ' ======= 4. შევქმნათ დროის სათაურის პანელი =======
            Dim timeHeaderPanel As New Panel()
            timeHeaderPanel.Name = "timeHeaderPanel"
            timeHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            timeHeaderPanel.Size = New Size(TIME_COLUMN_WIDTH, HEADER_HEIGHT)
            timeHeaderPanel.Location = New Point(DATE_COLUMN_WIDTH, 0)
            timeHeaderPanel.BackColor = Color.FromArgb(180, 180, 220)
            timeHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(timeHeaderPanel)

            ' דროის სათაურის ლეიბლი
            Dim timeHeaderLabel As New Label()
            timeHeaderLabel.Size = New Size(TIME_COLUMN_WIDTH - 2, HEADER_HEIGHT - 2)
            timeHeaderLabel.Location = New Point(1, 1)
            timeHeaderLabel.TextAlign = ContentAlignment.MiddleCenter
            timeHeaderLabel.Text = "დრო"
            timeHeaderLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            timeHeaderPanel.Controls.Add(timeHeaderLabel)

            ' ======= 5. შევქმნათ תერაპევტების სათაურების პანელი =======
            Dim therapistsHeaderPanel As New Panel()
            therapistsHeaderPanel.Name = "therapistsHeaderPanel"
            therapistsHeaderPanel.AutoScroll = False
            ' ზომა დამოკიდებული იქნება მასშტაბზე!
            Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim totalTherapistsWidth As Integer = THERAPIST_COLUMN_WIDTH * therapists.Count
            therapistsHeaderPanel.Size = New Size(totalTherapistsWidth, HEADER_HEIGHT)
            therapistsHeaderPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, 0)
            therapistsHeaderPanel.BackColor = Color.FromArgb(120, 180, 120) ' ოდნავ მწვანე ფერი თერაპევტებისთვის
            therapistsHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(therapistsHeaderPanel)

            ' ======= 6. შევქმнათ მთავარი גრიდის პანელი =======
            Dim mainGridPanel As New Panel()
            mainGridPanel.Name = "mainGridPanel"
            mainGridPanel.Size = New Size(totalWidth - TIME_COLUMN_WIDTH - DATE_COLUMN_WIDTH, totalHeight - HEADER_HEIGHT)
            mainGridPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            mainGridPanel.BackColor = Color.White
            mainGridPanel.BorderStyle = BorderStyle.FixedSingle
            mainGridPanel.AutoScroll = True
            pnlCalendarGrid.Controls.Add(mainGridPanel)

            Debug.WriteLine("InitializeDayViewPanelsForTherapists: პანელების ინიციალიზაცია دასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"InitializeDayViewPanelsForTherapists: שেცდომა - {ex.Message}")
            Debug.WriteLine($"InitializeDayViewPanelsForTherapists: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"תერაპევტების პანელების ინიციალიზაციის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' תერაპევтების სათაურების პანელის שევსება თერაპევტების სახელებით
    ''' </summary>
    Private Sub FillTherapistsHeaderPanel()
        Try
            Debug.WriteLine("FillTherapistsHeaderPanel: דაიწყო תერაპевტების სათაურების შევსება")

            ' ვიპოვოთ תერاპევტების סათაურების პანელი
            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)

            If therapistsHeaderPanel Is Nothing Then
                Debug.WriteLine("FillTherapistsHeaderPanel: תერაპევtების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' გავასუფთავოთ
            therapistsHeaderPanel.Controls.Clear()

            ' თუ תერაპევტები არ არის ჩატვირთული
            If therapists.Count = 0 Then
                LoadTherapistsForDate()
            End If

            Debug.WriteLine($"FillTherapistsHeaderPanel: ჩატვირთული תერაპევტების რაოდენობა: {therapists.Count}")

            ' თერაპევტის სვეტის სიგანე (იმასშტაბირდება)
            Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' גამოვთვალოთ סათაურების პანელის სრული სიგანე
            Dim totalHeaderWidth As Integer = THERAPIST_COLUMN_WIDTH * therapists.Count

            ' שევცვალოთ პანელის სიგანე მასშტაბის მიხედვით
            therapistsHeaderPanel.Width = totalHeaderWidth

            Debug.WriteLine($"FillTherapistsHeaderPanel: פנל-ის სიგანე: {therapistsHeaderPanel.Width}, THERAPIST_COLUMN_WIDTH: {THERAPIST_COLUMN_WIDTH}")

            ' ღია მწვანე ფონტი თერაპევتებისთვის
            Dim therapistFont As New Font("Sylfaen", 9, FontStyle.Bold)

            ' თერაპевტების סათაურების შექმნა
            For i As Integer = 0 To therapists.Count - 1
                Dim therapistHeader As New Label()
                therapistHeader.Size = New Size(THERAPIST_COLUMN_WIDTH - 1, therapistsHeaderPanel.Height - 2)
                therapistHeader.Location = New Point(i * THERAPIST_COLUMN_WIDTH, 1)
                therapistHeader.BackColor = Color.FromArgb(80, 120, 80) ' მუქი מწვანე
                therapistHeader.ForeColor = Color.White
                therapistHeader.TextAlign = ContentAlignment.MiddleCenter
                therapistHeader.Text = therapists(i)
                therapistHeader.Font = therapistFont
                therapistHeader.BorderStyle = BorderStyle.FixedSingle

                ' שעינها থাক თავდாყრო X პოზიცია Tag-ში სქროლის სინქრონიზაციისთვის
                therapistHeader.Tag = i * THERAPIST_COLUMN_WIDTH

                therapistsHeaderPanel.Controls.Add(therapistHeader)

                Debug.WriteLine($"FillTherapistsHeaderPanel: دамატა תერაפევტის სათაური [{i}]: {therapists(i)}, Tag={therapistHeader.Tag}")
            Next

            Debug.WriteLine($"FillTherapistsHeaderPanel: დაემატა {therapists.Count} თერাפევტის سათაური, მასштაბი: {hScale}")

        Catch ex As Exception
            Debug.WriteLine($"FillTherapistsHeaderPanel: שეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის შევსება תერაპევტებისთვის
    ''' </summary>
    Private Sub FillMainGridPanelForTherapists()
        Try
            Debug.WriteLine("FillMainGridPanelForTherapists: დაიწყოს מთავარი גრიდის شევსება תერაפევტებისთვის")

            ' ვიპოვოთ מთავარი گরიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("FillMainGridPanelForTherapists: მთავარი گرიდის پانელი ვერ მოიძებნა")
                Return
            End If

            ' გავासუფთავოთ მთავარი პანელი
            mainGridPanel.Controls.Clear()

            ' שामოწმოთ תერაპევტები და דრო
            If therapists.Count = 0 Then
                LoadTherapistsForDate()
            End If

            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' پארამეტრები - მასშטაბირების გამოყენებით
            Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)

            ' グриদის סრული სიგანе
            Dim gridWidth As Integer = THERAPIST_COLUMN_WIDTH * therapists.Count

            ' グריდის সრული সামაღლე
            Dim gridHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            ' მნიშვნელოვანი: დავაყენოთ mainGridPanel-ის AutoScrollMinSize
            mainGridPanel.AutoScrollMinSize = New Size(gridWidth, gridHeight)

            Debug.WriteLine($"FillMainGridPanelForTherapists: גრిდის ზომები - Width={gridWidth}, Height={gridHeight}, hScale={hScale}, vScale={vScale}")

            ' أуجრედების মাসিব (ახლა თერაपევტებისთვის)
            ReDim gridCells(therapists.Count - 1, timeIntervals.Count - 1)

            ' უজრედების शექმნა - პირდაపირ mainGridPanel-ზე
            For col As Integer = 0 To therapists.Count - 1
                For row As Integer = 0 To timeIntervals.Count - 1
                    Dim cell As New Panel()
                    cell.Size = New Size(THERAPIST_COLUMN_WIDTH, ROW_HEIGHT)
                    cell.Location = New Point(col * THERAPIST_COLUMN_WIDTH, row * ROW_HEIGHT)

                    ' אალטერნატიული ფერები
                    If row Mod 2 = 0 Then
                        cell.BackColor = Color.FromArgb(250, 255, 250) ' ოდნავ მწვანე ნაცარი
                    Else
                        cell.BackColor = Color.FromArgb(245, 250, 245) ' უფრო მწვანე ნაცარი
                    End If

                    ' უจრედის ছارჩო
                    AddHandler cell.Paint, AddressOf Cell_Paint

                    ' ვინახავთ უจრედის ობიექტس მასივში
                    gridCells(col, row) = cell

                    ' მნიშვნელოვანი: ვამატებთ უজრედს ფირდაপირ mainGridPanel-ზე
                    mainGridPanel.Controls.Add(cell)
                Next
            Next

            Debug.WriteLine($"FillMainGridPanelForTherapists: შეიქმნა {therapists.Count * timeIntervals.Count} უჯრედი תერაპევტებისთვის")

        Catch ex As Exception
            Debug.WriteLine($"FillMainGridPanelForTherapists: שეცდომა - {ex.Message}")
            Debug.WriteLine($"FillMainGridPanelForTherapists: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    Private Sub PlaceSessionsOnTherapistGrid()
        Try
            Debug.WriteLine("=== PlaceSessionsOnTherapistGrid: დაიწყო ===")

            ' בדীقות שემოწმებები
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnTherapistGrid: სესיები არ არის")
                Return
            End If

            If gridCells Is Nothing Then
                Debug.WriteLine("PlaceSessionsOnTherapistGrid: グრიდის უჯრედები არ არის")
                Return
            End If

            If timeIntervals.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnTherapistGrid: דროის ინტერვალები არ არის")
                Return
            End If

            If therapists.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnTherapistGrid: তেরাপევტები არ არის")
                Return
            End If

            ' مნיშვნელოვანი: ვიპოვოთ მთავარი গরিდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            If mainGridPanel Is Nothing Then
                Debug.WriteLine("❌ მთავარი গরიดის پานელი ვერ მოიძებნა!")
                Return
            End If

            ' মাসშტაბირებული პარამეტრები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' אრჩეული თারიღი
            Dim selectedDate As DateTime = DTPCalendar.Value.Date
            Debug.WriteLine($"არჩეული თარიღი: {selectedDate:dd.MM.yyyy}")

            ' ფილტრი - მხოლოდ אחรოული дღის සเสียები
            Dim daySessions = GetFilteredSessions()
            Debug.WriteLine($"ნაპოვნია {daySessions.Count} სესია ამ თарიღისთვის")

            ' სწيებიេბ ગানთავსება
            For Each session In daySessions
                Try
                    Debug.WriteLine($"--- ვამუશაবობთ סესიას ID={session.Id}, თერაפევტი='{session.TherapistName}', დრო={session.DateTime:HH:mm} ---")

                    ' 1. তেরাপევტის אינდექსის განসაზღვრა
                    Dim therapistIndex As Integer = -1
                    For i As Integer = 0 To therapists.Count - 1
                        If String.Equals(therapists(i).Trim(), session.TherapistName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                            therapistIndex = i
                            Exit For
                        End If
                    Next

                    If therapistIndex < 0 Then
                        Debug.WriteLine($"❌ تেরاপেβτი '{session.TherapistName}' ვერ მოიძებנა თერাפევტების სიაში!")
                        Continue For
                    End If

                    ' 2. דროის ინটერვალის განსაზღვრა (იგივე ლოგիკა, רაც სივრცეებისთვის)
                    Dim sessionTime As TimeSpan = session.DateTime.TimeOfDay
                    Dim timeIndex As Integer = -1

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
                    If therapistIndex < 0 OrElse timeIndex < 0 OrElse
              therapistIndex >= therapists.Count OrElse timeIndex >= timeIntervals.Count Then
                        Debug.WriteLine($"❌ არასწორი ინდექსები: therapist={therapistIndex}, time={timeIndex}")
                        Continue For
                    End If

                    ' 4. ბარათის პოზიციის გამოთვლა
                    Dim cardX As Integer = therapistIndex * THERAPIST_COLUMN_WIDTH + 4
                    Dim cardY As Integer = timeIndex * ROW_HEIGHT + 2

                    ' 5. პროპორციული ბარათის სიმაღლის გამოთვლა
                    Dim sessionDurationMinutes As Integer = session.Duration
                    Dim baseCardHeight As Double = ROW_HEIGHT * (sessionDurationMinutes / 30.0)
                    Dim minCardHeight As Integer = ROW_HEIGHT
                    Dim cardHeight As Integer = CInt(Math.Max(baseCardHeight, minCardHeight))
                    Dim maxCardHeight As Integer = ROW_HEIGHT * 4
                    If cardHeight > maxCardHeight Then cardHeight = maxCardHeight

                    ' 6. ფერების განსაზღვრა SessionStatusColors კლასიდან
                    Dim cardColor As Color = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                    Dim borderColor As Color = SessionStatusColors.GetStatusBorderColor(session.Status, session.DateTime)
                    Dim headerColor As Color = Color.FromArgb(
               Math.Max(0, cardColor.R - 50),
               Math.Max(0, cardColor.G - 50),
               Math.Max(0, cardColor.B - 50)
           )

                    Debug.WriteLine($"🎨 ფერის განსაზღვრა: სტატუსი='{session.Status}' -> ფონი={cardColor}, ჩარჩო={borderColor}, header={headerColor}")

                    ' 7. ბარათის შექმნა
                    Dim sessionCard As New Panel()
                    sessionCard.Size = New Size(THERAPIST_COLUMN_WIDTH - 8, cardHeight)
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
                    Dim currentY As Integer = HEADER_HEIGHT + 5
                    Dim labelSpacing As Integer = 12

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

                    ' სივრცე (თუ ბარათი საკმარისად დიდია)
                    If cardHeight > HEADER_HEIGHT + 30 AndAlso Not String.IsNullOrEmpty(session.Space) Then
                        Dim lblSpace As New Label()
                        lblSpace.Text = $"სივრცე: {session.Space}"
                        lblSpace.Size = New Size(sessionCard.Width - 10, 11)
                        lblSpace.Location = New Point(5, currentY)
                        lblSpace.Font = New Font("Sylfaen", 7, FontStyle.Regular)
                        lblSpace.ForeColor = Color.FromArgb(40, 40, 40)
                        lblSpace.BackColor = Color.Transparent
                        lblSpace.TextAlign = ContentAlignment.TopLeft
                        sessionCard.Controls.Add(lblSpace)
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

                    ' 14. ივენთების მიბმა (დრაგ ენდ დროპი თერაპევტების გრიდისთვისაც)
                    AddHandler sessionCard.MouseDown, AddressOf TherapistSessionCard_MouseDown
                    AddHandler sessionCard.MouseMove, AddressOf TherapistSessionCard_MouseMove
                    AddHandler sessionCard.MouseUp, AddressOf TherapistSessionCard_MouseUp
                    AddHandler sessionCard.Click, AddressOf SessionCard_Click
                    AddHandler sessionCard.DoubleClick, AddressOf SessionCard_DoubleClick

                    ' 15. ბარათის mainGridPanel-ზე დამატება
                    mainGridPanel.Controls.Add(sessionCard)
                    sessionCard.BringToFront()

                    Debug.WriteLine($"✅ ბარათი განთავსდა: ID={session.Id}, თერაპევტი='{session.TherapistName}', " &
                         $"ზომა={sessionCard.Width}x{sessionCard.Height}px, ფერი={cardColor}")

                Catch sessionEx As Exception
                    Debug.WriteLine($"❌ შეცდომა სესია ID={session.Id}: {sessionEx.Message}")
                    Continue For
                End Try
            Next

            Debug.WriteLine("=== PlaceSessionsOnTherapistGrid: დასრულება ===")

        Catch ex As Exception
            Debug.WriteLine($"❌ PlaceSessionsOnTherapistGrid: ზოგადი შეცდომა - {ex.Message}")
        End Try
    End Sub

#Region "თერაპევტების გრიდისთვის დრაგ ენდ დროპი"

    ''' <summary>
    ''' თერაპევტების გრიდზე სესიის ბარათის MouseDown ივენთი
    ''' </summary>
    Private Sub TherapistSessionCard_MouseDown(sender As Object, e As MouseEventArgs)
        Try
            Debug.WriteLine($"TherapistSessionCard_MouseDown: დაჭერა {e.Button}")

            If e.Button = MouseButtons.Left Then
                draggedCard = DirectCast(sender, Panel)
                isDragging = True
                dragStartPoint = e.Location
                originalCardPosition = draggedCard.Location

                ' შევინახოთ ორიგინალური სესიის მონაცემები
                Dim sessionId As Integer = CInt(draggedCard.Tag)
                originalSessionData = allSessions.FirstOrDefault(Function(s) s.Id = sessionId)

                ' ბარათი გადავიტანოთ ყველაზე წინ
                draggedCard.BringToFront()
                draggedCard.Cursor = Cursors.SizeAll

                ' მნიშვნელოვანი: Capture-ის დაყენება
                draggedCard.Capture = True

                Debug.WriteLine($"TherapistSessionCard_MouseDown: დაიწყო გადატანა სესია ID={sessionId}, მუშაობს თერაპევტების გრიდზე")
            End If
        Catch ex As Exception
            Debug.WriteLine($"TherapistSessionCard_MouseDown: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაپევტების გრიდზე სესიის ბარათის MouseMove ივენთი
    ''' </summary>
    Private Sub TherapistSessionCard_MouseMove(sender As Object, e As MouseEventArgs)
        Try
            If isDragging AndAlso draggedCard IsNot Nothing Then
                ' გამოვთვალოთ ახალი პოზიცია მაუსის კოორდინატების მიხედვით
                Dim currentLocation As Point = draggedCard.Location
                currentLocation.X += e.X - dragStartPoint.X
                currentLocation.Y += e.Y - dragStartPoint.Y

                ' მასშტაბირებული პარამეტრები გადატანისთვის
                Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
                Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

                ' ვამოწმებთ საზღვრებს - mainGridPanel-ის ფარგლებში
                Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
                If mainGridPanel IsNot Nothing Then
                    Dim maxX As Integer = (THERAPIST_COLUMN_WIDTH * therapists.Count) - draggedCard.Width
                    Dim maxY As Integer = (ROW_HEIGHT * timeIntervals.Count) - draggedCard.Height

                    ' შეზღუდავთ მოძრაობას გრიდის ფარგლებში
                    If currentLocation.X < 0 Then currentLocation.X = 0
                    If currentLocation.Y < 0 Then currentLocation.Y = 0
                    If currentLocation.X > maxX Then currentLocation.X = maxX
                    If currentLocation.Y > maxY Then currentLocation.Y = maxY

                    ' დავაყენოთ ახალი პოზიცია
                    draggedCard.Location = currentLocation

                    Debug.WriteLine($"TherapistSessionCard_MouseMove: ახალი პოზიცია - X={currentLocation.X}, Y={currentLocation.Y}")
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"TherapistSessionCard_MouseMove: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტების გრიდზე სესიის ბარათის MouseUp ივენთი
    ''' </summary>
    Private Sub TherapistSessionCard_MouseUp(sender As Object, e As MouseEventArgs)
        Try
            Debug.WriteLine($"TherapistSessionCard_MouseUp: მაუსის ღილაკი მოიხსნა, isDragging={isDragging}")

            If isDragging AndAlso draggedCard IsNot Nothing Then
                isDragging = False
                draggedCard.Cursor = Cursors.Hand
                draggedCard.Capture = False

                ' გამოვთვალოთ ახალი გრიდის პოზიცია (თერაპევტების მიხედვით)
                Dim newGridPosition = CalculateTherapistGridPosition(draggedCard.Location)
                Dim originalGridPosition = CalculateTherapistGridPosition(originalCardPosition)

                Debug.WriteLine($"TherapistSessionCard_MouseUp: ორიგინალური პოზიცია - Therapist={originalGridPosition.therapistIndex}, Time={originalGridPosition.timeIndex}")
                Debug.WriteLine($"TherapistSessionCard_MouseUp: ახალი პოზიცია - Therapist={newGridPosition.therapistIndex}, Time={newGridPosition.timeIndex}")

                ' ვამოწმებთ შეიცვალა თუ არა პოზიცია
                If newGridPosition.therapistIndex <> originalGridPosition.therapistIndex OrElse
           newGridPosition.timeIndex <> originalGridPosition.timeIndex Then

                    ' მესიჯბოქსი დადასტურებისთვის
                    Dim newTherapist As String = therapists(newGridPosition.therapistIndex)
                    Dim newTime As String = timeIntervals(newGridPosition.timeIndex).ToString("HH:mm")

                    Dim result As DialogResult = MessageBox.Show(
                $"გსურთ შევცვალოთ სეანსის პარამეტრები?{Environment.NewLine}" &
                $"ახალი თერაپევტი: {newTherapist}{Environment.NewLine}" &
                $"ახალი დრო: {newTime}",
                "სეანსის განახლება",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1
            )

                    If result = DialogResult.Yes Then
                        ' დავადასტუროთ ცვლილება თერაპევტისთვის
                        ConfirmTherapistSessionMove(newGridPosition)
                    Else
                        ' დავაბრუნოთ ძველ ადგილას
                        RevertSessionMove()
                    End If
                Else
                    ' პოზიცია არ შეცვლილა
                    Debug.WriteLine("TherapistSessionCard_MouseUp: პოზიცია არ შეცვლილა")
                End If

                ' გავასუფთავოთ ცვლადები
                draggedCard = Nothing
                originalSessionData = Nothing
            End If
        Catch ex As Exception
            Debug.WriteLine($"TherapistSessionCard_MouseUp: შეცდომა - {ex.Message}")
            ' შეცდომის შემთხვევაში დავაბრუნოთ ძველ ადგილას
            RevertSessionMove()
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტების გრიდის პოზიციის გამოთვლა pixel კოორდინატებიდან
    ''' </summary>
    Private Function CalculateTherapistGridPosition(pixelLocation As Point) As (therapistIndex As Integer, timeIndex As Integer)
        Try
            ' მასშტაბირებული პარამეტრები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' გამოვთვალოთ ინდექსები (მარჯინების გათვალისწინებით)
            Dim therapistIndex As Integer = pixelLocation.X \ THERAPIST_COLUMN_WIDTH
            Dim timeIndex As Integer = pixelLocation.Y \ ROW_HEIGHT

            ' შევამოწმოთ საზღვრები
            If therapistIndex < 0 Then therapistIndex = 0
            If therapistIndex >= therapists.Count Then therapistIndex = therapists.Count - 1
            If timeIndex < 0 Then timeIndex = 0
            If timeIndex >= timeIntervals.Count Then timeIndex = timeIntervals.Count - 1

            Debug.WriteLine($"CalculateTherapistGridPosition: Pixel({pixelLocation.X}, {pixelLocation.Y}) -> Grid({therapistIndex}, {timeIndex})")
            Return (therapistIndex, timeIndex)

        Catch ex As Exception
            Debug.WriteLine($"CalculateTherapistGridPosition: შეცდომა - {ex.Message}")
            Return (0, 0)
        End Try
    End Function

    ''' <summary>
    ''' თერაპევტების გრიდზე სესიის მოძრაობის დადასტურება და მონაცემების განახლება
    ''' </summary>
    Private Sub ConfirmTherapistSessionMove(newGridPosition As (therapistIndex As Integer, timeIndex As Integer))
        Try
            Debug.WriteLine($"ConfirmTherapistSessionMove: ახალი პოზიცია - Therapist={newGridPosition.therapistIndex}, Time={newGridPosition.timeIndex}")

            If originalSessionData Is Nothing OrElse dataService Is Nothing Then
                Debug.WriteLine("ConfirmTherapistSessionMove: originalSessionData ან dataService არის null")
                RevertSessionMove()
                Return
            End If

            ' ახალი თერაპევტი და დრო
            Dim newTherapist As String = therapists(newGridPosition.therapistIndex)
            Dim newDateTime As DateTime = timeIntervals(newGridPosition.timeIndex)

            ' შევქმნათ ახალი DateTime, რომელიც ინარჩუნებს ორიგინალურ თარიღს მაგრამ ცვლის დროს
            Dim originalDate As DateTime = originalSessionData.DateTime.Date
            Dim newFullDateTime As DateTime = originalDate.Add(newDateTime.TimeOfDay)

            ' განვაახლოთ სესიის ობიექტი
            originalSessionData.TherapistName = newTherapist
            originalSessionData.DateTime = newFullDateTime

            ' შევცდოთ მონაცემთა ბაზაში განახლება
            Try
                ' მივიღოთ ყველა სესიის მონაცემები
                Dim allSessionsData = dataService.GetData("DB-Schedule!A2:O")

                ' ვიპოვოთ ჩვენი სესიის მწკრივი
                For i As Integer = 0 To allSessionsData.Count - 1
                    Dim row = allSessionsData(i)
                    If row.Count > 0 AndAlso Integer.TryParse(row(0).ToString(), Nothing) Then
                        Dim rowId As Integer = Integer.Parse(row(0).ToString())
                        If rowId = originalSessionData.Id Then
                            ' განვაახლოთ თერაპევტი (I სვეტი - ინდექსი 8)
                            Dim updatedRow As New List(Of Object)(row)
                            If updatedRow.Count > 8 Then updatedRow(8) = newTherapist

                            ' განვაახლოთ თარიღი (F სვეტი - ინდექსი 5)
                            If updatedRow.Count > 5 Then updatedRow(5) = newFullDateTime.ToString("dd.MM.yyyy HH:mm")

                            ' განვაახლოთ მონაცემები Google Sheets-ში
                            Dim updateRange As String = $"DB-Schedule!A{i + 2}:O{i + 2}"
                            dataService.UpdateData(updateRange, updatedRow)

                            Debug.WriteLine($"ConfirmTherapistSessionMove: წარმატებით განახლდა სესია ID={originalSessionData.Id}")
                            Exit For
                        End If
                    End If
                Next

                ' განვაახლოთ ლოკალური მონაცემები
                LoadSessions()

                ' განვაახლოთ კალენდრის ხედი
                UpdateCalendarView()

                ' შეტყობინება წარმატებული განახლების შესახებ
                MessageBox.Show($"სესია წარმატებით გადატანილია თერაპევტზე {newTherapist} {newDateTime:HH:mm} დროზე",
                      "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch updateEx As Exception
                Debug.WriteLine($"ConfirmTherapistSessionMove: მონაცემების განახლების შეცდომა - {updateEx.Message}")
                MessageBox.Show($"სესიის განახლების შეცდომა: {updateEx.Message}",
                      "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' შეცდომის შემთხვევაში დავაბრუნოთ ძველ ადგილას
                RevertSessionMove()
            End Try

        Catch ex As Exception
            Debug.WriteLine($"ConfirmTherapistSessionMove: შეცდომა - {ex.Message}")
            RevertSessionMove()
        End Try
    End Sub

#End Region

    ''' <summary>
    ''' ბენეფიციარების გრიდის სათაურების პანელის გენერაცია - განახლებული ცარიელი კომბობოქსით
    ''' </summary>
    Private Sub InitializeBeneficiaryHeaderPanel()
        Try
            Debug.WriteLine("InitializeBeneficiaryHeaderPanel: დაიწყო ბენეფიციარების სათაურების ინიციალიზაცია")

            ' შევქმნათ ან ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                ' შევქმნათ ახალი პანელი
                beneficiaryHeaderPanel = New Panel()
                beneficiaryHeaderPanel.Name = "beneficiaryHeaderPanel"
                beneficiaryHeaderPanel.AutoScroll = False
                beneficiaryHeaderPanel.BackColor = Color.FromArgb(255, 230, 180) ' ღია ყაშაბისფერი ბენეფიციარების მიზანისთვის
                beneficiaryHeaderPanel.BorderStyle = BorderStyle.FixedSingle
                pnlCalendarGrid.Controls.Add(beneficiaryHeaderPanel)
            End If

            ' პანელის ზომა და პოზიცია
            Dim TIME_COLUMN_WIDTH As Integer = BASE_TIME_COLUMN_WIDTH
            Dim DATE_COLUMN_WIDTH As Integer = BASE_DATE_COLUMN_WIDTH
            Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT
            Dim COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' საწყისი მდგომარეობა - მხოლოდ ერთი სვეტი ცარიელი კომბობოქსით
            beneficiaryHeaderPanel.Size = New Size(COLUMN_WIDTH, HEADER_HEIGHT)
            beneficiaryHeaderPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, 0)

            ' გავასუფთავოთ პანელი
            beneficiaryHeaderPanel.Controls.Clear()

            ' შევქმნათ საწყისი კომბობოქსი ყველა ბენეფიციარით
            CreateBeneficiaryColumnWithComboBox(beneficiaryHeaderPanel, 0)

            Debug.WriteLine("InitializeBeneficiaryHeaderPanel: ბენეფიციარების სათაურების პანელი შეიქმნა ცარიელი კომბობოქსით")

        Catch ex As Exception
            Debug.WriteLine($"InitializeBeneficiaryHeaderPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის კომბობოქსის შექმნა მითითებულ სვეტში
    ''' </summary>
    ''' <param name="headerPanel">სათაურების პანელი</param>
    ''' <param name="columnIndex">სვეტის ინდექსი</param>
    Private Sub CreateBeneficiaryColumnWithComboBox(headerPanel As Panel, columnIndex As Integer)
        Try
            Debug.WriteLine($"CreateBeneficiaryColumnWithComboBox: ვქმნი კომბობოქსს სვეტისთვის {columnIndex}")

            Dim COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT

            ' კომბობოქსის ზომა და პოზიცია
            Dim comboBox As New ComboBox()
            comboBox.Name = $"comboBene_{columnIndex}"
            comboBox.Size = New Size(COLUMN_WIDTH - 30, HEADER_HEIGHT - 6) ' ღილაკისთვის ადგილის დატოვება
            comboBox.Location = New Point(columnIndex * COLUMN_WIDTH + 5, 3)
            comboBox.Font = New Font("Sylfaen", 8, FontStyle.Regular)
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList
            comboBox.FlatStyle = FlatStyle.Flat
            comboBox.BackColor = Color.White

            ' კომბობოქსის სია - ყველა ბენეფიციარი, რომელსაც ამ დღეს სესია აქვს
            FillBeneficiaryComboBox(comboBox)

            ' ივენთის მიბმა - ბენეფიციარის არჩევისას
            AddHandler comboBox.SelectedIndexChanged, AddressOf BeneficiaryComboBox_SelectedIndexChanged

            ' კომბობოქსის ტაგში შევინახოთ სვეტის ინდექსი
            comboBox.Tag = columnIndex

            ' კომბობოქსის დამატება პანელზე
            headerPanel.Controls.Add(comboBox)

            Debug.WriteLine($"CreateBeneficiaryColumnWithComboBox: კომბობოქსი შეიქმნა, ბენეფიციარების რაოდენობა: {comboBox.Items.Count}")

        Catch ex As Exception
            Debug.WriteLine($"CreateBeneficiaryColumnWithComboBox: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების კომბობოქსის შევსება - მხოლოდ ისინი, რომლებსაც ამ დღეს სესია აქვთ
    ''' </summary>
    ''' <param name="comboBox">შესავსები კომბობოქსი</param>
    Private Sub FillBeneficiaryComboBox(comboBox As ComboBox)
        Try
            Debug.WriteLine("FillBeneficiaryComboBox: ვივსებ კომბობოქსს ბენეფიციარებით")

            ' გავასუფთავოთ კომბობოქსი
            comboBox.Items.Clear()

            ' საწყისო ღია ეტიკეტი
            comboBox.Items.Add("-- აირჩიეთ ბენეფიციარი --")

            ' მიმდინარე თარიღი
            Dim selectedDate As DateTime = DTPCalendar.Value.Date

            ' ყველა ბენეფიციარი, რომელსაც მოცემულ დღეს სესია აქვს
            If allSessions IsNot Nothing Then
                Dim todaysBeneficiaries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each session In allSessions
                    If session.DateTime.Date = selectedDate Then
                        Dim fullName = $"{session.BeneficiaryName.Trim()} {session.BeneficiarySurname.Trim()}"
                        todaysBeneficiaries.Add(fullName)
                    End If
                Next

                ' უკვე არჩეული ბენეფიციარების გამოკლება (თუ არსებობენ)
                Dim selectedBeneficiaries = GetSelectedBeneficiaries()

                ' დანარჩენი ბენეფიციარების დამატება კომბობოქსში
                For Each beneficiary In todaysBeneficiaries.OrderBy(Function(b) b)
                    If Not selectedBeneficiaries.Contains(beneficiary) Then
                        comboBox.Items.Add(beneficiary)
                    End If
                Next
            End If

            ' საწყისი მითითება პირველ ელემენტზე
            If comboBox.Items.Count > 0 Then
                comboBox.SelectedIndex = 0
            End If

            Debug.WriteLine($"FillBeneficiaryComboBox: კომბობოქსი შევსებულია {comboBox.Items.Count} ელემენტით")

        Catch ex As Exception
            Debug.WriteLine($"FillBeneficiaryComboBox: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' უკვე არჩეული ბენეფიციარების სიის მოპოვება ყველა კომბობოქსიდან
    ''' </summary>
    ''' <returns>არჩეული ბენეფიციარების სია</returns>
    Private Function GetSelectedBeneficiaries() As HashSet(Of String)
        Try
            Dim selectedBeneficiaries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel IsNot Nothing Then
                ' გადავირბინოთ ყველა კომბობოქსი და ამოვკრიბოთ არჩეული მნიშვნელობები
                For Each ctrl As Control In beneficiaryHeaderPanel.Controls
                    If TypeOf ctrl Is ComboBox Then
                        Dim combo As ComboBox = DirectCast(ctrl, ComboBox)
                        If combo.SelectedIndex > 0 AndAlso combo.SelectedItem IsNot Nothing Then
                            Dim selectedText = combo.SelectedItem.ToString()
                            If Not selectedText.StartsWith("--") Then
                                selectedBeneficiaries.Add(selectedText)
                            End If
                        End If
                    End If
                Next
            End If

            Debug.WriteLine($"GetSelectedBeneficiaries: ნაპოვნია {selectedBeneficiaries.Count} არჩეული ბენეფიციარი")
            Return selectedBeneficiaries

        Catch ex As Exception
            Debug.WriteLine($"GetSelectedBeneficiaries: შეცდომა - {ex.Message}")
            Return New HashSet(Of String)()
        End Try
    End Function

    ''' <summary>
    ''' ბენეფიციარის კომბობოქსიდან არჩევის ივენთი
    ''' </summary>
    Private Sub BeneficiaryComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Dim comboBox As ComboBox = DirectCast(sender, ComboBox)
            Dim columnIndex As Integer = CInt(comboBox.Tag)

            Debug.WriteLine($"BeneficiaryComboBox_SelectedIndexChanged: არჩეულია {comboBox.SelectedItem} სვეტის {columnIndex} კომბობოქსში")

            ' თუ არჩეული არ არის ნამდვილი ბენეფიციარი, არ გავაგრძელოთ
            If comboBox.SelectedIndex <= 0 OrElse comboBox.SelectedItem.ToString().StartsWith("--") Then
                Return
            End If

            ' მივიღოთ არჩეული ბენეფიციარი
            Dim selectedBeneficiary As String = comboBox.SelectedItem.ToString()

            ' შევცვალოთ კომბობოქსი ლეიბლით + წაშლის ღილაკით
            ReplaceComboBoxWithLabelAndDeleteButton(comboBox, selectedBeneficiary)

            ' შევქმნათ ახალი კომბობოქსი შემდეგ სვეტში
            CreateNextBeneficiaryColumn()

            ' განვაახლოთ ბენეფიციარების სვეტები და მონაცემები
            UpdateBeneficiaryGrid()

        Catch ex As Exception
            Debug.WriteLine($"BeneficiaryComboBox_SelectedIndexChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კომბობოქსის ჩანაცვლება ლეიბლით და წაშლის ღილაკით
    ''' </summary>
    ''' <param name="comboBox">შესაცვლელი კომბობოქსი</param>
    ''' <param name="selectedBeneficiary">არჩეული ბენეფიციარის სახელი</param>
    Private Sub ReplaceComboBoxWithLabelAndDeleteButton(comboBox As ComboBox, selectedBeneficiary As String)
        Try
            Debug.WriteLine($"ReplaceComboBoxWithLabelAndDeleteButton: ვცვლი კომბობოქსს ლეიბლით - {selectedBeneficiary}")

            Dim columnIndex As Integer = CInt(comboBox.Tag)
            Dim parent As Panel = DirectCast(comboBox.Parent, Panel)

            ' კომბობოქსის პოზიცია და ზომები
            Dim comboLocation = comboBox.Location
            Dim comboSize = comboBox.Size

            ' ავშალოთ კომბობოქსი
            parent.Controls.Remove(comboBox)
            comboBox.Dispose()

            ' შევქმნათ ლეიბლი ბენეფიციარის სახელით
            Dim lblBeneficiary As New Label()
            lblBeneficiary.Name = $"lblBene_{columnIndex}"
            lblBeneficiary.Text = selectedBeneficiary
            lblBeneficiary.Location = comboLocation
            lblBeneficiary.Size = New Size(comboSize.Width - 25, comboSize.Height) ' ღილაკისთვის ადგილის დატოვება
            lblBeneficiary.Font = New Font("Sylfaen", 8, FontStyle.Bold)
            lblBeneficiary.TextAlign = ContentAlignment.MiddleCenter
            lblBeneficiary.BorderStyle = BorderStyle.FixedSingle
            lblBeneficiary.BackColor = Color.FromArgb(255, 248, 220) ' კრემისფერი
            lblBeneficiary.Tag = selectedBeneficiary ' ბენეფიციარის სახელი Tag-ში

            ' წაშლის ღილაკის შექმნა
            Dim btnDelete As New Button()
            btnDelete.Name = $"btnDelBene_{columnIndex}"
            btnDelete.Text = "✕" ' X სიმბოლო
            btnDelete.Size = New Size(20, comboSize.Height - 2)
            btnDelete.Location = New Point(comboLocation.X + comboSize.Width - 24, comboLocation.Y + 1)
            btnDelete.Font = New Font("Segoe UI Symbol", 8, FontStyle.Bold)
            btnDelete.ForeColor = Color.DarkRed
            btnDelete.BackColor = Color.FromArgb(255, 230, 230)
            btnDelete.FlatStyle = FlatStyle.Flat
            btnDelete.FlatAppearance.BorderSize = 1
            btnDelete.FlatAppearance.BorderColor = Color.Red
            btnDelete.Cursor = Cursors.Hand
            btnDelete.Tag = columnIndex ' სვეტის ინდექსი Tag-ში

            ' წაშლის ღილაკის ივენთის მიბმა
            AddHandler btnDelete.Click, AddressOf BeneficiaryDeleteButton_Click

            ' კონტროლების დამატება
            parent.Controls.Add(lblBeneficiary)
            parent.Controls.Add(btnDelete)

            Debug.WriteLine($"ReplaceComboBoxWithLabelAndDeleteButton: კომბობოქსი ჩანაცვლდა ლეიბლით და წაშლის ღილაკით")

        Catch ex As Exception
            Debug.WriteLine($"ReplaceComboBoxWithLabelAndDeleteButton: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' შემდეგი ბენეფიციარის კომბობოქსის სვეტის შექმნა
    ''' </summary>
    Private Sub CreateNextBeneficiaryColumn()
        Try
            Debug.WriteLine("CreateNextBeneficiaryColumn: ვქმნი შემდეგ სვეტს")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Debug.WriteLine("CreateNextBeneficiaryColumn: ბენეფიციარების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' დავითვალოთ მიმდინარე სვეტების რაოდენობა
            Dim currentColumns As Integer = 0
            For Each ctrl As Control In beneficiaryHeaderPanel.Controls
                If TypeOf ctrl Is Label AndAlso ctrl.Name.StartsWith("lblBene_") Then
                    currentColumns += 1
                End If
            Next

            ' გავაფართოოთ სათაურების პანელი ახალი სვეტისთვის
            Dim COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            beneficiaryHeaderPanel.Width += COLUMN_WIDTH

            ' შევქმნათ ახალი კომბობოქსი
            CreateBeneficiaryColumnWithComboBox(beneficiaryHeaderPanel, currentColumns)

            Debug.WriteLine($"CreateNextBeneficiaryColumn: შეიქმნა ახალი სვეტი ინდექსით {currentColumns}")

        Catch ex As Exception
            Debug.WriteLine($"CreateNextBeneficiaryColumn: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის წაშლის ღილაკზე დაჭერის ივენთი
    ''' </summary>
    Private Sub BeneficiaryDeleteButton_Click(sender As Object, e As EventArgs)
        Try
            Dim btnDelete As Button = DirectCast(sender, Button)
            Dim columnIndex As Integer = CInt(btnDelete.Tag)

            Debug.WriteLine($"BeneficiaryDeleteButton_Click: ვშლი სვეტს ინდექსით {columnIndex}")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Debug.WriteLine("BeneficiaryDeleteButton_Click: ბენეფიციარების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' წავშალოთ შესაბამისი ლეიბლი და ღილაკი
            Dim lblToRemove As Label = DirectCast(beneficiaryHeaderPanel.Controls.Find($"lblBene_{columnIndex}", False).FirstOrDefault(), Label)
            Dim btnToRemove As Button = DirectCast(beneficiaryHeaderPanel.Controls.Find($"btnDelBene_{columnIndex}", False).FirstOrDefault(), Button)

            If lblToRemove IsNot Nothing Then
                beneficiaryHeaderPanel.Controls.Remove(lblToRemove)
                lblToRemove.Dispose()
            End If

            If btnToRemove IsNot Nothing Then
                beneficiaryHeaderPanel.Controls.Remove(btnToRemove)
                btnToRemove.Dispose()
            End If

            ' ვხოცავთ სვეტებს და ვანახლებთ მათ პოზიციებს
            ReorganizeBeneficiaryColumns()

            ' განავალოთ ბენეფიციარების გრიდი
            UpdateBeneficiaryGrid()

            ' განვაახლოთ ყველა კომბობოქსი დარჩენილ ბენეფიციარების ჩვენებისთვის
            RefreshAllBeneficiaryComboBoxes()

            Debug.WriteLine($"BeneficiaryDeleteButton_Click: სვეტი {columnIndex} წარმატებით წაიშალა")

        Catch ex As Exception
            Debug.WriteLine($"BeneficiaryDeleteButton_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების სვეტების რეორგანიზაცია წაშლის შემდეგ
    ''' </summary>
    Private Sub ReorganizeBeneficiaryColumns()
        Try
            Debug.WriteLine("ReorganizeBeneficiaryColumns: ვირეორგანიზებ სვეტებს")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Return
            End If

            Dim COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' ყველა ლეიბლი და ღილაკი ხელახლა დალაგება
            Dim columnIndex As Integer = 0

            ' ყველა ლეიბლი (ბენეფიციარების სახელები)
            For Each ctrl As Control In beneficiaryHeaderPanel.Controls.OfType(Of Label)().OrderBy(Function(l) l.Name).ToList()
                If ctrl.Name.StartsWith("lblBene_") Then
                    ctrl.Name = $"lblBene_{columnIndex}"
                    ctrl.Location = New Point(columnIndex * COLUMN_WIDTH + 5, ctrl.Location.Y)

                    ' შესაბამისი წაშლის ღილაკი
                    Dim correspondingButton = beneficiaryHeaderPanel.Controls.Find($"btnDelBene_{ctrl.Name.Split("_"c)(1)}", False).FirstOrDefault()
                    If correspondingButton IsNot Nothing Then
                        correspondingButton.Name = $"btnDelBene_{columnIndex}"
                        correspondingButton.Location = New Point(columnIndex * COLUMN_WIDTH + COLUMN_WIDTH - 19, correspondingButton.Location.Y)
                        correspondingButton.Tag = columnIndex
                    End If

                    columnIndex += 1
                End If
            Next

            ' კომბობოქსების პოზიციის განახლება
            For Each ctrl As Control In beneficiaryHeaderPanel.Controls.OfType(Of ComboBox)().ToList()
                If ctrl.Name.StartsWith("comboBene_") Then
                    ctrl.Name = $"comboBene_{columnIndex}"
                    ctrl.Location = New Point(columnIndex * COLUMN_WIDTH + 5, ctrl.Location.Y)
                    ctrl.Tag = columnIndex
                    columnIndex += 1
                End If
            Next

            ' პანელის სიგანის განახლება
            beneficiaryHeaderPanel.Width = columnIndex * COLUMN_WIDTH

            Debug.WriteLine($"ReorganizeBeneficiaryColumns: რეორგანიზაცია დასრულდა, სულ სვეტი: {columnIndex}")

        Catch ex As Exception
            Debug.WriteLine($"ReorganizeBeneficiaryColumns: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ყველა ბენეფიციარის კომბობოქსის განახლება
    ''' </summary>
    Private Sub RefreshAllBeneficiaryComboBoxes()
        Try
            Debug.WriteLine("RefreshAllBeneficiaryComboBoxes: ვანახლებ ყველა კომბობოქსს")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Return
            End If

            ' განვაახლოთ ყველა კომბობოქსი
            For Each ctrl As Control In beneficiaryHeaderPanel.Controls.OfType(Of ComboBox)().ToList()
                If ctrl.Name.StartsWith("comboBene_") Then
                    FillBeneficiaryComboBox(DirectCast(ctrl, ComboBox))
                End If
            Next

            Debug.WriteLine("RefreshAllBeneficiaryComboBoxes: ყველა კომბობოქსი განახლდა")

        Catch ex As Exception
            Debug.WriteLine($"RefreshAllBeneficiaryComboBoxes: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების ხედის ჩვენება - სრულიად ახალი იმპლემენტაცია
    ''' </summary>
    Private Sub ShowDayViewByBeneficiary()
        Try
            Debug.WriteLine("ShowDayViewByBeneficiary: დაიწყო ბენეფიციარების ხედის ჩვენება")

            ' დავაყენოთ pnlFilter-ის ფერი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' ======= 1. პანელების ინიციალიზაცია =======
            InitializeBeneficiaryDayViewPanels()

            ' ======= 2. თარიღის ზოლის შევსება =======
            FillDateColumnPanel()

            ' ======= 3. დროის პანელის შევსება =======
            FillTimeColumnPanel()

            ' ======= 4. ბენეფიციარების სათაურების პანელის ინიციალიზაცია =======
            InitializeBeneficiaryHeaderPanel()

            ' ======= 5. მთავარი გრიდის პანელის შევსება ბენეფიციარებისთვის =======
            FillMainGridPanelForBeneficiaries()

            ' ======= 6. სინქრონიზაცია სქროლისთვის =======
            SetupScrollSynchronizationForBeneficiaries()

            ' ======= 7. ძველი ბარათების გასუფთავება =======
            ClearSessionCardsFromGrid()

            ' ======= 8. სესიების განთავსება ბენეფიციარების გრიდში =======
            PlaceSessionsOnBeneficiaryGrid()

            Debug.WriteLine("ShowDayViewByBeneficiary: ბენეფიციარების ხედის ჩვენება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewByBeneficiary: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewByBeneficiary: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"ბენეფიციარების ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების დღის ხედის პანელების ინიციალიზაცია
    ''' იგივე ლოგიკა, რაც სივრცეებისა და თერაპევტებისთვის
    ''' </summary>
    Private Sub InitializeBeneficiaryDayViewPanels()
        Try
            Debug.WriteLine("InitializeBeneficiaryDayViewPanels: დაიწყო ბენეფიციარების დღის ხედის პანელების ინიციალიზაცია")

            ' გავასუფთავოთ კალენდრის პანელი
            pnlCalendarGrid.Controls.Clear()

            ' pnlCalendarGrid-ს არ ჰქონდეს AutoScroll!
            pnlCalendarGrid.AutoScroll = False

            ' თუ დროის ინტერვალები არ არის ინიციალიზებული, დავამატოთ
            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' წავკითხოთ მშობელი პანელის ზომები
            Dim totalWidth As Integer = pnlCalendarGrid.ClientSize.Width
            Dim totalHeight As Integer = pnlCalendarGrid.ClientSize.Height

            Debug.WriteLine($"InitializeBeneficiaryDayViewPanels: კონტეინერის ზომები - Width={totalWidth}, Height={totalHeight}")

            ' კონსტანტები, რომლებიც არ იცვლება მასშტაბირებისას
            Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT
            Dim TIME_COLUMN_WIDTH As Integer = BASE_TIME_COLUMN_WIDTH
            Dim DATE_COLUMN_WIDTH As Integer = BASE_DATE_COLUMN_WIDTH

            ' გამოვთვალოთ დროისა და თარიღის პანელების ზომები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim totalRowsHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            Debug.WriteLine($"InitializeBeneficiaryDayViewPanels: ROW_HEIGHT={ROW_HEIGHT}, totalRowsHeight={totalRowsHeight}")

            ' ======= 1. შევქმნათ თარიღის ზოლის პანელი =======
            Dim dateColumnPanel As New Panel()
            dateColumnPanel.Name = "dateColumnPanel"
            dateColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            dateColumnPanel.AutoScroll = False
            dateColumnPanel.Size = New Size(DATE_COLUMN_WIDTH, totalRowsHeight)
            dateColumnPanel.Location = New Point(0, HEADER_HEIGHT)
            dateColumnPanel.BackColor = Color.FromArgb(60, 80, 150)
            dateColumnPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(dateColumnPanel)

            ' ======= 2. შევქმნათ საათების პანელი =======
            Dim timeColumnPanel As New Panel()
            timeColumnPanel.Name = "timeColumnPanel"
            timeColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            timeColumnPanel.AutoScroll = False
            timeColumnPanel.Size = New Size(TIME_COLUMN_WIDTH, totalRowsHeight)
            timeColumnPanel.Location = New Point(DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            timeColumnPanel.BackColor = Color.FromArgb(240, 240, 245)
            timeColumnPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(timeColumnPanel)

            ' ======= 3. შევქმნათ თარიღის სათაურის პანელი =======
            Dim dateHeaderPanel As New Panel()
            dateHeaderPanel.Name = "dateHeaderPanel"
            dateHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
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

            ' ======= 4. შევქმნათ დროის სათაურის პანელი =======
            Dim timeHeaderPanel As New Panel()
            timeHeaderPanel.Name = "timeHeaderPanel"
            timeHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
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

            ' ======= 5. შევქმნათ მთავარი გრიდის პანელი =======
            Dim mainGridPanel As New Panel()
            mainGridPanel.Name = "mainGridPanel"
            mainGridPanel.Size = New Size(totalWidth - TIME_COLUMN_WIDTH - DATE_COLUMN_WIDTH, totalHeight - HEADER_HEIGHT)
            mainGridPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            mainGridPanel.BackColor = Color.White
            mainGridPanel.BorderStyle = BorderStyle.FixedSingle
            mainGridPanel.AutoScroll = True
            pnlCalendarGrid.Controls.Add(mainGridPanel)

            Debug.WriteLine("InitializeBeneficiaryDayViewPanels: პანელების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"InitializeBeneficiaryDayViewPanels: შეცდომა - {ex.Message}")
            Debug.WriteLine($"InitializeBeneficiaryDayViewPanels: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"ბენეფიციარების პანელების ინიციალიზაციის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის შევსება ბენეფიციარებისთვის
    ''' დინამიკურად ქმნის უჯრედებს არჩეული ბენეფიციარების მიხედვით
    ''' </summary>
    Private Sub FillMainGridPanelForBeneficiaries()
        Try
            Debug.WriteLine("FillMainGridPanelForBeneficiaries: დაიწყო მთავარი გრიდის შევსება ბენეფიციარებისთვის")

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("FillMainGridPanelForBeneficiaries: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' გავასუფთავოთ მთავარი პანელი
            mainGridPanel.Controls.Clear()

            ' შევამოწმოთ დროის ინტერვალები
            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' მივიღოთ არჩეული ბენეფიციარების სია
            Dim selectedBeneficiaries = GetBeneficiaryColumns()

            If selectedBeneficiaries.Count = 0 Then
                Debug.WriteLine("FillMainGridPanelForBeneficiaries: არჩეული ბენეფიციარები ვერ მოიძებნა")
                Return
            End If

            ' პარამეტრები მასშტაბის გათვალისწინებით
            Dim BENEFICIARY_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)

            ' გრიდის სრული სიგანე
            Dim gridWidth As Integer = BENEFICIARY_COLUMN_WIDTH * selectedBeneficiaries.Count

            ' გრიდის სრული სიმაღლე
            Dim gridHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            ' მნიშვნელოვანი: დავაყენოთ mainGridPanel-ის AutoScrollMinSize
            mainGridPanel.AutoScrollMinSize = New Size(gridWidth, gridHeight)

            Debug.WriteLine($"FillMainGridPanelForBeneficiaries: გრიდის ზომები - Width={gridWidth}, Height={gridHeight}")

            ' უჯრედების მასივი (ახლა ბენეფიციარებისთვის)
            ReDim gridCells(selectedBeneficiaries.Count - 1, timeIntervals.Count - 1)

            ' უჯრედების შექმნა
            For col As Integer = 0 To selectedBeneficiaries.Count - 1
                For row As Integer = 0 To timeIntervals.Count - 1
                    Dim cell As New Panel()
                    cell.Size = New Size(BENEFICIARY_COLUMN_WIDTH, ROW_HEIGHT)
                    cell.Location = New Point(col * BENEFICIARY_COLUMN_WIDTH, row * ROW_HEIGHT)

                    ' ალტერნატიული ფერები (ღია ყვითელი ფერები ბენეფიციარებისთვის)
                    If row Mod 2 = 0 Then
                        cell.BackColor = Color.FromArgb(255, 255, 250) ' ღია კრემისფერი
                    Else
                        cell.BackColor = Color.FromArgb(255, 250, 245) ' უფრო კრემისფერი
                    End If

                    ' უჯრედის ჩარჩო
                    AddHandler cell.Paint, AddressOf Cell_Paint

                    ' ვინახავთ უჯრედის ობიექტს მასივში
                    gridCells(col, row) = cell

                    ' ვამატებთ უჯრედს პირდაპირ mainGridPanel-ზე
                    mainGridPanel.Controls.Add(cell)
                Next
            Next

            Debug.WriteLine($"FillMainGridPanelForBeneficiaries: შეიქმნა {selectedBeneficiaries.Count * timeIntervals.Count} უჯრედი ბენეფიციარებისთვის")

        Catch ex As Exception
            Debug.WriteLine($"FillMainGridPanelForBeneficiaries: შეცდომა - {ex.Message}")
            Debug.WriteLine($"FillMainGridPanelForBeneficiaries: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების ხედისთვის სქროლის სინქრონიზაცია
    ''' </summary>
    Private Sub SetupScrollSynchronizationForBeneficiaries()
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("SetupScrollSynchronizationForBeneficiaries: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            ' ჯერ მოვხსნათ არსებული ივენთის ჰენდლერები, თუ ისინი არსებობენ
            RemoveHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' მივაბათ Scroll ივენთი ხელახლა
            AddHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' ასევე, დავამატოთ ივენთი გორგოლჭის სქროლისთვის
            AddHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            Debug.WriteLine("SetupScrollSynchronizationForBeneficiaries: სქროლის სინქრონიზაცია დაყენებულია")
            Debug.WriteLine($"  beneficiaryHeaderPanel: {If(beneficiaryHeaderPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")

        Catch ex As Exception
            Debug.WriteLine($"SetupScrollSynchronizationForBeneficiaries: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების განთავსება ბენეფიციარების გრიდში
    ''' დაფუძნებულია PlaceSessionsOnGrid და PlaceSessionsOnTherapistGrid მეთოდებზე
    ''' </summary>
    Private Sub PlaceSessionsOnBeneficiaryGrid()
        Try
            Debug.WriteLine("=== PlaceSessionsOnBeneficiaryGrid: დაიწყო ===")

            ' ბაზისური შემოწმებები
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnBeneficiaryGrid: სესიები არ არის")
                Return
            End If

            If gridCells Is Nothing Then
                Debug.WriteLine("PlaceSessionsOnBeneficiaryGrid: გრიდის უჯრედები არ არის")
                Return
            End If

            If timeIntervals.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnBeneficiaryGrid: დროის ინტერვალები არ არის")
                Return
            End If

            ' მივიღოთ არჩეული ბენეფიციარები
            Dim selectedBeneficiaries = GetBeneficiaryColumns()

            If selectedBeneficiaries.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnBeneficiaryGrid: არჩეული ბენეფიციარები არ არის")
                Return
            End If

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            If mainGridPanel Is Nothing Then
                Debug.WriteLine("❌ მთავარი გრიდის პანელი ვერ მოიძებნა!")
                Return
            End If

            ' მასშტაბირებული პარამეტრები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim BENEFICIARY_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' ფილტრი - მხოლოდ ფილტრირებული სესიები
            Dim daySessions = GetFilteredSessions()
            Debug.WriteLine($"ნაპოვნია {daySessions.Count} ფილტრირებული სესია ბენეფიციარების ხედისთვის")

            ' სესიების განთავსება
            For Each session In daySessions
                Try
                    Debug.WriteLine($"--- ვამუშავებთ სესიას ID={session.Id}, ბენეფიციარი='{session.BeneficiaryName} {session.BeneficiarySurname}', დრო={session.DateTime:HH:mm} ---")

                    ' 1. ბენეფიციარის ინდექსის განსაზღვრა
                    Dim beneficiaryIndex As Integer = -1
                    Dim sessionBeneficiaryFullName As String = $"{session.BeneficiaryName.Trim()} {session.BeneficiarySurname.Trim()}"

                    For i As Integer = 0 To selectedBeneficiaries.Count - 1
                        If String.Equals(selectedBeneficiaries(i).Trim(), sessionBeneficiaryFullName, StringComparison.OrdinalIgnoreCase) Then
                            beneficiaryIndex = i
                            Exit For
                        End If
                    Next

                    If beneficiaryIndex < 0 Then
                        Debug.WriteLine($"❌ ბენეფიციარი '{sessionBeneficiaryFullName}' ვერ მოიძებნა არჩეულ ბენეფიციარებში!")
                        Continue For
                    End If

                    ' 2. დროის ინტერვალის განსაზღვრა (იგივე ლოგიკა, რაც სივრცეებისთვის)
                    Dim sessionTime As TimeSpan = session.DateTime.TimeOfDay
                    Dim timeIndex As Integer = -1

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
                    If beneficiaryIndex < 0 OrElse timeIndex < 0 OrElse
              beneficiaryIndex >= selectedBeneficiaries.Count OrElse timeIndex >= timeIntervals.Count Then
                        Debug.WriteLine($"❌ არასწორი ინდექსები: beneficiary={beneficiaryIndex}, time={timeIndex}")
                        Continue For
                    End If

                    ' 4. ბარათის პოზიციის გამოთვლა
                    Dim cardX As Integer = beneficiaryIndex * BENEFICIARY_COLUMN_WIDTH + 4
                    Dim cardY As Integer = timeIndex * ROW_HEIGHT + 2

                    ' 5. პროპორციული ბარათის სიმაღლის გამოთვლა
                    Dim sessionDurationMinutes As Integer = session.Duration
                    Dim baseCardHeight As Double = ROW_HEIGHT * (sessionDurationMinutes / 30.0)
                    Dim minCardHeight As Integer = ROW_HEIGHT
                    Dim cardHeight As Integer = CInt(Math.Max(baseCardHeight, minCardHeight))
                    Dim maxCardHeight As Integer = ROW_HEIGHT * 4
                    If cardHeight > maxCardHeight Then cardHeight = maxCardHeight

                    ' 6. ფერების განსაზღვრა SessionStatusColors კლასიდან
                    Dim cardColor As Color = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                    Dim borderColor As Color = SessionStatusColors.GetStatusBorderColor(session.Status, session.DateTime)
                    Dim headerColor As Color = Color.FromArgb(
               Math.Max(0, cardColor.R - 50),
               Math.Max(0, cardColor.G - 50),
               Math.Max(0, cardColor.B - 50)
           )

                    Debug.WriteLine($"🎨 ფერის განსაზღვრა: სტატუსი='{session.Status}' -> ფონი={cardColor}, ჩარჩო={borderColor}, header={headerColor}")

                    ' 7. ბარათის შექმნა
                    Dim sessionCard As New Panel()
                    sessionCard.Size = New Size(BENEFICIARY_COLUMN_WIDTH - 8, cardHeight)
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

                    ' 9. სესიის ID header-ში
                    Dim lblId As New Label()
                    lblId.Text = $"#{session.Id}"
                    lblId.Location = New Point(8, 2)
                    lblId.Size = New Size(headerPanel.Width - 16, 20)
                    lblId.Font = New Font("Sylfaen", 10, FontStyle.Bold)
                    lblId.ForeColor = Color.White
                    lblId.TextAlign = ContentAlignment.MiddleCenter
                    headerPanel.Controls.Add(lblId)

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

                    ' 11. ლეიბლების დინამიური განთავსება
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

                    ' სივრცე
                    If cardHeight > HEADER_HEIGHT + 50 AndAlso Not String.IsNullOrEmpty(session.Space) Then
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
                    End If

                    ' 12. რედაქტირების ღილაკი
                    Dim btnEdit As New Button()
                    btnEdit.Text = "✎"
                    btnEdit.Font = New Font("Segoe UI Symbol", 10, FontStyle.Bold)
                    btnEdit.ForeColor = Color.White
                    btnEdit.BackColor = Color.FromArgb(80, 80, 80)
                    btnEdit.Size = New Size(24, 24)
                    btnEdit.Location = New Point(sessionCard.Width - 30, sessionCard.Height - 30)
                    btnEdit.FlatStyle = FlatStyle.Flat
                    btnEdit.FlatAppearance.BorderSize = 0
                    btnEdit.Tag = session.Id
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

                    ' 14. ივენთების მიბმა (საშუალო მსგავსი დრაგ ენდ დროპისთვის)
                    AddHandler sessionCard.Click, AddressOf SessionCard_Click
                    AddHandler sessionCard.DoubleClick, AddressOf SessionCard_DoubleClick

                    ' 15. ბარათის mainGridPanel-ზე დამატება
                    mainGridPanel.Controls.Add(sessionCard)
                    sessionCard.BringToFront()

                    Debug.WriteLine($"✅ ბარათი განთავსდა: ID={session.Id}, ბენეფიციარი='{sessionBeneficiaryFullName}', " &
                         $"ზომა={sessionCard.Width}x{sessionCard.Height}px, ფერი={cardColor}")

                Catch sessionEx As Exception
                    Debug.WriteLine($"❌ შეცდომა სესია ID={session.Id}: {sessionEx.Message}")
                    Continue For
                End Try
            Next

            Debug.WriteLine("=== PlaceSessionsOnBeneficiaryGrid: დასრულება ===")

        Catch ex As Exception
            Debug.WriteLine($"❌ PlaceSessionsOnBeneficiaryGrid: ზოგადი შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების გრიდის განახლება - განაახლებს ყველა ლოგიკას
    ''' გამოიძახება ბენეფიციარის დამატებისა და წაშლის შემდეგ
    ''' </summary>
    Private Sub UpdateBeneficiaryGrid()
        Try
            Debug.WriteLine("UpdateBeneficiaryGrid: დაიწყო ბენეფიციარების გრიდის განახლება")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Debug.WriteLine("UpdateBeneficiaryGrid: ბენეფიციარების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' მივიღოთ არჩეული ბენეფიციარები
            Dim selectedBeneficiaries = GetBeneficiaryColumns()

            Debug.WriteLine($"UpdateBeneficiaryGrid: მოიძებნა {selectedBeneficiaries.Count} არჩეული ბენეფიციარი")

            ' განვაახლოთ ბენეფიციარების სათაურების პანელის სიგანე
            Dim BENEFICIARY_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim totalColumns As Integer = Math.Max(1, selectedBeneficiaries.Count + 1) ' +1 ცარიელი კომბობოქსისთვის
            beneficiaryHeaderPanel.Width = totalColumns * BENEFICIARY_COLUMN_WIDTH

            ' განვაახლოთ მთავარი გრიდი
            FillMainGridPanelForBeneficiaries()

            ' განვაახლოთ სქროლის სინქრონიზაცია
            SetupScrollSynchronizationForBeneficiaries()

            ' ძველი ბარათების გასუფთავება და ხელახლა განთავსება
            ClearSessionCardsFromGrid()
            PlaceSessionsOnBeneficiaryGrid()

            Debug.WriteLine("UpdateBeneficiaryGrid: ბენეფიციარების გრიდი წარმატებით განახლდა")

        Catch ex As Exception
            Debug.WriteLine($"UpdateBeneficiaryGrid: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მივიღოთ არჩეული ბენეფიციარების სია სათაურების პანელიდან
    ''' აბრუნებს მხოლოდ იმ ბენეფიციარებს, რომლებიც არჩეულია (არა კომბობოქსებს)
    ''' </summary>
    ''' <returns>არჩეული ბენეფიციარების სია</returns>
    Private Function GetBeneficiaryColumns() As List(Of String)
        Try
            Debug.WriteLine("GetBeneficiaryColumns: ვიღებ არჩეულ ბენეფიციარებს")

            Dim beneficiaries As New List(Of String)()

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Debug.WriteLine("GetBeneficiaryColumns: ბენეფიციარების სათაურების პანელი ვერ მოიძებნა")
                Return beneficiaries
            End If

            ' გავიაროთ ყველა ლეიბლი (არა კომბობოქსი) და ამოვკრიბოთ ბენეფიციარების სახელები
            For Each ctrl As Control In beneficiaryHeaderPanel.Controls.OfType(Of Label)().OrderBy(Function(l) l.Location.X).ToList()
                If ctrl.Name.StartsWith("lblBene_") AndAlso ctrl.Tag IsNot Nothing Then
                    Dim beneficiaryName As String = ctrl.Tag.ToString()
                    beneficiaries.Add(beneficiaryName)
                    Debug.WriteLine($"GetBeneficiaryColumns: დაემატა ბენეფიციარი: '{beneficiaryName}'")
                End If
            Next

            Debug.WriteLine($"GetBeneficiaryColumns: საბოლოოდ მოიძებნა {beneficiaries.Count} ბენეფიციარი")
            Return beneficiaries

        Catch ex As Exception
            Debug.WriteLine($"GetBeneficiaryColumns: შეცდომა - {ex.Message}")
            Return New List(Of String)()
        End Try
    End Function

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
    ''' UpdateCalendarView მეთოდის შესწორება - თერაპევტების ხედის მხარდაჭერით
    ''' (ეს უნდა ჩანაცვლდეს არსებული UpdateCalendarView მეთოდი)
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
            Dim daySessions = GetFilteredSessions()
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
                    AddHandler sessionCard.MouseDown, AddressOf SessionCard_MouseDown
                    AddHandler sessionCard.MouseMove, AddressOf SessionCard_MouseMove
                    AddHandler sessionCard.MouseUp, AddressOf SessionCard_MouseUp
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

    ' ═══════════════════════════════════════
    ' 🔧 მოდიფიკაცია: PlaceSessionsOnGrid (ფილტრაციით)
    ' ═══════════════════════════════════════

    ''' <summary>
    ''' სესიების განთავსება გრიდში - ფილტრაციით
    ''' ყველა კოდი იგივეა, მაგრამ daySessions-ის ნაცვლად ვიყენებთ GetFilteredSessions()-ს
    ''' </summary>
    Private Sub PlaceSessionsOnGridWithFilter()
        Try
            Debug.WriteLine("=== PlaceSessionsOnGrid: დაიწყო (ფილტრაციით) ===")

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

            ' ✨ ძირითადი ცვლილება: ფილტრირებული სესიების მიღება ჩეკბოქსების მიხედვით
            ' არის მაგივრად: Dim daySessions = allSessions.Where(Function(s) s.DateTime.Date = selectedDate).ToList()
            Dim daySessions = GetFilteredSessions()
            Debug.WriteLine($"ნაპოვნია {daySessions.Count} ფილტრირებული სესია")

            ' <--- აქ ყველა დანარჩენი კოდი იგივე რჩება, რაც PlaceSessionsOnGrid-ში არის --->
            ' სესიების განთავსება, ბარათების შექმნა, ფერების განსაზღვრა, ა.შ.

            For Each session In daySessions
                ' [თქვენი არსებული ყველა კოდი ბარათის შექმნისთვის]
                ' ...
            Next

            Debug.WriteLine("=== PlaceSessionsOnGrid: დასრულება ===")

        Catch ex As Exception
            Debug.WriteLine($"❌ PlaceSessionsOnGrid: ზოგადი შეცდომა - {ex.Message}")
        End Try
    End Sub

    ' ═══════════════════════════════════════
    ' 🔧 მოდიფიკაცია: PlaceSessionsOnTherapistGrid (ფილტრაციით)
    ' ═══════════════════════════════════════

    ''' <summary>
    ''' სესიების განთავსება თერაპევტების გრიდში - ფილტრაციით
    ''' ყველა კოდი იგივეა, მაგრამ daySessions-ის ნაცვლად ვიყენებთ GetFilteredSessions()-ს
    ''' </summary>
    Private Sub PlaceSessionsOnTherapistGridWithFilter()
        Try
            Debug.WriteLine("=== PlaceSessionsOnTherapistGrid: დაიწყო (ფილტრაციით) ===")

            ' ძირითადი შემოწმებები (იგივე)
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnTherapistGrid: სესიები არ არის")
                Return
            End If

            ' ... სხვა შემოწმებები ...

            ' ✨ ძირითადი ცვლილება: ფილტრირებული სესიების მიღება ჩეკბოქსების მიხედვით
            ' არის მაგივრად: Dim daySessions = allSessions.Where(Function(s) s.DateTime.Date = selectedDate).ToList()
            Dim daySessions = GetFilteredSessions()
            Debug.WriteLine($"ნაპოვნია {daySessions.Count} ფილტრირებული სესია თერაპევტების ხედისთვის")

            ' <--- აქ ყველა დანარჩენი კოდი იგივე რჩება, რაც PlaceSessionsOnTherapistGrid-ში არის --->
            ' სესიების განთავსება, ბარათების შექმნა, ფერების განსაზღვრა, ა.შ.

            For Each session In daySessions
                ' [თქვენი არსებული ყველა კოდი ბარათის შექმნისთვის]
                ' ...
            Next

            Debug.WriteLine("=== PlaceSessionsOnTherapistGrid: დასრულება ===")

        Catch ex As Exception
            Debug.WriteLine($"❌ PlaceSessionsOnTherapistGrid: ზოგადი შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის ბარათის Mouse ივენთების ცვლადები გადატანისთვის
    ''' </summary>
    Private isDragging As Boolean = False
    Private dragStartPoint As Point
    Private draggedCard As Panel = Nothing
    Private originalCardPosition As Point
    Private originalSessionData As SessionModel = Nothing

    ''' <summary>
    ''' სესიის ბარათზე MouseDown ივენთი
    ''' </summary>
    Private Sub SessionCard_MouseDown(sender As Object, e As MouseEventArgs)
        Try
            Debug.WriteLine($"SessionCard_MouseDown: დაჭერა {e.Button}")

            If e.Button = MouseButtons.Left Then
                draggedCard = DirectCast(sender, Panel)
                isDragging = True
                dragStartPoint = e.Location
                originalCardPosition = draggedCard.Location

                ' შევინახოთ ორიგინალური სესიის მონაცემები
                Dim sessionId As Integer = CInt(draggedCard.Tag)
                originalSessionData = allSessions.FirstOrDefault(Function(s) s.Id = sessionId)

                ' ბარათი გადავიტანოთ ყველაზე წინ
                draggedCard.BringToFront()
                draggedCard.Cursor = Cursors.SizeAll

                ' მნიშვნელოვანი: Capture-ის დაყენება
                draggedCard.Capture = True

                Debug.WriteLine($"SessionCard_MouseDown: დაიწყო გადატანა სესია ID={sessionId}, დარჩა Capturing={draggedCard.Capture}")
            End If
        Catch ex As Exception
            Debug.WriteLine($"SessionCard_MouseDown: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის ბარათზე MouseMove ივენთი
    ''' </summary>
    Private Sub SessionCard_MouseMove(sender As Object, e As MouseEventArgs)
        Try
            If isDragging AndAlso draggedCard IsNot Nothing Then
                ' გამოვთვალოთ ახალი პოზიცია მაუსის კოორდინატების მიხედვით
                Dim currentLocation As Point = draggedCard.Location
                currentLocation.X += e.X - dragStartPoint.X
                currentLocation.Y += e.Y - dragStartPoint.Y

                ' მასშტაბირებული პარამეტრები გადატანისთვის
                Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
                Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

                ' ვამოწმებთ საზღვრებს - mainGridPanel-ის ფარგლებში
                Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
                If mainGridPanel IsNot Nothing Then
                    Dim maxX As Integer = (SPACE_COLUMN_WIDTH * spaces.Count) - draggedCard.Width
                    Dim maxY As Integer = (ROW_HEIGHT * timeIntervals.Count) - draggedCard.Height

                    ' შეზღუდავთ მოძრაობას გრიდის ფარგლებში
                    If currentLocation.X < 0 Then currentLocation.X = 0
                    If currentLocation.Y < 0 Then currentLocation.Y = 0
                    If currentLocation.X > maxX Then currentLocation.X = maxX
                    If currentLocation.Y > maxY Then currentLocation.Y = maxY

                    ' დავაყენოთ ახალი პოზიცია
                    draggedCard.Location = currentLocation

                    Debug.WriteLine($"SessionCard_MouseMove: ახალი პოზიცია - X={currentLocation.X}, Y={currentLocation.Y}")
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"SessionCard_MouseMove: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის ბარათზე MouseUp ივენთი
    ''' </summary>
    Private Sub SessionCard_MouseUp(sender As Object, e As MouseEventArgs)
        Try
            Debug.WriteLine($"SessionCard_MouseUp: მაუსის ღილაკი მოიხსნა, isDragging={isDragging}")

            If isDragging AndAlso draggedCard IsNot Nothing Then
                isDragging = False
                draggedCard.Cursor = Cursors.Hand
                draggedCard.Capture = False

                ' გამოვთვალოთ ახალი გრიდის პოზიცია
                Dim newGridPosition = CalculateGridPosition(draggedCard.Location)
                Dim originalGridPosition = CalculateGridPosition(originalCardPosition)

                Debug.WriteLine($"SessionCard_MouseUp: ორიგინალური პოზიცია - Space={originalGridPosition.spaceIndex}, Time={originalGridPosition.timeIndex}")
                Debug.WriteLine($"SessionCard_MouseUp: ახალი პოზიცია - Space={newGridPosition.spaceIndex}, Time={newGridPosition.timeIndex}")

                ' ვამოწმებთ შეიცვალა თუ არა პოზიცია
                If newGridPosition.spaceIndex <> originalGridPosition.spaceIndex OrElse
               newGridPosition.timeIndex <> originalGridPosition.timeIndex Then

                    ' მესიჯბოქსი დადასტურებისთვის
                    Dim newSpace As String = spaces(newGridPosition.spaceIndex)
                    Dim newTime As String = timeIntervals(newGridPosition.timeIndex).ToString("HH:mm")

                    Dim result As DialogResult = MessageBox.Show(
                    $"გსურთ შევცვალოთ სეანსის პარამეტრები?{Environment.NewLine}" &
                    $"ახალი სივრცე: {newSpace}{Environment.NewLine}" &
                    $"ახალი დრო: {newTime}",
                    "სეანსის განახლება",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1
                )

                    If result = DialogResult.Yes Then
                        ' დავადასტუროთ ცვლილება
                        ConfirmSessionMove(newGridPosition)
                    Else
                        ' დავაბრუნოთ ძველ ადგილას
                        RevertSessionMove()
                    End If
                Else
                    ' პოზიცია არ შეცვლილა
                    Debug.WriteLine("SessionCard_MouseUp: პოზიცია არ შეცვლილა")
                End If

                ' გავასუფთავოთ ცვლადები
                draggedCard = Nothing
                originalSessionData = Nothing
            End If
        Catch ex As Exception
            Debug.WriteLine($"SessionCard_MouseUp: შეცდომა - {ex.Message}")
            ' შეცდომის შემთხვევაში დავაბრუნოთ ძველ ადგილას
            RevertSessionMove()
        End Try
    End Sub

    ''' <summary>
    ''' გრიდის პოზიციის გამოთვლა pixel კოორდინატებიდან
    ''' </summary>
    Private Function CalculateGridPosition(pixelLocation As Point) As (spaceIndex As Integer, timeIndex As Integer)
        Try
            ' მასშტაბირებული პარამეტრები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' გამოვთვალოთ ინდექსები (მარჯინების გათვალისწინებით)
            Dim spaceIndex As Integer = pixelLocation.X \ SPACE_COLUMN_WIDTH
            Dim timeIndex As Integer = pixelLocation.Y \ ROW_HEIGHT

            ' შევამოწმოთ საზღვრები
            If spaceIndex < 0 Then spaceIndex = 0
            If spaceIndex >= spaces.Count Then spaceIndex = spaces.Count - 1
            If timeIndex < 0 Then timeIndex = 0
            If timeIndex >= timeIntervals.Count Then timeIndex = timeIntervals.Count - 1

            Debug.WriteLine($"CalculateGridPosition: Pixel({pixelLocation.X}, {pixelLocation.Y}) -> Grid({spaceIndex}, {timeIndex})")
            Return (spaceIndex, timeIndex)

        Catch ex As Exception
            Debug.WriteLine($"CalculateGridPosition: შეცდომა - {ex.Message}")
            Return (0, 0)
        End Try
    End Function

    ''' <summary>
    ''' სესიის მოძრაობის დადასტურება და მონაცემების განახლება
    ''' </summary>
    Private Sub ConfirmSessionMove(newGridPosition As (spaceIndex As Integer, timeIndex As Integer))
        Try
            Debug.WriteLine($"ConfirmSessionMove: ახალი პოზიცია - Space={newGridPosition.spaceIndex}, Time={newGridPosition.timeIndex}")

            If originalSessionData Is Nothing OrElse dataService Is Nothing Then
                Debug.WriteLine("ConfirmSessionMove: originalSessionData ან dataService არის null")
                RevertSessionMove()
                Return
            End If

            ' ახალი სივრცე და დრო
            Dim newSpace As String = spaces(newGridPosition.spaceIndex)
            Dim newDateTime As DateTime = timeIntervals(newGridPosition.timeIndex)

            ' შევქმნათ ახალი DateTime, რომელიც ინარჩუნებს ორიგინალურ თარიღს მაგრამ ცვლის დროს
            Dim originalDate As DateTime = originalSessionData.DateTime.Date
            Dim newFullDateTime As DateTime = originalDate.Add(newDateTime.TimeOfDay)

            ' განვაახლოთ სესიის ობიექტი
            originalSessionData.Space = newSpace
            originalSessionData.DateTime = newFullDateTime

            ' შევცდოთ მონაცემთა ბაზაში განახლება
            Try
                ' მივიღოთ ყველა სესიის მონაცემები
                Dim allSessionsData = dataService.GetData("DB-Schedule!A2:O")

                ' ვიპოვოთ ჩვენი სესიის მწკრივი
                For i As Integer = 0 To allSessionsData.Count - 1
                    Dim row = allSessionsData(i)
                    If row.Count > 0 AndAlso Integer.TryParse(row(0).ToString(), Nothing) Then
                        Dim rowId As Integer = Integer.Parse(row(0).ToString())
                        If rowId = originalSessionData.Id Then
                            ' განვაახლოთ სივრცე (K სვეტი - ინდექსი 10)
                            Dim updatedRow As New List(Of Object)(row)
                            If updatedRow.Count > 10 Then updatedRow(10) = newSpace

                            ' განვაახლოთ თარიღი (F სვეტი - ინდექსი 5)
                            If updatedRow.Count > 5 Then updatedRow(5) = newFullDateTime.ToString("dd.MM.yyyy HH:mm")

                            ' განვაახლოთ მონაცემები Google Sheets-ში
                            Dim updateRange As String = $"DB-Schedule!A{i + 2}:O{i + 2}"
                            dataService.UpdateData(updateRange, updatedRow)

                            Debug.WriteLine($"ConfirmSessionMove: წარმატებით განახლდა სესია ID={originalSessionData.Id}")
                            Exit For
                        End If
                    End If
                Next

                ' განვაახლოთ ლოკალური მონაცემები
                LoadSessions()

                ' განვაახლოთ კალენდრის ხედი
                UpdateCalendarView()

                ' შეტყობინება წარმატებული განახლების შესახებ
                MessageBox.Show($"სესია წარმატებით გადატანილია {newSpace} სივრცეში {newDateTime:HH:mm} დროზე",
                          "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch updateEx As Exception
                Debug.WriteLine($"ConfirmSessionMove: მონაცემების განახლების შეცდომა - {updateEx.Message}")
                MessageBox.Show($"სესიის განახლების შეცდომა: {updateEx.Message}",
                          "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' შეცდომის შემთხვევაში დავაბრუნოთ ძველ ადგილას
                RevertSessionMove()
            End Try

        Catch ex As Exception
            Debug.WriteLine($"ConfirmSessionMove: შეცდომა - {ex.Message}")
            RevertSessionMove()
        End Try
    End Sub

    ''' <summary>
    ''' სესიის ძველ ადგილას დაბრუნება
    ''' </summary>
    Private Sub RevertSessionMove()
        Try
            If draggedCard IsNot Nothing Then
                draggedCard.Location = originalCardPosition
                Debug.WriteLine("RevertSessionMove: ბარათი დაბრუნდა ძველ ადგილას")
            End If
        Catch ex As Exception
            Debug.WriteLine($"RevertSessionMove: შეცდომა - {ex.Message}")
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
