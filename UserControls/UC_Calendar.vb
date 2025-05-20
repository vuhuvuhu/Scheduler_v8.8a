' ===========================================
' 📄 UserControls/UC_Calendar.vb
' -------------------------------------------
' კალენდრის კონტროლი: გაუმჯობესებული ვერსია მეხსიერების პრობლემების გარეშე
' ოპტიმიზებული რესურსების მენეჯმენტით და გაუმჯობესებული სკროლინგით
' ===========================================
Imports System.ComponentModel
Imports System.Text
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

    ' სივრცეები და დროის ინტერვალები
    Private spaces As List(Of String) = New List(Of String)()
    Private timeIntervals As List(Of DateTime) = New List(Of DateTime)()
    Private gridCells(,) As Panel ' უჯრედები სესიების განთავსებისთვის

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
    Private Const BASE_DATE_COLUMN_WIDTH As Integer = 40 ' თარიღის ზოლის საბაზისო სიგანე - არ იცვლება მასშტაბით

    ' თერაპევტების სია
    Private therapists As List(Of String) = New List(Of String)()

    ' შენახული არჩეული ბენეფიციარები - იგი შენარჩუნდება მასშტაბირებისას
    Private preservedBeneficiaries As List(Of String) = New List(Of String)()

    ' ბოლოს არჩეული ხედის ტიპი - გამოიყენება ხედის აღდგენისას
    Private lastViewType As String = "სივრცე"

    ' სესიის ბარათის Mouse ივენთების ცვლადები გადატანისთვის
    Private isDragging As Boolean = False
    Private dragStartPoint As Point
    Private draggedCard As Panel = Nothing
    Private originalCardPosition As Point
    Private originalSessionData As SessionModel = Nothing

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

        ' კომბობოქსების შევსება
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
        AddHandler BtnHUp.Click, AddressOf BtnHUp_Click_Improved
        AddHandler BtnHDown.Click, AddressOf BtnHDown_Click_Improved
        AddHandler BtnVUp.Click, AddressOf BtnVUp_Click_Improved
        AddHandler BtnVDown.Click, AddressOf BtnVDown_Click_Improved

        ' რადიობუტონების საწყისი მნიშვნელობები
        rbDay.Checked = True
        RBSpace.Checked = True
        lastViewType = "სივრცე"

        ' ფილტრების ინიციალიზაცია
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
    ''' გასწორებული ვერსია თარიღის სვეტით და სივრცეების სათაურებით
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

            ' გამოვთვალოთ დროისა და თარიღის პანელების ზომები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim totalRowsHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            Debug.WriteLine($"InitializeDayViewPanels: ROW_HEIGHT={ROW_HEIGHT}, totalRowsHeight={totalRowsHeight}, vScale={vScale}")

            ' ======= 1. შევქმნათ თარიღის ზოლის პანელი (ყველაზე მარცხნივ) =======
            Dim dateColumnPanel As New Panel()
            dateColumnPanel.Name = "dateColumnPanel"
            dateColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left ' მნიშვნელოვანი დამატება!
            dateColumnPanel.AutoScroll = False
            dateColumnPanel.Size = New Size(BASE_DATE_COLUMN_WIDTH, totalRowsHeight)
            dateColumnPanel.Location = New Point(0, BASE_HEADER_HEIGHT)
            dateColumnPanel.BackColor = Color.FromArgb(60, 80, 150)
            dateColumnPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(dateColumnPanel)

            ' ======= 2. შევქმნათ საათების პანელი (თარიღის ზოლის მარჯვნივ) =======
            Dim timeColumnPanel As New Panel()
            timeColumnPanel.Name = "timeColumnPanel"
            timeColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left ' მნიშვნელოვანი დამატება!
            timeColumnPanel.AutoScroll = False
            timeColumnPanel.Size = New Size(BASE_TIME_COLUMN_WIDTH, totalRowsHeight)
            timeColumnPanel.Location = New Point(BASE_DATE_COLUMN_WIDTH, BASE_HEADER_HEIGHT)
            timeColumnPanel.BackColor = Color.FromArgb(240, 240, 245)
            timeColumnPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(timeColumnPanel)

            ' ======= 3. შევქმნათ თარიღის სათაურის პანელი (მარცხენა ზედა კუთხეში) =======
            Dim dateHeaderPanel As New Panel()
            dateHeaderPanel.Name = "dateHeaderPanel"
            dateHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left ' მნიშვნელოვანი დამატება!
            dateHeaderPanel.Size = New Size(BASE_DATE_COLUMN_WIDTH, BASE_HEADER_HEIGHT)
            dateHeaderPanel.Location = New Point(0, 0)
            dateHeaderPanel.BackColor = Color.FromArgb(40, 60, 120)
            dateHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(dateHeaderPanel)

            ' თარიღის სათაურის ლეიბლი
            Dim dateHeaderLabel As New Label()
            dateHeaderLabel.Size = New Size(BASE_DATE_COLUMN_WIDTH - 2, BASE_HEADER_HEIGHT - 2)
            dateHeaderLabel.Location = New Point(1, 1)
            dateHeaderLabel.TextAlign = ContentAlignment.MiddleCenter
            dateHeaderLabel.Text = "თარიღი"
            dateHeaderLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            dateHeaderLabel.ForeColor = Color.White
            dateHeaderPanel.Controls.Add(dateHeaderLabel)

            ' ======= 4. შევქმნათ დროის სათაურის პანელი (თარიღის სათაურის მარჯვნივ) =======
            Dim timeHeaderPanel As New Panel()
            timeHeaderPanel.Name = "timeHeaderPanel"
            timeHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left ' მნიშვნელოვანი დამატება!
            timeHeaderPanel.Size = New Size(BASE_TIME_COLUMN_WIDTH, BASE_HEADER_HEIGHT)
            timeHeaderPanel.Location = New Point(BASE_DATE_COLUMN_WIDTH, 0)
            timeHeaderPanel.BackColor = Color.FromArgb(180, 180, 220)
            timeHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(timeHeaderPanel)

            ' დროის სათაურის ლეიბლი
            Dim timeHeaderLabel As New Label()
            timeHeaderLabel.Size = New Size(BASE_TIME_COLUMN_WIDTH - 2, BASE_HEADER_HEIGHT - 2)
            timeHeaderLabel.Location = New Point(1, 1)
            timeHeaderLabel.TextAlign = ContentAlignment.MiddleCenter
            timeHeaderLabel.Text = "დრო"
            timeHeaderLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            timeHeaderPanel.Controls.Add(timeHeaderLabel)

            ' ======= 5. შევქმნათ სივრცეების სათაურების პანელი (ზემოთ) =======
            Dim spacesHeaderPanel As New Panel()
            spacesHeaderPanel.Name = "spacesHeaderPanel"
            spacesHeaderPanel.AutoScroll = False
            ' ზომა დამოკიდებული იქნება მასშტაბზე!
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim totalSpacesWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count
            spacesHeaderPanel.Size = New Size(totalSpacesWidth, BASE_HEADER_HEIGHT)
            spacesHeaderPanel.Location = New Point(BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH, 0)
            spacesHeaderPanel.BackColor = Color.FromArgb(220, 220, 240)
            spacesHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(spacesHeaderPanel)

            ' ======= 6. შევქმნათ მთავარი გრიდის პანელი =======
            Dim mainGridPanel As New Panel()
            mainGridPanel.Name = "mainGridPanel"
            mainGridPanel.Size = New Size(totalWidth - BASE_TIME_COLUMN_WIDTH - BASE_DATE_COLUMN_WIDTH, totalHeight - BASE_HEADER_HEIGHT)
            mainGridPanel.Location = New Point(BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH, BASE_HEADER_HEIGHT)
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

            ' შევამოწმოთ მონაცემთა სერვისი
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

            ' დავალაგოთ თერაპევტები სახელით
            therapists.AddRange(therapistSet.OrderBy(Function(t) t))

            ' თუ არ მოიძებნა თერაპევტები, დავამატოთ საცდელი
            If therapists.Count = 0 Then
                Debug.WriteLine("LoadTherapistsForDate: თერაპევტები არ მოიძებნა, ვამატებთ საცდელ მონაცემებს")
                therapists.AddRange({"ჩანაწერი არ არის"})
            End If

            Debug.WriteLine($"LoadTherapistsForDate: ჩატვირთულია {therapists.Count} თერაპევტი")
        Catch ex As Exception
            Debug.WriteLine($"LoadTherapistsForDate: შეცდომა - {ex.Message}")
            ' შეცდომის შემთხვევაში დავაყენოთ საცდელი თერაპევტები
            therapists.Clear()
            therapists.AddRange({"ჩანაწერი არ არის"})
        End Try
    End Sub

    ''' <summary>
    ''' სქროლის სინქრონიზაცია - უნიფიცირებული მიდგომა ყველა ტიპის ხედისთვის
    ''' </summary>
    Private Sub SetupScrollSynchronization()
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("SetupScrollSynchronization: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' ვიპოვოთ ყველა შესაძლო სათაურების პანელი
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)
            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            ' ჯერ მოვხსნათ არსებული ივენთის ჰენდლერები, თუ ისინი არსებობენ
            RemoveHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll
            RemoveHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            ' მივაბათ Scroll ივენთი ხელახლა
            AddHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' ასევე, დავამატოთ ივენთი გორგოლჭის სქროლისთვის
            AddHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            Debug.WriteLine("SetupScrollSynchronization: სქროლის სინქრონიზაცია დაყენებულია ყველა ხედისთვის")
        Catch ex As Exception
            Debug.WriteLine($"SetupScrollSynchronization: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის სქროლის ივენთი - უნიფიცირებული მიდგომა
    ''' </summary>
    Private Sub MainGridPanel_Scroll(sender As Object, e As ScrollEventArgs)
        Try
            Dim mainGridPanel As Panel = DirectCast(sender, Panel)

            ' ვიპოვოთ ყველა საჭირო პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)
            Dim dateColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("dateColumnPanel", False).FirstOrDefault(), Panel)
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)
            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

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
            End If

            ' ჰორიზონტალური სქროლი
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                Dim scrollOffset As Integer = -mainGridPanel.HorizontalScroll.Value
                Dim fixedLeftPosition As Integer = BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH

                ' სივრცეების სათაურები (თუ არსებობს და ხილვადია)
                If spacesHeaderPanel IsNot Nothing AndAlso spacesHeaderPanel.Visible Then
                    spacesHeaderPanel.Left = fixedLeftPosition + scrollOffset
                End If

                ' თერაპევტების სათაურები (თუ არსებობს და ხილვადია)
                If therapistsHeaderPanel IsNot Nothing AndAlso therapistsHeaderPanel.Visible Then
                    therapistsHeaderPanel.Left = fixedLeftPosition + scrollOffset
                End If

                ' ბენეფიციარების სათაურები (თუ არსებობს და ხილვადია)
                If beneficiaryHeaderPanel IsNot Nothing AndAlso beneficiaryHeaderPanel.Visible Then
                    beneficiaryHeaderPanel.Left = fixedLeftPosition + scrollOffset
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"MainGridPanel_Scroll: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მაუსის გორგოლჭის ივენთი - უნიფიცირებული მიდგომა
    ''' </summary>
    Private Sub MainGridPanel_MouseWheel(sender As Object, e As MouseEventArgs)
        Try
            Dim mainGridPanel As Panel = DirectCast(sender, Panel)

            ' ვიპოვოთ ყველა საჭირო პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)
            Dim dateColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("dateColumnPanel", False).FirstOrDefault(), Panel)
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)
            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            ' ვერტიკალური სქროლის სინქრონიზაცია
            Dim verticalScrollOffset As Integer = -mainGridPanel.VerticalScroll.Value

            If timeColumnPanel IsNot Nothing Then
                timeColumnPanel.Top = BASE_HEADER_HEIGHT + verticalScrollOffset
            End If

            If dateColumnPanel IsNot Nothing Then
                dateColumnPanel.Top = BASE_HEADER_HEIGHT + verticalScrollOffset
            End If

            ' ჰორიზონტალური სქროლის სინქრონიზაცია
            Dim horizontalScrollOffset As Integer = -mainGridPanel.HorizontalScroll.Value
            Dim fixedLeftPosition As Integer = BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH

            ' სივრცეების სათაურები (თუ არსებობს და ხილვადია)
            If spacesHeaderPanel IsNot Nothing AndAlso spacesHeaderPanel.Visible Then
                spacesHeaderPanel.Left = fixedLeftPosition + horizontalScrollOffset
            End If

            ' თერაპევტების სათაურები (თუ არსებობს და ხილვადია)
            If therapistsHeaderPanel IsNot Nothing AndAlso therapistsHeaderPanel.Visible Then
                therapistsHeaderPanel.Left = fixedLeftPosition + horizontalScrollOffset
            End If

            ' ბენეფიციარების სათაურები (თუ არსებობს და ხილვადია)
            If beneficiaryHeaderPanel IsNot Nothing AndAlso beneficiaryHeaderPanel.Visible Then
                beneficiaryHeaderPanel.Left = fixedLeftPosition + horizontalScrollOffset
            End If
        Catch ex As Exception
            Debug.WriteLine($"MainGridPanel_MouseWheel: შეცდომა - {ex.Message}")
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
    ''' თარიღის ზოლის პანელის შევსება არჩეული თარიღის ინფორმაციით
    ''' RotatedLabel-ს ზომა და ფონტი ჩასწორდა მასშტაბირების შენარჩუნებისთვის
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

            ' RotatedLabel-ს ზომა დამოკიდებული უნდა იყოს მასშტაბზე
            Dim rotatedDateLabel As New RotatedLabel()

            ' შევინახოთ სათანადო ზომები - მასშტაბის დამოუკიდებლად
            rotatedDateLabel.Size = New Size(dateColumnPanel.Width - 4, dateColumnPanel.Height - 4)
            rotatedDateLabel.Location = New Point(2, 2)
            rotatedDateLabel.BackColor = Color.FromArgb(60, 80, 150)
            rotatedDateLabel.ForeColor = Color.White

            ' შრიფტის ზომა ვერტიკალური მასშტაბის დამოუკიდებლად
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
            Next

            Debug.WriteLine($"FillTimeColumnPanel: დასრულდა - დაემატა {timeIntervals.Count} დროის ეტიკეტი")

        Catch ex As Exception
            Debug.WriteLine($"FillTimeColumnPanel: შეცდომა - {ex.Message}")
            Debug.WriteLine($"FillTimeColumnPanel: StackTrace: {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' სივრცეების სათაურების პანელის შევსება და ზომის დაყენება
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
                ElseIf FontFamily.Families.Any(Function(f) f.Name = "ALK_Tall_Mtavruli") Then
                    mtavruliFont = New Font("ALK_Tall_Mtavruli", 9, FontStyle.Bold)
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
            Next

            Debug.WriteLine($"FillSpacesHeaderPanel: დაემატა {spaces.Count} სივრცის სათაური, მასშტაბი: {hScale}")

        Catch ex As Exception
            Debug.WriteLine($"FillSpacesHeaderPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის შევსება - გაუმჯობესებული ვერსია ორმაგი სკროლის თავიდან აცილებისთვის
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

            ' გავასუფთავოთ მთავარი პანელი
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

                    ' ვამატებთ უჯრედს პირდაპირ mainGridPanel-ზე
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
    ''' უჯრედის ჩარჩოს დახატვა
    ''' </summary>
    Private Sub Cell_Paint(sender As Object, e As PaintEventArgs)
        Dim cell As Panel = DirectCast(sender, Panel)

        Using pen As New Pen(Color.FromArgb(200, 200, 200), 1)
            ' მხოლოდ ვერტიკალური ხაზები - მარჯვენა ბორდერი
            e.Graphics.DrawLine(pen, cell.Width - 1, 0, cell.Width - 1, cell.Height - 1)
        End Using
    End Sub

    ''' <summary>
    ''' უჯრედების ძველი ბარათების გასუფთავება - გაუმჯობესებული ვერსია
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
                            ' გავასუფთავოთ ყველა სესიის ბარათის ივენთი
                            For Each ctrl In cell.Controls.OfType(Of Panel)().ToList()
                                RemoveAllCardHandlers(ctrl)
                                cell.Controls.Remove(ctrl)
                                ctrl.Dispose()
                            Next
                        End If
                    Next
                Next
            End If

            ' 2. ასევე ვშლით ბარათებს mainGridPanel-დანაც
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            If mainGridPanel IsNot Nothing Then
                ' ვშლით ყველა ბარათს, რომელთაც Tag აქვთ (სესიის ID)
                Dim cardsToRemove As New List(Of Control)()
                For Each ctrl As Control In mainGridPanel.Controls
                    If TypeOf ctrl Is Panel AndAlso ctrl.Tag IsNot Nothing Then
                        ' ეს სესიის ბარათია - მოვხსნათ ივენთები და დავამატოთ წასაშლელთა სიაში
                        RemoveAllCardHandlers(ctrl)
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
    ''' სესიის ბარათის ყველა ივენთის მოხსნა
    ''' </summary>
    Private Sub RemoveAllCardHandlers(card As Control)
        Try
            If card Is Nothing Then Return

            ' სესიის ბარათის ივენთების მოხსნა
            RemoveHandler card.MouseDown, AddressOf SessionCard_MouseDown
            RemoveHandler card.MouseMove, AddressOf SessionCard_MouseMove
            RemoveHandler card.MouseUp, AddressOf SessionCard_MouseUp
            RemoveHandler card.Click, AddressOf SessionCard_Click
            RemoveHandler card.DoubleClick, AddressOf SessionCard_DoubleClick

            ' რედაქტირების ღილაკის ივენთის მოხსნა
            For Each ctrl In card.Controls.OfType(Of Button)()
                RemoveHandler ctrl.Click, AddressOf BtnEditSession_Click
            Next

            ' უჯრედის დახატვის ივენთის მოხსნა
            RemoveHandler card.Paint, AddressOf SessionCard_Paint
        Catch ex As Exception
            Debug.WriteLine($"RemoveAllCardHandlers: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კალენდრის განახლებული ხედის მთავარი მეთოდი - გაუმჯობესებული ლოგიკით
    ''' არჩეული ბენეფიციარების შენახვით და სწორი განახლების ლოგიკით
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

            ' ბენეფიციარების შენახვა მასშტაბირებამდე თუ ბენეფიციარის ხედში ვიმყოფებით
            If rbDay.Checked AndAlso RBBene.Checked Then
                preservedBeneficiaries = GetBeneficiaryColumns()
                Debug.WriteLine($"UpdateCalendarView: შენახულია {preservedBeneficiaries.Count} ბენეფიციარი")
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
                    lastViewType = "სივრცე"
                ElseIf RBPer.Checked Then
                    ' თერაპევტების მიხედვით
                    ShowDayViewByTherapist()
                    lastViewType = "თერაპევტი"
                ElseIf RBBene.Checked Then
                    ' ბენეფიციარების მიხედვით
                    If preservedBeneficiaries.Count > 0 Then
                        ShowDayViewByBeneficiaryWithPreservation(preservedBeneficiaries)
                    Else
                        ShowDayViewByBeneficiary()
                    End If
                    lastViewType = "ბენეფიციარი"
                End If
            ElseIf rbWeek.Checked Then
                ' კვირის ხედი
                ShowWeekView()
                lastViewType = "კვირა"
            ElseIf rbMonth.Checked Then
                ' თვის ხედი
                ShowMonthView()
                lastViewType = "თვე"
            End If

            ' გამოვათავისუფლოთ სესიაზე მაუსის გადატანის რესურსები
            draggedCard = Nothing
            originalSessionData = Nothing
        Catch ex As Exception
            Debug.WriteLine($"UpdateCalendarView: შეცდომა - {ex.Message}")
            Debug.WriteLine($"UpdateCalendarView: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"კალენდრის ხედის განახლების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დღის ხედის ჩვენება სივრცეების მიხედვით - გასწორებული ვერსია
    ''' </summary>
    Private Sub ShowDayViewBySpace()
        Try
            Debug.WriteLine("ShowDayViewBySpace: დაიწყო დღის ხედის ჩვენება სივრცეების მიხედვით")

            ' დავაყენოთ pnlFilter-ის ფერი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' 1. პანელების ინიციალიზაცია
            InitializeDayViewPanels()

            ' 2. თარიღის ზოლის შევსება
            FillDateColumnPanel()

            ' 3. დროის პანელის შევსება
            FillTimeColumnPanel()

            ' 4. სივრცეების სათაურების პანელის შევსება
            FillSpacesHeaderPanel()

            ' 5. მთავარი გრიდის პანელის შევსება
            FillMainGridPanel()

            ' 6. სინქრონიზაცია სქროლისთვის
            SetupScrollSynchronization()

            ' 7. სესიების განთავსება გრიდში
            PlaceSessionsOnGrid()

            Debug.WriteLine("ShowDayViewBySpace: დღის ხედის ჩვენება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewBySpace: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewBySpace: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"დღის ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დღის ხედის ჩვენება თერაპევტების მიხედვით - გასწორებული ვერსია
    ''' </summary>
    Private Sub ShowDayViewByTherapist()
        Try
            Debug.WriteLine("ShowDayViewByTherapist: დაიწყო დღის ხედის ჩვენება თერაპევტების მიხედვით")

            ' 1. თერაპევტების ჩატვირთვა მოცემული თარიღისთვის
            LoadTherapistsForDate()

            ' დავაყენოთ pnlFilter-ის ფერი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' 2. პანელების ინიციალიზაცია
            InitializeDayViewPanelsForTherapists()

            ' 3. თარიღის ზოლის შევსება
            FillDateColumnPanel()

            ' 4. დროის პანელის შევსება
            FillTimeColumnPanel()

            ' 5. თერაპევტების სათაურების პანელის შევსება
            FillTherapistsHeaderPanel()

            ' 6. მთავარი გრიდის პანელის შევსება თერაპევტების მიხედვით
            FillMainGridPanelForTherapists()

            ' 7. სკროლის სინქრონიზაცია
            SetupScrollSynchronization()

            ' 8. სესიების განთავსება გრიდში თერაპევტების მიხედვით
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
    ''' </summary>
    Private Sub InitializeDayViewPanelsForTherapists()
        Try
            Debug.WriteLine("InitializeDayViewPanelsForTherapists: დაიწყო პანელების ინიციალიზაცია თერაპევტებისთვის")

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

            ' ======= 5. შევქმნათ თერაპევტების სათაურების პანელი =======
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

            ' ======= 6. შევქმნათ მთავარი გრიდის პანელი =======
            Dim mainGridPanel As New Panel()
            mainGridPanel.Name = "mainGridPanel"
            mainGridPanel.Size = New Size(totalWidth - TIME_COLUMN_WIDTH - DATE_COLUMN_WIDTH, totalHeight - HEADER_HEIGHT)
            mainGridPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            mainGridPanel.BackColor = Color.White
            mainGridPanel.BorderStyle = BorderStyle.FixedSingle
            mainGridPanel.AutoScroll = True
            pnlCalendarGrid.Controls.Add(mainGridPanel)

            Debug.WriteLine("InitializeDayViewPanelsForTherapists: პანელების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"InitializeDayViewPanelsForTherapists: შეცდომა - {ex.Message}")
            Debug.WriteLine($"InitializeDayViewPanelsForTherapists: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"თერაპევტების პანელების ინიციალიზაციის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტების სათაურების პანელის შევსება თერაპევტების სახელებით
    ''' </summary>
    Private Sub FillTherapistsHeaderPanel()
        Try
            Debug.WriteLine("FillTherapistsHeaderPanel: დაიწყო თერაპევტების სათაურების შევსება")

            ' ვიპოვოთ თერაპევტების სათაურების პანელი
            Dim therapistsHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("therapistsHeaderPanel", False).FirstOrDefault(), Panel)

            If therapistsHeaderPanel Is Nothing Then
                Debug.WriteLine("FillTherapistsHeaderPanel: თერაპევტების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' გავასუფთავოთ
            therapistsHeaderPanel.Controls.Clear()

            ' თუ თერაპევტები არ არის ჩატვირთული
            If therapists.Count = 0 Then
                LoadTherapistsForDate()
            End If

            Debug.WriteLine($"FillTherapistsHeaderPanel: ჩატვირთული თერაპევტების რაოდენობა: {therapists.Count}")

            ' თერაპევტის სვეტის სიგანე (იმასშტაბირდება)
            Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' გამოვთვალოთ სათაურების პანელის სრული სიგანე
            Dim totalHeaderWidth As Integer = THERAPIST_COLUMN_WIDTH * therapists.Count

            ' შევცვალოთ პანელის სიგანე მასშტაბის მიხედვით
            therapistsHeaderPanel.Width = totalHeaderWidth

            Debug.WriteLine($"FillTherapistsHeaderPanel: პანელის სიგანე: {therapistsHeaderPanel.Width}, THERAPIST_COLUMN_WIDTH: {THERAPIST_COLUMN_WIDTH}")

            ' ღია მწვანე ფონტი თერაპევტებისთვის
            Dim therapistFont As New Font("Sylfaen", 9, FontStyle.Bold)

            ' თერაპევტების სათაურების შექმნა
            For i As Integer = 0 To therapists.Count - 1
                Dim therapistHeader As New Label()
                therapistHeader.Size = New Size(THERAPIST_COLUMN_WIDTH - 1, therapistsHeaderPanel.Height - 2)
                therapistHeader.Location = New Point(i * THERAPIST_COLUMN_WIDTH, 1)
                therapistHeader.BackColor = Color.FromArgb(80, 120, 80) ' მუქი მწვანე
                therapistHeader.ForeColor = Color.White
                therapistHeader.TextAlign = ContentAlignment.MiddleCenter
                therapistHeader.Text = therapists(i)
                therapistHeader.Font = therapistFont
                therapistHeader.BorderStyle = BorderStyle.FixedSingle

                ' შევინახოთ თავდაყირო X პოზიცია Tag-ში სქროლის სინქრონიზაციისთვის
                therapistHeader.Tag = i * THERAPIST_COLUMN_WIDTH

                therapistsHeaderPanel.Controls.Add(therapistHeader)

                Debug.WriteLine($"FillTherapistsHeaderPanel: დაემატა თერაპევტის სათაური [{i}]: {therapists(i)}, Tag={therapistHeader.Tag}")
            Next

            Debug.WriteLine($"FillTherapistsHeaderPanel: დაემატა {therapists.Count} თერაპევტის სათაური, მასშტაბი: {hScale}")

        Catch ex As Exception
            Debug.WriteLine($"FillTherapistsHeaderPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის შევსება თერაპევტებისთვის
    ''' </summary>
    Private Sub FillMainGridPanelForTherapists()
        Try
            Debug.WriteLine("FillMainGridPanelForTherapists: დაიწყო მთავარი გრიდის შევსება თერაპევტებისთვის")

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("FillMainGridPanelForTherapists: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' გავასუფთავოთ მთავარი პანელი
            mainGridPanel.Controls.Clear()

            ' შევამოწმოთ თერაპევტები და დრო
            If therapists.Count = 0 Then
                LoadTherapistsForDate()
            End If

            If timeIntervals.Count = 0 Then
                InitializeTimeIntervals()
            End If

            ' პარამეტრები - მასშტაბირების გამოყენებით
            Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)

            ' გრიდის სრული სიგანე
            Dim gridWidth As Integer = THERAPIST_COLUMN_WIDTH * therapists.Count

            ' გრიდის სრული სიმაღლე
            Dim gridHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            ' მნიშვნელოვანი: დავაყენოთ mainGridPanel-ის AutoScrollMinSize
            mainGridPanel.AutoScrollMinSize = New Size(gridWidth, gridHeight)

            Debug.WriteLine($"FillMainGridPanelForTherapists: გრიდის ზომები - Width={gridWidth}, Height={gridHeight}, hScale={hScale}, vScale={vScale}")

            ' უჯრედების მასივი (ახლა თერაპევტებისთვის)
            ReDim gridCells(therapists.Count - 1, timeIntervals.Count - 1)

            ' უჯრედების შექმნა - პირდაპირ mainGridPanel-ზე
            For col As Integer = 0 To therapists.Count - 1
                For row As Integer = 0 To timeIntervals.Count - 1
                    Dim cell As New Panel()
                    cell.Size = New Size(THERAPIST_COLUMN_WIDTH, ROW_HEIGHT)
                    cell.Location = New Point(col * THERAPIST_COLUMN_WIDTH, row * ROW_HEIGHT)

                    ' ალტერნატიული ფერები
                    If row Mod 2 = 0 Then
                        cell.BackColor = Color.FromArgb(250, 255, 250) ' ოდნავ მწვანე ნაცარი
                    Else
                        cell.BackColor = Color.FromArgb(245, 250, 245) ' უფრო მწვანე ნაცარი
                    End If

                    ' უჯრედის ჩარჩო
                    AddHandler cell.Paint, AddressOf Cell_Paint

                    ' ვინახავთ უჯრედის ობიექტს მასივში
                    gridCells(col, row) = cell

                    ' ვამატებთ უჯრედს პირდაპირ mainGridPanel-ზე
                    mainGridPanel.Controls.Add(cell)
                Next
            Next

            Debug.WriteLine($"FillMainGridPanelForTherapists: შეიქმნა {therapists.Count * timeIntervals.Count} უჯრედი თერაპევტებისთვის")

        Catch ex As Exception
            Debug.WriteLine($"FillMainGridPanelForTherapists: შეცდომა - {ex.Message}")
            Debug.WriteLine($"FillMainGridPanelForTherapists: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის ხედის სათაურების პანელის ინიციალიზაცია - გაუმჯობესებული ვერსია
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
                beneficiaryHeaderPanel.BackColor = Color.FromArgb(255, 230, 180) ' ღია ყვითელი ფერი ბენეფიციარებისთვის
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

            ' გავასუფთავოთ პანელი - მოვხსნათ ივენთები
            For Each ctrl In beneficiaryHeaderPanel.Controls.OfType(Of ComboBox)().ToList()
                RemoveHandler ctrl.SelectedIndexChanged, AddressOf BeneficiaryComboBox_SelectedIndexChanged
            Next
            For Each ctrl In beneficiaryHeaderPanel.Controls.OfType(Of Button)().ToList()
                RemoveHandler ctrl.Click, AddressOf BeneficiaryDeleteButton_Click
            Next
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
            FillBeneficiaryComboBoxImproved(comboBox)

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
    ''' გაუმჯობესებული ფუნქცია ბენეფიციარების კომბობოქსის შევსებისთვის
    ''' გამორიცხავს უკვე არჩეულ ბენეფიციარებს
    ''' </summary>
    Private Sub FillBeneficiaryComboBoxImproved(comboBox As ComboBox)
        Try
            Debug.WriteLine("FillBeneficiaryComboBoxImproved: ვივსებ კომბობოქსს გაუმჯობესებული ლოგიკით")

            ' გავასუფთავოთ კომბობოქსი
            comboBox.Items.Clear()

            ' საწყისო ღია ეტიკეტი
            comboBox.Items.Add("-- აირჩიეთ ბენეფიციარი --")

            ' მიმდინარე თარიღი
            Dim selectedDate As DateTime = DTPCalendar.Value.Date

            ' ყველა ბენეფიციარი, რომელსაც მოცემულ დღეს სესია აქვს
            Dim todaysBeneficiaries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If allSessions IsNot Nothing Then
                For Each session In allSessions
                    If session.DateTime.Date = selectedDate Then
                        Dim fullName = $"{session.BeneficiaryName.Trim()} {session.BeneficiarySurname.Trim()}"
                        todaysBeneficiaries.Add(fullName)
                    End If
                Next
            End If

            ' მივიღოთ არჩეული ბენეფიციარები
            Dim selectedBeneficiaries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel = DirectCast(comboBox.Parent, Panel)

            If beneficiaryHeaderPanel IsNot Nothing Then
                ' პირველ რიგში ვკრიბავთ ლეიბლებიდან
                For Each ctrl As Control In beneficiaryHeaderPanel.Controls
                    If TypeOf ctrl Is Label AndAlso ctrl.Name.StartsWith("lblBene_") Then
                        Dim label As Label = DirectCast(ctrl, Label)
                        If label.Tag IsNot Nothing Then
                            selectedBeneficiaries.Add(label.Tag.ToString())
                        End If
                    End If
                Next
            End If

            Debug.WriteLine($"FillBeneficiaryComboBoxImproved: დღეს {todaysBeneficiaries.Count} ბენეფიციარი, უკვე არჩეულია {selectedBeneficiaries.Count}")

            ' დარჩენილი ბენეფიციარების დამატება კომბობოქსში
            For Each beneficiary In todaysBeneficiaries.OrderBy(Function(b) b)
                If Not selectedBeneficiaries.Contains(beneficiary) Then
                    comboBox.Items.Add(beneficiary)
                End If
            Next

            ' საწყისი მითითება პირველ ელემენტზე
            If comboBox.Items.Count > 0 Then
                comboBox.SelectedIndex = 0
            End If

            Debug.WriteLine($"FillBeneficiaryComboBoxImproved: კომბობოქსი შევსებულია {comboBox.Items.Count} ელემენტით")

        Catch ex As Exception
            Debug.WriteLine($"FillBeneficiaryComboBoxImproved: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კომბობოქსის ჩანაცვლება ლეიბლით და წაშლის ღილაკით
    ''' </summary>
    Private Sub ReplaceComboBoxWithLabelAndDeleteButton(comboBox As ComboBox, selectedBeneficiary As String)
        Try
            Debug.WriteLine($"ReplaceComboBoxWithLabelAndDeleteButton: ვცვლი კომბობოქსს ლეიბლით - {selectedBeneficiary}")

            Dim columnIndex As Integer = CInt(comboBox.Tag)
            Dim parent As Panel = DirectCast(comboBox.Parent, Panel)

            ' კომბობოქსის პოზიცია და ზომები
            Dim comboLocation = comboBox.Location
            Dim comboSize = comboBox.Size

            ' მოვხსნათ და გავასუფთავოთ კომბობოქსი
            RemoveHandler comboBox.SelectedIndexChanged, AddressOf BeneficiaryComboBox_SelectedIndexChanged
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
    ''' მთავარი გრიდის პანელის შევსება ბენეფიციარებისთვის
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

            ' მივიღოთ არჩეული ბენეფიციარების სია
            Dim selectedBeneficiaries = GetBeneficiaryColumns()

            If selectedBeneficiaries.Count = 0 AndAlso preservedBeneficiaries.Count > 0 Then
                selectedBeneficiaries = preservedBeneficiaries
                Debug.WriteLine($"FillMainGridPanelForBeneficiaries: გამოვიყენეთ შენახული ბენეფიციარები - {selectedBeneficiaries.Count}")
            End If

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
    ''' ბენეფიციარის ხედის ჩვენება - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Sub ShowDayViewByBeneficiary()
        Try
            Debug.WriteLine("ShowDayViewByBeneficiary: დაიწყო ბენეფიციარების ხედის ჩვენება")

            ' დავაყენოთ pnlFilter-ის ფერი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' 1. პანელების ინიციალიზაცია
            InitializeBeneficiaryDayViewPanels()

            ' 2. თარიღის ზოლის შევსება
            FillDateColumnPanel()

            ' 3. დროის პანელის შევსება
            FillTimeColumnPanel()

            ' 4. ბენეფიციარების სათაურების პანელის ინიციალიზაცია
            InitializeBeneficiaryHeaderPanel()

            ' 5. მთავარი გრიდის პანელის შევსება ბენეფიციარებისთვის
            FillMainGridPanelForBeneficiaries()

            ' 6. სინქრონიზაცია სქროლისთვის
            SetupScrollSynchronization()

            ' 7. სესიების განთავსება ბენეფიციარების გრიდში
            PlaceSessionsOnBeneficiaryGrid()

            Debug.WriteLine("ShowDayViewByBeneficiary: ბენეფიციარების ხედის ჩვენება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewByBeneficiary: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewByBeneficiary: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"ბენეფიციარების ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' ბენეფიციარის კომბობოქსის შევსება (გაუმჯობესებული ვერსია)
    ''' არსებულ კომბობოქსში ჩატვირთავს ბენეფიციარებს, გაფილტრავს არჩეულებს
    ''' </summary>
    ''' <param name="comboBox">შესავსები კომბობოქსი</param>
    Private Sub FillBeneficiaryComboBox(comboBox As ComboBox)
        Try
            Debug.WriteLine("FillBeneficiaryComboBox: ვივსებ კომბობოქსს ბენეფიციარებით")

            ' გავასუფთავოთ კომბობოქსი
            comboBox.Items.Clear()

            ' საწყისი ტექსტი
            comboBox.Items.Add("-- აირჩიეთ ბენეფიციარი --")

            ' არჩეული თარიღი კალენდრიდან
            Dim selectedDate As DateTime = DTPCalendar.Value.Date

            ' მოვძებნოთ ყველა ბენეფიციარი, ვისაც ამ თარიღში აქვს სესია
            Dim dateSessionBeneficiaries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If allSessions IsNot Nothing Then
                For Each session In allSessions
                    If session.DateTime.Date = selectedDate Then
                        Dim fullName = $"{session.BeneficiaryName.Trim()} {session.BeneficiarySurname.Trim()}"
                        dateSessionBeneficiaries.Add(fullName)
                    End If
                Next
            End If

            Debug.WriteLine($"FillBeneficiaryComboBox: ნაპოვნია {dateSessionBeneficiaries.Count} ბენეფიციარი არჩეული თარიღისთვის")

            ' მოვძებნოთ უკვე არჩეული ბენეფიციარები (არ უნდა გამეორდეს კომბობოქსში)
            Dim selectedBeneficiaries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel = DirectCast(comboBox.Parent, Panel)

            If beneficiaryHeaderPanel IsNot Nothing Then
                ' ვკრიბავთ ლეიბლებიდან ბენეფიციარებს
                For Each ctrl As Control In beneficiaryHeaderPanel.Controls
                    If TypeOf ctrl Is Label AndAlso ctrl.Name.StartsWith("lblBene_") Then
                        Dim label As Label = DirectCast(ctrl, Label)
                        If label.Tag IsNot Nothing Then
                            selectedBeneficiaries.Add(label.Tag.ToString())
                            Debug.WriteLine($"FillBeneficiaryComboBox: უკვე არჩეულია '{label.Tag.ToString()}'")
                        End If
                    End If
                Next
            End If

            ' დარჩენილი ბენეფიციარების დამატება კომბობოქსში ანბანური რიგით
            For Each beneficiary In dateSessionBeneficiaries.OrderBy(Function(b) b)
                ' თუ უკვე არ არის არჩეული, დავამატოთ
                If Not selectedBeneficiaries.Contains(beneficiary) Then
                    comboBox.Items.Add(beneficiary)
                    Debug.WriteLine($"FillBeneficiaryComboBox: დავამატე კომბობოქსში '{beneficiary}'")
                End If
            Next

            ' საწყისი მნიშვნელობა (უსაფრთხოდ)
            If comboBox.Items.Count > 0 Then
                comboBox.SelectedIndex = 0
            End If

            Debug.WriteLine($"FillBeneficiaryComboBox: კომბობოქსი შევსებულია, ელემენტების რაოდენობა: {comboBox.Items.Count}")

        Catch ex As Exception
            Debug.WriteLine($"FillBeneficiaryComboBox: შეცდომა - {ex.Message}")
            ' შეცდომის შემთხვევაში, ვცდილობთ მინიმუმ საწყისი ელემენტის დამატებას
            If comboBox.Items.Count = 0 Then
                comboBox.Items.Add("-- აირჩიეთ ბენეფიციარი --")
                comboBox.SelectedIndex = 0
            End If
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის კომბობოქსიდან არჩევის ივენთი
    ''' არჩევის შემდეგ კომბობოქსს ჩაანაცვლებს ლეიბლით და დაამატებს შემდეგ კომბობოქსს
    ''' </summary>
    Private Sub BeneficiaryComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Dim comboBox As ComboBox = DirectCast(sender, ComboBox)
            Dim columnIndex As Integer = CInt(comboBox.Tag)

            Debug.WriteLine($"BeneficiaryComboBox_SelectedIndexChanged: ინდექსი {comboBox.SelectedIndex}, ტექსტი: '{comboBox.SelectedItem}'")

            ' თუ არჩეული არ არის ნამდვილი ბენეფიციარი, გამოვიდეთ
            If comboBox.SelectedIndex <= 0 OrElse comboBox.SelectedItem.ToString().StartsWith("--") Then
                Debug.WriteLine("BeneficiaryComboBox_SelectedIndexChanged: არჩეული არ არის ბენეფიციარი, გამოვდივარ")
                Return
            End If

            ' არჩეული ბენეფიციარის მიღება
            Dim selectedBeneficiary As String = comboBox.SelectedItem.ToString()

            ' ვცვლით კომბობოქსს ლეიბლით + წაშლის ღილაკით
            ReplaceComboBoxWithLabelAndDeleteButton(comboBox, selectedBeneficiary)

            ' ახალი კომბობოქსის შექმნა შემდეგი სვეტისთვის
            CreateNextBeneficiaryColumn()

            ' განვაახლოთ გრიდი ახალი ბენეფიციარისთვის
            UpdateBeneficiaryGrid()

            Debug.WriteLine($"BeneficiaryComboBox_SelectedIndexChanged: დამატებულია ბენეფიციარი '{selectedBeneficiary}'")

        Catch ex As Exception
            Debug.WriteLine($"BeneficiaryComboBox_SelectedIndexChanged: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BeneficiaryComboBox_SelectedIndexChanged: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის წაშლის ღილაკის კლიკის დამუშავება
    ''' </summary>
    Private Sub BeneficiaryDeleteButton_Click(sender As Object, e As EventArgs)
        Try
            Dim btnDelete As Button = DirectCast(sender, Button)
            Dim columnIndex As Integer = CInt(btnDelete.Tag)

            Debug.WriteLine($"BeneficiaryDeleteButton_Click: წაშლის მოთხოვნა, სვეტი {columnIndex}")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Debug.WriteLine("BeneficiaryDeleteButton_Click: beneficiaryHeaderPanel ვერ მოიძებნა")
                Return
            End If

            ' ვიპოვოთ შესაბამისი ლეიბლი (ბენეფიციარის სახელი)
            Dim lblToRemove As Label = DirectCast(beneficiaryHeaderPanel.Controls.Find($"lblBene_{columnIndex}", False).FirstOrDefault(), Label)
            Dim btnToRemove As Button = DirectCast(beneficiaryHeaderPanel.Controls.Find($"btnDelBene_{columnIndex}", False).FirstOrDefault(), Button)

            ' შევინახოთ ბენეფიციარის სახელი დებაგირებისთვის
            Dim beneficiaryName As String = If(lblToRemove?.Tag?.ToString(), "უცნობი")

            ' ვშლით ლეიბლს და ღილაკს
            If lblToRemove IsNot Nothing Then
                beneficiaryHeaderPanel.Controls.Remove(lblToRemove)
                lblToRemove.Dispose()
            End If

            If btnToRemove IsNot Nothing Then
                beneficiaryHeaderPanel.Controls.Remove(btnToRemove)
                btnToRemove.Dispose()
            End If

            ' ვაკეთებთ სვეტების რეორგანიზაციას
            ReorganizeBeneficiaryColumns()

            ' ვანახლებთ გრიდს და სესიების ბარათებს
            UpdateBeneficiaryGrid()

            ' ვანახლებთ კომბობოქსებს ყველგან
            RefreshAllBeneficiaryComboBoxes()

            Debug.WriteLine($"BeneficiaryDeleteButton_Click: წაიშალა ბენეფიციარი '{beneficiaryName}', სვეტი {columnIndex}")

        Catch ex As Exception
            Debug.WriteLine($"BeneficiaryDeleteButton_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BeneficiaryDeleteButton_Click: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი ფუნქცია, რომელიც ჰორიზონტალურ მასშტაბირებას ახორციელებს
    ''' </summary>
    Private Sub BtnHUp_Click(sender As Object, e As EventArgs) Handles BtnHUp.Click
        Try
            Debug.WriteLine($"BtnHUp_Click: მიმდინარე მასშტაბი hScale={hScale:F2}")

            ' შევინახოთ ბენეფიციარის ხედი, თუ გვაქვს
            Dim preserveBeneficiaries As Boolean = False
            Dim beneficiaryList As New List(Of String)

            ' ვამოწმებთ, ვართ თუ არა ბენეფიციარის ხედში
            If rbDay.Checked AndAlso RBBene.Checked Then
                preserveBeneficiaries = True
                beneficiaryList = GetBeneficiaryColumns()
                Debug.WriteLine($"BtnHUp_Click: ბენეფიციარის ხედის შენახვა: {beneficiaryList.Count} ბენეფიციარი")
            End If

            ' მასშტაბის გაზრდა
            Dim oldHScale As Double = hScale
            hScale += SCALE_STEP
            If hScale > MAX_SCALE Then hScale = MAX_SCALE

            ' მასშტაბის ლეიბლების განახლება
            UpdateScaleLabels()

            Debug.WriteLine($"BtnHUp_Click: მასშტაბი შეიცვალა {oldHScale:F2} -> {hScale:F2}")

            ' ვანახლებთ კალენდრის ხედს შენახული ბენეფიციარებით, თუ საჭიროა
            If preserveBeneficiaries AndAlso beneficiaryList.Count > 0 Then
                ShowDayViewByBeneficiaryWithPreservation(beneficiaryList)
                Debug.WriteLine($"BtnHUp_Click: შენარჩუნებულია {beneficiaryList.Count} ბენეფიციარის ხედი")
            Else
                UpdateCalendarView()
                Debug.WriteLine("BtnHUp_Click: განახლებულია კალენდრის ხედი სტანდარტულად")
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnHUp_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnHUp_Click: StackTrace - {ex.StackTrace}")
            ' შეცდომის შემთხვევაში ვცდილობთ მაინც განახლებას
            UpdateCalendarView()
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი ფუნქცია, რომელიც ჰორიზონტალურ მასშტაბირებას ამცირებს
    ''' </summary>
    Private Sub BtnHDown_Click(sender As Object, e As EventArgs) Handles BtnHDown.Click
        Try
            Debug.WriteLine($"BtnHDown_Click: მიმდინარე მასშტაბი hScale={hScale:F2}")

            ' შევინახოთ ბენეფიციარის ხედი, თუ გვაქვს
            Dim preserveBeneficiaries As Boolean = False
            Dim beneficiaryList As New List(Of String)

            ' ვამოწმებთ, ვართ თუ არა ბენეფიციარის ხედში
            If rbDay.Checked AndAlso RBBene.Checked Then
                preserveBeneficiaries = True
                beneficiaryList = GetBeneficiaryColumns()
                Debug.WriteLine($"BtnHDown_Click: ბენეფიციარის ხედის შენახვა: {beneficiaryList.Count} ბენეფიციარი")
            End If

            ' მასშტაბის შემცირება
            Dim oldHScale As Double = hScale
            hScale -= SCALE_STEP
            If hScale < MIN_SCALE Then hScale = MIN_SCALE

            ' მასშტაბის ლეიბლების განახლება
            UpdateScaleLabels()

            Debug.WriteLine($"BtnHDown_Click: მასშტაბი შეიცვალა {oldHScale:F2} -> {hScale:F2}")

            ' ვანახლებთ კალენდრის ხედს შენახული ბენეფიციარებით, თუ საჭიროა
            If preserveBeneficiaries AndAlso beneficiaryList.Count > 0 Then
                ShowDayViewByBeneficiaryWithPreservation(beneficiaryList)
                Debug.WriteLine($"BtnHDown_Click: შენარჩუნებულია {beneficiaryList.Count} ბენეფიციარის ხედი")
            Else
                UpdateCalendarView()
                Debug.WriteLine("BtnHDown_Click: განახლებულია კალენდრის ხედი სტანდარტულად")
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnHDown_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnHDown_Click: StackTrace - {ex.StackTrace}")
            ' შეცდომის შემთხვევაში ვცდილობთ მაინც განახლებას
            UpdateCalendarView()
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი ფუნქცია, რომელიც ვერტიკალურ მასშტაბირებას ზრდის
    ''' </summary>
    Private Sub BtnVUp_Click(sender As Object, e As EventArgs) Handles BtnVUp.Click
        Try
            Debug.WriteLine($"BtnVUp_Click: მიმდინარე მასშტაბი vScale={vScale:F2}")

            ' შევინახოთ ბენეფიციარის ხედი, თუ გვაქვს
            Dim preserveBeneficiaries As Boolean = False
            Dim beneficiaryList As New List(Of String)

            ' ვამოწმებთ, ვართ თუ არა ბენეფიციარის ხედში
            If rbDay.Checked AndAlso RBBene.Checked Then
                preserveBeneficiaries = True
                beneficiaryList = GetBeneficiaryColumns()
                Debug.WriteLine($"BtnVUp_Click: ბენეფიციარის ხედის შენახვა: {beneficiaryList.Count} ბენეფიციარი")
            End If

            ' მასშტაბის გაზრდა
            Dim oldVScale As Double = vScale
            vScale += SCALE_STEP
            If vScale > MAX_SCALE Then vScale = MAX_SCALE

            ' მასშტაბის ლეიბლების განახლება
            UpdateScaleLabels()

            Debug.WriteLine($"BtnVUp_Click: მასშტაბი შეიცვალა {oldVScale:F2} -> {vScale:F2}")

            ' ვანახლებთ კალენდრის ხედს შენახული ბენეფიციარებით, თუ საჭიროა
            If preserveBeneficiaries AndAlso beneficiaryList.Count > 0 Then
                ShowDayViewByBeneficiaryWithPreservation(beneficiaryList)
                Debug.WriteLine($"BtnVUp_Click: შენარჩუნებულია {beneficiaryList.Count} ბენეფიციარის ხედი")
            Else
                UpdateCalendarView()
                Debug.WriteLine("BtnVUp_Click: განახლებულია კალენდრის ხედი სტანდარტულად")
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnVUp_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnVUp_Click: StackTrace - {ex.StackTrace}")
            ' შეცდომის შემთხვევაში ვცდილობთ მაინც განახლებას
            UpdateCalendarView()
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი ფუნქცია, რომელიც ვერტიკალურ მასშტაბირებას ამცირებს
    ''' </summary>
    Private Sub BtnVDown_Click(sender As Object, e As EventArgs) Handles BtnVDown.Click
        Try
            Debug.WriteLine($"BtnVDown_Click: მიმდინარე მასშტაბი vScale={vScale:F2}")

            ' შევინახოთ ბენეფიციარის ხედი, თუ გვაქვს
            Dim preserveBeneficiaries As Boolean = False
            Dim beneficiaryList As New List(Of String)

            ' ვამოწმებთ, ვართ თუ არა ბენეფიციარის ხედში
            If rbDay.Checked AndAlso RBBene.Checked Then
                preserveBeneficiaries = True
                beneficiaryList = GetBeneficiaryColumns()
                Debug.WriteLine($"BtnVDown_Click: ბენეფიციარის ხედის შენახვა: {beneficiaryList.Count} ბენეფიციარი")
            End If

            ' მასშტაბის შემცირება
            Dim oldVScale As Double = vScale
            vScale -= SCALE_STEP
            If vScale < MIN_SCALE Then vScale = MIN_SCALE

            ' მასშტაბის ლეიბლების განახლება
            UpdateScaleLabels()

            Debug.WriteLine($"BtnVDown_Click: მასშტაბი შეიცვალა {oldVScale:F2} -> {vScale:F2}")

            ' ვანახლებთ კალენდრის ხედს შენახული ბენეფიციარებით, თუ საჭიროა
            If preserveBeneficiaries AndAlso beneficiaryList.Count > 0 Then
                ShowDayViewByBeneficiaryWithPreservation(beneficiaryList)
                Debug.WriteLine($"BtnVDown_Click: შენარჩუნებულია {beneficiaryList.Count} ბენეფიციარის ხედი")
            Else
                UpdateCalendarView()
                Debug.WriteLine("BtnVDown_Click: განახლებულია კალენდრის ხედი სტანდარტულად")
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnVDown_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnVDown_Click: StackTrace - {ex.StackTrace}")
            ' შეცდომის შემთხვევაში ვცდილობთ მაინც განახლებას
            UpdateCalendarView()
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების შენარჩუნებით დღის ხედის ჩვენება
    ''' გამოიყენება ძირითადად მასშტაბირების შემდეგ
    ''' </summary>
    ''' <param name="preservedBeneficiaries">შენარჩუნებული ბენეფიციარების სია</param>
    Private Sub ShowDayViewByBeneficiaryWithPreservation(preservedBeneficiaries As List(Of String))
        Try
            Debug.WriteLine($"ShowDayViewByBeneficiaryWithPreservation: დაიწყო, ბენეფიციარების რაოდენობა: {preservedBeneficiaries.Count}")

            ' დავაყენოთ pnlFilter-ის ფერი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' ======= 1. პანელების ინიციალიზაცია =======
            InitializeBeneficiaryDayViewPanels()

            ' ======= 2. თარიღის ზოლის შევსება =======
            FillDateColumnPanel()

            ' ======= 3. დროის პანელის შევსება =======
            FillTimeColumnPanel()

            ' ======= 4. ბენეფიციარების სათაურების პანელის ინიციალიზაცია შენარჩუნებით =======
            If preservedBeneficiaries.Count > 0 Then
                InitializeBeneficiaryHeaderPanelWithPreservation(preservedBeneficiaries)
            Else
                InitializeBeneficiaryHeaderPanel()
            End If

            ' ======= 5. მთავარი გრიდის პანელის შევსება ბენეფიციარებისთვის =======
            FillMainGridPanelForBeneficiaries()

            ' ======= 6. სინქრონიზაცია სქროლისთვის =======
            SetupScrollSynchronizationForBeneficiaries()

            ' ======= 7. ძველი ბარათების გასუფთავება =======
            ClearSessionCardsFromGrid()

            ' ======= 8. სესიების განთავსება ბენეფიციარების გრიდში =======
            PlaceSessionsOnBeneficiaryGrid()

            Debug.WriteLine("ShowDayViewByBeneficiaryWithPreservation: დასრულდა შენარჩუნებული ბენეფიციარების ხედის ჩვენება")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewByBeneficiaryWithPreservation: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewByBeneficiaryWithPreservation: StackTrace - {ex.StackTrace}")

            ' შეცდომის შემთხვევაში ვცდილობთ სტანდარტულ ხედზე გადასვლას
            Try
                InitializeBeneficiaryHeaderPanel()
                UpdateCalendarView()
            Catch
                ' უკიდურეს შემთხვევაში გამოვაჩინოთ შეტყობინება
                MessageBox.Show($"ბენეფიციარების ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების სათაურების პანელის ინიციალიზაცია შენარჩუნებით
    ''' </summary>
    Private Sub InitializeBeneficiaryHeaderPanelWithPreservation(preservedBeneficiaries As List(Of String))
        Try
            Debug.WriteLine($"InitializeBeneficiaryHeaderPanelWithPreservation: დაიწყო, ბენეფიციარების რაოდენობა: {preservedBeneficiaries.Count}")

            ' ვიპოვოთ ან შევქმნათ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                ' შევქმნათ ახალი პანელი თუ არ არსებობს
                beneficiaryHeaderPanel = New Panel()
                beneficiaryHeaderPanel.Name = "beneficiaryHeaderPanel"
                beneficiaryHeaderPanel.AutoScroll = False
                beneficiaryHeaderPanel.BackColor = Color.FromArgb(255, 230, 180)
                beneficiaryHeaderPanel.BorderStyle = BorderStyle.FixedSingle
                pnlCalendarGrid.Controls.Add(beneficiaryHeaderPanel)
            End If

            ' პანელის საწყისი მდებარეობის და ზომების გამოთვლა
            Dim TIME_COLUMN_WIDTH As Integer = BASE_TIME_COLUMN_WIDTH
            Dim DATE_COLUMN_WIDTH As Integer = BASE_DATE_COLUMN_WIDTH
            Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT
            Dim BENEFICIARY_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' გავასუფთავოთ არსებული კონტროლები
            beneficiaryHeaderPanel.Controls.Clear()

            ' შევქმნათ ბენეფიციარების ლეიბლები და წაშლის ღილაკები
            For i As Integer = 0 To preservedBeneficiaries.Count - 1
                CreateBeneficiaryLabelWithDeleteButton(beneficiaryHeaderPanel, i, preservedBeneficiaries(i))
            Next

            ' შევქმნათ ახალი კომბობოქსი ბოლოში
            CreateBeneficiaryColumnWithComboBox(beneficiaryHeaderPanel, preservedBeneficiaries.Count)

            ' გამოვთვალოთ და დავაყენოთ პანელის ზომები
            Dim totalWidth As Integer = (preservedBeneficiaries.Count + 1) * BENEFICIARY_COLUMN_WIDTH
            beneficiaryHeaderPanel.Size = New Size(totalWidth, HEADER_HEIGHT)
            beneficiaryHeaderPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, 0)

            Debug.WriteLine($"InitializeBeneficiaryHeaderPanelWithPreservation: პანელის ზომები - Width: {totalWidth}, Height: {HEADER_HEIGHT}")
            Debug.WriteLine($"InitializeBeneficiaryHeaderPanelWithPreservation: დასრულდა, აღდგენილია {preservedBeneficiaries.Count} ბენეფიციარი")

        Catch ex As Exception
            Debug.WriteLine($"InitializeBeneficiaryHeaderPanelWithPreservation: შეცდომა - {ex.Message}")
            Debug.WriteLine($"InitializeBeneficiaryHeaderPanelWithPreservation: StackTrace - {ex.StackTrace}")

            ' შეცდომის შემთხვევაში ვცდილობთ სტანდარტულ ინიციალიზაციას
            InitializeBeneficiaryHeaderPanel()
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის ლეიბლისა და წაშლის ღილაკის შექმნა
    ''' </summary>
    Private Sub CreateBeneficiaryLabelWithDeleteButton(headerPanel As Panel, columnIndex As Integer, beneficiaryName As String)
        Try
            Debug.WriteLine($"CreateBeneficiaryLabelWithDeleteButton: დაიწყო სვეტისთვის {columnIndex}, ბენეფიციარი: '{beneficiaryName}'")

            ' გამოვთვალოთ სვეტის სიგანე და კონტროლების განლაგება
            Dim COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT

            ' ბენეფიციარის ლეიბლის შექმნა
            Dim lblBeneficiary As New Label()
            lblBeneficiary.Name = $"lblBene_{columnIndex}"
            lblBeneficiary.Text = beneficiaryName
            lblBeneficiary.Location = New Point(columnIndex * COLUMN_WIDTH + 5, 3)
            lblBeneficiary.Size = New Size(COLUMN_WIDTH - 30, HEADER_HEIGHT - 6)
            lblBeneficiary.Font = New Font("Sylfaen", 8, FontStyle.Bold)
            lblBeneficiary.TextAlign = ContentAlignment.MiddleCenter
            lblBeneficiary.BorderStyle = BorderStyle.FixedSingle
            lblBeneficiary.BackColor = Color.FromArgb(255, 248, 220) ' კრემისფერი ფონი
            lblBeneficiary.Tag = beneficiaryName ' შევინახოთ სახელი Tag-ში მომავალი გამოყენებისთვის

            ' წაშლის ღილაკის შექმნა
            Dim btnDelete As New Button()
            btnDelete.Name = $"btnDelBene_{columnIndex}"
            btnDelete.Text = "✕" ' X სიმბოლო
            btnDelete.Size = New Size(20, HEADER_HEIGHT - 8)
            btnDelete.Location = New Point(columnIndex * COLUMN_WIDTH + COLUMN_WIDTH - 24, 4)
            btnDelete.Font = New Font("Segoe UI Symbol", 8, FontStyle.Bold)
            btnDelete.ForeColor = Color.DarkRed
            btnDelete.BackColor = Color.FromArgb(255, 230, 230) ' წითელი ფონი
            btnDelete.FlatStyle = FlatStyle.Flat
            btnDelete.FlatAppearance.BorderSize = 1
            btnDelete.FlatAppearance.BorderColor = Color.Red
            btnDelete.Cursor = Cursors.Hand
            btnDelete.Tag = columnIndex ' სვეტის ინდექსის შენახვა გამოყენებისთვის

            ' მივაბათ წაშლის ივენთი
            AddHandler btnDelete.Click, AddressOf BeneficiaryDeleteButton_Click

            ' დავამატოთ ორივე კონტროლი პანელზე
            headerPanel.Controls.Add(lblBeneficiary)
            headerPanel.Controls.Add(btnDelete)

            Debug.WriteLine($"CreateBeneficiaryLabelWithDeleteButton: შეიქმნა ლეიბლი და წაშლის ღილაკი ბენეფიციარისთვის '{beneficiaryName}'")

        Catch ex As Exception
            Debug.WriteLine($"CreateBeneficiaryLabelWithDeleteButton: შეცდომა - {ex.Message}")
            Debug.WriteLine($"CreateBeneficiaryLabelWithDeleteButton: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების სვეტების რეორგანიზაცია (ინდექსების და პოზიციების განახლება)
    ''' გამოიყენება ბენეფიციარის წაშლის ან სხვა ცვლილების შემდეგ
    ''' </summary>
    Private Sub ReorganizeBeneficiaryColumns()
        Try
            Debug.WriteLine("ReorganizeBeneficiaryColumns: დაიწყო სვეტების რეორგანიზაცია")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Debug.WriteLine("ReorganizeBeneficiaryColumns: ბენეფიციარების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' სვეტის სიგანე მასშტაბის გათვალისწინებით
            Dim COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' მოვაგროვოთ ყველა ლეიბლი და დავალაგოთ პოზიციის მიხედვით
            Dim labels = beneficiaryHeaderPanel.Controls.OfType(Of Label)() _
                    .Where(Function(l) l.Name.StartsWith("lblBene_")) _
                    .OrderBy(Function(l) l.Location.X) _
                    .ToList()

            Debug.WriteLine($"ReorganizeBeneficiaryColumns: ნაპოვნია {labels.Count} ლეიბლი რეორგანიზაციისთვის")

            ' განვაახლოთ ლეიბლების პოზიციები და სახელები
            For i As Integer = 0 To labels.Count - 1
                Dim label = labels(i)
                Dim oldName = label.Name

                ' განვაახლოთ ლეიბლის სახელი და პოზიცია
                label.Name = $"lblBene_{i}"
                label.Location = New Point(i * COLUMN_WIDTH + 5, label.Location.Y)

                Debug.WriteLine($"ReorganizeBeneficiaryColumns: ლეიბლი {oldName} -> {label.Name}, X = {label.Location.X}")

                ' ვიპოვოთ შესაბამისი წაშლის ღილაკი
                Dim oldIndex = oldName.Split("_"c)(1)
                Dim button = beneficiaryHeaderPanel.Controls.OfType(Of Button)() _
                        .FirstOrDefault(Function(b) b.Name = $"btnDelBene_{oldIndex}")

                If button IsNot Nothing Then
                    ' განვაახლოთ ღილაკის სახელი, პოზიცია და Tag
                    button.Name = $"btnDelBene_{i}"
                    button.Location = New Point(i * COLUMN_WIDTH + COLUMN_WIDTH - 24, button.Location.Y)
                    button.Tag = i

                    Debug.WriteLine($"ReorganizeBeneficiaryColumns: ღილაკი {oldIndex} -> {i}, X = {button.Location.X}")
                End If
            Next

            ' მოვძებნოთ ყველა კომბობოქსი და დავალაგოთ
            Dim comboBoxes = beneficiaryHeaderPanel.Controls.OfType(Of ComboBox)() _
                        .Where(Function(c) c.Name.StartsWith("comboBene_")) _
                        .OrderBy(Function(c) c.Location.X) _
                        .ToList()

            Debug.WriteLine($"ReorganizeBeneficiaryColumns: ნაპოვნია {comboBoxes.Count} კომბობოქსი რეორგანიზაციისთვის")

            ' განვაახლოთ კომბობოქსების პოზიციები
            For i As Integer = 0 To comboBoxes.Count - 1
                Dim combo = comboBoxes(i)
                Dim newIndex = labels.Count + i

                ' განვაახლოთ კომბობოქსის სახელი, პოზიცია და Tag
                combo.Name = $"comboBene_{newIndex}"
                combo.Location = New Point(newIndex * COLUMN_WIDTH + 5, combo.Location.Y)
                combo.Tag = newIndex

                Debug.WriteLine($"ReorganizeBeneficiaryColumns: კომბობოქსი განახლდა -> {combo.Name}, X = {combo.Location.X}")
            Next

            ' განვაახლოთ პანელის სიგანე ყველა სვეტის გათვალისწინებით
            Dim totalColumns = labels.Count + comboBoxes.Count
            beneficiaryHeaderPanel.Width = totalColumns * COLUMN_WIDTH

            Debug.WriteLine($"ReorganizeBeneficiaryColumns: პანელის სიგანე განახლდა: {beneficiaryHeaderPanel.Width}px, სულ სვეტი: {totalColumns}")

        Catch ex As Exception
            Debug.WriteLine($"ReorganizeBeneficiaryColumns: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ReorganizeBeneficiaryColumns: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ახალი ცარიელი კომბობოქსის დამატება ბენეფიციარებისთვის
    ''' </summary>
    Private Sub CreateNextBeneficiaryColumn()
        Try
            Debug.WriteLine("CreateNextBeneficiaryColumn: დაიწყო ახალი კომბობოქსის დამატება")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Debug.WriteLine("CreateNextBeneficiaryColumn: ბენეფიციარების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' მოვითვალოთ მიმდინარე ლეიბლების რაოდენობა
            Dim labelCount = beneficiaryHeaderPanel.Controls.OfType(Of Label)() _
                        .Count(Function(l) l.Name.StartsWith("lblBene_"))

            Debug.WriteLine($"CreateNextBeneficiaryColumn: არსებული ლეიბლების რაოდენობა: {labelCount}")

            ' სვეტის სიგანე მასშტაბის გათვალისწინებით
            Dim COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' გავაფართოვოთ პანელი ახალი სვეტისთვის
            beneficiaryHeaderPanel.Width += COLUMN_WIDTH

            ' შევქმნათ ახალი კომბობოქსი
            CreateBeneficiaryColumnWithComboBox(beneficiaryHeaderPanel, labelCount)

            Debug.WriteLine($"CreateNextBeneficiaryColumn: დაემატა ახალი კომბობოქსი ინდექსით {labelCount}, პანელის ახალი სიგანე: {beneficiaryHeaderPanel.Width}px")

        Catch ex As Exception
            Debug.WriteLine($"CreateNextBeneficiaryColumn: შეცდომა - {ex.Message}")
            Debug.WriteLine($"CreateNextBeneficiaryColumn: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' არჩეული ბენეფიციარების სიის მიღება
    ''' </summary>
    ''' <returns>ბენეფიციარების სახელების სია</returns>
    Private Function GetBeneficiaryColumns() As List(Of String)
        Try
            Debug.WriteLine("GetBeneficiaryColumns: დაიწყო არჩეული ბენეფიციარების მოძიება")

            Dim beneficiaries As New List(Of String)()

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel IsNot Nothing Then
                ' გადავირბინოთ ყველა ლეიბლი და ამოვკრიბოთ ბენეფიციარების სახელები
                Dim labels = beneficiaryHeaderPanel.Controls.OfType(Of Label)() _
                        .Where(Function(l) l.Name.StartsWith("lblBene_")) _
                        .OrderBy(Function(l) l.Location.X) _
                        .ToList()

                For Each label In labels
                    If label.Tag IsNot Nothing Then
                        Dim beneficiaryName As String = label.Tag.ToString()
                        beneficiaries.Add(beneficiaryName)
                        Debug.WriteLine($"GetBeneficiaryColumns: ნაპოვნია ბენეფიციარი: '{beneficiaryName}'")
                    End If
                Next
            End If

            Debug.WriteLine($"GetBeneficiaryColumns: სულ ნაპოვნია {beneficiaries.Count} ბენეფიციარი")
            Return beneficiaries

        Catch ex As Exception
            Debug.WriteLine($"GetBeneficiaryColumns: შეცდომა - {ex.Message}")
            Debug.WriteLine($"GetBeneficiaryColumns: StackTrace - {ex.StackTrace}")
            Return New List(Of String)()
        End Try
    End Function

    ''' <summary>
    ''' ყველა ბენეფიციარის კომბობოქსის განახლება
    ''' თავიდან ავსებს მათ, რომ არ გამეორდეს არჩეული ბენეფიციარები
    ''' </summary>
    Private Sub RefreshAllBeneficiaryComboBoxes()
        Try
            Debug.WriteLine("RefreshAllBeneficiaryComboBoxes: დაიწყო კომბობოქსების განახლება")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Debug.WriteLine("RefreshAllBeneficiaryComboBoxes: ბენეფიციარების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' მოვძებნოთ ყველა კომბობოქსი პანელზე
            Dim comboBoxes = beneficiaryHeaderPanel.Controls.OfType(Of ComboBox)() _
                        .Where(Function(c) c.Name.StartsWith("comboBene_")) _
                        .ToList()

            Debug.WriteLine($"RefreshAllBeneficiaryComboBoxes: ნაპოვნია {comboBoxes.Count} კომბობოქსი განსაახლებლად")

            ' ყოველი კომბობოქსის განახლება
            For Each combo In comboBoxes
                ' შევინახოთ მიმდინარე არჩეული მნიშვნელობა
                Dim selectedIndex = combo.SelectedIndex
                Dim selectedText = If(selectedIndex >= 0, combo.SelectedItem?.ToString(), Nothing)

                ' განვაახლოთ კომბობოქსის შიგთავსი
                FillBeneficiaryComboBox(combo)

                ' ვცადოთ იგივე მნიშვნელობის არჩევა, თუ ჯერ არაფერი არ იყო არჩეული
                If selectedIndex > 0 AndAlso selectedText IsNot Nothing Then
                    Dim newIndex = combo.Items.IndexOf(selectedText)
                    If newIndex >= 0 Then
                        combo.SelectedIndex = newIndex
                    End If
                End If

                Debug.WriteLine($"RefreshAllBeneficiaryComboBoxes: განახლდა კომბობოქსი {combo.Name}, ელემენტების რაოდენობა: {combo.Items.Count}")
            Next

            Debug.WriteLine("RefreshAllBeneficiaryComboBoxes: ყველა კომბობოქსი განახლდა")

        Catch ex As Exception
            Debug.WriteLine($"RefreshAllBeneficiaryComboBoxes: შეცდომა - {ex.Message}")
            Debug.WriteLine($"RefreshAllBeneficiaryComboBoxes: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის სესიების განთავსება გრიდზე
    ''' </summary>
    Private Sub PlaceSessionsOnBeneficiaryGrid()
        Try
            Debug.WriteLine("=== PlaceSessionsOnBeneficiaryGrid: დაიწყო ===")

            ' ძირითადი შემოწმებები
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

            ' ვიპოვოთ არჩეული ბენეფიციარების სია
            Dim selectedBeneficiaries = GetBeneficiaryColumns()
            Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: არჩეული ბენეფიციარების რაოდენობა: {selectedBeneficiaries.Count}")

            If selectedBeneficiaries.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnBeneficiaryGrid: არჩეული ბენეფიციარები არ არის")
                Return
            End If

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            If mainGridPanel Is Nothing Then
                Debug.WriteLine("PlaceSessionsOnBeneficiaryGrid: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' მასშტაბირებული პარამეტრები
            Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
            Dim BENEFICIARY_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' მივიღოთ ფილტრირებული სესიები
            Dim filteredSessions = GetFilteredSessions()
            Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: ფილტრირებული სესიების რაოდენობა: {filteredSessions.Count}")

            ' განვათავსოთ სესიების ბარათები
            For Each session In filteredSessions
                Try
                    Debug.WriteLine($"--- PlaceSessionsOnBeneficiaryGrid: სესია ID={session.Id}, ბენეფიციარი='{session.BeneficiaryName} {session.BeneficiarySurname}', დრო={session.DateTime:HH:mm} ---")

                    ' 1. ვიპოვოთ ბენეფიციარის ინდექსი არჩეულებში
                    Dim sessionBeneficiaryFullName As String = $"{session.BeneficiaryName.Trim()} {session.BeneficiarySurname.Trim()}"
                    Dim beneficiaryIndex As Integer = selectedBeneficiaries.FindIndex(Function(b) String.Equals(b.Trim(), sessionBeneficiaryFullName, StringComparison.OrdinalIgnoreCase))

                    If beneficiaryIndex < 0 Then
                        Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: ბენეფიციარი '{sessionBeneficiaryFullName}' არ არის არჩეულებში")
                        Continue For
                    End If

                    ' 2. მოვძებნოთ უახლოესი დროის ინტერვალი
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

                    ' ალტერნატიული გამოთვლა, თუ ზემოთ ვერ ვიპოვეთ
                    If timeIndex < 0 OrElse minDifference.TotalMinutes > 15 Then
                        Dim startTime As TimeSpan = timeIntervals(0).TimeOfDay
                        Dim elapsedMinutes As Double = (sessionTime - startTime).TotalMinutes
                        timeIndex = CInt(Math.Round(elapsedMinutes / 30))
                        If timeIndex < 0 Then timeIndex = 0
                        If timeIndex >= timeIntervals.Count Then timeIndex = timeIntervals.Count - 1
                    End If

                    ' 3. შევამოწმოთ ინდექსები
                    If beneficiaryIndex < 0 OrElse timeIndex < 0 OrElse
                   beneficiaryIndex >= selectedBeneficiaries.Count OrElse timeIndex >= timeIntervals.Count Then
                        Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: არასწორი ინდექსები: beneficiary={beneficiaryIndex}, time={timeIndex}")
                        Continue For
                    End If

                    ' 4. გამოვთვალოთ ბარათის პოზიცია
                    Dim cardX As Integer = beneficiaryIndex * BENEFICIARY_COLUMN_WIDTH + 4
                    Dim cardY As Integer = timeIndex * ROW_HEIGHT + 2

                    ' 5. გამოვთვალოთ ბარათის სიმაღლე ხანგრძლივობის მიხედვით
                    Dim sessionDurationMinutes As Integer = session.Duration
                    Dim baseCardHeight As Double = ROW_HEIGHT * (sessionDurationMinutes / 30.0)
                    Dim cardHeight As Integer = CInt(Math.Max(baseCardHeight, ROW_HEIGHT))
                    ' შევზღუდოთ მაქსიმალური სიმაღლე
                    Dim maxCardHeight As Integer = ROW_HEIGHT * 4
                    If cardHeight > maxCardHeight Then cardHeight = maxCardHeight

                    ' 6. განვსაზღვროთ ბარათის ფერები სტატუსის მიხედვით
                    Dim cardColor As Color = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                    Dim borderColor As Color = SessionStatusColors.GetStatusBorderColor(session.Status, session.DateTime)
                    Dim headerColor As Color = Color.FromArgb(
                   Math.Max(0, cardColor.R - 50),
                   Math.Max(0, cardColor.G - 50),
                   Math.Max(0, cardColor.B - 50)
                )

                    ' 7. შევქმნათ სესიის ბარათი
                    Dim sessionCard As New Panel()
                    sessionCard.Size = New Size(BENEFICIARY_COLUMN_WIDTH - 8, cardHeight)
                    sessionCard.Location = New Point(cardX, cardY)
                    sessionCard.BackColor = cardColor
                    sessionCard.BorderStyle = BorderStyle.None
                    sessionCard.Tag = session.Id
                    sessionCard.Cursor = Cursors.Hand

                    ' 8. დავამატოთ სათაურის პანელი
                    Dim HEADER_HEIGHT As Integer = 24
                    Dim headerPanel As New Panel()
                    headerPanel.Size = New Size(sessionCard.Width, HEADER_HEIGHT)
                    headerPanel.Location = New Point(0, 0)
                    headerPanel.BackColor = headerColor
                    sessionCard.Controls.Add(headerPanel)

                    ' 9. დავამატოთ ID ლეიბლი სათაურის პანელზე
                    Dim lblId As New Label()
                    lblId.Text = $"#{session.Id}"
                    lblId.Location = New Point(8, 2)
                    lblId.Size = New Size(headerPanel.Width - 16, 20)
                    lblId.Font = New Font("Sylfaen", 10, FontStyle.Bold)
                    lblId.ForeColor = Color.White
                    lblId.TextAlign = ContentAlignment.MiddleCenter
                    headerPanel.Controls.Add(lblId)

                    ' 10. ბარათის ჩარჩო
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

                    ' 11. დავამატოთ დამატებითი ინფორმაცია ბარათზე (სტატუსი, თერაპევტი, სივრცე)
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

                    ' 12. დავამატოთ რედაქტირების ღილაკი
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

                    ' მრგვალი ღილაკი
                    Dim btnPath As New Drawing2D.GraphicsPath()
                    btnPath.AddEllipse(0, 0, btnEdit.Width, btnEdit.Height)
                    btnEdit.Region = New Region(btnPath)

                    ' ღილაკის კლიკის ივენთი
                    AddHandler btnEdit.Click, AddressOf BtnEditSession_Click

                    ' დავამატოთ ღილაკი ბარათზე
                    sessionCard.Controls.Add(btnEdit)
                    btnEdit.BringToFront()

                    ' 13. მომრგვალებული კუთხეები ბარათისთვის
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
                    Catch pathEx As Exception
                        ' თუ მომრგვალებული კუთხეების შექმნა ვერ მოხერხდა, გავაგრძელოთ მაინც
                        Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: მომრგვალებული კუთხეების შექმნის შეცდომა - {pathEx.Message}")
                    End Try

                    ' 14. მივაბათ ბარათის ივენთები
                    AddHandler sessionCard.Click, AddressOf SessionCard_Click
                    AddHandler sessionCard.DoubleClick, AddressOf SessionCard_DoubleClick

                    ' 15. შევქმნათ კონტექსტური მენიუ
                    CreateSessionCardContextMenu(sessionCard, session)

                    ' 16. დავამატოთ ბარათი მთავარ პანელზე
                    mainGridPanel.Controls.Add(sessionCard)
                    sessionCard.BringToFront()

                    Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: ბარათი განთავსდა: ID={session.Id}, ბენეფიციარი='{sessionBeneficiaryFullName}', X={cardX}, Y={cardY}")

                Catch sessionEx As Exception
                    Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: შეცდომა სესიის დამუშავებისას - {sessionEx.Message}")
                    Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: StackTrace - {sessionEx.StackTrace}")
                    Continue For
                End Try
            Next

            Debug.WriteLine("=== PlaceSessionsOnBeneficiaryGrid: დასრულდა ===")

        Catch ex As Exception
            Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: ზოგადი შეცდომა - {ex.Message}")
            Debug.WriteLine($"PlaceSessionsOnBeneficiaryGrid: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' საერთო მეთოდი, რომელიც შემდეგ ბენეფიციარის სვეტს ამატებს
    ''' </summary>
    Private Sub AddNextBeneficiary()
        Try
            Debug.WriteLine("AddNextBeneficiary: დაიწყო ახალი ბენეფიციარის დამატება")

            ' ვიპოვოთ ბენეფიციარების სათაურის პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Debug.WriteLine("AddNextBeneficiary: ბენეფიციარების სათაურების პანელი ვერ მოიძებნა")
                Return
            End If

            ' მოვძებნოთ ბოლო კომბობოქსი
            Dim lastComboBox As ComboBox = Nothing
            Dim maxIndex As Integer = -1

            For Each ctrl As Control In beneficiaryHeaderPanel.Controls
                If TypeOf ctrl Is ComboBox AndAlso ctrl.Name.StartsWith("comboBene_") Then
                    Dim comboBox As ComboBox = DirectCast(ctrl, ComboBox)
                    Dim indexStr As String = comboBox.Name.Split("_"c)(1)
                    Dim index As Integer

                    If Integer.TryParse(indexStr, index) AndAlso index > maxIndex Then
                        maxIndex = index
                        lastComboBox = comboBox
                    End If
                End If
            Next

            If lastComboBox IsNot Nothing AndAlso lastComboBox.SelectedIndex > 0 Then
                ' არჩეული ბენეფიციარი
                Dim selectedBeneficiary As String = lastComboBox.SelectedItem.ToString()

                ' შევცვალოთ კომბობოქსი ლეიბლით
                ReplaceComboBoxWithLabelAndDeleteButton(lastComboBox, selectedBeneficiary)

                ' შევქმნათ ახალი კომბობოქსი
                CreateNextBeneficiaryColumn()

                ' განვაახლოთ გრიდი
                UpdateBeneficiaryGrid()

                Debug.WriteLine($"AddNextBeneficiary: დაემატა ბენეფიციარი '{selectedBeneficiary}', შეიქმნა ახალი კომბობოქსი")
            Else
                Debug.WriteLine("AddNextBeneficiary: არ არის არჩეული ბენეფიციარი ბოლო კომბობოქსში")
            End If

        Catch ex As Exception
            Debug.WriteLine($"AddNextBeneficiary: შეცდომა - {ex.Message}")
            Debug.WriteLine($"AddNextBeneficiary: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ხედის ტიპის რადიო ღილაკის არჩევის დამუშავება (დღე/კვირა/თვე)
    ''' </summary>
    Private Sub ViewType_CheckedChanged(sender As Object, e As EventArgs) Handles rbDay.CheckedChanged, rbWeek.CheckedChanged, rbMonth.CheckedChanged
        Try
            Dim radioButton As RadioButton = DirectCast(sender, RadioButton)
            If Not radioButton.Checked Then Return ' მხოლოდ ჩართულ მდგომარეობაზე ვრეაგირებთ

            Debug.WriteLine($"ViewType_CheckedChanged: არჩეულია '{radioButton.Text}'")

            ' განვაახლოთ კალენდარი
            UpdateCalendarView()

        Catch ex As Exception
            Debug.WriteLine($"ViewType_CheckedChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფილტრის ტიპის რადიო ღილაკის არჩევის დამუშავება (სივრცე/თერაპევტი/ბენეფიციარი)
    ''' </summary>
    Private Sub FilterType_CheckedChanged(sender As Object, e As EventArgs) Handles RBSpace.CheckedChanged, RBPer.CheckedChanged, RBBene.CheckedChanged
        Try
            Dim radioButton As RadioButton = DirectCast(sender, RadioButton)
            If Not radioButton.Checked Then Return ' მხოლოდ ჩართულ მდგომარეობაზე ვრეაგირებთ

            Debug.WriteLine($"FilterType_CheckedChanged: არჩეულია '{radioButton.Text}'")

            ' განვაახლოთ კალენდარი
            UpdateCalendarView()

        Catch ex As Exception
            Debug.WriteLine($"FilterType_CheckedChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მასშტაბის ლეიბლების განახლება
    ''' </summary>
    Private Sub UpdateScaleLabels()
        Try
            ' შევქმნათ ან განვაახლოთ ჰორიზონტალური მასშტაბის ლეიბლი
            Dim lblHScale As Label = DirectCast(Controls.Find("lblHScale", True).FirstOrDefault(), Label)
            If lblHScale Is Nothing Then
                ' შევქმნათ ახალი ლეიბლი, თუ არ არსებობს
                lblHScale = New Label()
                lblHScale.Name = "lblHScale"
                lblHScale.AutoSize = True
                lblHScale.Location = New Point(BtnHUp.Left + BtnHUp.Width + 5, BtnHUp.Top + 3)
                lblHScale.Font = New Font("Segoe UI", 8)
                Me.Controls.Add(lblHScale)
            End If
            lblHScale.Text = $"×{hScale:F2}"

            ' შევქმნათ ან განვაახლოთ ვერტიკალური მასშტაბის ლეიბლი
            Dim lblVScale As Label = DirectCast(Controls.Find("lblVScale", True).FirstOrDefault(), Label)
            If lblVScale Is Nothing Then
                ' შევქმნათ ახალი ლეიბლი, თუ არ არსებობს
                lblVScale = New Label()
                lblVScale.Name = "lblVScale"
                lblVScale.AutoSize = True
                lblVScale.Location = New Point(BtnVUp.Left + BtnVUp.Width + 5, BtnVUp.Top + 3)
                lblVScale.Font = New Font("Segoe UI", 8)
                Me.Controls.Add(lblVScale)
            End If
            lblVScale.Text = $"×{vScale:F2}"

            Debug.WriteLine($"UpdateScaleLabels: განახლდა მასშტაბის ლეიბლები - Horizontal: ×{hScale:F2}, Vertical: ×{vScale:F2}")

        Catch ex As Exception
            Debug.WriteLine($"UpdateScaleLabels: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' განახლებული კალენდრის ხედის მეთოდი
    ''' ამოწმებს არჩეულ რადიო ღილაკებს და აჩვენებს შესაბამის ხედს
    ''' </summary>
    Public Sub UpdateCalendarView()
        Try
            Debug.WriteLine("UpdateCalendarView: დაიწყო კალენდრის ხედის განახლება")

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

            ' შევამოწმოთ რომელი ხედია არჩეული
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

            Debug.WriteLine("UpdateCalendarView: კალენდრის ხედის განახლება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UpdateCalendarView: შეცდომა - {ex.Message}")
            Debug.WriteLine($"UpdateCalendarView: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"კალენდრის ხედის განახლების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' შემოწმება რომელი სტატუსის ფილტრებია ჩართული
    ''' </summary>
    ''' <returns>True თუ სტატუსი უნდა გამოჩნდეს</returns>
    Private Function IsStatusFilteredIn(status As String) As Boolean
        Try
            ' სტატუსის დაფორმატება
            Dim normalizedStatus As String = status.Trim().ToLower()

            ' ვამოწმებთ სტატუსის მიხედვით
            Select Case normalizedStatus
                Case "დაგეგმილი"
                    Return CheckBox1.Checked
                Case "შესრულებული"
                    Return CheckBox2.Checked
                Case "გაცდენა საპატიო"
                    Return CheckBox3.Checked
                Case "გაცდენა არასაპატიო"
                    Return CheckBox4.Checked
                Case "აღდგენა"
                    Return CheckBox5.Checked
                Case "პროგრამით გატარება"
                    Return CheckBox6.Checked
                Case "გაუქმებული", "გაუქმება"
                    Return CheckBox7.Checked
                Case Else
                    ' უცნობი სტატუსებისთვის ნაგულისხმევი მნიშვნელობა - გამოჩნდეს
                    Return True
            End Select

        Catch ex As Exception
            Debug.WriteLine($"IsStatusFilteredIn: შეცდომა - {ex.Message}")
            ' შეცდომის შემთხვევაში ვაჩვენოთ
            Return True
        End Try
    End Function

    ''' <summary>
    ''' სესიების ფილტრაცია სტატუსის მიხედვით
    ''' </summary>
    ''' <returns>ფილტრირებული სესიების სია</returns>
    Private Function GetFilteredSessions() As List(Of SessionModel)
        Try
            Debug.WriteLine("GetFilteredSessions: დაიწყო სესიების ფილტრაცია")

            ' არჩეული თარიღი
            Dim selectedDate As DateTime = DTPCalendar.Value.Date

            ' მხოლოდ არჩეული დღის სესიები
            Dim daySessions = allSessions.Where(Function(s) s.DateTime.Date = selectedDate).ToList()

            ' ფილტრაცია სტატუსების მიხედვით
            Dim filteredSessions As New List(Of SessionModel)()

            For Each session In daySessions
                ' შემოწმება შესაბამისი სტატუსის ფილტრის ჩეკბოქსის მიხედვით
                If IsStatusFilteredIn(session.Status) Then
                    filteredSessions.Add(session)
                End If
            Next

            Debug.WriteLine($"GetFilteredSessions: {daySessions.Count} სესიიდან ფილტრაციის შემდეგ დარჩა {filteredSessions.Count} სესია")
            Return filteredSessions

        Catch ex As Exception
            Debug.WriteLine($"GetFilteredSessions: შეცდომა - {ex.Message}")
            Debug.WriteLine($"GetFilteredSessions: StackTrace - {ex.StackTrace}")
            Return New List(Of SessionModel)()
        End Try
    End Function

    ''' <summary>
    ''' დროის დიაპაზონის ცვლილებაზე რეაგირება
    ''' </summary>
    Private Sub TimeRange_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbStart.SelectedIndexChanged, cbFinish.SelectedIndexChanged
        Try
            ' ვამოწმებთ ორივე დრო არის თუ არა არჩეული
            If cbStart.SelectedIndex < 0 OrElse cbFinish.SelectedIndex < 0 Then
                Debug.WriteLine("TimeRange_SelectedIndexChanged: დროების ერთ-ერთი არ არის არჩეული")
                Return
            End If

            Debug.WriteLine($"TimeRange_SelectedIndexChanged: არჩეულია დროის დიაპაზონი {cbStart.SelectedItem} - {cbFinish.SelectedItem}")

            ' განვაახლოთ დროის ინტერვალები
            InitializeTimeIntervals()

            ' განვაახლოთ კალენდრის ხედი
            UpdateCalendarView()

        Catch ex As Exception
            Debug.WriteLine($"TimeRange_SelectedIndexChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კალენდრის თარიღის ცვლილებაზე რეაგირება
    ''' </summary>
    Private Sub DTPCalendar_ValueChanged(sender As Object, e As EventArgs) Handles DTPCalendar.ValueChanged
        Try
            Debug.WriteLine($"DTPCalendar_ValueChanged: არჩეულია თარიღი {DTPCalendar.Value:dd.MM.yyyy}")

            ' განვაახლოთ კალენდრის ხედი
            UpdateCalendarView()

        Catch ex As Exception
            Debug.WriteLine($"DTPCalendar_ValueChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სტატუსის ფილტრის ჩეკბოქსის ცვლილების დამუშავება
    ''' </summary>
    Private Sub StatusFilter_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged, CheckBox3.CheckedChanged, CheckBox4.CheckedChanged, CheckBox5.CheckedChanged, CheckBox6.CheckedChanged, CheckBox7.CheckedChanged
        Try
            Dim checkBox As CheckBox = DirectCast(sender, CheckBox)

            Debug.WriteLine($"StatusFilter_CheckedChanged: {checkBox.Text} = {checkBox.Checked}")

            ' განვაახლოთ კალენდრის ხედი ფილტრაციის შემდეგ
            UpdateCalendarView()

        Catch ex As Exception
            Debug.WriteLine($"StatusFilter_CheckedChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ახალი სესიის დამატების ღილაკის დამუშავება
    ''' </summary>
    Private Sub BtnAddSchedule_Click(sender As Object, e As EventArgs) Handles BtnAddSchedule.Click
        Try
            Debug.WriteLine("BtnAddSchedule_Click: ახალი სესიის დამატების მოთხოვნა")

            ' შევამოწმოთ უკვე გახსნილია თუ არა NewRecordForm
            For Each frm As Form In Application.OpenForms
                If TypeOf frm Is NewRecordForm Then
                    Debug.WriteLine("BtnAddSchedule_Click: NewRecordForm უკვე გახსნილია, გადავიტანოთ წინა პლანზე")
                    frm.Focus()
                    Return
                End If
            Next

            ' შევამოწმოთ არის თუ არა მომხმარებელი ავტორიზებული
            If String.IsNullOrEmpty(userEmail) Then
                MessageBox.Show("ახალი სესიის დასამატებლად საჭიროა ავტორიზაცია", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Debug.WriteLine("BtnAddSchedule_Click: მომხმარებელი არ არის ავტორიზებული")
                Return
            End If

            ' შევამოწმოთ გვაქვს თუ არა dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Debug.WriteLine("BtnAddSchedule_Click: მონაცემთა სერვისი არ არის ინიციალიზებული")
                Return
            End If

            ' შევქმნათ და გავხსნათ ახალი ჩანაწერის ფორმა
            Dim newRecordForm As New NewRecordForm(dataService, "სესია", userEmail, "UC_Calendar")

            ' დავამატოთ ფორმის დახურვის ივენთის დამუშავება
            AddHandler newRecordForm.FormClosed, Sub(s, args)
                                                     ' თუ ფორმა დაიხურა OK რეზულტატით, განვაახლოთ მონაცემები
                                                     If newRecordForm.DialogResult = DialogResult.OK Then
                                                         Debug.WriteLine("BtnAddSchedule_Click: ახალი სესია წარმატებით დაემატა")

                                                         ' ჩავტვირთოთ სესიები თავიდან
                                                         LoadSessions()

                                                         ' განვაახლოთ კალენდრის ხედი
                                                         UpdateCalendarView()
                                                     End If
                                                 End Sub

            ' გავხსნათ ფორმა
            newRecordForm.Show()

            Debug.WriteLine("BtnAddSchedule_Click: ახალი სესიის ფორმა გაიხსნა")

        Catch ex As Exception
            Debug.WriteLine($"BtnAddSchedule_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnAddSchedule_Click: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"ახალი სესიის ფორმის გახსნის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სესიების განახლების (Refresh) ღილაკის დამუშავება
    ''' </summary>
    Private Sub BtnRef_Click(sender As Object, e As EventArgs) Handles BtnRef.Click
        Try
            Debug.WriteLine("BtnRef_Click: იწყება მონაცემების განახლება")

            ' გამოვაჩინოთ კურსორი როგორც დატვირთვის კურსორი
            Cursor = Cursors.WaitCursor

            ' გავასუფთავოთ სესიების ქეში
            allSessions = Nothing

            ' ჩავტვირთოთ სესიები თავიდან
            LoadSessions()

            ' გავასუფთავოთ კალენდარი და ხელახლა დავტვირთოთ ყველაფერი
            ClearSessionCardsFromGrid()
            UpdateCalendarView()

            ' დავაბრუნოთ ჩვეულებრივი კურსორი
            Cursor = Cursors.Default

            Debug.WriteLine("BtnRef_Click: მონაცემები წარმატებით განახლდა")

        Catch ex As Exception
            Debug.WriteLine($"BtnRef_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnRef_Click: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"მონაცემების განახლებისას მოხდა შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)

            ' დავაბრუნოთ ჩვეულებრივი კურსორი შეცდომის შემთხვევაშიც
            Cursor = Cursors.Default
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

            ' შევამოწმოთ არის თუ არა მომხმარებელი ავტორიზებული
            If String.IsNullOrEmpty(userEmail) Then
                MessageBox.Show("სესიის რედაქტირებისთვის საჭიროა ავტორიზაცია", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Debug.WriteLine("BtnEditSession_Click: მომხმარებელი არ არის ავტორიზებული")
                Return
            End If

            ' შევამოწმოთ გვაქვს თუ არა dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Debug.WriteLine("BtnEditSession_Click: მონაცემთა სერვისი არ არის ინიციალიზებული")
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
            Debug.WriteLine($"BtnEditSession_Click: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"სესიის რედაქტირების ფორმის გახსნის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' UserControl-ის ჩატვირთვის ივენთი
    ''' </summary>
    Private Sub UC_Calendar_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            ' მნიშვნელოვანი: UserControl-ს და pnlCalendarGrid-ს არ ჰქონდეს AutoScroll
            Me.AutoScroll = False

            If pnlCalendarGrid IsNot Nothing Then
                pnlCalendarGrid.AutoScroll = False
            End If

            ' საწყისი პარამეტრების დაყენება (თუ ისინი არ არის დაყენებული)
            InitializeTimeIntervals()

            ' კალენდრის ხედის განახლება
            UpdateCalendarView()

            Debug.WriteLine("UC_Calendar_Load: კალენდრის კონტროლი წარმატებით ჩაიტვირთა")

        Catch ex As Exception
            Debug.WriteLine($"UC_Calendar_Load: შეცდომა - {ex.Message}")
            Debug.WriteLine($"UC_Calendar_Load: StackTrace - {ex.StackTrace}")
        End Try
    End Sub
    ''' <summary>
    ''' UserControl-ის ზომის ცვლილებაზე რეაგირება
    ''' განაახლებს კალენდრის ზომას კონტეინერის მიხედვით
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

            ' განვაახლოთ მთავარი გრიდის პანელი, თუ ის უკვე ინიციალიზებულია
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            If mainGridPanel IsNot Nothing Then
                ' განვაახლოთ მთავარი გრიდის პანელის ზომები
                mainGridPanel.Width = pnlCalendarGrid.Width - mainGridPanel.Left - 2
                mainGridPanel.Height = pnlCalendarGrid.Height - mainGridPanel.Top - 2
            End If

            Debug.WriteLine($"UC_Calendar_Resize: კალენდრის ზომები განახლდა, ახალი ზომები: {Me.Width}x{Me.Height}")

        Catch ex As Exception
            Debug.WriteLine($"UC_Calendar_Resize: შეცდომა - {ex.Message}")
            Debug.WriteLine($"UC_Calendar_Resize: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის ბარათზე დაწკაპუნების დამუშავება
    ''' აჩვენებს სესიის დეტალებს მესიჯბოქსში
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
                ' გამოვაჩინოთ სესიის დეტალები მესიჯბოქსში
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
                sb.AppendLine($"ფასი: {session.Price:N2}")
                sb.AppendLine($"დაფინანსება: {session.Funding}")

                ' თუ კომენტარები არ არის ცარიელი, დავამატოთ
                If Not String.IsNullOrEmpty(session.Comments) Then
                    sb.AppendLine()
                    sb.AppendLine("კომენტარები:")
                    sb.AppendLine(session.Comments)
                End If

                MessageBox.Show(sb.ToString(), "სესიის დეტალები", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Debug.WriteLine($"SessionCard_Click: ნაჩვენებია სესიის დეტალები, ID={sessionId}")
            Else
                Debug.WriteLine($"SessionCard_Click: სესია ID={sessionId} ვერ მოიძებნა")
            End If

        Catch ex As Exception
            Debug.WriteLine($"SessionCard_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"SessionCard_Click: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის ბარათზე ორჯერ დაწკაპუნების დამუშავება
    ''' ხსნის სესიის რედაქტირების ფორმას
    ''' </summary>
    Private Sub SessionCard_DoubleClick(sender As Object, e As EventArgs)
        Try
            ' ვიპოვოთ სესიის ID
            Dim sessionCard As Panel = DirectCast(sender, Panel)
            Dim sessionId As Integer = CInt(sessionCard.Tag)

            Debug.WriteLine($"SessionCard_DoubleClick: სესიის რედაქტირების მოთხოვნა, ID={sessionId}")

            ' შევამოწმოთ არის თუ არა მომხმარებელი ავტორიზებული
            If String.IsNullOrEmpty(userEmail) Then
                MessageBox.Show("სესიის რედაქტირებისთვის საჭიროა ავტორიზაცია",
                             "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ხელმისაწვდომი",
                             "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' შევქმნათ და გავხსნათ რედაქტირების ფორმა
            Dim editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_Calendar")

            ' დავამატოთ ფორმის დახურვის ივენთის დამუშავება
            AddHandler editForm.FormClosed, Sub(s, args)
                                                ' თუ ფორმა დაიხურა OK რეზულტატით, განვაახლოთ მონაცემები
                                                If editForm.DialogResult = DialogResult.OK Then
                                                    Debug.WriteLine($"SessionCard_DoubleClick: სესია ID={sessionId} წარმატებით განახლდა")

                                                    ' ჩავტვირთოთ სესიები თავიდან
                                                    LoadSessions()

                                                    ' განვაახლოთ კალენდრის ხედი
                                                    UpdateCalendarView()
                                                End If
                                            End Sub

            ' გავხსნათ ფორმა
            editForm.ShowDialog()

            Debug.WriteLine($"SessionCard_DoubleClick: სესიის რედაქტირების ფორმა გაიხსნა ID={sessionId}-სთვის")

        Catch ex As Exception
            Debug.WriteLine($"SessionCard_DoubleClick: შეცდომა - {ex.Message}")
            Debug.WriteLine($"SessionCard_DoubleClick: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"სესიის რედაქტირების ფორმის გახსნის შეცდომა: {ex.Message}",
                         "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სესიის სტატუსის სწრაფი ცვლილება
    ''' გამოიყენება კონტექსტური მენიუდან
    ''' </summary>
    Private Sub QuickChangeSessionStatus(session As SessionModel, newStatus As String)
        Try
            Debug.WriteLine($"QuickChangeSessionStatus: სესიის ID={session.Id} სტატუსის შეცვლა '{session.Status}' -> '{newStatus}'")

            ' შევამოწმოთ არის თუ არა მომხმარებელი ავტორიზებული
            If String.IsNullOrEmpty(userEmail) Then
                MessageBox.Show("სესიის სტატუსის შესაცვლელად საჭიროა ავტორიზაცია",
                             "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი ვერ მოიძებნა",
                             "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' დადასტურების მესიჯბოქსი
            Dim result = MessageBox.Show(
                $"გსურთ სესიის #{session.Id} ({session.BeneficiaryName} {session.BeneficiarySurname}) სტატუსის შეცვლა?" & Environment.NewLine &
                $"ძველი სტატუსი: {session.Status}" & Environment.NewLine &
                $"ახალი სტატუსი: {newStatus}",
                "სტატუსის შეცვლა",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1)

            If result = DialogResult.Yes Then
                ' მივიღოთ ყველა სესიის მონაცემები
                Dim allSessionsData = dataService.GetData("DB-Schedule!A2:O")

                ' ვიპოვოთ ჩვენი სესიის მწკრივი
                For i As Integer = 0 To allSessionsData.Count - 1
                    Dim row = allSessionsData(i)
                    If row.Count > 0 AndAlso Integer.TryParse(row(0).ToString(), Nothing) Then
                        Dim rowId As Integer = Integer.Parse(row(0).ToString())
                        If rowId = session.Id Then
                            ' განვაახლოთ სტატუსი (M სვეტი - ინდექსი 12)
                            Dim updatedRow As New List(Of Object)(row)
                            If updatedRow.Count > 12 Then updatedRow(12) = newStatus

                            ' განვაახლოთ მონაცემები Google Sheets-ში
                            Dim updateRange As String = $"DB-Schedule!A{i + 2}:O{i + 2}"
                            dataService.UpdateData(updateRange, updatedRow)

                            Debug.WriteLine($"QuickChangeSessionStatus: წარმატებით განახლდა სესია ID={session.Id}")

                            ' ჩავტვირთოთ სესიები თავიდან
                            LoadSessions()

                            ' განვაახლოთ კალენდრის ხედი
                            UpdateCalendarView()

                            ' შეტყობინება
                            MessageBox.Show($"სესიის სტატუსი წარმატებით შეიცვალა: {newStatus}",
                                         "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)

                            Exit For
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            Debug.WriteLine($"QuickChangeSessionStatus: შეცდომა - {ex.Message}")
            Debug.WriteLine($"QuickChangeSessionStatus: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"სტატუსის შეცვლისას მოხდა შეცდომა: {ex.Message}",
                         "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სესიის ბარათისთვის კონტექსტური მენიუს შექმნა
    ''' </summary>
    Private Sub CreateSessionCardContextMenu(sessionCard As Panel, session As SessionModel)
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

            ' რედაქტირების პუნქტი
            Dim editMenuItem As New ToolStripMenuItem("რედაქტირება")
            AddHandler editMenuItem.Click, Sub(sender, e)
                                               Try
                                                   ' ვხსნით რედაქტირების ფორმას
                                                   Dim editForm As New NewRecordForm(dataService, "სესია", session.Id, userEmail, "UC_Calendar")
                                                   Dim result = editForm.ShowDialog()

                                                   ' ფორმის დახურვის შემდეგ ვანახლებთ მონაცემებს
                                                   If result = DialogResult.OK Then
                                                       LoadSessions()
                                                       UpdateCalendarView()
                                                   End If
                                               Catch ex As Exception
                                                   Debug.WriteLine($"EditMenuItem_Click: შეცდომა - {ex.Message}")
                                                   MessageBox.Show($"რედაქტირების შეცდომა: {ex.Message}",
                                                               "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                               End Try
                                           End Sub
            contextMenu.Items.Add(editMenuItem)

            ' სტატუსის შეცვლის ქვემენიუ
            Dim statusMenuItem As New ToolStripMenuItem("სტატუსის შეცვლა")

            ' ყველა შესაძლო სტატუსი
            Dim statuses As String() = {"დაგეგმილი", "შესრულებული", "გაუქმებული", "გაცდენა საპატიო", "გაცდენა არასაპატიო", "აღდგენა", "პროგრამით გატარება"}

            ' დავამატოთ ყველა სტატუსი ქვემენიუში, გარდა ამჟამინდელი სტატუსისა
            For Each status In statuses
                ' ამჟამინდელი სტატუსი არ გვინდა მენიუში
                If Not String.Equals(status, session.Status, StringComparison.OrdinalIgnoreCase) Then
                    Dim statusItem As New ToolStripMenuItem(status)

                    ' ვუყენებთ შესაბამის ფერს
                    statusItem.BackColor = SessionStatusColors.GetStatusColor(status, DateTime.Now)

                    ' ვამატებთ ქმედებას
                    Dim statusName As String = status ' ლოკალური ცვლადი closure-ისთვის
                    AddHandler statusItem.Click, Sub(sender, e) QuickChangeSessionStatus(session, statusName)

                    statusMenuItem.DropDownItems.Add(statusItem)
                End If
            Next

            contextMenu.Items.Add(statusMenuItem)

            ' გამყოფი
            contextMenu.Items.Add(New ToolStripSeparator())

            ' ინფორმაციის ჩვენების პუნქტი
            Dim showInfoMenuItem As New ToolStripMenuItem("ინფორმაციის ნახვა")
            AddHandler showInfoMenuItem.Click, Sub(sender, e) SessionCard_Click(sessionCard, New EventArgs())
            contextMenu.Items.Add(showInfoMenuItem)

            ' კონტექსტური მენიუს მიბმა ბარათზე
            sessionCard.ContextMenuStrip = contextMenu

            Debug.WriteLine($"CreateSessionCardContextMenu: შეიქმნა კონტექსტური მენიუ სესიისთვის ID={session.Id}")

        Catch ex As Exception
            Debug.WriteLine($"CreateSessionCardContextMenu: შეცდომა - {ex.Message}")
            Debug.WriteLine($"CreateSessionCardContextMenu: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ყველა გრიდის უჯრედებიდან სესიების ბარათების წაშლა
    ''' </summary>
    Private Sub ClearSessionCardsFromGrid()
        Try
            Debug.WriteLine("ClearSessionCardsFromGrid: იწყება ყველა სესიის ბარათის წაშლა")

            ' 1. თუ გრიდის უჯრედები ინიციალიზებულია, გავასუფთავოთ
            If gridCells IsNot Nothing Then
                For col As Integer = 0 To gridCells.GetLength(0) - 1
                    For row As Integer = 0 To gridCells.GetLength(1) - 1
                        Dim cell = gridCells(col, row)
                        If cell IsNot Nothing Then
                            cell.Controls.Clear()
                        End If
                    Next
                Next

                Debug.WriteLine("ClearSessionCardsFromGrid: gridCells მასივიდან წაშლილია ყველა ბარათი")
            End If

            ' 2. ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel IsNot Nothing Then
                ' მოვძებნოთ ყველა ბარათი, რომლებსაც Tag აქვთ (სესიის ID)
                Dim cardsToRemove As New List(Of Control)()

                For Each ctrl As Control In mainGridPanel.Controls
                    If TypeOf ctrl Is Panel AndAlso ctrl.Tag IsNot Nothing AndAlso TypeOf ctrl.Tag Is Integer Then
                        ' ეს არის სესიის ბარათი
                        cardsToRemove.Add(ctrl)
                    End If
                Next

                ' წავშალოთ ყველა ნაპოვნი ბარათი
                For Each card In cardsToRemove
                    mainGridPanel.Controls.Remove(card)
                    card.Dispose()
                Next

                Debug.WriteLine($"ClearSessionCardsFromGrid: mainGridPanel-დან წაიშალა {cardsToRemove.Count} ბარათი")
            End If

            Debug.WriteLine("ClearSessionCardsFromGrid: ყველა ბარათი წაიშალა")

        Catch ex As Exception
            Debug.WriteLine($"ClearSessionCardsFromGrid: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ClearSessionCardsFromGrid: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' დიაგნოსტიკური მეთოდი სკროლის პრობლემების შესამოწმებლად
    ''' </summary>
    Public Sub DiagnoseScrollIssue()
        Try
            Debug.WriteLine("=== DiagnoseScrollIssue: დაიწყო ===")

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            Debug.WriteLine($"mainGridPanel: {If(mainGridPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")

            If mainGridPanel IsNot Nothing Then
                Debug.WriteLine($"  ზომა: {mainGridPanel.Size}")
                Debug.WriteLine($"  სკროლის მნიშვნელობები: H={mainGridPanel.HorizontalScroll.Value}, V={mainGridPanel.VerticalScroll.Value}")
                Debug.WriteLine($"  AutoScrollMinSize: {mainGridPanel.AutoScrollMinSize}")
                Debug.WriteLine($"  AutoScroll: {mainGridPanel.AutoScroll}")
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

            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)
            Debug.WriteLine($"beneficiaryHeaderPanel: {If(beneficiaryHeaderPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")
            If beneficiaryHeaderPanel IsNot Nothing Then
                Debug.WriteLine($"  ზომა: {beneficiaryHeaderPanel.Size}, ადგილმდებარეობა: {beneficiaryHeaderPanel.Location}, ხილვადი: {beneficiaryHeaderPanel.Visible}")
            End If

            ' ვიპოვოთ დროის პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)
            Debug.WriteLine($"timeColumnPanel: {If(timeColumnPanel IsNot Nothing, "მოიძებნა", "არ მოიძებნა")}")
            If timeColumnPanel IsNot Nothing Then
                Debug.WriteLine($"  ზომა: {timeColumnPanel.Size}, ადგილმდებარეობა: {timeColumnPanel.Location}, ხილვადი: {timeColumnPanel.Visible}")
            End If

            Debug.WriteLine("=== DiagnoseScrollIssue: დასრულდა ===")

        Catch ex As Exception
            Debug.WriteLine($"DiagnoseScrollIssue: შეცდომა - {ex.Message}")
            Debug.WriteLine($"DiagnoseScrollIssue: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ახალი ბენეფიციარის ხედის გამართვა
    ''' მოიხმარება აპლიკაციის გაშვებისთანავე
    ''' </summary>
    Public Sub SetupBeneficiaryView()
        Try
            Debug.WriteLine("SetupBeneficiaryView: ვამზადებთ ბენეფიციარის ხედს")

            ' ინიციალიზაცია
            If RBBene.Checked Then
                ' უკვე არჩეულია ბენეფიციარის ხედი, მაგრამ შეიძლება საჭიროა სრული განახლება
                ShowDayViewByBeneficiary()
            End If

            Debug.WriteLine("SetupBeneficiaryView: ბენეფიციარის ხედი მომზადებულია")

        Catch ex As Exception
            Debug.WriteLine($"SetupBeneficiaryView: შეცდომა - {ex.Message}")
            Debug.WriteLine($"SetupBeneficiaryView: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' მონაცემების მომზადება კალენდრის მომხმარებლისთვის ჩვენებამდე
    ''' </summary>
    Public Sub PrepareCalendarData()
        Try
            Debug.WriteLine("PrepareCalendarData: მიმდინარეობს კალენდრის მონაცემების მომზადება")

            ' სივრცეების ჩატვირთვა
            LoadSpaces()

            ' თერაპევტების ჩატვირთვა
            LoadTherapistsForDate()

            ' ყველა სესიის ჩატვირთვა მონაცემთა წყაროდან
            LoadSessions()

            Debug.WriteLine("PrepareCalendarData: კალენდრის მონაცემები წარმატებით მომზადდა")

        Catch ex As Exception
            Debug.WriteLine($"PrepareCalendarData: შეცდომა - {ex.Message}")
            Debug.WriteLine($"PrepareCalendarData: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' კლასი, რომელიც იძლევა სტანდარტიზებულ ფერებს სესიის სტატუსების მიხედვით
    ''' </summary>
    Public Class SessionStatusColors
        ''' <summary>
        ''' სტატუსის მიხედვით სესიის ბარათის ფერის დაბრუნება
        ''' </summary>
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
                    ' უცნობი სტატუსებისთვის ნაგულისხმევი ფერი
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
                Return Color.FromArgb(255, 182, 193) ' ღ
    