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
                "მწვანე აბა", "ლურჯი აბა", "სენსორი", "ფიზიკური", "მეტყველება",
                "მუსიკა", "თერაპევტი", "არტი", "სხვა", "საკონფერენციო",
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
    ''' კალენდრის დღის ხედის პანელების ინიციალიზაცია - გაუმჯობესებული მასშტაბირებისთვის
    ''' </summary>
    Private Sub InitializeDayViewPanels()
        Try
            Debug.WriteLine("InitializeDayViewPanels: დაიწყო პანელების ინიციალიზაცია")

            ' გავასუფთავოთ კალენდრის პანელი
            pnlCalendarGrid.Controls.Clear()

            ' დავაყენოთ კალენდრის პანელის სკროლი
            pnlCalendarGrid.AutoScroll = False

            ' წავკითხოთ მშობელი პანელის ზომები
            Dim totalWidth As Integer = pnlCalendarGrid.ClientSize.Width - 5 ' 5 პიქსელი დაშორებისთვის
            Dim totalHeight As Integer = pnlCalendarGrid.ClientSize.Height - 5

            Debug.WriteLine($"InitializeDayViewPanels: კონტეინერის ზომები - Width={totalWidth}, Height={totalHeight}")

            ' კონსტანტები, რომლებიც არ იცვლება მასშტაბირებისას
            Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT
            Dim TIME_COLUMN_WIDTH As Integer = BASE_TIME_COLUMN_WIDTH

            ' სივრცის სვეტის სიგანე (იმასშტაბირდება)
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' სივრცეების სათაურების პანელის სრული სიგანე
            Dim totalHeaderWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count

            ' ======= 1. შევქმნათ საათების პანელი (მარცხნივ) =======
            Dim timeColumnPanel As New Panel()
            timeColumnPanel.Name = "timeColumnPanel"
            timeColumnPanel.Size = New Size(TIME_COLUMN_WIDTH, totalHeight - HEADER_HEIGHT)
            timeColumnPanel.Location = New Point(0, HEADER_HEIGHT)
            timeColumnPanel.BackColor = Color.FromArgb(240, 240, 245)
            timeColumnPanel.BorderStyle = BorderStyle.FixedSingle
            timeColumnPanel.AutoScroll = False ' მნიშვნელოვანია - გავთიშოთ სკროლი
            pnlCalendarGrid.Controls.Add(timeColumnPanel)

            ' ======= 2. შევქმნათ სივრცეების სათაურების პანელი (ზემოთ) =======
            Dim spacesHeaderPanel As New Panel()
            spacesHeaderPanel.Name = "spacesHeaderPanel"
            ' მნიშვნელოვანი: სივრცეების სათაურების პანელის სიგანე ემთხვევა სრულ სიგანეს
            spacesHeaderPanel.Size = New Size(totalHeaderWidth, HEADER_HEIGHT)
            spacesHeaderPanel.Location = New Point(TIME_COLUMN_WIDTH, 0)
            spacesHeaderPanel.BackColor = Color.FromArgb(220, 220, 240)
            spacesHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            spacesHeaderPanel.AutoScroll = False ' მნიშვნელოვანია - გავთიშოთ სკროლი
            pnlCalendarGrid.Controls.Add(spacesHeaderPanel)

            ' ======= 3. შევქმნათ დროის სათაურის პანელი (მარცხენა ზედა კუთხეში) =======
            Dim timeHeaderPanel As New Panel()
            timeHeaderPanel.Name = "timeHeaderPanel"
            timeHeaderPanel.Size = New Size(TIME_COLUMN_WIDTH, HEADER_HEIGHT)
            timeHeaderPanel.Location = New Point(0, 0)
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

            ' ======= 4. შევქმნათ მთავარი გრიდის პანელი =======
            Dim mainGridPanel As New Panel()
            mainGridPanel.Name = "mainGridPanel"
            mainGridPanel.Size = New Size(totalWidth - TIME_COLUMN_WIDTH, totalHeight - HEADER_HEIGHT)
            mainGridPanel.Location = New Point(TIME_COLUMN_WIDTH, HEADER_HEIGHT)
            mainGridPanel.BackColor = Color.White
            mainGridPanel.BorderStyle = BorderStyle.FixedSingle
            mainGridPanel.AutoScroll = True ' მხოლოდ მთავარ პანელს აქვს სკროლი ჩართული
            pnlCalendarGrid.Controls.Add(mainGridPanel)

            Debug.WriteLine("InitializeDayViewPanels: პანელების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"InitializeDayViewPanels: შეცდომა - {ex.Message}")
            Debug.WriteLine($"InitializeDayViewPanels: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"პანელების ინიციალიზაციის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დროის პანელის შევსება ინტერვალებით
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

            ' დროის ეტიკეტები
            For i As Integer = 0 To timeIntervals.Count - 1
                Dim timeLabel As New Label()
                timeLabel.Size = New Size(timeColumnPanel.Width - 5, ROW_HEIGHT)
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

                timeColumnPanel.Controls.Add(timeLabel)
            Next

            ' მთლიანი სიმაღლის დაყენება
            timeColumnPanel.Height = timeIntervals.Count * ROW_HEIGHT

            Debug.WriteLine($"FillTimeColumnPanel: დაემატა {timeIntervals.Count} დროის ეტიკეტი")

        Catch ex As Exception
            Debug.WriteLine($"FillTimeColumnPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სივრცეების სათაურების პანელის შევსება და ზომის დაყენება მასშტაბის გათვალისწინებით
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

            ' სივრცის სვეტის სიგანე (იმასშტაბირდება)
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

            ' გამოვთვალოთ სათაურების პანელის სრული სიგანე
            Dim totalHeaderWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count

            ' შევცვალოთ პანელის სიგანე მასშტაბის მიხედვით
            ' ეს არის მთავარი ცვლილება - ვზრდით სათაურების პანელის სიგანეს
            spacesHeaderPanel.Width = totalHeaderWidth

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

            ' სივრცეების სათაურების შექმნა
            For i As Integer = 0 To spaces.Count - 1
                Dim spaceHeader As New Label()
                spaceHeader.Size = New Size(SPACE_COLUMN_WIDTH, spacesHeaderPanel.Height - 2)
                spaceHeader.Location = New Point(i * SPACE_COLUMN_WIDTH, 1)
                spaceHeader.BackColor = Color.FromArgb(60, 80, 150)
                spaceHeader.ForeColor = Color.White
                spaceHeader.TextAlign = ContentAlignment.MiddleCenter
                spaceHeader.Text = spaces(i).ToUpper()
                spaceHeader.Font = mtavruliFont
                spacesHeaderPanel.Controls.Add(spaceHeader)
            Next

            Debug.WriteLine($"FillSpacesHeaderPanel: დაემატა {spaces.Count} სივრცის სათაური, მასშტაბი: {hScale}")
            Debug.WriteLine($"FillSpacesHeaderPanel: სათაურების პანელის სიგანე: {totalHeaderWidth}")

        Catch ex As Exception
            Debug.WriteLine($"FillSpacesHeaderPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის შევსება გაუმჯობესებული ჰორიზონტალური მასშტაბირებით
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

            ' გავასუფთავოთ
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

            ' გამოვთვალოთ ხელმისაწვდომი სიგანე
            Dim availableWidth As Integer = mainGridPanel.Width

            ' გრიდის სრული სიგანე (გავითვალისწინოთ მასშტაბი)
            Dim gridWidth As Integer = SPACE_COLUMN_WIDTH * spaces.Count

            ' გრიდის სრული სიმაღლე
            Dim gridHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            ' დავაყენოთ მთავარი გრიდის პანელის AutoScrollMinSize საკმარისად დიდი ზომით
            ' ეს უზრუნველყოფს, რომ სკროლი გამოჩნდეს და შესაძლებელი იყოს ყველა სვეტის ნახვა
            mainGridPanel.AutoScrollMinSize = New Size(gridWidth, gridHeight)

            Debug.WriteLine($"FillMainGridPanel: გრიდის ზომები - Width={gridWidth}, Height={gridHeight}, hScale={hScale}")

            ' უჯრედების მასივი
            ReDim gridCells(spaces.Count - 1, timeIntervals.Count - 1)

            ' უჯრედების შექმნა
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

                    ' ვამატებთ უჯრედს პანელზე
                    mainGridPanel.Controls.Add(cell)
                Next
            Next

            Debug.WriteLine($"FillMainGridPanel: შეიქმნა {spaces.Count * timeIntervals.Count} უჯრედი")

        Catch ex As Exception
            Debug.WriteLine($"FillMainGridPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' დღის ხედის ჩვენება სივრცეების მიხედვით - ახალი ვერსია სამი პანელით
    ''' </summary>
    Private Sub ShowDayViewBySpace()
        Try
            Debug.WriteLine("ShowDayViewBySpace: დაიწყო დღის ხედის ჩვენება სივრცეების მიხედვით")

            ' დავაყენოთ pnlFilter-ის ფერი - ნახევრად გამჭვირვალე თეთრი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' ======= 1. პანელების ინიციალიზაცია =======
            InitializeDayViewPanels()

            ' ======= 2. დროის პანელის შევსება =======
            FillTimeColumnPanel()

            ' ======= 3. სივრცეების სათაურების პანელის შევსება =======
            FillSpacesHeaderPanel()

            ' ======= 4. მთავარი გრიდის პანელის შევსება =======
            FillMainGridPanel()

            ' ======= 5. სინქრონიზაცია სქროლისთვის =======
            SetupScrollSynchronization()

            Debug.WriteLine("ShowDayViewBySpace: დღის ხედის ჩვენება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewBySpace: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewBySpace: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"დღის ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სკროლის სინქრონიზაცია - გაუმჯობესებული ვერსია ორმაგი სკროლის პრობლემის გადასაჭრელად
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

            ' გავთიშოთ დროის პანელის სკროლი
            timeColumnPanel.AutoScroll = False

            ' გავთიშოთ სივრცეების სათაურების პანელის სკროლი
            spacesHeaderPanel.AutoScroll = False

            ' მივაბათ Scroll ივენთი
            AddHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            Debug.WriteLine("SetupScrollSynchronization: სკროლის სინქრონიზაცია დაყენებულია")

        Catch ex As Exception
            Debug.WriteLine($"SetupScrollSynchronization: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის სკროლის ივენთი - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Sub MainGridPanel_Scroll(sender As Object, e As ScrollEventArgs)
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(sender, Panel)

            ' ვიპოვოთ დროის პანელი
            Dim timeColumnPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("timeColumnPanel", False).FirstOrDefault(), Panel)

            If timeColumnPanel Is Nothing Then
                Return
            End If

            ' ვიპოვოთ სივრცეების სათაურების პანელი
            Dim spacesHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("spacesHeaderPanel", False).FirstOrDefault(), Panel)

            If spacesHeaderPanel Is Nothing Then
                Return
            End If

            ' ვერტიკალური სქროლი - სინქრონიზაცია დროის პანელთან
            If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
                ' დავაყენოთ დროის პანელის ვერტიკალური პოზიცია
                timeColumnPanel.Location = New Point(timeColumnPanel.Location.X, BASE_HEADER_HEIGHT - e.NewValue)
            End If

            ' ჰორიზონტალური სქროლი - სინქრონიზაცია სივრცეების სათაურების პანელთან
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                ' დავაყენოთ სივრცეების სათაურების პანელის ჰორიზონტალური პოზიცია
                ' პანელის პოზიცია იქნება BASE_TIME_COLUMN_WIDTH - სკროლის პოზიცია
                spacesHeaderPanel.Location = New Point(BASE_TIME_COLUMN_WIDTH - e.NewValue, spacesHeaderPanel.Location.Y)

                Debug.WriteLine($"MainGridPanel_Scroll: ჰორიზონტალური სკროლი - {e.NewValue}, სათაურების პანელის X={spacesHeaderPanel.Location.X}")
            End If

        Catch ex As Exception
            Debug.WriteLine($"MainGridPanel_Scroll: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' უჯრედის ჩარჩოს დახატვა
    ''' </summary>
    Private Sub Cell_Paint(sender As Object, e As PaintEventArgs)
        Dim cell As Panel = DirectCast(sender, Panel)

        Using pen As New Pen(Color.FromArgb(200, 200, 200), 1)
            ' ბორდერის დახატვა
            e.Graphics.DrawRectangle(pen, 0, 0, cell.Width - 1, cell.Height - 1)
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
    ''' კალენდრის ხედის განახლება - გაუმჯობესებული ვერსია მასშტაბირებისთვის
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

                    ' ზომის დაყენება მასშტაბის გათვალისწინებით
                    SetMainGridPanelSize()

                    ' პანელების ზომების და პოზიციების განახლება
                    UpdatePanelSizesAndPositions()

                    ' დავამატოთ სესიების ჩვენება
                    ShowSessionsInCalendar()
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
    ''' </summary>
    Private Sub UC_Calendar_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            ' საწყისი პარამეტრების დაყენება (თუ ისინი არ არის დაყენებული)
            InitializeTimeIntervals()

            ' კალენდრის განახლება
            UpdateCalendarView()

            Debug.WriteLine("UC_Calendar_Load: კალენდრის კონტროლი წარმატებით ჩაიტვირთა")
        Catch ex As Exception
            Debug.WriteLine($"UC_Calendar_Load: შეცდომა - {ex.Message}")
        End Try
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
    ''' ჰორიზონტალური მასშტაბის გაზრდა - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Sub BtnHUp_Click(sender As Object, e As EventArgs) Handles BtnHUp.Click
        Try
            ' მიმდინარე მასშტაბის შენახვა დებაგინგისთვის
            Dim oldHScale As Double = hScale

            ' მასშტაბის გაზრდა დადგენილი ნაბიჯით
            hScale += SCALE_STEP

            ' მაქსიმალური მასშტაბის შეზღუდვა
            If hScale > MAX_SCALE Then hScale = MAX_SCALE

            Debug.WriteLine($"BtnHUp_Click: ჰორიზონტალური მასშტაბი შეიცვალა {oldHScale:F2} -> {hScale:F2}")

            ' კალენდრის განახლება ახალი მასშტაბით მხოლოდ მაშინ, თუ მასშტაბი შეიცვალა
            If oldHScale <> hScale Then
                ' განვაახლოთ მხოლოდ გრიდის ზომები და პანელები
                InitializeDayViewPanels()
                FillTimeColumnPanel()
                FillSpacesHeaderPanel()
                FillMainGridPanel()
                SetupScrollSynchronization()
                ShowSessionsInCalendar()

                ' განვაახლოთ პანელების ზომები და პოზიციები
                UpdatePanelSizesAndPositions()
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnHUp_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnHUp_Click: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ჰორიზონტალური მასშტაბის შემცირება - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Sub BtnHDown_Click(sender As Object, e As EventArgs) Handles BtnHDown.Click
        Try
            ' მიმდინარე მასშტაბის შენახვა დებაგინგისთვის
            Dim oldHScale As Double = hScale

            ' მასშტაბის შემცირება დადგენილი ნაბიჯით
            hScale -= SCALE_STEP

            ' მინიმალური მასშტაბის შეზღუდვა
            If hScale < MIN_SCALE Then hScale = MIN_SCALE

            Debug.WriteLine($"BtnHDown_Click: ჰორიზონტალური მასშტაბი შეიცვალა {oldHScale:F2} -> {hScale:F2}")

            ' კალენდრის განახლება ახალი მასშტაბით მხოლოდ მაშინ, თუ მასშტაბი შეიცვალა
            If oldHScale <> hScale Then
                ' განვაახლოთ მხოლოდ გრიდის ზომები და პანელები
                InitializeDayViewPanels()
                FillTimeColumnPanel()
                FillSpacesHeaderPanel()
                FillMainGridPanel()
                SetupScrollSynchronization()
                ShowSessionsInCalendar()

                ' განვაახლოთ პანელების ზომები და პოზიციები
                UpdatePanelSizesAndPositions()
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnHDown_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnHDown_Click: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ვერტიკალური მასშტაბის გაზრდა
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
                UpdateCalendarView()
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnVUp_Click: შეცდომა - {ex.Message}")
            Debug.WriteLine($"BtnVUp_Click: StackTrace - {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ვერტიკალური მასშტაბის შემცირება
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
End Class