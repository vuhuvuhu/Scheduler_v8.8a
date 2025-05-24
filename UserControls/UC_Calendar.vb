' ===========================================
' 📄 UserControls/UC_Calendar.vb (ოპტიმიზირებული)
' -------------------------------------------
' კალენდრის User Control - ოპტიმიზირებული, მოდულარული ვერსია
' ===========================================
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Scheduler_v8._8a.Scheduler_v8_8a.Controls
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services
Imports Scheduler_v8._8a.Scheduler_v8_8a.Utils
'Imports Scheduler_v8_8a.Controls
'Imports Scheduler_v8_8a.Models
'Imports Scheduler_v8_8a.Services
'Imports Scheduler_v8_8a.Utils

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
    Private therapists As List(Of String) = New List(Of String)()

    ' გრიდის მენეჯერი
    Private gridManager As CalendarGridManager = Nothing

    ' მასშტაბირების კოეფიციენტები
    Private hScale As Double = 1.0 ' ჰორიზონტალური მასშტაბი
    Private vScale As Double = 1.0 ' ვერტიკალური მასშტაბი

    ' მასშტაბირების ნაბიჯები
    Private Const SCALE_STEP As Double = 0.1 ' მასშტაბის ცვლილების ნაბიჯი
    Private Const MIN_SCALE As Double = 0.5 ' მინიმალური მასშტაბი
    Private Const MAX_SCALE As Double = 4 ' მაქსიმალური მასშტაბი

    ' Drag & Drop ცვლადები
    Private isDragging As Boolean = False
    Private dragStartPoint As Point
    Private draggedCard As Panel = Nothing
    Private originalCardPosition As Point
    Private originalSessionData As SessionModel = Nothing

    ' გრიდის უჯრედები (ლოკალური)
    Private gridCells(,) As Panel

    ''' <summary>
    ''' კონსტრუქტორი კალენდრის ViewModel-ით
    ''' </summary>
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

        ' რადიობუტონების საწყისი მნიშვნელობები
        rbDay.Checked = True
        RBSpace.Checked = True

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
    ''' კალენდრის ხედის განახლება
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
    ''' დღის ხედის ჩვენება სივრცეების მიხედვით
    ''' </summary>
    Private Sub ShowDayViewBySpace()
        Try
            Debug.WriteLine("ShowDayViewBySpace: დაიწყო დღის ხედის ჩვენება სივრცეების მიხედვით")

            ' დავაყენოთ pnlFilter-ის ფერი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' შევქმნათ გრიდის მენეჯერი
            gridManager = New CalendarGridManager(pnlCalendarGrid, spaces, Nothing, timeIntervals, hScale, vScale)

            ' დავაყენოთ ხედვის ტიპი "space"
            gridManager.SetViewType("space")

            ' ინიციალიზაცია დღის ხედისთვის
            gridManager.InitializeDayViewPanels()

            ' თარიღის ზოლის შევსება
            gridManager.FillDateColumnPanel(DTPCalendar.Value.Date)

            ' დროის პანელის შევსება
            gridManager.FillTimeColumnPanel()

            ' სივრცეების სათაურების პანელის შევსება
            gridManager.FillSpacesHeaderPanel()

            ' მთავარი გრიდის პანელის შევსება
            gridManager.FillMainGridPanel()

            ' სქროლის სინქრონიზაცია
            SetupScrollSynchronization()

            ' შევინახოთ გრიდის უჯრედები ლოკალურად
            gridCells = gridManager.gridCells

            ' სესიების განთავსება გრიდში
            PlaceSessionsOnGrid()

            Debug.WriteLine("ShowDayViewBySpace: დღის ხედის ჩვენება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewBySpace: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewBySpace: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"დღის ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დღის ხედის ჩვენება თერაპევტების მიხედვით
    ''' </summary>
    Private Sub ShowDayViewByTherapist()
        Try
            Debug.WriteLine("ShowDayViewByTherapist: დაიწყო დღის ხედის ჩვენება თერაპევტების მიხედვით")

            ' თერაპევტების ჩატვირთვა
            LoadTherapistsForDate()

            ' დავაყენოთ pnlFilter-ის ფერი
            pnlFIlter.BackColor = Color.FromArgb(200, Color.White)

            ' შევქმნათ გრიდის მენეჯერი თერაპევტებისთვის
            gridManager = New CalendarGridManager(pnlCalendarGrid, Nothing, therapists, timeIntervals, hScale, vScale)

            ' დავაყენოთ ხედვის ტიპი "therapist"
            gridManager.SetViewType("therapist")

            ' ინიციალიზაცია დღის ხედისთვის
            gridManager.InitializeDayViewPanels()

            ' თარიღის ზოლის შევსება
            gridManager.FillDateColumnPanel(DTPCalendar.Value.Date)

            ' დროის პანელის შევსება
            gridManager.FillTimeColumnPanel()

            ' თერაპევტების სათაურების პანელის შევსება
            gridManager.FillTherapistsHeaderPanel()

            ' მთავარი გრიდის პანელის შევსება თერაპევტებისთვის
            gridManager.FillMainGridPanelForTherapists()

            ' სქროლის სინქრონიზაცია
            SetupScrollSynchronization()

            ' შევინახოთ გრიდის უჯრედები ლოკალურად
            gridCells = gridManager.gridCells

            ' სესიების განთავსება გრიდში
            PlaceSessionsOnTherapistGrid()

            Debug.WriteLine("ShowDayViewByTherapist: დღის ხედის ჩვენება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewByTherapist: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewByTherapist: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"თერაპევტების ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დღის ხედის ჩვენება ბენეფიციარების მიხედვით
    ''' </summary>
    Private Sub ShowDayViewByBeneficiary()
        Try
            Debug.WriteLine("ShowDayViewByBeneficiary: დაიწყო ბენეფიციარების ხედის ჩვენება")

            ' ბენეფიციარების ხედის პანელების ინიციალიზაცია
            InitializeBeneficiaryDayViewPanels()

            Debug.WriteLine("ShowDayViewByBeneficiary: ბენეფიციარების ხედის ჩვენება დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowDayViewByBeneficiary: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowDayViewByBeneficiary: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"ბენეფიციარების ხედის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

            ' ვალაგებთ თერაპევტებს სახელოვნად
            therapists.AddRange(therapistSet.OrderBy(Function(t) t))

            ' თუ არ მოიძებნა თერაპევტები, დავამატოთ საცდელი
            If therapists.Count = 0 Then
                Debug.WriteLine("LoadTherapistsForDate: თერაპევტები არ მოიძებნა, ვამატებთ საცდელ მონაცემებს")
                therapists.AddRange({"ჩანაწერი არ არის"})
            End If

            Debug.WriteLine($"LoadTherapistsForDate: ჩატვირთულია {therapists.Count} თერაპევტი")
        Catch ex As Exception
            Debug.WriteLine($"LoadTherapistsForDate: შეცდომა - {ex.Message}")

            ' შეცდომის შემთხვევაში გამოვიყენოთ საცდელი თერაპევტები
            therapists.Clear()
            therapists.AddRange({"ჩანაწერი არ არის"})
        End Try
    End Sub

    ''' <summary>
    ''' სქროლის სინქრონიზაცია
    ''' </summary>
    Private Sub SetupScrollSynchronization()
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("SetupScrollSynchronization: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' ჯერ მოვხსნათ არსებული ივენთის ჰენდლერები, თუ ისინი არსებობენ
            RemoveHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll
            RemoveHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            ' მივაბათ Scroll ივენთი ხელახლა
            AddHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' დავამატოთ ივენთი გორგოლჭის სქროლისთვის
            AddHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            Debug.WriteLine("SetupScrollSynchronization: სქროლის სინქრონიზაცია დაყენებულია")
        Catch ex As Exception
            Debug.WriteLine($"SetupScrollSynchronization: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთავარი გრიდის პანელის სქროლის ივენთი
    ''' </summary>
    Private Sub MainGridPanel_Scroll(sender As Object, e As ScrollEventArgs)
        Try
            Dim mainGridPanel As Panel = DirectCast(sender, Panel)

            ' სქროლის პოზიცია
            Dim scrollPosition As Point = New Point(
                -mainGridPanel.HorizontalScroll.Value,
                -mainGridPanel.VerticalScroll.Value
            )

            ' გრიდის მენეჯერის გამოყენება სქროლის სინქრონიზაციისთვის
            If gridManager IsNot Nothing Then
                gridManager.SynchronizeScroll(scrollPosition)
            End If
        Catch ex As Exception
            Debug.WriteLine($"MainGridPanel_Scroll: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მაუსის გორგოლჭის ივენთი
    ''' </summary>
    Private Sub MainGridPanel_MouseWheel(sender As Object, e As MouseEventArgs)
        Try
            Dim mainGridPanel As Panel = DirectCast(sender, Panel)

            ' სქროლის პოზიცია
            Dim scrollPosition As Point = New Point(
                -mainGridPanel.HorizontalScroll.Value,
                -mainGridPanel.VerticalScroll.Value
            )

            ' გრიდის მენეჯერის გამოყენება სქროლის სინქრონიზაციისთვის
            If gridManager IsNot Nothing Then
                gridManager.SynchronizeScroll(scrollPosition)
            End If
        Catch ex As Exception
            Debug.WriteLine($"MainGridPanel_MouseWheel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების განთავსება გრიდში
    ''' </summary>
    Private Sub PlaceSessionsOnGrid()
        Try
            Debug.WriteLine("=== PlaceSessionsOnGrid: დაიწყო ===")

            ' ძირითადი შემოწმებები
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnGrid: სესიები არ არის")
                Return
            End If

            If gridManager Is Nothing Then
                Debug.WriteLine("PlaceSessionsOnGrid: გრიდის მენეჯერი არ არის")
                Return
            End If

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            If mainGridPanel Is Nothing Then
                Debug.WriteLine("❌ მთავარი გრიდის პანელი ვერ მოიძებნა!")
                Return
            End If

            ' გავასუფთავოთ არსებული სესიის ბარათები მაინპანელიდან
            For Each ctrl In mainGridPanel.Controls.OfType(Of Panel)().ToList()
                If ctrl.Tag IsNot Nothing AndAlso TypeOf ctrl.Tag Is Integer Then
                    mainGridPanel.Controls.Remove(ctrl)
                    ctrl.Dispose()
                End If
            Next

            ' მასშტაბირებული პარამეტრები
            Dim scalingFactors = gridManager.GetScaleFactors()
            Dim ROW_HEIGHT As Integer = CInt(CalendarGridManager.BASE_ROW_HEIGHT * scalingFactors.vScale)
            Dim SPACE_COLUMN_WIDTH As Integer = CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * scalingFactors.hScale)

            ' ფილტრირებული სესიები
            Dim filteredSessions = GetFilteredSessions()
            Debug.WriteLine($"ნაპოვნია {filteredSessions.Count} ფილტრირებული სესია ამ თარიღისთვის")

            ' სესიების განთავსება
            For Each session In filteredSessions
                Try
                    Debug.WriteLine($"--- ვამუშავებთ სესიას ID={session.Id}, სივრცე='{session.Space}', დრო={session.DateTime:HH:mm} ---")

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
                    Dim timeIndex As Integer = CalendarUtils.FindNearestTimeIndex(sessionTime, timeIntervals)

                    ' 3. საზღვრების შემოწმება
                    If spaceIndex < 0 OrElse timeIndex < 0 OrElse
                   spaceIndex >= spaces.Count OrElse timeIndex >= timeIntervals.Count Then
                        Debug.WriteLine($"❌ არასწორი ინდექსები: space={spaceIndex}, time={timeIndex}")
                        Continue For
                    End If

                    ' 4. ბარათის პოზიციის გამოთვლა
                    Dim cardX As Integer = spaceIndex * SPACE_COLUMN_WIDTH + 4
                    Dim cardY As Integer = timeIndex * ROW_HEIGHT + 2

                    ' 5. პროპორციული ბარათის სიმაღლის გამოთვლა
                    Dim sessionDurationMinutes As Integer = session.Duration
                    Dim baseCardHeight As Double = ROW_HEIGHT * (sessionDurationMinutes / 30.0)
                    Dim minCardHeight As Integer = ROW_HEIGHT
                    Dim cardHeight As Integer = CInt(Math.Max(baseCardHeight, minCardHeight))
                    Dim maxCardHeight As Integer = ROW_HEIGHT * 4
                    If cardHeight > maxCardHeight Then cardHeight = maxCardHeight

                    ' 6. ბარათის შექმნა SessionCardFactory-ის გამოყენებით
                    Dim sessionCard = SessionCardFactory.CreateSessionCard(
    session,
    SPACE_COLUMN_WIDTH - 8,
    cardHeight,
    New Point(cardX, cardY),
    AddressOf EditSession ' დამატებული: რედაქტირების დელეგატი
)

                    ' 7. კონტექსტური მენიუს შექმნა
                    SessionCardFactory.CreateContextMenu(sessionCard, session, AddressOf QuickChangeStatus_Click)

                    ' 8. მაუსის ივენთების მიბმა
                    AddHandler sessionCard.MouseDown, AddressOf SessionCard_MouseDown
                    AddHandler sessionCard.MouseMove, AddressOf SessionCard_MouseMove
                    AddHandler sessionCard.MouseUp, AddressOf SessionCard_MouseUp
                    AddHandler sessionCard.Click, AddressOf SessionCard_Click
                    AddHandler sessionCard.DoubleClick, AddressOf SessionCard_DoubleClick

                    ' კონტექსტური მენიუს ივენთების დამატება
                    If sessionCard.ContextMenuStrip IsNot Nothing Then
                        ' რედაქტირების მენიუს პუნქტი
                        Dim editMenuItem = sessionCard.ContextMenuStrip.Items.Cast(Of ToolStripItem)().FirstOrDefault(Function(item) TypeOf item Is ToolStripMenuItem AndAlso item.Text = "რედაქტირება")
                        If editMenuItem IsNot Nothing Then
                            AddHandler editMenuItem.Click, Sub(menuSender, menuE)
                                                               Dim menuItem = DirectCast(menuSender, ToolStripMenuItem)
                                                               Dim menuSessionId As Integer = CInt(menuItem.Tag)
                                                               EditSession(menuSessionId)
                                                           End Sub
                        End If

                        ' ინფორმაციის ნახვა მენიუს პუნქტი
                        Dim infoMenuItem = sessionCard.ContextMenuStrip.Items.Cast(Of ToolStripItem)().FirstOrDefault(Function(item) TypeOf item Is ToolStripMenuItem AndAlso item.Text = "ინფორმაციის ნახვა")
                        ' კონტექსტური მენიუს ივენთების დამატებისას:
                        ' ინფორმაციის ნახვა მენიუს პუნქტი განახლება
                        If infoMenuItem IsNot Nothing Then
                            AddHandler infoMenuItem.Click, Sub(menuSender, menuE)
                                                               Dim menuItem = DirectCast(menuSender, ToolStripMenuItem)
                                                               Dim menuSessionId As Integer = CInt(menuItem.Tag)
                                                               ' ახალი მეთოდის გამოძახება
                                                               Dim clickedSession = allSessions.FirstOrDefault(Function(s) s.Id = menuSessionId)
                                                               If clickedSession IsNot Nothing Then
                                                                   ShowSessionDetails(clickedSession)
                                                               End If
                                                           End Sub
                        End If
                    End If
                    ' ბარათის გრიდზე დამატება
                    mainGridPanel.Controls.Add(sessionCard)
                    sessionCard.BringToFront()

                    Debug.WriteLine($"✅ ბარათი განთავსდა: ID={session.Id}, სივრცე='{session.Space}', " &
                     $"ზომა={sessionCard.Width}x{sessionCard.Height}px")

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
    ''' სესიების განთავსება თერაპევტების გრიდში
    ''' </summary>
    Private Sub PlaceSessionsOnTherapistGrid()
        Try
            Debug.WriteLine("=== PlaceSessionsOnTherapistGrid: დაიწყო ===")

            ' ძირითადი შემოწმებები
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnTherapistGrid: სესიები არ არის")
                Return
            End If

            If gridManager Is Nothing Then
                Debug.WriteLine("PlaceSessionsOnTherapistGrid: გრიდის მენეჯერი არ არის")
                Return
            End If

            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)
            If mainGridPanel Is Nothing Then
                Debug.WriteLine("❌ მთავარი გრიდის პანელი ვერ მოიძებნა!")
                Return
            End If

            ' გავასუფთავოთ არსებული სესიის ბარათები მაინპანელიდან
            For Each ctrl In mainGridPanel.Controls.OfType(Of Panel)().ToList()
                If ctrl.Tag IsNot Nothing AndAlso TypeOf ctrl.Tag Is Integer Then
                    mainGridPanel.Controls.Remove(ctrl)
                    ctrl.Dispose()
                End If
            Next

            ' მასშტაბირებული პარამეტრები
            Dim scalingFactors = gridManager.GetScaleFactors()
            Dim ROW_HEIGHT As Integer = CInt(CalendarGridManager.BASE_ROW_HEIGHT * scalingFactors.vScale)
            Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * scalingFactors.hScale)

            ' ფილტრი - ფილტრირებული სესიები
            Dim filteredSessions = GetFilteredSessions()
            Debug.WriteLine($"ნაპოვნია {filteredSessions.Count} ფილტრირებული სესია ამ თარიღისთვის")

            ' სესიების განთავსება
            For Each session In filteredSessions
                Try
                    Debug.WriteLine($"--- ვამუშავებთ სესიას ID={session.Id}, თერაპევტი='{session.TherapistName}', დრო={session.DateTime:HH:mm} ---")

                    ' 1. თერაპევტის ინდექსის განსაზღვრა
                    Dim therapistIndex As Integer = -1
                    For i As Integer = 0 To therapists.Count - 1
                        If String.Equals(therapists(i).Trim(), session.TherapistName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                            therapistIndex = i
                            Exit For
                        End If
                    Next

                    If therapistIndex < 0 Then
                        Debug.WriteLine($"❌ თერაპევტი '{session.TherapistName}' ვერ მოიძებნა!")
                        Continue For
                    End If

                    ' 2. დროის ინტერვალის განსაზღვრა
                    Dim sessionTime As TimeSpan = session.DateTime.TimeOfDay
                    Dim timeIndex As Integer = CalendarUtils.FindNearestTimeIndex(sessionTime, timeIntervals)

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

                    ' 6. ბარათის შექმნა SessionCardFactory-ის გამოყენებით
                    Dim sessionCard = SessionCardFactory.CreateSessionCard(
    session,
    THERAPIST_COLUMN_WIDTH - 8,
    cardHeight,
    New Point(cardX, cardY),
    AddressOf EditSession ' დამატებული: რედაქტირების დელეგატი
)

                    ' 7. კონტექსტური მენიუს შექმნა
                    SessionCardFactory.CreateContextMenu(sessionCard, session, AddressOf QuickChangeStatus_Click)

                    ' 8. მაუსის ივენთების მიბმა
                    AddHandler sessionCard.MouseDown, AddressOf TherapistSessionCard_MouseDown
                    AddHandler sessionCard.MouseMove, AddressOf TherapistSessionCard_MouseMove
                    AddHandler sessionCard.MouseUp, AddressOf TherapistSessionCard_MouseUp
                    AddHandler sessionCard.Click, AddressOf SessionCard_Click
                    AddHandler sessionCard.DoubleClick, AddressOf SessionCard_DoubleClick

                    ' კონტექსტური მენიუს ივენთების დამატება
                    If sessionCard.ContextMenuStrip IsNot Nothing Then
                        ' რედაქტირების მენიუს პუნქტი
                        Dim editMenuItem = sessionCard.ContextMenuStrip.Items.Cast(Of ToolStripItem)().FirstOrDefault(Function(item) TypeOf item Is ToolStripMenuItem AndAlso item.Text = "რედაქტირება")
                        If editMenuItem IsNot Nothing Then
                            AddHandler editMenuItem.Click, Sub(menuSender, menuE)
                                                               Dim menuItem = DirectCast(menuSender, ToolStripMenuItem)
                                                               Dim menuSessionId As Integer = CInt(menuItem.Tag)
                                                               EditSession(menuSessionId)
                                                           End Sub
                        End If

                        ' ინფორმაციის ნახვა მენიუს პუნქტი
                        Dim infoMenuItem = sessionCard.ContextMenuStrip.Items.Cast(Of ToolStripItem)().FirstOrDefault(Function(item) TypeOf item Is ToolStripMenuItem AndAlso item.Text = "ინფორმაციის ნახვა")
                        If infoMenuItem IsNot Nothing Then
                            AddHandler infoMenuItem.Click, Sub(menuSender, menuE)
                                                               Dim menuItem = DirectCast(menuSender, ToolStripMenuItem)
                                                               Dim menuSessionId As Integer = CInt(menuItem.Tag)
                                                               ' ახალი მეთოდის გამოძახება
                                                               Dim clickedSession = allSessions.FirstOrDefault(Function(s) s.Id = menuSessionId)
                                                               If clickedSession IsNot Nothing Then
                                                                   ShowSessionDetails(clickedSession)
                                                               End If
                                                           End Sub
                        End If
                    End If

                    ' ბარათის გრიდზე დამატება
                    mainGridPanel.Controls.Add(sessionCard)
                    sessionCard.BringToFront()

                    Debug.WriteLine($"✅ ბარათი განთავსდა: ID={session.Id}, თერაპევტი='{session.TherapistName}', " &
                     $"ზომა={sessionCard.Width}x{sessionCard.Height}px")

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

                Debug.WriteLine($"SessionCard_MouseDown: დაიწყო გადატანა სესია ID={sessionId}")
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
                Dim ROW_HEIGHT As Integer = CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale)
                Dim SPACE_COLUMN_WIDTH As Integer = CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale)

                ' ვიპოვოთ მთავარი გრიდის პანელი
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
                Dim newGridPosition = CalendarUtils.CalculateGridPosition(
                    draggedCard.Location,
                    CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale),
                    CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale),
                    spaces.Count,
                    timeIntervals.Count
                )

                Dim originalGridPosition = CalendarUtils.CalculateGridPosition(
                    originalCardPosition,
                    CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale),
                    CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale),
                    spaces.Count,
                    timeIntervals.Count
                )

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
                        ConfirmSessionMove(newGridPosition, "space")
                    Else
                        ' დავაბრუნოთ ძველ ადგილას
                        RevertSessionMove()
                    End If
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

                Debug.WriteLine($"TherapistSessionCard_MouseDown: დაიწყო გადატანა სესია ID={sessionId}")
            End If
        Catch ex As Exception
            Debug.WriteLine($"TherapistSessionCard_MouseDown: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტების გრიდზე სესიის ბარათის MouseMove ივენთი
    ''' </summary>
    Private Sub TherapistSessionCard_MouseMove(sender As Object, e As MouseEventArgs)
        Try
            If isDragging AndAlso draggedCard IsNot Nothing Then
                ' გამოვთვალოთ ახალი პოზიცია მაუსის კოორდინატების მიხედვით
                Dim currentLocation As Point = draggedCard.Location
                currentLocation.X += e.X - dragStartPoint.X
                currentLocation.Y += e.Y - dragStartPoint.Y

                ' მასშტაბირებული პარამეტრები გადატანისთვის
                Dim ROW_HEIGHT As Integer = CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale)
                Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale)

                ' ვიპოვოთ მთავარი გრიდის პანელი
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

                ' გამოვთვალოთ ახალი გრიდის პოზიცია
                Dim newGridPosition = CalendarUtils.CalculateTherapistGridPosition(
                    draggedCard.Location,
                    CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale),
                    CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale),
                    therapists.Count,
                    timeIntervals.Count
                )

                Dim originalGridPosition = CalendarUtils.CalculateTherapistGridPosition(
                    originalCardPosition,
                    CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale),
                    CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale),
                    therapists.Count,
                    timeIntervals.Count
                )

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
                        $"ახალი თერაპევტი: {newTherapist}{Environment.NewLine}" &
                        $"ახალი დრო: {newTime}",
                        "სეანსის განახლება",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1
                    )

                    If result = DialogResult.Yes Then
                        ' დავადასტუროთ ცვლილება
                        ConfirmSessionMove(newGridPosition, "therapist")
                    Else
                        ' დავაბრუნოთ ძველ ადგილას
                        RevertSessionMove()
                    End If
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
    ''' სესიის მოძრაობის დადასტურება და მონაცემების განახლება
    ''' </summary>
    Private Sub ConfirmSessionMove(newGridPosition As (spaceIndex As Integer, timeIndex As Integer), gridType As String)
        Try
            Debug.WriteLine($"ConfirmSessionMove: ახალი პოზიცია - Index1={newGridPosition.spaceIndex}, Time={newGridPosition.timeIndex}, GridType={gridType}")

            If originalSessionData Is Nothing OrElse dataService Is Nothing Then
                Debug.WriteLine("ConfirmSessionMove: originalSessionData ან dataService არის null")
                RevertSessionMove()
                Return
            End If

            ' ვადგენთ ახალ მნიშვნელობებს გრიდის ტიპის მიხედვით
            Dim newFieldValue As String
            Dim fieldIndex As Integer

            If gridType = "space" Then
                ' სივრცის გრიდი
                newFieldValue = spaces(newGridPosition.spaceIndex)
                fieldIndex = 10 ' K სვეტი - სივრცე
            ElseIf gridType = "therapist" Then
                ' თერაპევტის გრიდი
                newFieldValue = therapists(newGridPosition.spaceIndex)
                fieldIndex = 8 ' I სვეტი - თერაპევტი
            Else
                Debug.WriteLine($"ConfirmSessionMove: არასწორი გრიდის ტიპი: {gridType}")
                RevertSessionMove()
                Return
            End If

            ' დროის განახლება (ორივე შემთხვევაში)
            Dim newDateTime As DateTime = timeIntervals(newGridPosition.timeIndex)

            ' შევქმნათ ახალი DateTime, რომელიც ინარჩუნებს ორიგინალურ თარიღს მაგრამ ცვლის დროს
            Dim originalDate As DateTime = originalSessionData.DateTime.Date
            Dim newFullDateTime As DateTime = originalDate.Add(newDateTime.TimeOfDay)

            ' განვაახლოთ სესიის ობიექტი
            If gridType = "space" Then
                originalSessionData.Space = newFieldValue
            ElseIf gridType = "therapist" Then
                originalSessionData.TherapistName = newFieldValue
            End If
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
                            ' განვაახლოთ საჭირო ველები
                            Dim updatedRow As New List(Of Object)(row)

                            ' განვაახლოთ შესაბამისი ველი (სივრცე ან თერაპევტი)
                            If updatedRow.Count > fieldIndex Then
                                updatedRow(fieldIndex) = newFieldValue
                            End If

                            ' განვაახლოთ თარიღი (F სვეტი - ინდექსი 5)
                            If updatedRow.Count > 5 Then
                                updatedRow(5) = newFullDateTime.ToString("dd.MM.yyyy HH:mm")
                            End If

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
                Dim successMessage As String
                If gridType = "space" Then
                    successMessage = $"სესია წარმატებით გადატანილია {newFieldValue} სივრცეში {newDateTime:HH:mm} დროზე"
                Else
                    successMessage = $"სესია წარმატებით გადატანილია თერაპევტზე {newFieldValue} {newDateTime:HH:mm} დროზე"
                End If

                MessageBox.Show(successMessage, "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)

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
    ''' სესიის ბარათზე დაჭერის ივენთი - განახლებული ვერსია
    ''' ახლა აჩვენებს თანხისა და დაფინანსების ინფორმაციასაც
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
                ' გამოვაჩინოთ სესიის დეტალები MessageBox-ში - გაფართოებული ვერსია
                ShowSessionDetails(session)
            End If
        Catch ex As Exception
            Debug.WriteLine($"SessionCard_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' გამოაჩენს სესიის დეტალურ ინფორმაციას MessageBox-ში
    ''' ჩათვლით თანხისა და დაფინანსების მონაცემებისაც
    ''' </summary>
    ''' <param name="session">სესიის მოდელი</param>
    Private Sub ShowSessionDetails(session As SessionModel)
        Try
            Debug.WriteLine($"ShowSessionDetails: ვაჩვენებ დეტალებს სესია ID={session.Id}")

            ' დეტალური ინფორმაციის მომზადება
            Dim sb As New System.Text.StringBuilder()

            ' სათაური
            sb.AppendLine($"🏥 სესიის ინფორმაცია (ID: {session.Id})")
            sb.AppendLine("════════════════════════════════════")
            sb.AppendLine()

            ' ძირითადი ინფორმაცია
            sb.AppendLine("👤 ბენეფიციარი:")
            sb.AppendLine($"   {session.BeneficiaryName} {session.BeneficiarySurname}")
            sb.AppendLine()

            ' თარიღი და დრო - ფორმატირებული ვერსია
            sb.AppendLine("📅 თარიღი და დრო:")
            sb.AppendLine($"   {session.DateTime:dddd, dd MMMM yyyy}")
            sb.AppendLine($"   🕐 {session.DateTime:HH:mm}")
            sb.AppendLine()

            ' ხანგრძლივობა - ფორმატირებული დრო
            sb.AppendLine("⏱️ ხანგრძლივობა:")
            Dim hours As Integer = session.Duration \ 60
            Dim minutes As Integer = session.Duration Mod 60
            If hours > 0 Then
                sb.AppendLine($"   {hours} სთ {minutes} წთ ({session.Duration} წუთი)")
            Else
                sb.AppendLine($"   {session.Duration} წუთი")
            End If
            sb.AppendLine()

            ' თერაპევტი
            sb.AppendLine("👨‍⚕️ თერაპევტი:")
            If Not String.IsNullOrEmpty(session.TherapistName.Trim()) Then
                sb.AppendLine($"   {session.TherapistName}")
            Else
                sb.AppendLine("   [მითითებული არ არის]")
            End If
            sb.AppendLine()

            ' თერაპიის ტიპი
            sb.AppendLine("🔬 თერაპიის ტიპი:")
            If Not String.IsNullOrEmpty(session.TherapyType.Trim()) Then
                sb.AppendLine($"   {session.TherapyType}")
            Else
                sb.AppendLine("   [მითითებული არ არის]")
            End If
            sb.AppendLine()

            ' სივრცე
            sb.AppendLine("🏠 სივრცე:")
            If Not String.IsNullOrEmpty(session.Space.Trim()) Then
                sb.AppendLine($"   {session.Space}")
            Else
                sb.AppendLine("   [მითითებული არ არის]")
            End If
            sb.AppendLine()

            ' ჯგუფური სესია
            sb.AppendLine("👥 ნახვა:")
            sb.AppendLine($"   {If(session.IsGroup, "ჯგუფური სესია", "ინდივიდუალური სესია")}")
            sb.AppendLine()

            ' თანხა - ახალი ანაზღაურების ინფორმაცია
            sb.AppendLine("💰 ანაზღაურება:")
            If session.Price > 0 Then
                ' ფასის ფორმატირება ლარით
                sb.AppendLine($"   {session.Price:F2} ლარი")

                ' ჯგუფური სესიის შემთხვევაში ინფორმაცია
                If session.IsGroup Then
                    sb.AppendLine("   (ჯგუფური სესიის მიხედვით)")
                End If
            Else
                sb.AppendLine("   უფასო სესია")
            End If
            sb.AppendLine()

            ' დაფინანსება - ახალი დაფინანსების ინფორმაცია
            sb.AppendLine("🏦 დაფინანსება:")
            If Not String.IsNullOrEmpty(session.Funding.Trim()) Then
                sb.AppendLine($"   {session.Funding}")

                ' თუ დაფინანსება არის სახელმწიფო, დამატებითი ინფო
                If session.Funding.ToLower().Contains("სახელმწიფო") OrElse
               session.Funding.ToLower().Contains("უშპ") Then
                    sb.AppendLine("   (სახელმწიფო პროგრამა)")
                ElseIf session.Funding.ToLower().Contains("კერძო") Then
                    sb.AppendLine("   (კერძო გადახდა)")
                End If
            Else
                sb.AppendLine("   [მითითებული არ არის]")
            End If
            sb.AppendLine()

            ' სტატუსი - ფერწერის შესაბამისი აღწერა
            sb.AppendLine("📊 მდგომარეობა:")
            sb.AppendLine($"   {session.Status}")

            ' სტატუსის მიხედვით დამატებითი ინფორმაცია
            Select Case session.Status.ToLower().Trim()
                Case "დაგეგმილი"
                    If session.DateTime > DateTime.Now Then
                        sb.AppendLine("   ⏳ უახლოვეს მომავალში")
                    Else
                        sb.AppendLine("   ⚠️ ვადაგადაცილებული")
                    End If
                Case "შესრულებული"
                    sb.AppendLine("   ✅ წარმატებით დასრულებული")
                Case "გაუქმებული", "გაუქმება"
                    sb.AppendLine("   ❌ გაუქმებული")
                Case "გაცდენა საპატიო"
                    sb.AppendLine("   😌 საპატიო მიზეზით გაცდენა")
                Case "გაცდენა არასაპატიო"
                    sb.AppendLine("   😞 გამოუცხადებლად გაცდენა")
            End Select
            sb.AppendLine()

            ' კომენტარები (თუ არის)
            If Not String.IsNullOrEmpty(session.Comments.Trim()) Then
                sb.AppendLine("📝 კომენტარი:")
                sb.AppendLine($"   {session.Comments}")
                sb.AppendLine()
            End If

            ' დამატებითი სტატისტიკა
            sb.AppendLine("📈 დამატებითი ინფო:")

            ' ვადაგადაცილების შემოწმება
            If session.IsOverdue Then
                Dim overdueDays As Integer = CInt((DateTime.Now.Date - session.DateTime.Date).TotalDays)
                If overdueDays = 0 Then
                    sb.AppendLine("   📅 დღევანდელი ვადაგადაცილებული")
                ElseIf overdueDays = 1 Then
                    sb.AppendLine("   📅 გუშინიდან ვადაგადაცილებული")
                Else
                    sb.AppendLine($"   📅 {overdueDays} დღით ვადაგადაცილებული")
                End If
            ElseIf session.DateTime > DateTime.Now Then
                Dim daysUntil As Integer = CInt((session.DateTime.Date - DateTime.Now.Date).TotalDays)
                If daysUntil = 0 Then
                    sb.AppendLine("   📅 დღევანდელი")
                ElseIf daysUntil = 1 Then
                    sb.AppendLine("   📅 ხვალ")
                Else
                    sb.AppendLine($"   📅 {daysUntil} დღეში")
                End If
            End If

            ' გამოვაჩინოთ MessageBox სპეციალური ფორმატირებით
            MessageBox.Show(sb.ToString(),
                       "🏥 სესიის დეტალური ინფორმაცია",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Information)

            Debug.WriteLine($"ShowSessionDetails: სესიის დეტალები გამოჩნდა ID={session.Id}")

        Catch ex As Exception
            Debug.WriteLine($"ShowSessionDetails: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიის დეტალების ჩვენებისას დაფიქსირდა შეცდომა: {ex.Message}",
                       "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' სესიის ბარათზე ორჯერ დაწკაპუნების ივენთი
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
    ''' სქროლის სინქრონიზაცია ბენეფიციარების ხედისთვის
    ''' </summary>
    Private Sub SetupScrollSynchronizationForBeneficiaries()
        Try
            ' ვიპოვოთ მთავარი გრიდის პანელი
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If mainGridPanel Is Nothing Then
                Debug.WriteLine("SetupScrollSynchronizationForBeneficiaries: მთავარი გრიდის პანელი ვერ მოიძებნა")
                Return
            End If

            ' ჯერ მოვხსნათ არსებული ივენთის ჰენდლერები, თუ ისინი არსებობენ
            RemoveHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll
            RemoveHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            ' მივაბათ Scroll ივენთი ხელახლა
            AddHandler mainGridPanel.Scroll, AddressOf MainGridPanel_Scroll

            ' ასევე, დავამატოთ ივენთი გორგოლჭის სქროლისთვის
            AddHandler mainGridPanel.MouseWheel, AddressOf MainGridPanel_MouseWheel

            Debug.WriteLine("SetupScrollSynchronizationForBeneficiaries: სქროლის სინქრონიზაცია დაყენებულია")
        Catch ex As Exception
            Debug.WriteLine($"SetupScrollSynchronizationForBeneficiaries: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების გრიდის ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeBeneficiaryDayViewPanels()
        Try
            Debug.WriteLine("InitializeBeneficiaryDayViewPanels: დაიწყო ბენეფიციარების დღის ხედის პანელების ინიციალიზაცია")

            ' გავასუფთავოთ კალენდრის პანელი
            pnlCalendarGrid.Controls.Clear()

            ' პანელების ინიციალიზაცია როგორც სივრცეების შემთხვევაში
            ' თუმცა ამჯერად არ გამოვიყენებთ გრიდის მენეჯერს, რადგან ბენეფიციარების ხედი სპეციფიკურია

            ' კონსტანტები, რომლებიც არ იცვლება მასშტაბირებისას
            Dim HEADER_HEIGHT As Integer = CalendarGridManager.BASE_HEADER_HEIGHT
            Dim TIME_COLUMN_WIDTH As Integer = CalendarGridManager.BASE_TIME_COLUMN_WIDTH
            Dim DATE_COLUMN_WIDTH As Integer = CalendarGridManager.BASE_DATE_COLUMN_WIDTH

            ' გამოვთვალოთ დროისა და თარიღის პანელების ზომები
            Dim ROW_HEIGHT As Integer = CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale)
            Dim totalRowsHeight As Integer = ROW_HEIGHT * timeIntervals.Count

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

            ' ======= 5. შევქმნათ ბენეფიციარების სათაურების პანელი =======
            Dim beneficiaryHeaderPanel As New Panel()
            beneficiaryHeaderPanel.Name = "beneficiaryHeaderPanel"
            beneficiaryHeaderPanel.AutoScroll = False
            beneficiaryHeaderPanel.Size = New Size(CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale), HEADER_HEIGHT)
            beneficiaryHeaderPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, 0)
            beneficiaryHeaderPanel.BackColor = Color.FromArgb(255, 230, 180) ' ღია ყაშაბისფერი ბენეფიციარების მიზანისთვის
            beneficiaryHeaderPanel.BorderStyle = BorderStyle.FixedSingle
            pnlCalendarGrid.Controls.Add(beneficiaryHeaderPanel)

            ' ======= 6. შევქმნათ მთავარი გრიდის პანელი =======
            Dim mainGridPanel As New Panel()
            mainGridPanel.Name = "mainGridPanel"
            mainGridPanel.Size = New Size(pnlCalendarGrid.ClientSize.Width - TIME_COLUMN_WIDTH - DATE_COLUMN_WIDTH, pnlCalendarGrid.ClientSize.Height - HEADER_HEIGHT)
            mainGridPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, HEADER_HEIGHT)
            mainGridPanel.BackColor = Color.White
            mainGridPanel.BorderStyle = BorderStyle.FixedSingle
            mainGridPanel.AutoScroll = True
            pnlCalendarGrid.Controls.Add(mainGridPanel)

            ' თარიღის ზოლის შევსება
            FillDateColumnPanel(dateColumnPanel)

            ' დროის პანელის შევსება
            FillTimeColumnPanel(timeColumnPanel)

            ' ბენეფიციარების კომბობოქსის შექმნა
            CreateBeneficiaryColumnWithComboBox(beneficiaryHeaderPanel, 0)

            Debug.WriteLine("InitializeBeneficiaryDayViewPanels: პანელების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"InitializeBeneficiaryDayViewPanels: შეცდომა - {ex.Message}")
            Debug.WriteLine($"InitializeBeneficiaryDayViewPanels: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"ბენეფიციარების პანელების ინიციალიზაციის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის კომბობოქსის შექმნა მითითებულ სვეტში
    ''' </summary>
    Private Sub CreateBeneficiaryColumnWithComboBox(headerPanel As Panel, columnIndex As Integer)
        Try
            Debug.WriteLine($"CreateBeneficiaryColumnWithComboBox: ვქმნი კომბობოქსს სვეტისთვის {columnIndex}")

            Dim COLUMN_WIDTH As Integer = CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale)
            Dim HEADER_HEIGHT As Integer = CalendarGridManager.BASE_HEADER_HEIGHT

            ' კომბობოქსის ზომა და პოზიცია
            Dim comboBox As New ComboBox()
            comboBox.Name = $"comboBene_{columnIndex}"
            comboBox.Size = New Size(COLUMN_WIDTH - 30, HEADER_HEIGHT - 6) ' ღილაკისთვის ადგილის დატოვება
            comboBox.Location = New Point(columnIndex * COLUMN_WIDTH + 5, 3)
            comboBox.Font = New Font("Sylfaen", 8, FontStyle.Regular)
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList
            comboBox.FlatStyle = FlatStyle.Flat
            comboBox.BackColor = Color.White

            ' კომბობოქსის სია - ყველა ბენეფიციარი
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
    ''' ბენეფიციარების კომბობოქსის შევსება
    ''' </summary>
    Private Sub FillBeneficiaryComboBox(comboBox As ComboBox)
        Try
            Debug.WriteLine("FillBeneficiaryComboBox: კომბობოქსის შევსება")

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

            Debug.WriteLine($"FillBeneficiaryComboBox: დღეს {todaysBeneficiaries.Count} ბენეფიციარი, უკვე არჩეულია {selectedBeneficiaries.Count}")

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

            Debug.WriteLine($"FillBeneficiaryComboBox: კომბობოქსი შევსებულია {comboBox.Items.Count} ელემენტით")

        Catch ex As Exception
            Debug.WriteLine($"FillBeneficiaryComboBox: შეცდომა - {ex.Message}")
        End Try
    End Sub

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

            ' განვაახლოთ ბენეფიციარების გრიდი
            UpdateBeneficiaryGrid()

            ' განვაახლოთ ყველა კომბობოქსი
            RefreshAllBeneficiaryComboBoxesImproved()

        Catch ex As Exception
            Debug.WriteLine($"BeneficiaryComboBox_SelectedIndexChanged: შეცდომა - {ex.Message}")
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
            Dim COLUMN_WIDTH As Integer = CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale)
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

            ' ვარეორგანიზებთ სვეტებს და ვანახლებთ მათ პოზიციებს
            ReorganizeBeneficiaryColumnsImproved()

            ' განვაახლოთ ბენეფიციარების გრიდი
            UpdateBeneficiaryGrid()

            ' განვაახლოთ ყველა კომბობოქსი ახალი ლოგიკით
            RefreshAllBeneficiaryComboBoxesImproved()

            Debug.WriteLine($"BeneficiaryDeleteButton_Click: სვეტი {columnIndex} წარმატებით წაიშალა")

        Catch ex As Exception
            Debug.WriteLine($"BeneficiaryDeleteButton_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების სვეტების რეორგანიზაცია
    ''' </summary>
    Private Sub ReorganizeBeneficiaryColumnsImproved()
        Try
            Debug.WriteLine("ReorganizeBeneficiaryColumnsImproved: ვირეორგანიზებ სვეტებს")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Return
            End If

            Dim COLUMN_WIDTH As Integer = CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale)

            ' ვაგროვებთ ყველა ლეიბლს ორდენირებულად
            Dim allLabels = beneficiaryHeaderPanel.Controls.OfType(Of Label)() _
                          .Where(Function(l) l.Name.StartsWith("lblBene_")) _
                          .OrderBy(Function(l) l.Location.X) _
                          .ToList()

            ' ყველა სვეტს ვანაწილებთ ახალ ინდექსებს
            For i As Integer = 0 To allLabels.Count - 1
                Dim label = allLabels(i)

                ' განვაახლოთ ლეიბლის სახელი, პოზიცია
                label.Name = $"lblBene_{i}"
                label.Location = New Point(i * COLUMN_WIDTH + 5, label.Location.Y)

                ' ვიპოვოთ და განვაახლოთ შესაბამისი წაშლის ღილაკი
                Dim correspondingButton = beneficiaryHeaderPanel.Controls.OfType(Of Button)() _
                                        .FirstOrDefault(Function(b) b.Name.Contains($"btnDelBene_") AndAlso
                                                      Math.Abs(b.Location.X - (label.Location.X + label.Width)) < 30)

                If correspondingButton IsNot Nothing Then
                    correspondingButton.Name = $"btnDelBene_{i}"
                    correspondingButton.Location = New Point(i * COLUMN_WIDTH + COLUMN_WIDTH - 24, correspondingButton.Location.Y)
                    correspondingButton.Tag = i
                End If
            Next

            ' კომბობოქსების განახლება
            Dim allComboBoxes = beneficiaryHeaderPanel.Controls.OfType(Of ComboBox)() _
                              .Where(Function(c) c.Name.StartsWith("comboBene_")) _
                              .OrderBy(Function(c) c.Location.X) _
                              .ToList()

            For i As Integer = 0 To allComboBoxes.Count - 1
                Dim combo = allComboBoxes(i)
                Dim newIndex = allLabels.Count + i

                combo.Name = $"comboBene_{newIndex}"
                combo.Location = New Point(newIndex * COLUMN_WIDTH + 5, combo.Location.Y)
                combo.Tag = newIndex
            Next

            ' პანელის სიგანის განახლება
            Dim totalColumns As Integer = allLabels.Count + allComboBoxes.Count
            beneficiaryHeaderPanel.Width = totalColumns * COLUMN_WIDTH

            Debug.WriteLine($"ReorganizeBeneficiaryColumnsImproved: რეორგანიზაცია დასრულდა, სულ სვეტი: {totalColumns}")

        Catch ex As Exception
            Debug.WriteLine($"ReorganizeBeneficiaryColumnsImproved: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მივიღოთ არჩეული ბენეფიციარების სია სათაურების პანელიდან
    ''' </summary>
    Private Function GetBeneficiaryColumns() As List(Of String)
        Try
            Debug.WriteLine("GetBeneficiaryColumns: ვიღებ არჩეულ ბენეფიციარებს")

            Dim beneficiaries As New List(Of String)()

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel IsNot Nothing Then
                ' გადავირბინოთ ყველა ლეიბლი და ამოვკრიბოთ ბენეფიციარები
                For Each ctrl As Control In beneficiaryHeaderPanel.Controls.OfType(Of Label)().OrderBy(Function(l) l.Location.X).ToList()
                    If ctrl.Name.StartsWith("lblBene_") Then
                        If ctrl.Tag IsNot Nothing Then
                            beneficiaries.Add(ctrl.Tag.ToString())
                            Debug.WriteLine($"GetBeneficiaryColumns: დაემატა ბენეფიციარი: '{ctrl.Tag}'")
                        End If
                    End If
                Next
            End If

            Debug.WriteLine($"GetBeneficiaryColumns: საბოლოოდ მოიძებნა {beneficiaries.Count} ბენეფიციარი")
            Return beneficiaries

        Catch ex As Exception
            Debug.WriteLine($"GetBeneficiaryColumns: შეცდომა - {ex.Message}")
            Return New List(Of String)()
        End Try
    End Function

    ''' <summary>
    ''' ყველა ბენეფიციარის კომბობოქსის განახლება
    ''' </summary>
    Private Sub RefreshAllBeneficiaryComboBoxesImproved()
        Try
            Debug.WriteLine("RefreshAllBeneficiaryComboBoxesImproved: ვანახლებ ყველა კომბობოქსს")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Then
                Return
            End If

            ' განვაახლოთ ყველა კომბობოქსი ახალი ლოგიკით
            For Each ctrl As Control In beneficiaryHeaderPanel.Controls.OfType(Of ComboBox)().ToList()
                If ctrl.Name.StartsWith("comboBene_") Then
                    ' შევინახოთ მიმდინარე არჩეული ინდექსი
                    Dim currentSelection As Integer = DirectCast(ctrl, ComboBox).SelectedIndex

                    ' განვაახლოთ კომბობოქსის შინაარსი
                    FillBeneficiaryComboBox(DirectCast(ctrl, ComboBox))

                    ' თუ შესაძლებელია, დავაბრუნოთ არჩეული ინდექსი
                    If currentSelection >= 0 AndAlso currentSelection < DirectCast(ctrl, ComboBox).Items.Count Then
                        DirectCast(ctrl, ComboBox).SelectedIndex = currentSelection
                    End If
                End If
            Next

            Debug.WriteLine("RefreshAllBeneficiaryComboBoxesImproved: ყველა კომბობოქსი განახლდა")

        Catch ex As Exception
            Debug.WriteLine($"RefreshAllBeneficiaryComboBoxesImproved: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის გრიდის განახლება
    ''' </summary>
    Private Sub UpdateBeneficiaryGrid()
        Try
            Debug.WriteLine("UpdateBeneficiaryGrid: დაიწყო ბენეფიციარების გრიდის განახლება")

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim beneficiaryHeaderPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)
            Dim mainGridPanel As Panel = DirectCast(pnlCalendarGrid.Controls.Find("mainGridPanel", False).FirstOrDefault(), Panel)

            If beneficiaryHeaderPanel Is Nothing Or mainGridPanel Is Nothing Then
                Debug.WriteLine("UpdateBeneficiaryGrid: ბენეფიციარების სათაურების პანელი ან მთავარი გრიდი ვერ მოიძებნა")
                Return
            End If

            ' მივიღოთ არჩეული ბენეფიციარები
            Dim selectedBeneficiaries = GetBeneficiaryColumns()

            Debug.WriteLine($"UpdateBeneficiaryGrid: მოიძებნა {selectedBeneficiaries.Count} არჩეული ბენეფიციარი")

            ' გავასუფთავოთ და ხელახლა შევავსოთ მთავარი გრიდი
            mainGridPanel.Controls.Clear()

            ' მასშტაბირებული პარამეტრები
            Dim ROW_HEIGHT As Integer = CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale)
            Dim BENEFICIARY_COLUMN_WIDTH As Integer = CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale)

            ' უჯრედების მასივი
            ReDim gridCells(selectedBeneficiaries.Count - 1, timeIntervals.Count - 1)

            ' გრიდის სრული სიგანე და სიმაღლე
            Dim gridWidth As Integer = BENEFICIARY_COLUMN_WIDTH * Math.Max(1, selectedBeneficiaries.Count)
            Dim gridHeight As Integer = ROW_HEIGHT * timeIntervals.Count

            ' AutoScrollMinSize დაყენება
            mainGridPanel.AutoScrollMinSize = New Size(gridWidth, gridHeight)

            ' უჯრედების შექმნა არჩეული ბენეფიციარებისთვის
            For col As Integer = 0 To Math.Max(0, selectedBeneficiaries.Count - 1)
                For row As Integer = 0 To timeIntervals.Count - 1
                    Try
                        Dim cell As New Panel()
                        cell.Size = New Size(BENEFICIARY_COLUMN_WIDTH, ROW_HEIGHT)
                        cell.Location = New Point(col * BENEFICIARY_COLUMN_WIDTH, row * ROW_HEIGHT)

                        ' ალტერნატიული ფერები
                        If row Mod 2 = 0 Then
                            cell.BackColor = Color.FromArgb(255, 255, 250) ' ღია კრემისფერი
                        Else
                            cell.BackColor = Color.FromArgb(255, 250, 245) ' უფრო კრემისფერი
                        End If

                        ' უჯრედის ჩარჩო
                        AddHandler cell.Paint, AddressOf Cell_Paint

                        ' ვინახავთ უჯრედის ობიექტს მასივში
                        gridCells(col, row) = cell

                        ' ვამატებთ უჯრედს გრიდზე
                        mainGridPanel.Controls.Add(cell)
                    Catch cellEx As Exception
                        Debug.WriteLine($"UpdateBeneficiaryGrid: უჯრედის შექმნის შეცდომა [{col},{row}]: {cellEx.Message}")
                    End Try
                Next
            Next

            ' ხელახლა ვაყენებთ სქროლის სინქრონიზაციას
            SetupScrollSynchronizationForBeneficiaries()

            ' ვათავსებთ სესიებს
            If selectedBeneficiaries.Count > 0 Then
                PlaceSessionsOnBeneficiaryGrid()
            End If

            Debug.WriteLine("UpdateBeneficiaryGrid: ბენეფიციარების გრიდი წარმატებით განახლდა")

        Catch ex As Exception
            Debug.WriteLine($"UpdateBeneficiaryGrid: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თარიღის ზოლის პანელის შევსება
    ''' </summary>
    Private Sub FillDateColumnPanel(datePanel As Panel)
        Try
            Debug.WriteLine("FillDateColumnPanel: დაიწყო თარიღის ზოლის შევსება")

            ' გავასუფთავოთ
            datePanel.Controls.Clear()

            ' არჩეული თარიღი კალენდრიდან
            Dim selectedDate As Date = DTPCalendar.Value.Date

            ' ქართული კულტურა თარიღის ფორმატირებისთვის
            Dim georgianCulture As New Globalization.CultureInfo("ka-GE")

            ' კვირის დღე, რიცხვი, თვე და წელი
            Dim weekDay As String = selectedDate.ToString("dddd", georgianCulture)
            Dim dayOfMonth As String = selectedDate.ToString("dd", georgianCulture)
            Dim month As String = selectedDate.ToString("MMMM", georgianCulture)
            Dim year As String = selectedDate.ToString("yyyy", georgianCulture)

            ' შემობრუნებული ლეიბლის შექმნა
            Dim rotatedDateLabel As New RotatedLabel()
            rotatedDateLabel.Size = New Size(datePanel.Width - 4, datePanel.Height - 4)
            rotatedDateLabel.Location = New Point(2, 2)
            rotatedDateLabel.BackColor = Color.FromArgb(60, 80, 150)
            rotatedDateLabel.ForeColor = Color.White
            rotatedDateLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)

            ' ვერტიკალური ტექსტის ფორმირება
            Dim dateText As New System.Text.StringBuilder()
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
            datePanel.Controls.Add(rotatedDateLabel)

            Debug.WriteLine($"FillDateColumnPanel: თარიღის ზოლი შევსებულია, პანელის ზომა: {datePanel.Size}")

        Catch ex As Exception
            Debug.WriteLine($"FillDateColumnPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' დროის პანელის შევსება ინტერვალებით
    ''' </summary>
    Private Sub FillTimeColumnPanel(timePanel As Panel)
        Try
            Debug.WriteLine("FillTimeColumnPanel: დაიწყო დროის პანელის შევსება")

            ' გავასუფთავოთ
            timePanel.Controls.Clear()

            ' მწკრივის სიმაღლე (იმასშტაბირდება)
            Dim ROW_HEIGHT As Integer = CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale)

            ' განვაახლოთ დროის პანელის სიმაღლე
            timePanel.Height = ROW_HEIGHT * timeIntervals.Count

            Debug.WriteLine($"FillTimeColumnPanel: დროის პანელის სიმაღლე: {timePanel.Height}, ROW_HEIGHT: {ROW_HEIGHT}, ინტერვალები: {timeIntervals.Count}")

            ' დროის ეტიკეტები
            For i As Integer = 0 To timeIntervals.Count - 1
                Dim timeLabel As New Label()
                timeLabel.Size = New Size(timePanel.Width - 5, ROW_HEIGHT)
                timeLabel.Location = New Point(0, i * ROW_HEIGHT)
                timeLabel.Text = timeIntervals(i).ToString("HH:mm")
                timeLabel.TextAlign = ContentAlignment.MiddleRight
                timeLabel.Padding = New Padding(0, 0, 5, 0)
                timeLabel.Font = New Font("Segoe UI", 8, FontStyle.Regular)

                ' შევინახოთ Y პოზიცია Tag-ში
                timeLabel.Tag = i * ROW_HEIGHT

                ' ალტერნატიული ფერები მწკრივებისთვის
                If i Mod 2 = 0 Then
                    timeLabel.BackColor = Color.FromArgb(245, 245, 250)
                Else
                    timeLabel.BackColor = Color.FromArgb(250, 250, 255)
                End If

                ' ჩარჩოს დამატება
                timeLabel.BorderStyle = BorderStyle.FixedSingle

                timePanel.Controls.Add(timeLabel)
            Next

            Debug.WriteLine($"FillTimeColumnPanel: დასრულდა - დაემატა {timeIntervals.Count} დროის ეტიკეტი")

        Catch ex As Exception
            Debug.WriteLine($"FillTimeColumnPanel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' დროებითი მეთოდი - კვირის ხედის ჩვენება (ჯერ არ არის იმპლემენტირებული)
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
    ''' დროებითი მეთოდი - თვის ხედის ჩვენება (ჯერ არ არის იმპლემენტირებული)
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
    ''' ჰორიზონტალური მასშტაბის გაზრდა
    ''' </summary>
    Private Sub BtnHUp_Click(sender As Object, e As EventArgs) Handles BtnHUp.Click
        Try
            ' მიმდინარე მდგომარეობის შენახვა
            Dim wasInBeneficiaryView = rbDay.Checked AndAlso RBBene.Checked
            Dim preservedBeneficiaries As List(Of String) = Nothing

            If wasInBeneficiaryView Then
                preservedBeneficiaries = GetBeneficiaryColumns()
            End If

            Dim oldHScale As Double = hScale
            hScale += SCALE_STEP
            If hScale > MAX_SCALE Then hScale = MAX_SCALE

            Debug.WriteLine($"BtnHUp_Click: ჰორიზონტალური მასშტაბი შეიცვალა {oldHScale:F2} -> {hScale:F2}")

            ' თუ გრიდის მენეჯერი არსებობს, განვაახლოთ მასშტაბი
            If gridManager IsNot Nothing Then
                gridManager.UpdateScaleFactors(hScale, vScale)
            End If

            ' მასშტაბის შეცვლის შემდეგ განვაახლოთ ხედი
            UpdateCalendarView()

            ' თუ ბენეფიციარების ხედში ვიყავით, ვცადოთ აღდგენა
            If wasInBeneficiaryView AndAlso preservedBeneficiaries IsNot Nothing Then
                RestoreBeneficiarySelections(preservedBeneficiaries)
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnHUp_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ჰორიზონტალური მასშტაბის შემცირება
    ''' </summary>
    Private Sub BtnHDown_Click(sender As Object, e As EventArgs) Handles BtnHDown.Click
        Try
            ' მიმდინარე მდგომარეობის შენახვა
            Dim wasInBeneficiaryView = rbDay.Checked AndAlso RBBene.Checked
            Dim preservedBeneficiaries As List(Of String) = Nothing

            If wasInBeneficiaryView Then
                preservedBeneficiaries = GetBeneficiaryColumns()
            End If

            Dim oldHScale As Double = hScale
            hScale -= SCALE_STEP
            If hScale < MIN_SCALE Then hScale = MIN_SCALE

            Debug.WriteLine($"BtnHDown_Click: ჰორიზონტალური მასშტაბი შეიცვალა {oldHScale:F2} -> {hScale:F2}")

            ' თუ გრიდის მენეჯერი არსებობს, განვაახლოთ მასშტაბი
            If gridManager IsNot Nothing Then
                gridManager.UpdateScaleFactors(hScale, vScale)
            End If

            ' მასშტაბის შეცვლის შემდეგ განვაახლოთ ხედი
            UpdateCalendarView()

            ' თუ ბენეფიციარების ხედში ვიყავით, ვცადოთ აღდგენა
            If wasInBeneficiaryView AndAlso preservedBeneficiaries IsNot Nothing Then
                RestoreBeneficiarySelections(preservedBeneficiaries)
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnHDown_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ვერტიკალური მასშტაბის გაზრდა
    ''' </summary>
    Private Sub BtnVUp_Click(sender As Object, e As EventArgs) Handles BtnVUp.Click
        Try
            ' მიმდინარე მდგომარეობის შენახვა
            Dim wasInBeneficiaryView = rbDay.Checked AndAlso RBBene.Checked
            Dim preservedBeneficiaries As List(Of String) = Nothing

            If wasInBeneficiaryView Then
                preservedBeneficiaries = GetBeneficiaryColumns()
            End If

            Dim oldVScale As Double = vScale
            vScale += SCALE_STEP
            If vScale > MAX_SCALE Then vScale = MAX_SCALE

            Debug.WriteLine($"BtnVUp_Click: ვერტიკალური მასშტაბი შეიცვალა {oldVScale:F2} -> {vScale:F2}")

            ' თუ გრიდის მენეჯერი არსებობს, განვაახლოთ მასშტაბი
            If gridManager IsNot Nothing Then
                gridManager.UpdateScaleFactors(hScale, vScale)
            End If

            ' მასშტაბის შეცვლის შემდეგ განვაახლოთ ხედი
            UpdateCalendarView()

            ' თუ ბენეფიციარების ხედში ვიყავით, ვცადოთ აღდგენა
            If wasInBeneficiaryView AndAlso preservedBeneficiaries IsNot Nothing Then
                RestoreBeneficiarySelections(preservedBeneficiaries)
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnVUp_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ვერტიკალური მასშტაბის შემცირება
    ''' </summary>
    Private Sub BtnVDown_Click(sender As Object, e As EventArgs) Handles BtnVDown.Click
        Try
            ' მიმდინარე მდგომარეობის შენახვა
            Dim wasInBeneficiaryView = rbDay.Checked AndAlso RBBene.Checked
            Dim preservedBeneficiaries As List(Of String) = Nothing

            If wasInBeneficiaryView Then
                preservedBeneficiaries = GetBeneficiaryColumns()
            End If

            Dim oldVScale As Double = vScale
            vScale -= SCALE_STEP
            If vScale < MIN_SCALE Then vScale = MIN_SCALE

            Debug.WriteLine($"BtnVDown_Click: ვერტიკალური მასშტაბი შეიცვალა {oldVScale:F2} -> {vScale:F2}")

            ' თუ გრიდის მენეჯერი არსებობს, განვაახლოთ მასშტაბი
            If gridManager IsNot Nothing Then
                gridManager.UpdateScaleFactors(hScale, vScale)
            End If

            ' მასშტაბის შეცვლის შემდეგ განვაახლოთ ხედი
            UpdateCalendarView()

            ' თუ ბენეფიციარების ხედში ვიყავით, ვცადოთ აღდგენა
            If wasInBeneficiaryView AndAlso preservedBeneficiaries IsNot Nothing Then
                RestoreBeneficiarySelections(preservedBeneficiaries)
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnVDown_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების აღდგენა მასშტაბირების შემდეგ
    ''' </summary>
    Private Sub RestoreBeneficiarySelections(beneficiaryNames As List(Of String))
        Try
            If beneficiaryNames Is Nothing OrElse beneficiaryNames.Count = 0 Then
                Return
            End If

            ' ვიპოვოთ ბენეფიციარების სათაურების პანელი
            Dim headerPanel = DirectCast(pnlCalendarGrid.Controls.Find("beneficiaryHeaderPanel", False).FirstOrDefault(), Panel)
            If headerPanel Is Nothing Then Return

            ' ვიპოვოთ კომბობოქსი
            Dim comboBox As ComboBox = Nothing
            For Each ctrl In headerPanel.Controls
                If TypeOf ctrl Is ComboBox AndAlso ctrl.Name.StartsWith("comboBene_") Then
                    comboBox = DirectCast(ctrl, ComboBox)
                    Exit For
                End If
            Next

            If comboBox Is Nothing Then Return

            ' დავამატოთ თითოეული ბენეფიციარი
            For Each beneficiaryName In beneficiaryNames
                If comboBox Is Nothing Then
                    Exit For
                End If

                ' ვიპოვნოთ ბენეფიციარი კომბობოქსში
                Dim index = comboBox.Items.IndexOf(beneficiaryName)
                If index > 0 Then  ' 0 ინდექსი არის "-- აირჩიეთ ბენეფიციარი --"
                    ' სიმულირებული არჩევა
                    comboBox.SelectedIndex = index

                    ' დავიცადოთ UI განახლება
                    Application.DoEvents()

                    ' ვიპოვოთ შემდეგი კომბობოქსი
                    Dim nextComboFound = False
                    For Each ctrl In headerPanel.Controls
                        If TypeOf ctrl Is ComboBox AndAlso ctrl.Name.StartsWith("comboBene_") AndAlso Not Object.ReferenceEquals(ctrl, comboBox) Then
                            comboBox = DirectCast(ctrl, ComboBox)
                            nextComboFound = True
                            Exit For
                        End If
                    Next

                    If Not nextComboFound Then
                        comboBox = Nothing
                    End If
                End If
            Next

        Catch ex As Exception
            Debug.WriteLine($"RestoreBeneficiarySelections: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიის სტატუსის სწრაფი ცვლილება კონტექსტური მენიუდან
    ''' </summary>
    Private Sub QuickChangeStatus_Click(sender As Object, e As EventArgs)
        Try
            Dim menuItem = DirectCast(sender, ToolStripMenuItem)
            Dim sessionData = DirectCast(menuItem.Tag, Tuple(Of Integer, String))

            Dim sessionId As Integer = sessionData.Item1
            Dim newStatus As String = sessionData.Item2

            Debug.WriteLine($"QuickChangeStatus_Click: სესიის ID={sessionId} სტატუსის შეცვლა '{newStatus}'")

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

            ' ვიპოვოთ სესია
            Dim session = allSessions.FirstOrDefault(Function(s) s.Id = sessionId)
            If session Is Nothing Then
                MessageBox.Show($"სესია ID={sessionId} ვერ მოიძებნა",
                               "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' დადასტურების მოთხოვნა
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

                            Debug.WriteLine($"QuickChangeStatus_Click: წარმატებით განახლდა სესია ID={session.Id}")

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
            Debug.WriteLine($"QuickChangeStatus_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"სტატუსის შეცვლისას მოხდა შეცდომა: {ex.Message}",
                           "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ახალი სესიის დამატების ღილაკზე დაჭერის დამუშავება
    ''' </summary>
    Private Sub BtnAddSchedule_Click(sender As Object, e As EventArgs) Handles BtnAddSchedule.Click
        Debug.WriteLine("BtnAddSchedule_Click: ახალი ჩანაწერის დამატების მოთხოვნა")

        Try
            ' შევამოწმოთ უკვე გახსნილია თუ არა NewRecordForm
            For Each frm As Form In Application.OpenForms
                If TypeOf frm Is NewRecordForm Then
                    ' თუ უკვე გახსნილია, მოვიტანოთ წინ და გამოვიდეთ მეთოდიდან
                    Debug.WriteLine("BtnAddSchedule_Click: NewRecordForm უკვე გახსნილია, ფოკუსის გადატანა")
                    frm.Focus()
                    Return
                End If
            Next

            ' შევამოწმოთ გვაქვს თუ არა dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Debug.WriteLine("BtnAddSchedule_Click: dataService არ არის ინიციალიზებული")
                Return
            End If

            ' ნაგულისხმევად "სესია" ტიპი
            Dim recordType As String = "სესია"

            ' NewRecordForm-ის გახსნა Add რეჟიმში
            Dim newRecordForm As New NewRecordForm(dataService, recordType, userEmail, "UC_Calendar")
            Dim result = newRecordForm.ShowDialog()

            ' თუ ფორმა დაიხურა OK რეზულტატით, განვაახლოთ მონაცემები
            If result = DialogResult.OK Then
                Debug.WriteLine("BtnAddSchedule_Click: სესია წარმატებით დაემატა")

                ' ჩავტვირთოთ სესიები თავიდან
                LoadSessions()

                ' განვაახლოთ კალენდრის ხედი
                UpdateCalendarView()
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnAddSchedule_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ახალი ჩანაწერის ფორმის გახსნის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სესიების განახლების (refresh) ღილაკზე დაჭერის დამუშავება
    ''' </summary>
    Private Sub BtnRef_Click(sender As Object, e As EventArgs) Handles BtnRef.Click
        Try
            Debug.WriteLine("BtnRef_Click: იწყება მონაცემების განახლება")

            ' გამოვაჩინოთ მაუსის მაჩვენებელი როგორც WaitCursor
            Cursor = Cursors.WaitCursor

            ' გავასუფთავოთ სესიების ქეში
            allSessions = Nothing

            ' ჩავტვირთოთ სესიები თავიდან
            LoadSessions()

            ' განვაახლოთ კალენდარი
            UpdateCalendarView()

            ' დავაბრუნოთ მაუსის მაჩვენებელი
            Cursor = Cursors.Default

            Debug.WriteLine("BtnRef_Click: მონაცემები წარმატებით განახლდა")

        Catch ex As Exception
            Debug.WriteLine($"BtnRef_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"მონაცემების განახლებისას მოხდა შეცდომა: {ex.Message}",
                       "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)

            ' დავაბრუნოთ მაუსის მაჩვენებელი შეცდომის შემთხვევაშიც
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' კალენდრის თარიღის ცვლილების დამუშავება
    ''' </summary>
    Private Sub DTPCalendar_ValueChanged(sender As Object, e As EventArgs)
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
        ' ვამოწმებთ ორივე დრო არის თუ არა არჩეული
        If cbStart.SelectedIndex < 0 OrElse cbFinish.SelectedIndex < 0 Then
            Return
        End If

        ' განვაახლოთ კალენდრის ხედი
        UpdateCalendarView()
    End Sub

    ''' <summary>
    ''' UserControl-ის ჩატვირთვის ივენთი
    ''' </summary>
    Private Sub UC_Calendar_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            ' მნიშვნელოვანი: უზრუნველვყოფთ, რომ UserControl-ს და pnlCalendarGrid-ს არ ჰქონდეს AutoScroll
            Me.AutoScroll = False
            If pnlCalendarGrid IsNot Nothing Then
                pnlCalendarGrid.AutoScroll = False
            End If

            ' საწყისი პარამეტრების დაყენება
            InitializeTimeIntervals()

            ' კალენდრის განახლება
            UpdateCalendarView()

            Debug.WriteLine("UC_Calendar_Load: კალენდრის კონტროლი წარმატებით ჩაიტვირთა")
        Catch ex As Exception
            Debug.WriteLine($"UC_Calendar_Load: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' UserControl-ის Resize ივენთი
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

            ' თუ კალენდრის ხედი უკვე გაშლილია, განვაახლოთ ხედი
            UpdateCalendarView()

            Debug.WriteLine($"UC_Calendar_Resize: კალენდრის ზომები განახლდა, ახალი ზომები: {Me.Width}x{Me.Height}")
        Catch ex As Exception
            Debug.WriteLine($"UC_Calendar_Resize: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების განთავსება ბენეფიციარების გრიდში
    ''' </summary>
    Private Sub PlaceSessionsOnBeneficiaryGrid()
        Try
            Debug.WriteLine("=== PlaceSessionsOnBeneficiaryGrid: დაიწყო ===")

            ' ბაზისური შემოწმებები
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                Debug.WriteLine("PlaceSessionsOnBeneficiaryGrid: სესიები არ არის")
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

            ' გავასუფთავოთ არსებული სესიის ბარათები მაინპანელიდან
            For Each ctrl In mainGridPanel.Controls.OfType(Of Panel)().ToList()
                If ctrl.Tag IsNot Nothing AndAlso TypeOf ctrl.Tag Is Integer Then
                    mainGridPanel.Controls.Remove(ctrl)
                    ctrl.Dispose()
                End If
            Next

            ' მასშტაბირებული პარამეტრები
            Dim ROW_HEIGHT As Integer = CInt(CalendarGridManager.BASE_ROW_HEIGHT * vScale)
            Dim BENEFICIARY_COLUMN_WIDTH As Integer = CInt(CalendarGridManager.BASE_SPACE_COLUMN_WIDTH * hScale)

            ' ფილტრირებული სესიები
            Dim filteredSessions = GetFilteredSessions()
            Debug.WriteLine($"ნაპოვნია {filteredSessions.Count} ფილტრირებული სესია ამ თარიღისთვის")

            ' სესიების განთავსება
            For Each session In filteredSessions
                Try
                    Debug.WriteLine($"--- ვამუშავებთ სესიას ID={session.Id}, ბენეფიციარი='{session.BeneficiaryName} {session.BeneficiarySurname}' ---")

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

                    ' 2. დროის ინტერვალის განსაზღვრა
                    Dim sessionTime As TimeSpan = session.DateTime.TimeOfDay
                    Dim timeIndex As Integer = CalendarUtils.FindNearestTimeIndex(sessionTime, timeIntervals)

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

                    ' 6. ბარათის შექმნა SessionCardFactory-ის გამოყენებით
                    Dim sessionCard = SessionCardFactory.CreateSessionCard(
                    session,
                    BENEFICIARY_COLUMN_WIDTH - 8,
                    cardHeight,
                    New Point(cardX, cardY),
                    AddressOf EditSession ' დამატებული: რედაქტირების დელეგატი
                )

                    ' 7. კონტექსტური მენიუს შექმნა
                    SessionCardFactory.CreateContextMenu(sessionCard, session, AddressOf QuickChangeStatus_Click)

                    ' 8. მაუსის ივენთების მიბმა
                    AddHandler sessionCard.Click, AddressOf SessionCard_Click
                    AddHandler sessionCard.DoubleClick, AddressOf SessionCard_DoubleClick

                    ' კონტექსტური მენიუს ივენთების დამატება
                    If sessionCard.ContextMenuStrip IsNot Nothing Then
                        ' რედაქტირების მენიუს პუნქტი
                        Dim editMenuItem = sessionCard.ContextMenuStrip.Items.Cast(Of ToolStripItem)().FirstOrDefault(Function(item) TypeOf item Is ToolStripMenuItem AndAlso item.Text = "რედაქტირება")
                        If editMenuItem IsNot Nothing Then
                            AddHandler editMenuItem.Click, Sub(menuSender, menuE)
                                                               Dim menuItem = DirectCast(menuSender, ToolStripMenuItem)
                                                               Dim menuSessionId As Integer = CInt(menuItem.Tag)
                                                               EditSession(menuSessionId)
                                                           End Sub
                        End If

                        ' ინფორმაციის ნახვა მენიუს პუნქტი
                        Dim infoMenuItem = sessionCard.ContextMenuStrip.Items.Cast(Of ToolStripItem)().FirstOrDefault(Function(item) TypeOf item Is ToolStripMenuItem AndAlso item.Text = "ინფორმაციის ნახვა")
                        If infoMenuItem IsNot Nothing Then
                            AddHandler infoMenuItem.Click, Sub(menuSender, menuE)
                                                               Dim menuItem = DirectCast(menuSender, ToolStripMenuItem)
                                                               Dim menuSessionId As Integer = CInt(menuItem.Tag)
                                                               ' ახალი მეთოდის გამოძახება
                                                               Dim clickedSession = allSessions.FirstOrDefault(Function(s) s.Id = menuSessionId)
                                                               If clickedSession IsNot Nothing Then
                                                                   ShowSessionDetails(clickedSession)
                                                               End If
                                                           End Sub
                        End If
                    End If

                    ' ბარათის გრიდზე დამატება
                    mainGridPanel.Controls.Add(sessionCard)
                    sessionCard.BringToFront()

                    Debug.WriteLine($"✅ ბარათი განთავსდა: ID={session.Id}, ბენეფიციარი='{sessionBeneficiaryFullName}', " &
                     $"ზომა={sessionCard.Width}x{sessionCard.Height}px")

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
    ''' სესიის რედაქტირების ღილაკზე დაჭერის დამუშავება
    ''' </summary>
    ''' <param name="sessionId"></param>
    Private Sub EditSession(sessionId As Integer)
        Try
            Debug.WriteLine($"EditSession: სესიის რედაქტირების მოთხოვნა, ID={sessionId}")

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
                Debug.WriteLine($"EditSession: სესია ID={sessionId} წარმატებით განახლდა")

                ' ჩავტვირთოთ სესიები თავიდან
                LoadSessions()

                ' განვაახლოთ კალენდრის ხედი
                UpdateCalendarView()
            End If

        Catch ex As Exception
            Debug.WriteLine($"EditSession: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიის რედაქტირების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
