﻿' ===========================================
' 📄 UserControls/UC_Home.vb
' -------------------------------------------
' მთავარი გვერდი - გამოყენებული HomeViewModel-ით და Data Binding-ით
' ===========================================
Imports System.Windows.Forms
Imports System.Drawing
'Imports Scheduler_v8_8a.Models
Imports System.Globalization
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_Home
    Inherits UserControl

    ' ViewModel რომელზეც ხდება მიბმა
    Private ReadOnly viewModel As HomeViewModel

    ' პაგინაციის ცვლადები
    Private CurrentPage As Integer = 0
    Private TotalPages As Integer = 0
    Private CardsPerPage As Integer = 0

    ' სესიების სია მეხსიერებაში
    Private AllSessions As List(Of SessionModel) = New List(Of SessionModel)
    Private IsAuthorizedUser As Boolean = False
    Private UserRoleValue As String = ""
    ' მონაცემთა სერვისი
    Private dataService As Scheduler_v8_8a.Services.IDataService = Nothing

    Private TodaySessions As List(Of SessionModel) = New List(Of SessionModel)()

    ''' <summary>
    ''' კონსტრუქტორი: იღებს HomeViewModel-ს Data Binding-ისთვის
    ''' </summary>
    ''' <param name="homeVm">HomeViewModel ობიექტი</param>
    Public Sub New(homeVm As HomeViewModel)
        ' UI ელემენტების ინიციალიზაცია
        InitializeComponent()

        ' ViewModel შემოწმება - თუ null-ია, შევქმნათ ახალი
        If homeVm IsNot Nothing Then
            viewModel = homeVm
        Else
            ' homeVm არის null, შევქმნათ ახალი ინსტანცია
            viewModel = New HomeViewModel()
            Debug.WriteLine("UC_Home: გადმოცემული viewModel არის null, შეიქმნა ახალი ინსტანცია")
        End If

        ' UI კომპონენტების ინიციალიზაცია
        GBTools.Visible = False
        GBTools.Enabled = False

        ' Timer-ის დაყენება
        Timer1.Interval = 1000
        AddHandler Timer1.Tick, AddressOf Timer1_Tick
        Timer1.Start()

        ' GroupBox-ების გამჭვირვალე ფონის დაყენება (50%)
        GBGreeting.BackColor = Color.FromArgb(200, Color.White)
        GBNow.BackColor = Color.FromArgb(200, Color.White)
        GBTools.BackColor = Color.FromArgb(200, Color.White)
        GBRedTasks.BackColor = Color.FromArgb(200, Color.White)
        GB_Today.BackColor = Color.FromArgb(200, Color.White)
        GBActiveTasks.BackColor = Color.FromArgb(200, Color.White)
        GBBD.BackColor = Color.FromArgb(200, Color.White)

        ' პაგინაციის ღილაკების მომზადება
        BtnPrev.Enabled = False  ' საწყის მდგომარეობაში უკან ღილაკი გამორთულია
        LPage.Text = ""

        ' მიბმა ViewModel-ზე
        BindToViewModel()

        ' საწყისი მონაცემების ჩატვირთვა
        LoadData()

        ' საწყისი მდგომარეობაში ინსტრუმენტები დამალულია
        SetToolsVisibility(False)

        ' განვაახლოთ დღევანდელი სტატისტიკა
        UpdateTodayStatistics()
        UpdateOverdueStatistics()

        ' Resize ივენთის მიბმა
        AddHandler Me.Resize, AddressOf UC_Home_Resize
    End Sub

    ''' <summary>
    ''' ხილვადობის დაყენება ინსტრუმენტების პანელისთვის (GBTools)
    ''' </summary>
    ''' <param name="visible">უნდა იყოს თუ არა ხილული</param>
    Public Sub SetToolsVisibility(visible As Boolean)
        ' ინსტრუმენტების პანელი მხოლოდ მაშინ ჩანს, როცა მომხმარებელს აქვს შესაბამისი როლი
        GBTools.Visible = visible

        ' დავარეგულიროთ სხვა კონტეინერების ზომები, რათა არ იყოს ცარიელი ადგილი
        If visible Then
            ' ჩვეულებრივი მდგომარეობა, როცა ყველა პანელი ჩანს
            GBGreeting.Width = 253 ' საწყისი სიგანე
        Else
            ' როცა GBTools არ ჩანს, GBGreeting-ი უფრო ფართო უნდა იყოს
            GBGreeting.Width = 253 ' გაფართოებული სიგანე
        End If
    End Sub

    ''' <summary>
    ''' Timer-ის მოვლენის უმარტივესი დამმუშავებელი - მხოლოდ საათის განახლება
    ''' </summary>
    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        Try
            ' უბრალოდ მიმდინარე დროის ჩვენება
            Dim now = DateTime.Now
            LTime.Text = now.ToString("HH:mm:ss")

            ' გამოვაჩინოთ დებაგ შეტყობინება დროის განახლების შესახებ
            'Debug.WriteLine($"Timer1_Tick: დრო განახლდა - {now.ToString("HH:mm:ss")}")
        Catch ex As Exception
            Debug.WriteLine($"Timer1_Tick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' UI ელემენტების მიბმა ViewModel-ზე
    ''' </summary>
    Private Sub BindToViewModel()
        ' დავრწმუნდეთ, რომ viewModel არ არის null
        If viewModel IsNot Nothing Then
            ' მიბმა ViewModel-ის PropertyChanged ივენთზე
            AddHandler viewModel.PropertyChanged, AddressOf ViewModel_PropertyChanged

            ' მისალმების დაყენება
            LUserName.Text = viewModel.UserName
            LWish.Text = viewModel.Greeting

            ' დროის დაყენება
            UpdateTimeDisplay()
        Else
            Debug.WriteLine("BindToViewModel: viewModel არის null, ვერ ხერხდება დაბაინდვა")

            ' დავაყენოთ საწყისი მნიშვნელობები
            LUserName.Text = "მომხმარებელო"
            LWish.Text = "დილა მშვიდობისა"
        End If
    End Sub
    ''' <summary>
    ''' მიუთითებს მონაცემთა სერვისს
    ''' </summary>
    ''' <param name="service">მონაცემთა სერვისი</param>
    Public Sub SetDataService(service As Scheduler_v8_8a.Services.IDataService)
        dataService = service
        Debug.WriteLine("UC_Home.SetDataService: მითითებულია მონაცემთა სერვისი")
    End Sub
    ''' <summary>
    ''' მომხმარებლის სახელის განახლება
    ''' </summary>
    ''' <param name="userName">მომხმარებლის სახელი</param>
    Public Sub UpdateUserName(userName As String)
        Try
            Debug.WriteLine($"UpdateUserName: მომხმარებლის სახელის განახლება '{userName}'")

            ' პირდაპირ ლეიბლში დავაყენოთ სახელი
            LUserName.Text = userName

            ' ასევე, განვაახლოთ ViewModel
            If viewModel IsNot Nothing Then
                viewModel.UserName = userName
                Debug.WriteLine($"UpdateUserName: ViewModel განახლდა, UserName='{viewModel.UserName}'")
            End If

            ' ძალით განვაახლოთ LUserName
            LUserName.Refresh()
            Application.DoEvents()

            Debug.WriteLine($"UpdateUserName: LUserName.Text = '{LUserName.Text}'")
        Catch ex As Exception
            Debug.WriteLine($"UpdateUserName: შეცდომა - {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' დროის ჩვენების განახლება ViewModel-დან
    ''' </summary>
    Private Sub UpdateTimeDisplay()
        ' UI-ს ვანახლებთ მხოლოდ თუ კონტროლი ხილვადია და ინიციალიზებულია
        If Me.IsHandleCreated AndAlso Me.Visible Then
            ' UI ნაკადის შემოწმება
            If Me.InvokeRequired Then
                Me.Invoke(Sub() UpdateTimeDisplay())
            Else
                ' UI-ს განახლება ViewModel-ის მონაცემებით
                LTime.Text = viewModel.FormattedTime
                LDate.Text = viewModel.FormattedDate
                LWeekDay.Text = viewModel.WeekDayName
            End If
        End If
    End Sub
    ''' <summary>
    ''' განაახლებს დღევანდელი სესიების სტატისტიკის ლეიბლებს
    ''' </summary>
    Private Sub UpdateTodayStatistics()
        Try
            ' შევამოწმოთ, არის თუ არა მონაცემები ხელმისაწვდომი
            If AllSessions Is Nothing Then
                Debug.WriteLine("UpdateTodayStatistics: AllSessions არის Nothing")
                Return
            End If

            ' მიმდინარე თარიღი თარიღის კომპონენტით (დრო 00:00)
            Dim today As DateTime = DateTime.Today

            ' გავფილტროთ მხოლოდ დღევანდელი სესიები
            Dim todaySessions As New List(Of SessionModel)
            For Each session As SessionModel In AllSessions
                If session.DateTime.Date = today Then
                    todaySessions.Add(session)
                End If
            Next

            Debug.WriteLine("UpdateTodayStatistics: ნაპოვნია " & todaySessions.Count & " დღევანდელი სესია")

            ' დავითვალოთ სტატისტიკა
            Dim totalSessions As Integer = todaySessions.Count

            ' შესრულებული სესიები (სტატუსით "შესრულებული")
            Dim completedSessions As Integer = 0
            For Each session As SessionModel In todaySessions
                If session.Status.Trim().ToLower() = "შესრულებული" Then
                    completedSessions += 1
                End If
            Next

            ' დაგეგმილი სესიები (სტატუსით "დაგეგმილი")
            Dim plannedSessions As Integer = 0
            For Each session As SessionModel In todaySessions
                If session.Status.Trim().ToLower() = "დაგეგმილი" Then
                    plannedSessions += 1
                End If
            Next

            ' უნიკალური ბენეფიციარების რაოდენობა
            Dim beneficiarySet As New HashSet(Of String)
            For Each session As SessionModel In todaySessions
                Dim fullName As String = session.BeneficiaryName & " " & session.BeneficiarySurname
                beneficiarySet.Add(fullName)
            Next
            Dim uniqueBeneficiaries As Integer = beneficiarySet.Count

            ' უნიკალური პერსონალის რაოდენობა (თერაპევტები)
            Dim therapistSet As New HashSet(Of String)
            For Each session As SessionModel In todaySessions
                If Not String.IsNullOrWhiteSpace(session.TherapistName) Then
                    therapistSet.Add(session.TherapistName)
                End If
            Next
            Dim uniqueTherapists As Integer = therapistSet.Count

            ' განვაახლოთ ლეიბლები - დიზაინის შეუცვლელად
            LSe.Text = totalSessions.ToString()
            LDone.Text = completedSessions.ToString()
            LNDone.Text = plannedSessions.ToString()
            LBenes.Text = uniqueBeneficiaries.ToString()
            LPers.Text = uniqueTherapists.ToString()

            Debug.WriteLine("UpdateTodayStatistics: სტატისტიკა განახლდა - " &
            "სულ: " & totalSessions & ", " &
            "შესრულებული: " & completedSessions & ", " &
            "დაგეგმილი: " & plannedSessions & ", " &
            "ბენეფიციარები: " & uniqueBeneficiaries & ", " &
            "პერსონალი: " & uniqueTherapists)

        Catch ex As Exception
            Debug.WriteLine("UpdateTodayStatistics: შეცდომა - " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' მოახლოებული დაბადების დღეების ჩვენება GBBD გრუპბოქსში
    ''' </summary>
    Public Sub PopulateUpcomingBirthdays(birthdays As List(Of BirthdayModel))
        Try
            Debug.WriteLine($"UC_Home.PopulateUpcomingBirthdays: დაიწყო, მოახლოებული დაბადების დღეების ჩვენებას")

            ' გავასუფთავოთ GBBD გრუპბოქსი
            GBBD.Controls.Clear()

            ' შევამოწმოთ გვაქვს თუ არა მონაცემები
            If birthdays Is Nothing OrElse birthdays.Count = 0 Then
                ' საჩვენებელი ლეიბლი: არ არის დაბადების დღეები
                Dim lblNoData As New Label()
                lblNoData.Text = "მოახლოებული დაბადების დღეები არ არის"
                lblNoData.AutoSize = True
                lblNoData.Location = New Point(10, 20)
                GBBD.Controls.Add(lblNoData)

                Debug.WriteLine("UC_Home.PopulateUpcomingBirthdays: birthdays სია ცარიელია ან NULL")
                Return
            End If

            ' დავალაგოთ დაბადების დღეები დღეების რაოდენობის მიხედვით
            Dim sortedBirthdays = birthdays.OrderBy(Function(b) b.DaysUntilBirthday).ToList()
            Debug.WriteLine($"UC_Home.PopulateUpcomingBirthdays: მიღებულია {sortedBirthdays.Count} დაბადების დღე")

            ' გამოვიტანოთ პირველი 5 დაბადების დღე (ან ნაკლები თუ არ გვყავს 5)
            Dim yPos As Integer = 20
            Dim displayCount As Integer = 0
            Dim maxToShow As Integer = Math.Min(5, sortedBirthdays.Count)

            For i As Integer = 0 To maxToShow - 1
                Dim birthday = sortedBirthdays(i)

                Debug.WriteLine($"UC_Home.PopulateUpcomingBirthdays: დაბადების დღე - ID={birthday.Id}, " &
                       $"სახელი={birthday.PersonName}, გვარი={birthday.PersonSurname}, " &
                       $"თარიღი={birthday.BirthDate:dd.MM.yyyy}, დარჩენილი დღეები={birthday.DaysUntilBirthday}")

                ' ტექსტი: "სახელი გვარი - დარჩა X დღე"
                Dim daysText As String
                Select Case birthday.DaysUntilBirthday
                    Case 0
                        daysText = "დღეს"
                    Case 1
                        daysText = "ხვალ"
                    Case Else
                        daysText = $"დარჩა {birthday.DaysUntilBirthday} დღე"
                End Select

                Dim birthdayText As String = $"{birthday.PersonName} {birthday.PersonSurname} - {daysText}"

                ' ლეიბლის შექმნა
                Dim lblBirthday As New Label()
                lblBirthday.Text = birthdayText
                lblBirthday.AutoSize = True
                lblBirthday.Location = New Point(10, yPos)

                ' თუ დღეს აქვს დაბადების დღე, გამოვყოთ წითლად
                If birthday.DaysUntilBirthday = 0 Then
                    lblBirthday.ForeColor = Color.Red
                    lblBirthday.Font = New Font(lblBirthday.Font, FontStyle.Bold)
                End If

                ' დავამატოთ ლეიბლი
                GBBD.Controls.Add(lblBirthday)

                ' გადავიდეთ შემდეგ პოზიციაზე
                yPos += 20
                displayCount += 1
            Next

            ' თუ 5-ზე მეტი დაბადების დღეა, ვაჩვენოთ "...და კიდევ X"
            If sortedBirthdays.Count > 5 Then
                Dim lblMore As New Label()
                lblMore.Text = $"...და კიდევ {sortedBirthdays.Count - 5}"
                lblMore.AutoSize = True
                lblMore.Location = New Point(10, yPos)
                lblMore.ForeColor = Color.Gray
                GBBD.Controls.Add(lblMore)
            End If

            Debug.WriteLine($"UC_Home.PopulateUpcomingBirthdays: წარმატებით გამოჩნდა {displayCount} დაბადების დღე")

        Catch ex As Exception
            Debug.WriteLine($"UC_Home.PopulateUpcomingBirthdays: შეცდომა - {ex.Message}")
            Debug.WriteLine($"UC_Home.PopulateUpcomingBirthdays: Stack Trace - {ex.StackTrace}")

            ' შეცდომის შემთხვევაში მაინც ვაჩვენოთ შეტყობინება
            Dim lblError As New Label()
            lblError.Text = "შეცდომა დაბადების დღეების ჩვენებისას"
            lblError.AutoSize = True
            lblError.Location = New Point(10, 20)
            lblError.ForeColor = Color.Red
            GBBD.Controls.Add(lblError)
        End Try
    End Sub
    ''' <summary>
    ''' განაახლებს ვადაგადაცილებული სესიების სტატისტიკის ლეიბლებს
    ''' </summary>
    Private Sub UpdateOverdueStatistics()
        Try
            ' შევამოწმოთ, არის თუ არა მონაცემები ხელმისაწვდომი
            If AllSessions Is Nothing Then
                Debug.WriteLine("UpdateOverdueStatistics: AllSessions არის Nothing")
                Return
            End If

            ' მიმდინარე თარიღი
            Dim today As DateTime = DateTime.Today

            ' ვთვლით ყველა ვადაგადაცილებულ სესიას
            Dim allOverdueSessions As Integer = 0

            ' ვთვლით დღევანდელ ვადაგადაცილებულ სესიებს
            Dim todayOverdueSessions As Integer = 0

            For Each session As SessionModel In AllSessions
                ' ვიყენებთ SessionModel-ის IsOverdue თვისებას
                ' რათა გავარკვიოთ არის თუ არა სესია ვადაგადაცილებული
                If session.IsOverdue Then
                    allOverdueSessions += 1

                    ' ვამოწმებთ არის თუ არა დღევანდელი
                    If session.DateTime.Date = today Then
                        todayOverdueSessions += 1
                    End If
                End If
            Next

            ' განვაახლოთ ლეიბლები - დიზაინის შეუცვლელად
            LNeerReaction.Text = allOverdueSessions.ToString()
            LNeedReactionToday.Text = todayOverdueSessions.ToString()

            Debug.WriteLine("UpdateOverdueStatistics: სტატისტიკა განახლდა - " &
                       "ყველა ვადაგადაცილებული: " & allOverdueSessions & ", " &
                       "დღევანდელი ვადაგადაცილებული: " & todayOverdueSessions)

        Catch ex As Exception
            Debug.WriteLine("UpdateOverdueStatistics: შეცდომა - " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' ViewModel-ის PropertyChanged ივენთის დამმუშავებელი
    ''' </summary>
    Private Sub ViewModel_PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs)
        ' ივენთის მართვა UI-ის მთავარ ნაკადში
        If Me.InvokeRequired Then
            Me.Invoke(Sub() ViewModel_PropertyChanged(sender, e))
            Return
        End If
        UpdateTodayStatistics()
        ' შესაბამისი თვისების განახლება UI-ში
        Select Case e.PropertyName
            Case NameOf(viewModel.UserName)
                LUserName.Text = viewModel.UserName
                Debug.WriteLine($"ViewModel_PropertyChanged: UserName განახლდა '{viewModel.UserName}'")
            Case NameOf(viewModel.Greeting)
                LWish.Text = viewModel.Greeting
            Case NameOf(viewModel.FormattedTime)
                LTime.Text = viewModel.FormattedTime
            Case NameOf(viewModel.FormattedDate)
                LDate.Text = viewModel.FormattedDate
            Case NameOf(viewModel.WeekDayName)
                LWeekDay.Text = viewModel.WeekDayName
        End Select
    End Sub

    ''' <summary>
    ''' მონაცემების ჩატვირთვა
    ''' </summary>
    Private Sub LoadData()
        ' ViewModel-ის RefreshData მეთოდის გამოძახება
        ' მონაცემების ჩატვირთვა (დაბადების დღეები, სესიები, დავალებები)
        viewModel.RefreshData()
        ' განვაახლოთ დღევანდელი სტატისტიკა
        UpdateTodayStatistics()
        UpdateOverdueStatistics()
        ' მოახლოებული დაბადების დღეების განახლება

    End Sub

    ''' <summary>
    ''' Refresh ღილაკის Click ივენთის დამმუშავებელი - სრულად განაახლებს ყველა მონაცემს
    ''' </summary>
    Private Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles BtnRefresh.Click
        Debug.WriteLine("BtnRefresh_Click: განახლების მოთხოვნა მიღებულია")

        ' ღილაკის დროებით გამორთვა ორჯერ დაჭერის პრევენციისთვის
        BtnRefresh.Enabled = False
        Cursor = Cursors.WaitCursor

        Try
            ' გაუქმება ყველა ქეშის, თუ შესაძლებელია
            If dataService IsNot Nothing Then
                Try
                    ' SheetDataService-ისთვის
                    If TypeOf dataService Is SheetDataService Then
                        DirectCast(dataService, SheetDataService).InvalidateAllCache()
                        Debug.WriteLine("BtnRefresh_Click: SheetDataService ქეში გასუფთავდა")
                    ElseIf TypeOf dataService Is GoogleSheetsDataService Then
                        ' თუ ამ სერვისს აქვს ქეშის გასუფთავების მეთოდი
                        Debug.WriteLine("BtnRefresh_Click: GoogleSheetsDataService ქეშირება")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"BtnRefresh_Click: შეცდომა ქეშის გაუქმებისას - {ex.Message}")
                End Try
            Else
                Debug.WriteLine("BtnRefresh_Click: dataService არ არის ინიციალიზებული")
            End If

            ' 1. განვაახლოთ ViewModel-ის მონაცემები
            viewModel.RefreshData()

            ' 2. ჩავტვირთოთ ყველა მონაცემი ხელახლა
            RefreshAllData()

            ' 3. ვაცნობოთ მომხმარებელს
            MessageBox.Show("მონაცემები წარმატებით განახლდა", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Debug.WriteLine($"BtnRefresh_Click: განახლების შეცდომა - {ex.Message}")
            MessageBox.Show($"მონაცემების განახლების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ღილაკის და კურსორის აღდგენა
            BtnRefresh.Enabled = True
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' UserControl-ის დატვირთვის ივენთი, რომელიც გაეშვება იგი პირველად ჩნდება
    ''' </summary>
    Private Sub UC_Home_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Timer-ის ხელახლა დაწყება თუ გაჩერებული იყო
        If Not Timer1.Enabled Then
            Timer1.Start()
        End If
    End Sub

    ''' <summary>
    ''' UserControl-ის დამალვის ივენთი, რომელიც გაეშვება როდესაც კონტროლი იმალება
    ''' </summary>
    Private Sub UC_Home_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        ' Timer-ი ვმუშაობს მხოლოდ როცა ხილულია
        Timer1.Enabled = Me.Visible
    End Sub

    ''' <summary>
    ''' UserControl-ის Resize ივენთის დამმუშავებელი
    ''' </summary>
    Private Sub UC_Home_Resize(sender As Object, e As EventArgs)
        ' შევამოწმოთ არის თუ არა ფორმა ინიციალიზებული
        If Not Me.IsHandleCreated OrElse GBRedTasks Is Nothing Then
            Return
        End If

        ' თუ სესიები გვაქვს, განვაახლოთ ბარათების ჩვენება
        If AllSessions.Count > 0 Then
            ' ვიანგარიშოთ თავიდან რამდენი ეტევა
            CalculateCardsPerPage()

            ' თუ მიმდინარე გვერდი აღარ არსებობს (მაგ. ფანჯარა გაიზარდა), დავაბრუნოთ პირველ გვერდზე
            If CurrentPage >= TotalPages Then
                CurrentPage = Math.Max(0, TotalPages - 1)
            End If

            ' ჩავტვირთოთ მიმდინარე გვერდის ბარათები
            ShowCurrentPageCards()

            ' განვაახლოთ პაგინაციის კონტროლები
            UpdatePaginationControls()
        End If
    End Sub

    ''' <summary>
    ''' ანგარიშობს რამდენი ბარათი ეტევა ერთ გვერდზე და საერთო გვერდების რაოდენობას
    ''' </summary>
    Private Sub CalculateCardsPerPage()
        Try
            ' ბარათის ზომები და დაშორებები - ზუსტად იგივე კონსტანტები რაც ShowCurrentPageCards-ში
            Const CARD_WIDTH As Integer = 250
            Const CARD_HEIGHT As Integer = 185
            Const CARD_MARGIN As Integer = 15

            ' გამოვთვალოთ რამდენი ბარათი დაეტევა ერთ მწკრივში
            Dim availableWidth As Integer = GBRedTasks.ClientSize.Width - (2 * CARD_MARGIN)
            Dim cardsPerRow As Integer = Math.Max(1, availableWidth \ (CARD_WIDTH + CARD_MARGIN))

            ' გამოვთვალოთ რამდენი მწკრივი დაეტევა
            ' პაგინაციის ღილაკებისთვის დარჩება ადგილი ბოლოში, მაგრამ არ შეიცვლება მათი მდებარეობა
            Dim paginationHeight As Integer = 50 ' სავარაუდო სიმაღლე პაგინაციის ღილაკებისთვის
            Dim availableHeight As Integer = GBRedTasks.ClientSize.Height - paginationHeight - (2 * CARD_MARGIN)

            ' გამოვითვალოთ რამდენი მწკრივი ჩაეტევა
            Dim rowsPerPage As Integer = Math.Max(1, availableHeight \ (CARD_HEIGHT + CARD_MARGIN))

            ' უსაფრთხოებისთვის, თუ დათვლილი მწკრივების სიმაღლე აჭარბებს ხელმისაწვდომ სიმაღლეს
            If (rowsPerPage * (CARD_HEIGHT + CARD_MARGIN)) > availableHeight Then
                rowsPerPage = Math.Max(1, rowsPerPage - 1)
            End If

            ' გამოვთვალოთ რამდენი ბარათი ეტევა ერთ გვერდზე
            CardsPerPage = cardsPerRow * rowsPerPage

            ' AllSessions უკვე შეიცავს მხოლოდ ვადაგადაცილებულ სესიებს
            ' საერთო გვერდების რაოდენობა
            TotalPages = Math.Max(1, Math.Ceiling(AllSessions.Count / CDbl(CardsPerPage)))

            Debug.WriteLine($"CalculateCardsPerPage: GBRedTasks.Size={GBRedTasks.Size}")
            Debug.WriteLine($"CalculateCardsPerPage: availableWidth={availableWidth}, availableHeight={availableHeight}")
            Debug.WriteLine($"CalculateCardsPerPage: cardsPerRow={cardsPerRow}, rowsPerPage={rowsPerPage}")
            Debug.WriteLine($"CalculateCardsPerPage: CardsPerPage={CardsPerPage}, TotalPages={TotalPages}")
            Debug.WriteLine($"CalculateCardsPerPage: AllSessions.Count={AllSessions.Count}")

        Catch ex As Exception
            Debug.WriteLine($"CalculateCardsPerPage: შეცდომა - {ex.Message}")
            CardsPerPage = 8 ' ნაგულისხმები მნიშვნელობა
            TotalPages = Math.Ceiling(AllSessions.Count / CDbl(CardsPerPage))
        End Try
    End Sub
    ''' <summary>
    ''' აჩვენებს მიმდინარე გვერდის ბარათებს - განახლებული ვერსია 
    ''' ღილაკის ფერის და ფორმის დაბრუნებით და ვადაგადაცილებული დღეების ჩვენებით
    ''' </summary>
    Private Sub ShowCurrentPageCards()
        ' გასუფთავება ძველი ბარათებისგან
        GBRedTasks.Controls.Clear()

        ' თუ სესიები არ არის, დავამატოთ შეტყობინება
        If AllSessions Is Nothing OrElse AllSessions.Count = 0 Then
            Dim lblNoSessions As New Label()
            lblNoSessions.Text = "ვადაგადაცილებული სესიები არ არის"
            lblNoSessions.AutoSize = True
            lblNoSessions.Location = New Point(10, 20)
            GBRedTasks.Controls.Add(lblNoSessions)
            Debug.WriteLine("ShowCurrentPageCards: დაემატა შეტყობინება 'ვადაგადაცილებული სესიები არ არის'")
            Return
        End If

        ' ბარათის ზომები და დაშორებები - დავაბრუნეთ ორიგინალური ზომები
        Const CARD_WIDTH As Integer = 250
        Const CARD_HEIGHT As Integer = 185  ' გაზრდილი სიმაღლე 
        Const CARD_MARGIN As Integer = 15
        Const HEADER_HEIGHT As Integer = 24

        ' გამოვთვალოთ რამდენი ბარათი დაეტევა ერთ მწკრივში
        Dim availableWidth As Integer = GBRedTasks.ClientSize.Width - (2 * CARD_MARGIN) - 5
        Dim cardsPerRow As Integer = Math.Max(1, availableWidth \ (CARD_WIDTH + CARD_MARGIN))

        ' დავაკორექტიროთ ჰორიზონტალური დაშორება თუ კი საჭიროა
        Dim horizontalSpacing As Integer = CARD_MARGIN
        If cardsPerRow > 1 Then
            horizontalSpacing = Math.Max(CARD_MARGIN, (availableWidth - (cardsPerRow * CARD_WIDTH)) \ (cardsPerRow - 1))
        End If

        ' გვერდზე ბარათების დიაპაზონი
        Dim startIndex As Integer = CurrentPage * CardsPerPage
        Dim endIndex As Integer = Math.Min(startIndex + CardsPerPage - 1, AllSessions.Count - 1)

        Debug.WriteLine($"ShowCurrentPageCards: გვერდი {CurrentPage + 1}/{TotalPages}, " &
            $"დიაპაზონი: {startIndex}-{endIndex}, " &
            $"cardsPerRow={cardsPerRow}, horizontalSpacing={horizontalSpacing}")

        ' ვადაგადაცილებული სესიების ბარათები
        Dim xPos As Integer = CARD_MARGIN
        Dim yPos As Integer = CARD_MARGIN + 10 ' დამატებითი 10 პიქსელი ქვემოთ
        Dim cardCount As Integer = 0

        For i As Integer = startIndex To endIndex
                Dim session = AllSessions(i)

                ' დებაგინფო სესიის შესახებ
                'Debug.WriteLine($"ShowCurrentPageCards: სესია #{i}: ID={session.Id}, " &
                '$"ბენეფიციარი={session.BeneficiaryName} {session.BeneficiarySurname}, " &
                '$"თარიღი={session.FormattedDateTime}, " &
                '$"სტატუსი={session.Status}")

                ' შევქმნათ ბარათი
                Dim card As New Panel()
                card.Size = New Size(CARD_WIDTH, CARD_HEIGHT)
                card.Location = New Point(xPos, yPos)
                card.BorderStyle = BorderStyle.None
                card.BackColor = Color.FromArgb(255, 235, 235) ' უფრო ღია წითელი

                ' მოვუმრგვალოთ კუთხეები ბარათს
                Dim path As New Drawing2D.GraphicsPath()
                Dim cornerRadius As Integer = 10
                path.AddArc(0, 0, cornerRadius * 2, cornerRadius * 2, 180, 90)
                path.AddArc(card.Width - cornerRadius * 2, 0, cornerRadius * 2, cornerRadius * 2, 270, 90)
                path.AddArc(card.Width - cornerRadius * 2, card.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 0, 90)
                path.AddArc(0, card.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 90, 90)
                path.CloseFigure()
                card.Region = New Region(path)

                ' ზედა მუქი ზოლი
                Dim headerPanel As New Panel()
                headerPanel.Size = New Size(CARD_WIDTH, HEADER_HEIGHT)
                headerPanel.Location = New Point(0, 0)
                headerPanel.BackColor = Color.FromArgb(180, 0, 0) ' მუქი წითელი

                ' მოვუმრგვალოთ მხოლოდ ზედა კუთხეები header-ს
                Dim headerPath As New Drawing2D.GraphicsPath()
                headerPath.AddArc(0, 0, cornerRadius * 2, cornerRadius * 2, 180, 90)
                headerPath.AddArc(headerPanel.Width - cornerRadius * 2, 0, cornerRadius * 2, cornerRadius * 2, 270, 90)
                headerPath.AddLine(headerPanel.Width, cornerRadius, headerPanel.Width, headerPanel.Height)
                headerPath.AddLine(headerPanel.Width, headerPanel.Height, 0, headerPanel.Height)
                headerPath.AddLine(0, headerPanel.Height, 0, cornerRadius)
                headerPath.CloseFigure()
                headerPanel.Region = New Region(headerPath)
                card.Controls.Add(headerPanel)

                ' თარიღი თეთრი ფონტით header-ში
                Dim lblDateTime As New Label()
                lblDateTime.Text = session.FormattedDateTime
                lblDateTime.Location = New Point(8, 4)
                lblDateTime.AutoSize = True
                lblDateTime.Font = New Font(lblDateTime.Font.FontFamily, 9, FontStyle.Regular)
                lblDateTime.ForeColor = Color.White
                headerPanel.Controls.Add(lblDateTime)

                ' ID მარჯვენა კუთხეში header-ში
                Dim lblId As New Label()
                lblId.Text = $"#{session.Id}"
                lblId.Location = New Point(headerPanel.Width - 40, 4)
                lblId.AutoSize = True
                lblId.Font = New Font(lblId.Font.FontFamily, 9, FontStyle.Regular)
                lblId.ForeColor = Color.White
                headerPanel.Controls.Add(lblId)

                ' ბენეფიციარის სახელი და გვარი - მთავრული ასოებით
                ' მთავრული ქართული შრიფტი
                Dim mtavruliFont As String = "Sylfaen" ' ნაგულისხმევი ფონტი თუ ვერ მოიძებნა მოთხოვნილი

                ' ვცადოთ ვიპოვოთ მოთხოვნილი ფონტი
                Try
                    Using testFont As New Font("KA_LITERATURULI_MT", 10)
                        mtavruliFont = "KA_LITERATURULI_MT"
                        Debug.WriteLine("ShowCurrentPageCards: მოიძებნა KA_LITERATURULI_MT ფონტი")
                    End Using
                Catch
                    ' თუ მოთხოვნილი ფონტი ვერ მოიძებნა, შევამოწმოთ ალტერნატიული ფონტები
                    Dim altFonts As String() = {"BPG_Nino_Mtavruli", "ALK_Tall_Mtavruli", "Sylfaen"}
                    For Each fontName In altFonts
                        Try
                            Using testFont As New Font(fontName, 10)
                                mtavruliFont = fontName
                                Debug.WriteLine($"ShowCurrentPageCards: მოიძებნა ალტერნატიული ფონტი: {fontName}")
                                Exit For
                            End Using
                        Catch
                            ' ეს ფონტი ვერ მოიძებნა, შემდეგს ვცდით
                            Debug.WriteLine($"ShowCurrentPageCards: ვერ მოიძებნა ფონტი: {fontName}")
                        End Try
                    Next
                End Try

                ' ბენეფიციარის სახელის ლეიბლი - მთავრული ასოებით
                Dim lblBeneficiary As New Label()
                ' გარდავქმნათ ტექსტი დიდ ასოებად - ToUpper() გამოიყენება მთავრული ასოებისთვის
                Dim beneficiaryText As String = $"{session.BeneficiaryName.ToUpper()} {session.BeneficiarySurname.ToUpper()}"
                lblBeneficiary.Text = beneficiaryText
                lblBeneficiary.Location = New Point(8, HEADER_HEIGHT + 8)
                lblBeneficiary.Size = New Size(CARD_WIDTH - 16, 30) ' გაზრდილი სიმაღლე
                lblBeneficiary.Font = New Font(mtavruliFont, 10, FontStyle.Bold) ' მთავრული ფონტი
                lblBeneficiary.TextAlign = ContentAlignment.MiddleCenter ' ცენტრში მოთავსება
                card.Controls.Add(lblBeneficiary)

                ' თერაპევტის სახელი
                Dim lblTherapist As New Label()
                lblTherapist.Text = $"{session.TherapistName}"
                lblTherapist.Location = New Point(8, HEADER_HEIGHT + 42) ' ქვემოთ ჩაწეული
                lblTherapist.Size = New Size(CARD_WIDTH - 16, 20) ' ფიქსირებული სიგანე
                card.Controls.Add(lblTherapist)

                ' თერაპიის ტიპი
                Dim lblTherapyType As New Label()
                lblTherapyType.Text = $"{session.TherapyType}"
                lblTherapyType.Location = New Point(8, HEADER_HEIGHT + 64) ' ქვემოთ ჩაწეული
                lblTherapyType.Size = New Size(CARD_WIDTH - 16, 20) ' ფიქსირებული სიგანე
                card.Controls.Add(lblTherapyType)

                ' სივრცე
                Dim lblSpace As New Label()
                lblSpace.Text = $"{session.Space}"
                lblSpace.Location = New Point(8, HEADER_HEIGHT + 86) ' ქვემოთ ჩაწეული
                lblSpace.Size = New Size(CARD_WIDTH - 16, 20) ' ფიქსირებული სიგანე
                card.Controls.Add(lblSpace)

                ' დაფინანსება
                Dim lblFunding As New Label()
                lblFunding.Text = $"{session.Funding}"
                lblFunding.Location = New Point(8, HEADER_HEIGHT + 108) ' ქვემოთ ჩაწეული
                lblFunding.Size = New Size(CARD_WIDTH - 16, 20) ' ფიქსირებული სიგანე
                card.Controls.Add(lblFunding)

                ' ვადაგადაცილების დღეები - დავამატეთ ვადაგადაცილების შესახებ ინფორმაცია
                Dim overdueText As String = "ვადაგადაცილება: "
                Dim daysOverdue As Integer = Math.Abs((DateTime.Today - session.DateTime.Date).Days)
                overdueText += $"{daysOverdue} დღე"

                Dim lblOverdue As New Label()
                lblOverdue.Text = overdueText
                lblOverdue.Location = New Point(8, HEADER_HEIGHT + 132) ' ბოლო რიგი
                lblOverdue.AutoSize = True
                lblOverdue.ForeColor = Color.DarkRed
                lblOverdue.Font = New Font(lblOverdue.Font.FontFamily, 9, FontStyle.Regular) ' Bold სტილი
                card.Controls.Add(lblOverdue)

                ' რედაქტირების ღილაკი - დავაბრუნეთ ორიგინალური ფერი და სტილი
                If IsAuthorizedUser AndAlso (UserRoleValue = "1" OrElse UserRoleValue = "2" OrElse UserRoleValue = "3") Then
                    Dim btnEdit As New Button()
                    btnEdit.Text = "✎" ' ფანქრის სიმბოლო
                    btnEdit.Font = New Font("Segoe UI Symbol", 10, FontStyle.Bold)
                    btnEdit.ForeColor = Color.White
                    btnEdit.BackColor = Color.FromArgb(220, 0, 0) ' მუქი წითელი - ორიგინალური ფერი
                    btnEdit.Size = New Size(30, 30) ' ორიგინალური ზომა
                    btnEdit.Location = New Point(CARD_WIDTH - 40, HEADER_HEIGHT + 123) ' ორიგინალური პოზიცია
                    btnEdit.FlatStyle = FlatStyle.Flat ' ორიგინალური სტილი
                    btnEdit.FlatAppearance.BorderSize = 0
                    btnEdit.Tag = session.Id ' შევინახოთ სესიის ID ღილაკის Tag-ში
                    btnEdit.Cursor = Cursors.Hand ' ხელის კურსორი

                    ' მრგვალი ფორმის ღილაკი
                    Dim btnPath As New Drawing2D.GraphicsPath()
                    btnPath.AddEllipse(0, 0, btnEdit.Width, btnEdit.Height)
                    btnEdit.Region = New Region(btnPath)

                    ' ღილაკის მიბმა ფუნქციაზე
                    AddHandler btnEdit.Click, AddressOf BtnEditSession_Click

                    ' ღილაკის დამატება ბარათზე
                    card.Controls.Add(btnEdit)

                    ' წინა პლანზე გამოტანა - მნიშვნელოვანია
                    btnEdit.BringToFront()

                    Debug.WriteLine($"ShowCurrentPageCards: ღილაკი დაემატა ბარათზე, ID={session.Id}")
                End If

                ' დავამატოთ ბარათი GroupBox-ზე
                GBRedTasks.Controls.Add(card)

                ' განვაახლოთ ბარათის პოზიცია შემდეგი ბარათისთვის
                cardCount += 1
                If cardCount Mod cardsPerRow = 0 Then
                    ' ახალი მწკრივი
                    xPos = CARD_MARGIN
                    yPos += CARD_HEIGHT + CARD_MARGIN
                Else
                    ' იგივე მწკრივი, შემდეგი სვეტი
                    xPos += CARD_WIDTH + horizontalSpacing
                End If
            Next

        ' განვაახლოთ პაგინაციის ღილაკების მდებარეობა
        'BtnPrev.Location = New Point(10, GBRedTasks.Height - 40)
        'LPage.Location = New Point(50, GBRedTasks.Height - 35)
        'BtnNext.Location = New Point(100, GBRedTasks.Height - 40)

        ' წინა პლანზე გამოტანა
        'BtnPrev.BringToFront()
        'LPage.BringToFront()
        'BtnNext.BringToFront()
    End Sub

    ''' <summary>
    ''' განაახლებს პაგინაციის კონტროლებს
    ''' </summary>
    Private Sub UpdatePaginationControls()
        ' LPage-ში ვაჩვენებთ მიმდინარე გვერდს და საერთო რაოდენობას
        LPage.Text = $"{CurrentPage + 1}/{TotalPages}"

        ' BtnPrev (წინა გვერდი) გამორთულია თუ პირველ გვერდზე ვართ
        BtnPrev.Enabled = (CurrentPage > 0)

        ' BtnNext (შემდეგი გვერდი) გამორთულია თუ ბოლო გვერდზე ვართ
        BtnNext.Enabled = (CurrentPage < TotalPages - 1)

        Debug.WriteLine($"UpdatePaginationControls: გვერდი {CurrentPage + 1}/{TotalPages}, " &
                        $"BtnPrev.Enabled={BtnPrev.Enabled}, BtnNext.Enabled={BtnNext.Enabled}")
    End Sub

    ''' <summary>
    ''' წინა გვერდის ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnPrev_Click(sender As Object, e As EventArgs) Handles BtnPrev.Click
        If CurrentPage > 0 Then
            CurrentPage -= 1
            ShowCurrentPageCards()
            UpdatePaginationControls()
        End If
    End Sub

    ''' <summary>
    ''' შემდეგი გვერდის ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles BtnNext.Click
        If CurrentPage < TotalPages - 1 Then
            CurrentPage += 1
            ShowCurrentPageCards()
            UpdatePaginationControls()
        End If
    End Sub

    ''' <summary>
    ''' ვადაგადაცილებული სესიების ბარათების შექმნა და GBRedTasks-ში განთავსება
    ''' </summary>
    Public Sub PopulateOverdueSessions(sessions As List(Of SessionModel), isAuthorized As Boolean, userRole As String)
        Try
            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: დაიწყო, სესიების რაოდენობა: {If(sessions Is Nothing, 0, sessions.Count)}")

            ' შევინახოთ მომხმარებლის მდგომარეობა
            IsAuthorizedUser = isAuthorized
            UserRoleValue = userRole

            ' მნიშვნელოვანი: განვაახლოთ ვადაგადაცილებული სესიების ლეიბლები აქვე
            ' რადგან სესიები უკვე გვაქვს
            Dim allOverdueSessions As Integer = If(sessions Is Nothing, 0, sessions.Count)

            ' დღევანდელი ვადაგადაცილებული სესიები
            Dim today As DateTime = DateTime.Today
            Dim todayOverdueSessions As Integer = 0

            If sessions IsNot Nothing Then
                For Each session In sessions
                    If session.DateTime.Date = today Then
                        todayOverdueSessions += 1
                    End If
                Next
            End If

            ' განვაახლოთ ლეიბლები
            LNeerReaction.Text = allOverdueSessions.ToString()
            LNeedReactionToday.Text = todayOverdueSessions.ToString()
            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: სტატისტიკა განახლდა - სულ: {allOverdueSessions}, დღევანდელი: {todayOverdueSessions}")

            ' გავასუფთავოთ ძველი სესიები და ჩავსვათ ახალი
            AllSessions = New List(Of SessionModel)()

            ' გადავატაროთ ყველა სესია, მოვიპოვოთ ცხრილიდან არსებული raw data და შევასწოროთ ველები
            If sessions IsNot Nothing Then
                For Each session In sessions
                    AllSessions.Add(session)
                Next
            End If

            ' თუ სესიები არ არის, დავამატოთ შეტყობინება
            If AllSessions Is Nothing OrElse AllSessions.Count = 0 Then
                ' გასუფთავება ძველი ბარათებისგან
                GBRedTasks.Controls.Clear()
                Dim lblNoSessions As New Label()
                lblNoSessions.Text = "ვადაგადაცილებული სესიები არ არის"
                lblNoSessions.AutoSize = True
                lblNoSessions.Location = New Point(10, 20)
                GBRedTasks.Controls.Add(lblNoSessions)
                ' პაგინაციის კონტროლები გამოვრთოთ
                BtnPrev.Enabled = False
                BtnNext.Enabled = False
                LPage.Text = ""
                Debug.WriteLine("UC_Home.PopulateOverdueSessions: დაემატა შეტყობინება 'ვადაგადაცილებული სესიები არ არის'")
                Return
            End If

            ' დავადგინოთ რამდენი ბარათი ეტევა თითო გვერდზე
            CalculateCardsPerPage()

            ' დავრწმუნდეთ რომ მიმდინარე გვერდი ვალიდურია
            CurrentPage = Math.Min(CurrentPage, Math.Max(0, TotalPages - 1))

            ' ვაჩვენოთ ბარათები
            ShowCurrentPageCards()

            ' განვაახლოთ პაგინაციის კონტროლები
            UpdatePaginationControls()

        Catch ex As Exception
            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: შეცდომა - {ex.Message}")
            MessageBox.Show($"შეცდომა ვადაგადაცილებული სესიების ასახვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' დღევანდელი სესიების სიის განახლება და სტატისტიკის დათვლა
    ''' </summary>
    ''' <param name="sessions">ყველა დღევანდელი სესია</param>
    Public Sub UpdateTodaySessionsStatistics(sessions As List(Of SessionModel))
        Try
            Debug.WriteLine($"UC_Home.UpdateTodaySessionsStatistics: დაიწყო, სესიების რაოდენობა: {If(sessions Is Nothing, 0, sessions.Count)}")
            Debug.WriteLine($"UC_Home.UpdateTodaySessionsStatistics: დღევანდელი თარიღი: {DateTime.Today:dd.MM.yyyy}")

            ' შევამოწმოთ და დავაფიქსიროთ ყველა თარიღი
            If sessions IsNot Nothing AndAlso sessions.Count > 0 Then
                For i As Integer = 0 To Math.Min(10, sessions.Count - 1)
                    Debug.WriteLine($"UC_Home.UpdateTodaySessionsStatistics: სესია #{i + 1} - " &
                           $"ID={sessions(i).Id}, " &
                           $"თარიღი={sessions(i).DateTime:dd.MM.yyyy HH:mm}, " &
                           $"DateTime.ToString()={sessions(i).DateTime.ToString()}, " &
                           $"DateTime.Date={sessions(i).DateTime.Date:dd.MM.yyyy}")
                Next
            End If

            ' დავრწმუნდეთ, რომ ვმუშაობთ მხოლოდ დღევანდელ სესიებზე
            Dim currentDateStr = DateTime.Today.ToString("dd.MM.yyyy") ' მაგ: "11.05.2025"
            Dim confirmedTodaySessions = New List(Of SessionModel)()

            If sessions IsNot Nothing Then
                For Each session In sessions
                    Dim sessionDateStr = session.DateTime.ToString("dd.MM.yyyy")

                    ' სტრიქონების შედარება უფრო საიმედოა ამ შემთხვევაში
                    If sessionDateStr = currentDateStr Then
                        confirmedTodaySessions.Add(session)
                        Debug.WriteLine($"UC_Home.UpdateTodaySessionsStatistics: ✓ დამატებულია დღევანდელი სესია - ID={session.Id}, " &
                               $"თარიღი={session.DateTime:dd.MM.yyyy}, " &
                               $"სტრიქონი={sessionDateStr}")
                    Else
                        Debug.WriteLine($"UC_Home.UpdateTodaySessionsStatistics: ✗ გამოტოვებულია არა-დღევანდელი სესია - ID={session.Id}, " &
                               $"თარიღი={session.DateTime:dd.MM.yyyy}, " &
                               $"სტრიქონი={sessionDateStr}, " &
                               $"(დღევანდელი={currentDateStr})")
                    End If
                Next
            End If

            Debug.WriteLine($"UC_Home.UpdateTodaySessionsStatistics: დადასტურებული დღევანდელი სესიები: {confirmedTodaySessions.Count}")

            ' განაახლეთ დღევანდელი სესიების სია
            TodaySessions.Clear()
            For Each session In confirmedTodaySessions
                TodaySessions.Add(session)
            Next

            ' დავითვალოთ სტატისტიკა
            Dim totalSessions As Integer = TodaySessions.Count

            ' შესრულებული სესიები (სტატუსით "შესრულებული")
            Dim completedSessions As Integer = 0
            For Each session As SessionModel In TodaySessions
                If session.Status.Trim().ToLower() = "შესრულებული" Then
                    completedSessions += 1
                End If
            Next

            ' დაგეგმილი სესიები (სტატუსით "დაგეგმილი")
            Dim plannedSessions As Integer = 0
            For Each session As SessionModel In TodaySessions
                If session.Status.Trim().ToLower() = "დაგეგმილი" Then
                    plannedSessions += 1
                End If
            Next

            ' უნიკალური ბენეფიციარები
            Dim beneficiarySet As New HashSet(Of String)
            For Each session As SessionModel In TodaySessions
                Dim fullName As String = session.BeneficiaryName & " " & session.BeneficiarySurname
                beneficiarySet.Add(fullName)
            Next
            Dim uniqueBeneficiaries As Integer = beneficiarySet.Count

            ' უნიკალური თერაპევტები
            Dim therapistSet As New HashSet(Of String)
            For Each session As SessionModel In TodaySessions
                If Not String.IsNullOrWhiteSpace(session.TherapistName) Then
                    therapistSet.Add(session.TherapistName)
                End If
            Next
            Dim uniqueTherapists As Integer = therapistSet.Count

            ' განვაახლოთ ლეიბლები
            LSe.Text = totalSessions.ToString()
            LDone.Text = completedSessions.ToString()
            LNDone.Text = plannedSessions.ToString()
            LBenes.Text = uniqueBeneficiaries.ToString()
            LPers.Text = uniqueTherapists.ToString()

            Debug.WriteLine("UC_Home.UpdateTodaySessionsStatistics: სტატისტიკა განახლდა - " &
            "სულ: " & totalSessions & ", " &
            "შესრულებული: " & completedSessions & ", " &
            "დაგეგმილი: " & plannedSessions & ", " &
            "ბენეფიციარები: " & uniqueBeneficiaries & ", " &
            "პერსონალი: " & uniqueTherapists)

        Catch ex As Exception
            Debug.WriteLine("UC_Home.UpdateTodaySessionsStatistics: შეცდომა - " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' რედაქტირების ღილაკზე დაჭერის დამმუშავებელი - 
    ''' ხსნის NewRecordForm-ს რედაქტირების რეჟიმში
    ''' </summary>
    Private Sub BtnEditSession_Click(sender As Object, e As EventArgs)
        Try
            Dim btn = DirectCast(sender, Button)
            Dim sessionId As Integer = CInt(btn.Tag)

            Debug.WriteLine($"BtnEditSession_Click: დაიწყო სესიის რედაქტირება, ID={sessionId}")

            ' შევამოწმოთ გვაქვს თუ არა dataService
            If dataService Is Nothing Then
                ' ვეცადოთ მოვიპოვოთ Form1-დან
                Dim mainForm = TryCast(Application.OpenForms("Form1"), Form1)
                If mainForm IsNot Nothing Then
                    ' ვცადოთ ავაღოთ მონაცემთა სერვისი Form1-დან რეფლექსიით
                    Dim dataServiceField = mainForm.GetType().GetField("dataService", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                    If dataServiceField IsNot Nothing Then
                        Dim formDataService = dataServiceField.GetValue(mainForm)
                        If TypeOf formDataService Is IDataService Then
                            dataService = DirectCast(formDataService, IDataService)
                            Debug.WriteLine("BtnEditSession_Click: dataService წარმატებით მიღებულია Form1-დან")
                        End If
                    End If
                End If

                ' თუ ver მოვიპოვეთ dataService, ვაჩვენოთ შეტყობინება და გამოვიდეთ
                If dataService Is Nothing Then
                    MessageBox.Show("მონაცემთა სერვისი არ არის ხელმისაწვდომი", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            ' მოვიპოვოთ მომხმარებლის ელფოსტა
            Dim userEmail As String = "უცნობი"
            ' ვცადოთ Form1-დან მოვიპოვოთ მომხმარებლის ელფოსტა
            Dim mainForm2 = TryCast(Application.OpenForms("Form1"), Form1)
            If mainForm2 IsNot Nothing AndAlso mainForm2.GetType().GetMethod("GetUserEmail") IsNot Nothing Then
                ' თუ GetUserEmail მეთოდი არსებობს, გამოვიყენოთ
                userEmail = CType(mainForm2.GetType().GetMethod("GetUserEmail").Invoke(mainForm2, Nothing), String)
                Debug.WriteLine($"BtnEditSession_Click: მომხმარებლის ელფოსტა = {userEmail}")
            End If

            ' შევქმნათ და გავხსნათ რედაქტირების ფორმა
            Dim editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_Home")

            ' გავხსნათ ფორმა
            Dim result = editForm.ShowDialog()

            ' თუ ფორმა დაიხურა OK რეზულტატით, განვაახლოთ მონაცემები
            If result = DialogResult.OK Then
                Debug.WriteLine($"BtnEditSession_Click: სესია ID={sessionId} წარმატებით განახლდა")

                ' თუ გვაქვს RefreshAllData მეთოდი, გამოვიძახოთ
                Try
                    Dim refreshMethod = Me.GetType().GetMethod("RefreshAllData", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                    If refreshMethod IsNot Nothing Then
                        refreshMethod.Invoke(Me, Nothing)
                        Debug.WriteLine("BtnEditSession_Click: RefreshAllData მეთოდი გამოძახებულია")
                    Else
                        ' ალტერნატიულად, გამოვიძახოთ LoadData
                        Dim loadDataMethod = Me.GetType().GetMethod("LoadData", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                        If loadDataMethod IsNot Nothing Then
                            loadDataMethod.Invoke(Me, Nothing)
                            Debug.WriteLine("BtnEditSession_Click: LoadData მეთოდი გამოძახებულია")
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"BtnEditSession_Click: შეცდომა მონაცემების განახლებისას: {ex.Message}")
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnEditSession_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიის რედაქტირების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' სრულიად განაახლებს ყველა მონაცემს და UI ელემენტს
    ''' </summary>
    Private Sub RefreshAllData()
        Debug.WriteLine("UC_Home.RefreshAllData: დაიწყო ყველა მონაცემის განახლება")

        If dataService Is Nothing Then
            Debug.WriteLine("UC_Home.RefreshAllData: dataService არ არის ინიციალიზებული")
            Return
        End If

        Try
            ' 1. განვაახლოთ მომხმარებლის ინფორმაცია და მისალმება
            ' ეს ველები მოგვეწოდება მთავარი Form1-დან viewModel-ის საშუალებით
            LUserName.Text = viewModel.UserName
            LWish.Text = viewModel.GetGreetingByTime() ' მისალმების განახლება მიმდინარე დროის მიხედვით

            ' 2. განვაახლოთ დროის ინფორმაცია
            viewModel.CurrentTime = DateTime.Now
            LTime.Text = viewModel.FormattedTime
            LDate.Text = viewModel.FormattedDate
            LWeekDay.Text = viewModel.WeekDayName

            ' 3. დღევანდელი სტატისტიკის განახლება - GB_Today ლეიბლები
            Dim todaySessions = dataService.GetTodaySessions()
            Debug.WriteLine($"UC_Home.RefreshAllData: მიღებულია {todaySessions.Count} დღევანდელი სესია")

            ' შესრულებული სესიები (სტატუსით "შესრულებული")
            Dim completedSessions As Integer = 0
            For Each session In todaySessions
                If session.Status.Trim().ToLower() = "შესრულებული" Then
                    completedSessions += 1
                End If
            Next

            ' დაგეგმილი სესიები (სტატუსით "დაგეგმილი")
            Dim plannedSessions As Integer = 0
            For Each session In todaySessions
                If session.Status.Trim().ToLower() = "დაგეგმილი" Then
                    plannedSessions += 1
                End If
            Next

            ' უნიკალური ბენეფიციარები
            Dim beneficiarySet As New HashSet(Of String)
            For Each session In todaySessions
                beneficiarySet.Add(session.FullName)
            Next

            ' უნიკალური თერაპევტები
            Dim therapistSet As New HashSet(Of String)
            For Each session In todaySessions
                If Not String.IsNullOrWhiteSpace(session.TherapistName) Then
                    therapistSet.Add(session.TherapistName)
                End If
            Next

            ' ლეიბლების განახლება
            LSe.Text = todaySessions.Count.ToString()
            LDone.Text = completedSessions.ToString()
            LNDone.Text = plannedSessions.ToString()
            LBenes.Text = beneficiarySet.Count.ToString()
            LPers.Text = therapistSet.Count.ToString()

            ' 4. ვადაგადაცილებული სესიების სტატისტიკა - GBActiveTasks ლეიბლები
            Dim overdueSessions = dataService.GetOverdueSessions()
            Debug.WriteLine($"UC_Home.RefreshAllData: მიღებულია {overdueSessions.Count} ვადაგადაცილებული სესია")

            ' დღევანდელი ვადაგადაცილებული სესიები
            Dim today = DateTime.Today
            Dim todayOverdueSessions As Integer = 0
            For Each session In overdueSessions
                If session.DateTime.Date = today Then
                    todayOverdueSessions += 1
                End If
            Next

            ' ლეიბლების განახლება
            LNeerReaction.Text = overdueSessions.Count.ToString()
            LNeedReactionToday.Text = todayOverdueSessions.ToString()

            ' 5. დაბადების დღეები - GBBD ლეიბლები
            Dim birthdays = dataService.GetUpcomingBirthdays(30) ' 30 დღის განმავლობაში
            Debug.WriteLine($"UC_Home.RefreshAllData: მიღებულია {birthdays.Count} დაბადების დღე")

            ' UI-ზე გამოტანა
            PopulateUpcomingBirthdays(birthdays)

            ' 6. ვადაგადაცილებული სესიების ბარათები - GBRedTasks
            ' წამოვიღოთ ყველა სესია
            Dim allSessions = dataService.GetAllSessions()
            Debug.WriteLine($"UC_Home.RefreshAllData: მიღებულია {allSessions.Count} სესია (სულ)")

            ' ავტორიზაციის ინფორმაცია უცვლელი რჩება
            ' გამოვიყენოთ ვადაგადაცილებული სესიების გამოტანის ფუნქცია
            PopulateOverdueSessions(overdueSessions, IsAuthorizedUser, UserRoleValue)

            Debug.WriteLine("UC_Home.RefreshAllData: ყველა მონაცემი წარმატებით განახლდა")
        Catch ex As Exception
            Debug.WriteLine($"UC_Home.RefreshAllData: შეცდომა - {ex.Message}")
            Debug.WriteLine($"UC_Home.RefreshAllData: StackTrace - {ex.StackTrace}")
            Throw ' გადავაგდოთ შეცდომა ზედა დონეზე დამუშავებისთვის
        End Try
    End Sub
    ''' <summary>
    ''' ტესტური მეთოდი კონტროლების ხილვადობის დიაგნოსტიკისთვის
    ''' </summary>
    Public Sub TestDirectControls()
        Try
            Debug.WriteLine("UC_Home.TestDirectControls: დაიწყო კონტროლების ტესტირება")

            ' კონტროლების მდგომარეობის შემოწმება
            Debug.WriteLine($"GBRedTasks ხილვადობა: {GBRedTasks.Visible}, " &
                          $"ზომები: {GBRedTasks.Width}x{GBRedTasks.Height}, " &
                          $"კონტროლების რაოდენობა: {GBRedTasks.Controls.Count}")

            ' პაგინაციის კონტროლების შემოწმება
            Debug.WriteLine($"BtnPrev ხილვადობა: {BtnPrev.Visible}, მდგომარეობა: {BtnPrev.Enabled}")
            Debug.WriteLine($"BtnNext ხილვადობა: {BtnNext.Visible}, მდგომარეობა: {BtnNext.Enabled}")
            Debug.WriteLine($"LPage ხილვადობა: {LPage.Visible}, ტექსტი: {LPage.Text}")

            ' სესიების ინფორმაცია
            Debug.WriteLine($"AllSessions რაოდენობა: {AllSessions.Count}, " &
                          $"CardsPerPage: {CardsPerPage}, " &
                          $"CurrentPage: {CurrentPage}, " &
                          $"TotalPages: {TotalPages}")
        Catch ex As Exception
            Debug.WriteLine($"UC_Home.TestDirectControls: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ' ViewModel-ის თვისებების წაკითხვის მეთოდები UI-დან
    Public ReadOnly Property UserName() As String
        Get
            Return viewModel.UserName
        End Get
    End Property
    ''' <summary>
    ''' BtnAddAray ღილაკზე დაჭერა - ახალი ჩანაწერის ფორმის გამოჩენა
    ''' შეზღუდვით, რომ მხოლოდ ერთი ფორმა გაიხსნას
    ''' </summary>
    Private Sub BtnAddAray_Click(sender As Object, e As EventArgs) Handles BtnAddAray.Click
        Debug.WriteLine("UC_Home.BtnAddAray_Click: ახალი ჩანაწერის დამატების მოთხოვნა")

        Try
            ' შევამოწმოთ უკვე გახსნილია თუ არა NewRecordForm
            For Each frm As Form In Application.OpenForms
                If TypeOf frm Is NewRecordForm Then
                    ' თუ უკვე გახსნილია, მოვიტანოთ წინ და გამოვიდეთ მეთოდიდან
                    Debug.WriteLine("UC_Home.BtnAddAray_Click: NewRecordForm უკვე გახსნილია, ფოკუსის გადატანა")
                    frm.Focus()
                    Return
                End If
            Next

            ' შევამოწმოთ გვაქვს თუ არა dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Debug.WriteLine("UC_Home.BtnAddAray_Click: dataService არ არის ინიციალიზებული")
                Return
            End If

            ' მომხმარებლის email მიღება
            Dim userEmail As String = "უცნობი"
            ' მოვძებნოთ მთავარი ფორმა და თუ იქ არის GetUserEmail მეთოდი, გამოვიყენოთ
            Dim mainForm = TryCast(Application.OpenForms("Form1"), Form1)
            If mainForm IsNot Nothing AndAlso mainForm.GetType().GetMethod("GetUserEmail") IsNot Nothing Then
                ' თუ GetUserEmail მეთოდი არსებობს, გამოვიყენოთ
                userEmail = CType(mainForm.GetType().GetMethod("GetUserEmail").Invoke(mainForm, Nothing), String)
            End If

            ' ნაგულისხმევად "სესია" ტიპი
            Dim recordType As String = "სესია"

            ' NewRecordForm-ის გახსნა Add რეჟიმში
            Dim newRecordForm As New NewRecordForm(dataService, recordType, userEmail, "UC_Home")
            newRecordForm.Show()

        Catch ex As Exception
            Debug.WriteLine($"UC_Home.BtnAddAray_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ახალი ჩანაწერის ფორმის გახსნის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class