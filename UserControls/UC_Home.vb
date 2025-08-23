Option Infer On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Globalization
Imports System.Linq
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services
Imports Scheduler_v8_8a.Services

' Ensure Option Infer for type inference and fix loops

Public Class UC_Home
    Inherits UserControl

    ' ViewModel სივრცისთვისც ხდება მიბმა
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

    ' ========= Performance fields =========
    Private overdueCardCache As New Dictionary(Of Integer, Panel) ' sessionId -> card panel
    Private sharedHeaderFont As Font = Nothing
    Private sharedBoldFont As Font = Nothing
    Private sharedNormalFont As Font = Nothing
    Private contextMenuForEditUsers As Boolean = False

    Private Sub InitializePerformanceResources()
        If sharedHeaderFont Is Nothing Then
            sharedHeaderFont = New Font(Font.FontFamily, 9, FontStyle.Regular)
            sharedBoldFont = New Font("Sylfaen", 10, FontStyle.Bold)
            sharedNormalFont = New Font(Font.FontFamily, 9, FontStyle.Regular)
        End If
        ' Enable double buffering on GroupBox containers to reduce flicker & painting cost
        EnableDoubleBuffer(GBRedTasks)
    End Sub

    Private Sub EnableDoubleBuffer(ctrl As Control)
        Try
            Dim t = ctrl.GetType()
            Dim pi = t.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            If pi IsNot Nothing Then pi.SetValue(ctrl, True, Nothing)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' კონსტრუქტორი: იღებს HomeViewModel-ს Data Binding-ისთვის
    ''' </summary>
    ''' <param name="homeVm">HomeViewModel ობიექტი</param>
    Public Sub New(homeVm As HomeViewModel)
        ' UI ელემენტების ინიციალიზაცია
        InitializeComponent()

        ' ViewModel შემოწმება - თუ null-აა, შევქმნათ nueva
        If homeVm IsNot Nothing Then
            viewModel = homeVm
        Else
            ' homeVm არის null, შევქმნათ nueva ინსტანცია
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

        ' მიდის ViewModel-ზე
        BindToViewModel()

        ' საწყისი მონაცემების ჩატვირთვა
        LoadData()

        ' საწყისი მდგომარეობაში ინსტრუმენტები დამალულია
        SetToolsVisibility(False)

        ' განვასახიერებთ დღევანდელი სტატისტიკა
        UpdateTodayStatistics()
        UpdateOverdueStatistics()

        ' Resize ივენტის მიბმა
        AddHandler Me.Resize, AddressOf UC_Home_Resize
    End Sub

    ''' <summary>
    ''' ხილვადობის დაყენება ინსტურმენტების პანელისთვის (GBTools)
    ''' </summary>
    ''' <param name="visible">უნდა იყოს თუ არა ხილული</param>
    Public Sub SetToolsVisibility(visible As Boolean)
        ' ინსტრუმენტების პანელი მხოლოდ მაშინ visibles, როცა მომხმარებელს აქვთ შესაბამისი როლი
        GBTools.Visible = visible

        ' დავარეგულიროთ სხვა კონტეინერების ზომები, რომ არ იყოს ცარიელი ადგა
        If visible Then
            ' ჩვეულებრივი მდგომარეობა, როცა ყველა პანელი ხდება
            GBGreeting.Width = 253 ' საწყისი სიგანე
        Else
            ' როცა GBTools არ ჩანს, GBGreeting-ი შეიძლება იყოს უფრო ფართო
            GBGreeting.Width = 253 ' გაფართოებული სიგანე
        End If
    End Sub

    ''' <summary>
    ''' Timer-ის მოვლენების უმარტივესი დამმუშავი - მხოლოდ საათის განახლება
    ''' </summary>
    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        Try
            ' უბრალოდ მიმდინარე დროის გაწხრვა
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
            ' მიბმა ViewModel-ის PropertyChanged ივენტზე
            AddHandler viewModel.PropertyChanged, AddressOf ViewModel_PropertyChanged

            ' მისალმების დაყენება
            LUserName.Text = viewModel.UserName
            LWish.Text = viewModel.Greeting

            ' დროის დაყენება
            UpdateTimeDisplay()
        Else
            Debug.WriteLine("BindToViewModel: viewModel არის null, ვერ ხერხდება დაბაინდვა")

            ' დავაყენოთ საწყისი მნიშვნელობები
            LUserName.Text = "მომხმარებლო"
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

            ' ასევე, განვასახიროთ ViewModel
            If viewModel IsNot Nothing Then
                viewModel.UserName = userName
                Debug.WriteLine($"UpdateUserName: ViewModel განახლებულია, UserName='{viewModel.UserName}'")
            End If

            ' ძალით გან[:-] LUserName
            LUserName.Refresh()
            Application.DoEvents()

            Debug.WriteLine($"UpdateUserName: LUserName.Text = '{LUserName.Text}'")
        Catch ex As Exception
            Debug.WriteLine($"UpdateUserName: შეცდომა - {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' დროისShown განახლება ViewModel-დან
    ''' </summary>
    Private Sub UpdateTimeDisplay()
        ' UI-ს ვანახლებთ მხოლოდ თუ კონტროლი ხილვადია და ინციალიზებულია
        If Me.IsHandleCreated AndAlso Me.Visible Then
            ' UI ნაკრების შემოწმება
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
    ''' განაახლებთ დღევანდელი სესიების_STATისტიკის ლეიბლებს
    ''' </summary>
    Private Sub UpdateTodayStatistics()
        Try
            If AllSessions Is Nothing Then Return
            Dim today = DateTime.Today
            Dim todayList = AllSessions.Where(Function(s) s.DateTime.Date = today).ToList()
            Dim completedSessions = todayList.Where(Function(s) s.Status.Trim().ToLower() = "შესრულებული").Count()
            Dim plannedSessions = todayList.Where(Function(s) s.Status.Trim().ToLower() = "დაგეგმილი").Count()
            Dim beneficiarySet = New HashSet(Of String)(todayList.Select(Function(s) s.BeneficiaryName & " " & s.BeneficiarySurname))
            Dim therapistSet = New HashSet(Of String)(todayList.Where(Function(s) Not String.IsNullOrWhiteSpace(s.TherapistName)).Select(Function(s) s.TherapistName))
            LSe.Text = todayList.Count.ToString()
            LDone.Text = completedSessions.ToString()
            LNDone.Text = plannedSessions.ToString()
            LBenes.Text = beneficiarySet.Count.ToString()
            LPers.Text = therapistSet.Count.ToString()
        Catch ex As Exception
            Debug.WriteLine("UpdateTodayStatistics შეცდომა: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' მოახლოებული დაბადების დღეებისShows GBBD გრუპბოქსში
    ''' </summary>
    Public Sub PopulateUpcomingBirthdays(birthdays As List(Of BirthdayModel))
        Try
            Debug.WriteLine($"UC_Home.PopulateUpcomingBirthdays: დაიწყო, მოახლოებული დაბადების დღეების ჩვენსებას")

            ' გავასუფთავოთ GBBD გრუპბოქსი
            GBBD.Controls.Clear()

            ' შევამოწმოთ máme თუ არა მონაცემები
            If birthdays Is Nothing OrElse birthdays.Count = 0 Then
                ' საჩვენებელი ლეიბლი: არ არის დაბადების დღის მონაცემები
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

                ' გავისტუმროთ ლეიბლი
                GBBD.Controls.Add(lblBirthday)

                ' გადავიდეთ შემდეგ პოზიციისთვის
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

            ' შეცდომის შემთხვევაში ჯერ კიდევ ვაჩვენოთ შეტყობინება
            Dim lblError As New Label()
            lblError.Text = "შეცდომა დაბადების დღეების_show-ზე"
            lblError.AutoSize = True
            lblError.Location = New Point(10, 20)
            lblError.ForeColor = Color.Red
            GBBD.Controls.Add(lblError)
        End Try
    End Sub
    ''' <summary>
    ''' განაახლებთ ვადაგადაცილებულ სესიების სტატისტიკის ლეიბლებს
    ''' </summary>
    Private Sub UpdateOverdueStatistics()
        Try
            If AllSessions Is Nothing Then Return
            Dim today = DateTime.Today
            Dim overdueList = AllSessions.Where(Function(s) s.IsOverdue).ToList()
            Dim todayOver = overdueList.Where(Function(s) s.DateTime.Date = today).Count()
            LNeerReaction.Text = overdueList.Count.ToString()
            LNeedReactionToday.Text = todayOver.ToString()
        Catch ex As Exception
            Debug.WriteLine("UpdateOverdueStatistics შეცდომა: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' ViewModel-ის PropertyChanged ივენტის დამმუშავი
    ''' </summary>
    Private Sub ViewModel_PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs)
        ' ივენტის მართვა UI-ის მთავარ ნაკლებში
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
        ' ViewModel-ის RefreshData მეთოდების გამოძახება
        ' მონაცემების ჩატვირთვა (დაბადების დღეები, სესიები, დავალებები)
        viewModel.RefreshData()
        ' განვასახიროთ დღევანდელი პერსონალური სტატისტიკა
        UpdateTodayStatistics()
        ' განვსახიროთ ვადაგადაცილებული სესიები
        UpdateOverdueStatistics()
        ' მოახლოვებული დაბადების დღეების განახლება ყოველდღიური განახლების ნაწილად

    End Sub

    ''' <summary>
    ''' Refresh ღილაკის Click ივენტის დამმუშავი - სრულად განახლებს ყველა მონაცემს
    ''' </summary>
    Private Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles BtnRefresh.Click
        Debug.WriteLine("BtnRefresh_Click: განახლების მოთხოვნა მიღებულია")

        ' ღილაკის დროებითი გამორთვა ორჯერ დაჭერის პრევენციისთვის
        BtnRefresh.Enabled = False
        Cursor = Cursors.WaitCursor

        Try
            ' გაუქმება ყველა ქეშის, თუ შესაძლებელი
            If dataService IsNot Nothing Then
                Try
                    ' SheetDataService-isათვის
                    If TypeOf dataService Is SheetDataService Then
                        DirectCast(dataService, SheetDataService).InvalidateAllCache()
                        Debug.WriteLine("BtnRefresh_Click: SheetDataService ქეში გასუფთავდა")
                    ElseIf TypeOf dataService Is GoogleSheetsDataService Then
                        ' თუ ამ სერვისს აქვს ქეშის გასუფთლების მეთოდი
                        Debug.WriteLine("BtnRefresh_Click: GoogleSheetsDataService ქეშირება")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"BtnRefresh_Click: შეცდომა ქეშის გაუქმებისას - {ex.Message}")
                End Try
            Else
                Debug.WriteLine("BtnRefresh_Click: dataService არ არის ინიციალიზებული")
            End If

            ' 1. განახლეთ ViewModel-ის მონაცემები
            viewModel.RefreshData()

            ' 2. ჩავტვირთეთ ყველა მონაცემი ხელახლა
            RefreshAllData()

            ' 3. ვაცნობოთ მომხმარებელს
            MessageBox.Show("მონაცემები წარმატებით განახლდა", "ინფორმირება", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Debug.WriteLine($"BtnRefresh_Click: განახლების შეცდომა - {ex.Message}")
            MessageBox.Show($"მონაცემების განახლებანის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ღილაკის და კურსორის აღდგენა
            BtnRefresh.Enabled = True
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' UserControl-ის დატვირთვის ივენტი, რომელიც გაეშვება იგი პირველად ჩნდება
    ''' </summary>
    Private Sub UC_Home_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Timer-ის ხელახლა დაწყება თუ გაჩერებული იყო
        If Not Timer1.Enabled Then
            Timer1.Start()
        End If
    End Sub

    ''' <summary>
    ''' UserControl-ის დამალვის ივენტი, რომელიც გაეშვებოდა როცა კონტროლი იმალება
    ''' </summary>
    Private Sub UC_Home_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        ' Timer-ი ვმუშაობს მხოლოდ როცა ხილულია
        Timer1.Enabled = Me.Visible
    End Sub

    ''' <summary>
    ''' UserControl-ის Resize ივენტის დამმუშავი
    ''' </summary>
    Private Sub UC_Home_Resize(sender As Object, e As EventArgs)
        ' შევამოწმოთ არის თუ არა ფორმა ინციალიზირებული
        If Not Me.IsHandleCreated OrElse GBRedTasks Is Nothing Then
            Return
        End If

        ' თუ სესიები გვაქვს, განవరით ბარათების_show
        If AllSessions.Count > 0 Then
            ' ვიანგარიშოთ ისევ რამდენი ერევა
            CalculateCardsPerPage()

            ' თუ მიმდინარე გვერდი აღარ არსებობს (მაგ. ფანჯარა გაიზარდა), დავაბრუნოთ პირველ გვერდზე
            If CurrentPage >= TotalPages Then
                CurrentPage = Math.Max(0, TotalPages - 1)
            End If

            ' ჩავატმართოთ_current გვერდის ბარათები
            ShowCurrentPageCards()

            ' განვაახლოთ პაგინაციის კონტროლები
            UpdatePaginationControls()
        End If
    End Sub

    ''' <summary>
    ''' ანგარიშობთ რამდენი ბარათი ეტევა ერთ გვერდზე და საერთო გვერდების რაოდენობას
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
            ' პაგინაციის ღილაკებისათვის დარჩება ადგილი ბოლოში, მაგრამ არ შეიცვლება მათი მდებარეობა
            Dim paginationHeight As Integer = 50 ' სავარაუდო სიმაღლე პაგინაციის ღილაკებისთვის
            Dim availableHeight As Integer = GBRedTasks.ClientSize.Height - paginationHeight - (2 * CARD_MARGIN)

            ' გამოვითვრამდე რამდენი მწკრივი ჩაეტევა
            If (cardsPerRow * (CARD_WIDTH + CARD_MARGIN)) > availableWidth Then
                cardsPerRow = Math.Max(1, cardsPerRow - 1)
            End If
            Dim rowsPerPage As Integer = Math.Max(1, availableHeight \ (CARD_HEIGHT + CARD_MARGIN))

            ' უსაფრთხოებისთვის, თუ დათვლილი მწ მუნავხე სიმაღლე აჭარბებს ხელმისაწვდელ სიმაღლეს
            If (rowsPerPage * (CARD_HEIGHT + CARD_MARGIN)) > availableHeight Then
                rowsPerPage = Math.Max(1, rowsPerPage - 1)
            End If

            ' გამოითვალეთ რაოდენობა ბარათი ეტევა ერთ გვერდზე
            CardsPerPage = cardsPerRow * rowsPerPage

            ' AllSessions უკვე შეიცავს მხოლოდ ვადაგადაცილებული სესიები
            ' საერთო გვერდების რაოდენობა
            TotalPages = Math.Max(1, Math.Ceiling(AllSessions.Count / CDbl(CardsPerPage)))

            Debug.WriteLine($"CalculateCardsPerPage: GBRedTasks.Size={GBRedTasks.Size}")
            Debug.WriteLine($"CalculateCardsPerPage: availableWidth={availableWidth}, availableHeight={availableHeight}")
            Debug.WriteLine($"CalculateCardsPerPage: cardsPerRow={cardsPerRow}, rowsPerPage={rowsPerPage}")
            Debug.WriteLine($"CalculateCardsPerPage: CardsPerPage={CardsPerPage}, TotalPages={TotalPages}")
            Debug.WriteLine($"CalculateCardsPerPage: AllSessions.Count={AllSessions.Count}")

        Catch ex As Exception
            Debug.WriteLine($"CalculateCardsPerPage: შეცდომა - {ex.Message}")
            CardsPerPage = 8 ' ნაგულისხმენი მნიშვნელობა
            TotalPages = Math.Ceiling(AllSessions.Count / CDbl(CardsPerPage))
        End Try
    End Sub
    ''' <summary>
    ''' აჩვენებს_current გვერდის ბარათებს - განახლებული ვერსია 
    ''' ღილაკის ფერის და ფორმის დაბრუნებით და ვადაგადაცილებული დღეების რამოდენიმე თვის გამოყენებით
    ''' </summary>
    Private Sub ShowCurrentPageCards()
        If AllSessions Is Nothing Then Return
        GBRedTasks.SuspendLayout()
        Try
            GBRedTasks.Controls.Clear()
            If AllSessions.Count = 0 Then
                GBRedTasks.Controls.Add(New Label() With {.Text = "ვადაგადაცილებული სესიები არ არის", .AutoSize = True, .Location = New Point(10, 20)})
                Return
            End If
            Const CARD_WIDTH As Integer = 250
            Const CARD_HEIGHT As Integer = 185
            Const CARD_MARGIN As Integer = 15
            Dim availableWidth As Integer = GBRedTasks.ClientSize.Width - (2 * CARD_MARGIN) - 5
            Dim cardsPerRow As Integer = Math.Max(1, availableWidth \ (CARD_WIDTH + CARD_MARGIN))
            Dim horizontalSpacing As Integer = CARD_MARGIN
            If cardsPerRow > 1 Then
                horizontalSpacing = Math.Max(CARD_MARGIN, (availableWidth - (cardsPerRow * CARD_WIDTH)) \ (cardsPerRow - 1))
            End If
            If CardsPerPage <= 0 Then CardsPerPage = cardsPerRow ' safety
            Dim startIndex As Integer = CurrentPage * CardsPerPage
            Dim endIndex As Integer = Math.Min(startIndex + CardsPerPage - 1, AllSessions.Count - 1)
            Dim xPos As Integer = CARD_MARGIN
            Dim yPos As Integer = CARD_MARGIN + 10
            Dim cardCount As Integer = 0
            For i As Integer = startIndex To endIndex
                Dim session = AllSessions(i)
                Dim card = GetOrCreateOverdueCard(session)
                card.Location = New Point(xPos, yPos)
                GBRedTasks.Controls.Add(card)
                cardCount += 1
                If cardCount Mod cardsPerRow = 0 Then
                    xPos = CARD_MARGIN
                    yPos += CARD_HEIGHT + CARD_MARGIN
                Else
                    xPos += CARD_WIDTH + horizontalSpacing
                End If
            Next
        Finally
            GBRedTasks.ResumeLayout()
        End Try
    End Sub

    ' Create or update single overdue session card (reuse to avoid allocations on every refresh)
    Private Function GetOrCreateOverdueCard(session As SessionModel) As Panel
        Const CARD_WIDTH As Integer = 250
        Const CARD_HEIGHT As Integer = 185
        Const HEADER_HEIGHT As Integer = 24
        InitializePerformanceResources()
        Dim card As Panel = Nothing
        If overdueCardCache.TryGetValue(session.Id, card) AndAlso Not card Is Nothing AndAlso Not card.IsDisposed Then
            UpdateCardContent(card, session)
            Return card
        End If
        Dim isOverdue = session.IsOverdue
        Dim cardBack = ColorPalette.GetSessionCardColor(session.Status, isOverdue)
        Dim headerBack = ColorPalette.GetSessionHeaderColor(session.Status, isOverdue)
        card = New Panel() With {.Size = New Size(CARD_WIDTH, CARD_HEIGHT), .BackColor = cardBack, .Tag = session.Id, .Cursor = Cursors.Hand}
        Dim path As New Drawing2D.GraphicsPath()
        Dim r As Integer = 10
        path.AddArc(0, 0, r * 2, r * 2, 180, 90)
        path.AddArc(CARD_WIDTH - r * 2, 0, r * 2, r * 2, 270, 90)
        path.AddArc(CARD_WIDTH - r * 2, CARD_HEIGHT - r * 2, r * 2, r * 2, 0, 90)
        path.AddArc(0, CARD_HEIGHT - r * 2, r * 2, r * 2, 90, 90)
        path.CloseFigure()
        card.Region = New Region(path)
        Dim headerPanel As New Panel() With {.Size = New Size(CARD_WIDTH, HEADER_HEIGHT), .BackColor = headerBack}
        Dim headerPath As New Drawing2D.GraphicsPath()
        headerPath.AddArc(0, 0, r * 2, r * 2, 180, 90)
        headerPath.AddArc(CARD_WIDTH - r * 2, 0, r * 2, r * 2, 270, 90)
        headerPath.AddLine(CARD_WIDTH, r, CARD_WIDTH, HEADER_HEIGHT)
        headerPath.AddLine(CARD_WIDTH, HEADER_HEIGHT, 0, HEADER_HEIGHT)
        headerPath.AddLine(0, HEADER_HEIGHT, 0, r)
        headerPanel.Region = New Region(headerPath)
        card.Controls.Add(headerPanel)
        headerPanel.Controls.Add(New Label() With {.Name = "LDate", .AutoSize = True, .Location = New Point(8, 4), .Font = sharedHeaderFont, .ForeColor = Color.White})
        headerPanel.Controls.Add(New Label() With {.Name = "LId", .AutoSize = True, .Location = New Point(CARD_WIDTH - 40, 4), .Font = sharedHeaderFont, .ForeColor = Color.White})
        card.Controls.Add(New Label() With {.Name = "LFullName", .Location = New Point(8, HEADER_HEIGHT + 8), .Size = New Size(CARD_WIDTH - 16, 30), .Font = sharedBoldFont, .TextAlign = ContentAlignment.MiddleCenter})
        card.Controls.Add(New Label() With {.Name = "LTherapist", .Location = New Point(8, HEADER_HEIGHT + 42), .Size = New Size(CARD_WIDTH - 16, 20), .Font = sharedNormalFont})
        card.Controls.Add(New Label() With {.Name = "LTherapy", .Location = New Point(8, HEADER_HEIGHT + 64), .Size = New Size(CARD_WIDTH - 16, 20), .Font = sharedNormalFont})
        card.Controls.Add(New Label() With {.Name = "LSpace", .Location = New Point(8, HEADER_HEIGHT + 86), .Size = New Size(CARD_WIDTH - 16, 20), .Font = sharedNormalFont})
        card.Controls.Add(New Label() With {.Name = "LFunding", .Location = New Point(8, HEADER_HEIGHT + 108), .Size = New Size(CARD_WIDTH - 16, 20), .Font = sharedNormalFont})
        card.Controls.Add(New Label() With {.Name = "LOverdue", .Location = New Point(8, HEADER_HEIGHT + 132), .AutoSize = True, .Font = sharedNormalFont, .ForeColor = Color.DarkRed})
        ' ============================================
        ' New color overrides for overdue cards
        If session.IsOverdue Then
            card.BackColor = Color.FromArgb(255, 220, 235) ' ვარდისფერი ფონისთვის
            headerPanel.BackColor = Color.FromArgb(200, 60, 120) ' მუქი ვარდისფერი თავზე
        End If
        ' ============================================

        BuildOverdueContextMenu(card, session)

        overdueCardCache(session.Id) = card
        UpdateCardContent(card, session)
        Return card
    End Function

    Private Sub UpdateCardContent(card As Panel, session As SessionModel)
        Try
            Dim isOverdue = session.IsOverdue
            If isOverdue Then
                card.BackColor = OverdueCardPink
            Else
                card.BackColor = ColorPalette.GetSessionCardColor(session.Status, isOverdue)
            End If
            Dim hdrPanel = card.Controls.OfType(Of Panel)().First()
            hdrPanel.BackColor = If(isOverdue, OverdueHeaderPink, ColorPalette.GetSessionHeaderColor(session.Status, isOverdue))
            Dim lDate = CType(hdrPanel.Controls("LDate"), Label)
            Dim lId = CType(hdrPanel.Controls("LId"), Label)
            lDate.Text = session.FormattedDateTime
            lId.Text = "#" & session.Id.ToString()
            CType(card.Controls("LFullName"), Label).Text = (session.BeneficiaryName & " " & session.BeneficiarySurname).ToUpper()
            CType(card.Controls("LTherapist"), Label).Text = session.TherapistName
            CType(card.Controls("LTherapy"), Label).Text = session.TherapyType
            CType(card.Controls("LSpace"), Label).Text = session.Space
            CType(card.Controls("LFunding"), Label).Text = session.Funding
            Dim lOver = CType(card.Controls("LOverdue"), Label)
            If isOverdue Then
                Dim days = Math.Abs((DateTime.Today - session.DateTime.Date).Days)
                lOver.Text = $"ვადაგადაცილება: {days} დღე"
                lOver.Visible = True
            Else
                lOver.Visible = False
            End If
            BuildOverdueContextMenu(card, session)
            Dim inline = card.Controls.Find("BtnInlineEdit", False).FirstOrDefault()
            If inline IsNot Nothing Then inline.Tag = session.Id
        Catch ex As Exception
            Debug.WriteLine($"UpdateCardContent შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' დამატებული: დეტალების ჩვენების_helper
    ''' </summary>
    Private Sub ShowSessionDetails(session As SessionModel)
        Try
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine($"სესიის ID: {session.Id}")
            sb.AppendLine($"ბენეფიციარი: {session.BeneficiaryName} {session.BeneficiarySurname}")
            sb.AppendLine($"თერაპევტი: {session.TherapistName}")
            sb.AppendLine($"თერაპიის ტიპი: {session.TherapyType}")
            sb.AppendLine($"სივრცე: {session.Space}")
            sb.AppendLine($"დრო: {session.DateTime:dd.MM.yyyy HH:mm}")
            sb.AppendLine($"ხანგრძლივობა: {session.Duration} წთ")
            sb.AppendLine($"ფასი: {If(session.Price > 0, session.Price.ToString("F2") & " ₾", "-")}")
            sb.AppendLine($"დაფინანსება: {session.Funding}")
            sb.AppendLine($"სტატუსი: {session.Status}")
            If session.IsOverdue Then sb.AppendLine("ვადაგადაცილებულია")
            If Not String.IsNullOrWhiteSpace(session.Comments) Then
                sb.AppendLine("კომენტარი:")
                sb.AppendLine(session.Comments)
            End If
            MessageBox.Show(sb.ToString(), "სესიის დეტალური ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Debug.WriteLine($"ShowSessionDetails(Home):錯誤 - {ex.Message}")
        End Try
    End Sub

    Private Function FindSessionById(id As Integer) As SessionModel
        Return AllSessions.FirstOrDefault(Function(s) s.Id = id)
    End Function

    ''' <summary>
    ''' რედაქტირების ღილაკზე დაჭერის დამმუშავი - 
    ''' ხსნის NewRecordForm-ს რედაქტირების რეჟიმში
    ''' </summary>
    Private Sub BtnEditSession_Click(sender As Object, e As EventArgs)
        Try
            Dim btn = DirectCast(sender, Button)
            Dim sessionId As Integer = CInt(btn.Tag)

            Debug.WriteLine($"BtnEditSession_Click: დაიწყო სესიის რედაქტირება, ID={sessionId}")

            If dataService Is Nothing Then
                Dim mainForm = TryCast(Application.OpenForms("Form1"), Form1)
                If mainForm IsNot Nothing Then
                    Dim dataServiceField = mainForm.GetType().GetField("dataService", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                    If dataServiceField IsNot Nothing Then
                        Dim formDataService = dataServiceField.GetValue(mainForm)
                        If TypeOf formDataService Is IDataService Then
                            dataService = DirectCast(formDataService, IDataService)
                            Debug.WriteLine("BtnEditSession_Click: dataService მიღებული Form1-დან")
                        End If
                    End If
                End If
                If dataService Is Nothing Then
                    MessageBox.Show("მონაცემთა სერვისი არ არის ხელმისაწვდელი", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            Dim userEmail As String = "უცნობი"
            Dim mainForm2 = TryCast(Application.OpenForms("Form1"), Form1)
            If mainForm2 IsNot Nothing AndAlso mainForm2.GetType().GetMethod("GetUserEmail") IsNot Nothing Then
                userEmail = CType(mainForm2.GetType().GetMethod("GetUserEmail").Invoke(mainForm2, Nothing), String)
            End If

            Dim editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_Home")

            Dim result = editForm.ShowDialog()

            If result = DialogResult.OK Then
                Debug.WriteLine($"BtnEditSession_Click: სესია ID={sessionId} განახლდა")

                Try
                    Dim refreshMethod = Me.GetType().GetMethod("RefreshAllData", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                    If refreshMethod IsNot Nothing Then
                        refreshMethod.Invoke(Me, Nothing)
                    Else
                        Dim loadDataMethod = Me.GetType().GetMethod("LoadData", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                        If loadDataMethod IsNot Nothing Then loadDataMethod.Invoke(Me, Nothing)
                    End If
                Catch innerEx As Exception
                    Debug.WriteLine($"BtnEditSession_Click: Refresh შეცდომა - {innerEx.Message}")
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnEditSession_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიის რედაქტირების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' სრულად განახლებულ მონაცემებს და UI ელემენტს
    ''' </summary>
    Private Sub RefreshAllData()
        Debug.WriteLine("UC_Home.RefreshAllData: დაწყებული ყველა მონაცემის განახლება")

        If dataService Is Nothing Then
            Debug.WriteLine("UC_Home.RefreshAllData: dataService არ არის ინიციალებელი")
            Return
        End If

        Try
            LUserName.Text = viewModel.UserName
            LWish.Text = viewModel.GetGreetingByTime()
            viewModel.CurrentTime = DateTime.Now
            LTime.Text = viewModel.FormattedTime
            LDate.Text = viewModel.FormattedDate
            LWeekDay.Text = viewModel.WeekDayName

            Dim todaySessions = dataService.GetTodaySessions()
            Debug.WriteLine($"UC_Home.RefreshAllData: მიღებულია {todaySessions.Count} დღევანდელი სესია")
            Dim completedSessions = todaySessions.Where(Function(s) s.Status.Trim().ToLower() = "შესრულებული").Count()
            Dim plannedSessions = todaySessions.Where(Function(s) s.Status.Trim().ToLower() = "დაგეგმილი").Count()
            Dim beneficiarySet As New HashSet(Of String)(todaySessions.Select(Function(s) s.FullName))
            Dim therapistSet As New HashSet(Of String)(todaySessions.Where(Function(s) Not String.IsNullOrWhiteSpace(s.TherapistName)).Select(Function(s) s.TherapistName))
            LSe.Text = todaySessions.Count.ToString()
            LDone.Text = completedSessions.ToString()
            LNDone.Text = plannedSessions.ToString()
            LBenes.Text = beneficiarySet.Count.ToString()
            LPers.Text = therapistSet.Count.ToString()

            Dim overdueSessions = dataService.GetOverdueSessions()
            Debug.WriteLine($"UC_Home.RefreshAllData: მიღებულია {overdueSessions.Count} ვადაგადაცილებული სესიები")
            Dim today = DateTime.Today
            Dim todayOverdueSessions = overdueSessions.Where(Function(s) s.DateTime.Date = today).Count()
            LNeerReaction.Text = overdueSessions.Count.ToString()
            LNeedReactionToday.Text = todayOverdueSessions.ToString()

            ' Clean old dynamic controls from GBActiveTasks (correct casing)
            If GBActiveTasks.Controls.Count > 0 Then
                ' Keep static labels only (optional) else clear all
                For i = GBActiveTasks.Controls.Count - 1 To 0 Step -1
                    Dim c = GBActiveTasks.Controls(i)
                    If TypeOf c Is Label AndAlso c.Name.StartsWith("Static", StringComparison.OrdinalIgnoreCase) Then Continue For
                    GBActiveTasks.Controls.RemoveAt(i)
                Next
            End If

            ' (Optional) Add simple info label without external links
            If overdueSessions.Count > 0 Then
                Dim lbl As New Label() With {
                    .AutoSize = True,
                    .Location = New Point(8, 8),
                    .ForeColor = Color.Red,
                    .Font = New Font(Font, FontStyle.Bold),
                    .Text = $"ვადაგადაცილებული სესიები: {overdueSessions.Count} (დღეს: {todayOverdueSessions})"
                }
                GBActiveTasks.Controls.Add(lbl)
            End If

            Dim birthdays = dataService.GetUpcomingBirthdays(30)
            PopulateUpcomingBirthdays(birthdays)

            Dim allSessions = dataService.GetAllSessions() ' (Not used directly here but retained for potential future use)
            PopulateOverdueSessions(overdueSessions, IsAuthorizedUser, UserRoleValue)

            Debug.WriteLine("UC_Home.RefreshAllData: ყველა მონაცემი წარმატებით განახლდა")
        Catch ex As Exception
            Debug.WriteLine($"UC_Home.RefreshAllData: შეცდომა - {ex.Message}")
            Debug.WriteLine($"UC_Home.RefreshAllData: StackTrace - {ex.StackTrace}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' ტესტური მეთოდი კონტროლების ხილვადობის დიაგნოსტიკისთვის
    ''' </summary>
    Public Sub TestDirectControls()
        Try
            Debug.WriteLine("UC_Home.TestDirectControls: დაიწყო ქალაქების ტესტირება")

            ' კონტროლების მდგომარეობის შემოწოცი
            Debug.WriteLine($"GBRedTasks ხილვადობა: {GBRedTasks.Visible}, " &
                          $"ზომები: {GBRedTasks.Width}x{GBRedTasks.Height}, " &
                          $"კონტროლის რაოდენობა: {GBRedTasks.Controls.Count}")

            ' პაგინაციის კონტროლების მწარმოებელი
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
    ''' BtnAddAray ღილაკზე დასაჭერი - новая форма ჩანაწერის გამოჩენა
    ''' შეზღუდვით, რომ მხოლოდ ერთი ფორმა გაიხსნას
    ''' </summary>
    Private Sub BtnAddAray_Click(sender As Object, e As EventArgs) Handles BtnAddAray.Click
        Debug.WriteLine("UC_Home.BtnAddAray_Click: ახალი ჩანაწერის დამატების მოთხოვნა")

        Try
            ' შევამოწმოთ უკვე გახსნილი თუ არა NewRecordForm
            For Each frm As Form In Application.OpenForms
                If TypeOf frm Is NewRecordForm Then
                    ' თუ უკვე გახსნილია, მოგვიტანეთ წინ და გამოსვლის მეთოდიდან
                    Debug.WriteLine("UC_Home.BtnAddAray_Click: NewRecordForm უკვე გახსნილია, ფოკუსის გადატანა")
                    frm.Focus()
                    Return
                End If
            Next

            ' გავუსწოროთ Cursor-ი მოლოდინზე
            Cursor = Cursors.WaitCursor

            ' შევამოწმოთ máme თუ არა dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიშნულება", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Debug.WriteLine("UC_Home.BtnAddAray_Click: dataService არ არის ინიშნულება")
                Return
            End If

            ' მომხმარებლის email მიღება
            Dim userEmail As String = "უცნობი"
            ' მოიძიოთ მთავარი ფორმა და წამყვანი მეტოდების გამოყენებით
            Dim mainForm = TryCast(Application.OpenForms("Form1"), Form1)
            If mainForm IsNot Nothing AndAlso mainForm.GetType().GetMethod("GetUserEmail") IsNot Nothing Then
                ' თუ GetUserEmail მეთოდი existe, გამოვიყენოთ
                userEmail = CType(mainForm.GetType().GetMethod("GetUserEmail").Invoke(mainForm, Nothing), String)
            End If

            ' ნაგულისხმევად "სესია" ტიპი
            Dim recordType As String = "სესია"

            ' NewRecordForm-ის გაწვდვა Add რეჟიმში
            Dim newRecordForm As New NewRecordForm(dataService, recordType, userEmail, "UC_Home")
            newRecordForm.Show()

        Catch ex As Exception
            Debug.WriteLine($"UC_Home.BtnAddAray_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ახალი ჩანაწer's ფორმის გახს面的 შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Cursor-ების აღდგენა
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub UpdatePaginationControls()
        LPage.Text = $"{CurrentPage + 1}/{Math.Max(1, TotalPages)}"
        BtnPrev.Enabled = CurrentPage > 0
        BtnNext.Enabled = CurrentPage < TotalPages - 1
    End Sub

    Public Sub PopulateOverdueSessions(sessions As List(Of SessionModel), isAuthorized As Boolean, userRole As String)
        IsAuthorizedUser = isAuthorized
        UserRoleValue = userRole
        AllSessions = If(sessions, New List(Of SessionModel))
        InitializePerformanceResources()
        CalculateCardsPerPage()
        CurrentPage = 0
        ShowCurrentPageCards()
        UpdatePaginationControls()
    End Sub

    ' Custom colors for overdue cards
    Private ReadOnly OverdueCardPink As Color = Color.FromArgb(255, 220, 235) ' ვარდისფერი ფონისთვის
    Private ReadOnly OverdueHeaderPink As Color = Color.FromArgb(200, 60, 120)

    ' Update / add helper to rebuild context menu
    Private Sub BuildOverdueContextMenu(card As Panel, session As SessionModel)
        Dim cm As New ContextMenuStrip()
        ' Details
        Dim miDetails As New ToolStripMenuItem("დეტალები") With {.Tag = session.Id}
        AddHandler miDetails.Click, Sub(senderObj, args)
                                        Dim sid = CInt(DirectCast(senderObj, ToolStripMenuItem).Tag)
                                        Dim s = FindSessionById(sid)
                                        If s IsNot Nothing Then ShowSessionDetails(s)
                                    End Sub
        cm.Items.Add(miDetails)

        ' Status submenu
        Dim miStatusRoot As New ToolStripMenuItem("სტატუსის ცვლილება") With {.Tag = session.Id}
        Dim statuses As String() = {"შესრულებული", "დაგეგმილი", "გაუქმებული", "გაცდენა არასაპატიო", "გაცდენა საპატიო", "აღდგენა"}
        For Each st In statuses
            Dim miSt As New ToolStripMenuItem(st) With {.Tag = New With {Key .Id = session.Id, Key .Status = st}}
            AddHandler miSt.Click, Sub(senderObj, args)
                                       Dim payload = DirectCast(DirectCast(senderObj, ToolStripMenuItem).Tag, Object)
                                       ChangeSessionStatus(CInt(payload.Id), CStr(payload.Status))
                                   End Sub
            miStatusRoot.DropDownItems.Add(miSt)
        Next
        cm.Items.Add(miStatusRoot)

        If IsAuthorizedUser AndAlso (UserRoleValue = Roles.Admin OrElse UserRoleValue = Roles.Manager OrElse UserRoleValue = Roles.Staff) Then
            Dim miEdit As New ToolStripMenuItem("რედაქტირება") With {.Tag = session.Id}
            AddHandler miEdit.Click, Sub(senderObj, args)
                                         Dim sid = CInt(DirectCast(senderObj, ToolStripMenuItem).Tag)
                                         Dim tempBtn As New Button() With {.Tag = sid}
                                         BtnEditSession_Click(tempBtn, EventArgs.Empty)
                                     End Sub
            cm.Items.Add(miEdit)
        End If

        card.ContextMenuStrip = cm
    End Sub

    ' Helper to change status quickly (updates sheet and local model)
    Private Sub ChangeSessionStatus(sessionId As Integer, newStatus As String)
        If dataService Is Nothing Then
            MessageBox.Show("მონაცემთა სერვისი მიუწვდომელია", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Try
            ' Load raw sheet to find row (A2:O)
            Dim raw = dataService.GetData("DB-Schedule!A2:O")
            Dim targetRowIndex As Integer = -1 ' zero-based within raw
            For i = 0 To raw.Count - 1
                Dim row = raw(i)
                If row.Count > 0 AndAlso Integer.TryParse(row(0).ToString(), Nothing) Then
                    Dim idVal As Integer
                    If Integer.TryParse(row(0).ToString(), idVal) AndAlso idVal = sessionId Then
                        targetRowIndex = i
                        Exit For
                    End If
                End If
            Next
            If targetRowIndex = -1 Then
                MessageBox.Show("სესია ვერ მოიძებნა ცხრილში", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Dim sheetRowNumber = targetRowIndex + 2 ' account for header row
            ' Column M is 13th (A=1) so M{row}:M{row}
            dataService.UpdateData($"DB-Schedule!M{sheetRowNumber}:M{sheetRowNumber}", New List(Of Object) From {newStatus})

            ' Update local model & card
            Dim local = FindSessionById(sessionId)
            If local IsNot Nothing Then
                local.Status = newStatus
                ' If status change affects overdue logic (e.g., marking completed) maybe remove from list
                If Not local.IsOverdue Then
                    ' remove from overdue collection
                    AllSessions.Remove(local)
                End If
            End If
            ' Refresh cards visually
            ShowCurrentPageCards()
            UpdateOverdueStatistics()
        Catch ex As Exception
            Debug.WriteLine($"ChangeSessionStatus შეცდომა: {ex.Message}")
            MessageBox.Show("სტატუსის შეცვლის შეცდომა: " & ex.Message, "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class