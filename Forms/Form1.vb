' ===========================================
' 📄 Forms/Form1.vb (განახლებული MVVM-თვის და სერვისების დამოკიდებულების ინვერსიისთვის)
' -------------------------------------------
' მიზანი: UI გამიჯნული ViewModel-ზე, დამოკიდებულების ინვერსიით
' ===========================================
Imports System.IO
Imports System.ComponentModel
Imports Google.Apis.Auth.OAuth2
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services
Imports System.Text
Imports System.Threading.Tasks

Public Class Form1

    ' ViewModel-ები
    Private viewModel As MainViewModel
    Private homeViewModel As HomeViewModel

    ' სერვისები
    Private authService As GoogleOAuthService
    Private dataService As IDataService

    ' UI მენეჯერი და რეიუზადი კონტროლები
    Private menuMgr As MenuManager
    Private homeControl As UC_Home

    ' კონფიგურაცია
    Private ReadOnly spreadsheetId As String = "1SrBc4vLKPui6467aNmF5Hw-WZEd7dfGhkeFjfcnUqog"
    Private ReadOnly utilsFolder As String = Path.Combine(Application.StartupPath, "Utils")
    Private ReadOnly serviceAccountKeyPath As String = Path.Combine(utilsFolder, "google-service-account-key8_7a.json")
    Private ReadOnly secretsFile As String = Path.Combine(utilsFolder, "client_secret_v8_7.json")
    Private ReadOnly tokenStorePath As String = Path.Combine(utilsFolder, "TokenStore")

    'მეილი გავაპაბლიკოთ
    Public Function GetUserEmail() As String
        Return If(viewModel?.Email, "უცნობი")
    End Function

    ''' <summary>
    ''' კონსტრუქტორი - განახლებული ივენტების მიმსმენები ჩართულია
    ''' </summary>
    Public Sub New()
        InitializeComponent()

        ' საქაღალდეების შექმნა თუ არ არსებობს
        If Not Directory.Exists(utilsFolder) Then Directory.CreateDirectory(utilsFolder)
        If Not Directory.Exists(tokenStorePath) Then Directory.CreateDirectory(tokenStorePath)

        ' მენიუს მენეჯერის ინიციალიზაცია
        menuMgr = New MenuManager(mainMenu)

        ' ✨ ივენთების მსმენელების დამატება - განახლებული!
        AddHandler menuMgr.ScheduleMenuSelected, AddressOf OnScheduleMenuSelected
        AddHandler menuMgr.BeneficiaryReportSelected, AddressOf OnBeneficiaryReportSelected
        AddHandler menuMgr.TherapistReportSelected, AddressOf OnTherapistReportSelected

        ' შენიშვნა: homeControl აღარ იქმნება აქ, ვიცით ჯერ ViewModel/DataService неактивен.
    End Sub

    ''' <summary>
    ''' ViewModel-ის PropertyChanged ივენთის დამუშავება
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OnScheduleMenuSelected(sender As Object, e As EventArgs)
        Debug.WriteLine("OnScheduleMenuSelected: 'განრიგი' არჩეულია მენიუში")
        ShowSchedule()
    End Sub

    ''' <summary>
    ''' Form Load: ინიციალიზაცია და Google Sheets სერვისთან დაკავშირება
    ''' </summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' მენიუს საწყისი მდგომარეობა - მხოლოდ საწყისი
        menuMgr.ShowOnlyHomeMenu()

        ' ViewModel-ების ინიციალიზაცია სანამ UI კონტროლს შევქმნით
        viewModel = New MainViewModel()
        homeViewModel = New HomeViewModel()

        ' ივენთის ჰენდლერის დამატება
        AddHandler viewModel.PropertyChanged, AddressOf OnViewModelPropertyChanged

        ' GoogleServiceAccountClient-ის ინიციალიზაცია და მონაცემების სერვისის შექმნა
        Try
            ' შეამოწმე, ხარ უკართულად ხარ თუ არა სერვის აკაუნტის ფაილი
            If Not File.Exists(serviceAccountKeyPath) Then
                MessageBox.Show($"სერვის აკაუნტის JSON ფაილი ვერ მოიძებნა: {serviceAccountKeyPath}",
                              "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                ' შევქმნათ მონაცემთა სერვისი სერვის აკაუნტის გამოყენებით
                dataService = New SheetDataService(serviceAccountKeyPath, spreadsheetId)
                Debug.WriteLine("Form1_Load: მონაცემთა სერვისი წარმატებით ინიციალიზებულია")
            End If

            ' OAuth სერვისის ინიციალიზაცია მომხმარებლის ავტორიზაციისთვის
            authService = New GoogleOAuthService(secretsFile, tokenStorePath)

        Catch ex As Exception
            MessageBox.Show($"შეცდომა სერვისების ინიციალიზაციისსათ: {ex.Message}",
                          "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' HomeControl შექმნა ერთხელ (რეიუზი)
        EnsureHomeControl()
        ShowHome()

        ' UI-ის საწყისი ინსტრუქციები ViewModel-იდან
        LUser.Text = If(String.IsNullOrEmpty(viewModel.Email), "გთხოვთ გაიაროთ ავტორიზაცია", viewModel.Email)
        BtnLogin.Text = If(viewModel.IsAuthorized, "გასვლა", "ავტორიზაცია")
        'ბექგრაუნდის სურათის ჩატვირთვა
        Dim imagePath As String = Path.Combine(Application.StartupPath, "Resources", "AppImages", "bg1.jpg")
        If File.Exists(imagePath) Then
            Me.BackgroundImage = Image.FromFile(imagePath)
            Me.BackgroundImageLayout = ImageLayout.Stretch
        End If
        ' დავამატოთ მისალმება homeViewModel-ში
        homeViewModel.Greeting = homeViewModel.GetGreetingByTime()

    End Sub

    Private Sub EnsureHomeControl()
        If homeControl Is Nothing OrElse homeControl.IsDisposed Then
            homeControl = New UC_Home(homeViewModel)
            homeControl.Dock = DockStyle.Fill
            If dataService IsNot Nothing Then homeControl.SetDataService(dataService)
        End If
    End Sub

    ''' <summary>
    ''' BtnLogin Click: მხოლოდ მომხმარებლის ავტორიზაცია მისი ვინაობის დასადგენად
    ''' </summary>
    Private Async Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        ' დავბლოკოთ ღილაკი, რომ თავიდან ავირიდოს მრავალჯერადი დაჭერა
        BtnLogin.Enabled = False

        Try
            If Not viewModel.IsAuthorized Then
                Try
                    ' 1) Google OAuth ავტორიზაცია ᲛᲮᲝᲚᲝᲓ მომხმარებლის ინფორმაციისთვის
                    Dim credential = Await authService.AuthorizeAsync(New String() {
                    Google.Apis.Oauth2.v2.Oauth2Service.Scope.UserinfoEmail,
                    Google.Apis.Oauth2.v2.Oauth2Service.Scope.UserinfoProfile
                })

                    ' 2) OAuth სერვისის შექმნა მომხმარებლის ინფორმაციას მისიაღებად
                    Dim oauthService = New Google.Apis.Oauth2.v2.Oauth2Service(
                    New Google.Apis.Services.BaseClientService.Initializer() With {
                        .HttpClientInitializer = credential,
                        .ApplicationName = "Scheduler_v8.8a"
                    })

                    ' 3) მომხმარებლის ინფორმაციის მიღება
                    Dim userInfo = Await oauthService.Userinfo.Get().ExecuteAsync()
                    Dim email = userInfo.Email
                    Dim fName = GetFirstName(userInfo.Name)
                    ' 4) მომხმარებლის როლის მიიღება dataService-დან (რომელიც იყენებს სერვის აკაუნტს)
                    Dim role = dataService.GetOrCreateUserRole(email)

                    ' 5) ViewModel განახლება
                    viewModel.Email = email
                    viewModel.Role = role
                    viewModel.IsAuthorized = True

                    ' 6) მხოლოდ სახელის გამოყოფა
                    homeViewModel.UserName = If(String.IsNullOrEmpty(fName), email, fName)

                    ' 7) UI და მენიუს განახლება
                    BtnLogin.Text = "გასვლა"
                    LUser.Text = email
                    menuMgr.ShowMenuByRole(role)

                    ' 8) Home გვერდის ჩვენება და მონაცემების ჩატვირთვა
                    EnsureHomeControl()
                    homeControl.UpdateUserName(homeViewModel.UserName)
                    Await ReloadHomeDataAsync()
                Catch ex As Exception
                    MessageBox.Show($"ავტორიზაცია ვერ შესრულდა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Else
                Try
                    ' გასვლის ლოგიკა
                    Await authService.RevokeAsync()
                    viewModel.IsAuthorized = False
                    viewModel.Email = String.Empty
                    viewModel.Role = String.Empty
                    homeViewModel.UserName = String.Empty

                    ' UI განახლება
                    BtnLogin.Text = "ავტორიზაცია"
                    LUser.Text = "გთხოვთ გაიაროთ ავტორიზაცია"

                    ' მენიუს განახლება
                    menuMgr.ShowOnlyHomeMenu()

                    ' Home გვერდის განახლება
                    ShowHome()
                Catch ex As Exception
                    MessageBox.Show($"გასვლა ვერ განხორციელდა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Finally
            ' გავააქტიუროთ ღილაკი ისევ
            BtnLogin.Enabled = True
        End Try
    End Sub
    ''' <summary>
    ''' სრული სახელიდან პირველი სახელის გამორჩევა
    ''' </summary>
    ''' <param name="fullName">სრული სახელი (სახელი და გვარი)</param>
    ''' <returns>მხოლოდ პირველი სახელის</returns>
    Private Function GetFirstName(fullName As String) As String
        ' თუ ცარიელია, დავაბრუნოთ ცარიელი სტრიქონი
        If String.IsNullOrEmpty(fullName) Then
            Return String.Empty
        End If

        ' გამოვყოთ სიტყვები სახელიდან (გამოყოფილი ჰარით)
        Dim nameParts As String() = fullName.Trim().Split(" "c)

        ' დავაბრუნოთ პირველი სიტყვა, რომელიც უნდა იყოს სახელი
        If nameParts.Length > 0 Then
            Return nameParts(0)
        End If

        ' თუ ვერ დავშალეთ, დავაბრუნოთ სრული სახელი
        Return fullName
    End Function

    ''' <summary>
    ''' ინსტრუმენტების ხილვადობის მართვა ავტორიზაციის სტატუსისა და როლის მიხედვით
    ''' </summary>
    ''' <param name="isAuthorized">არის თუ არა მომხმარებელი ავტორიზებული</param>
    ''' <param name="role">მომხმარებლის როლი (არასავალდებულო)</param>
    ''' <param name="additionalParam">დამატებითი პარამეტრი (არასავალდებულო)</param>
    Private Sub SetToolsVisibility(isAuthorized As Boolean, Optional role As String = "", Optional additionalParam As Object = Nothing)
        Try
            ' თუ homeControl არ არის ინიციალიზებული, გამოვიდეთ მეთოდიდან
            If homeControl Is Nothing OrElse homeControl.IsDisposed Then
                Return
            End If

            ' როლის მნიშვნელობის დადგენა
            Dim userRole As String = If(String.IsNullOrEmpty(role), viewModel.Role, role)

            ' ვიპოვოთ GBTools პანელი, თუ არსებობს
            Dim toolsPanel = TryCast(homeControl.Controls.Find("GBTools", True).FirstOrDefault(), GroupBox)
            If toolsPanel IsNot Nothing Then
                ' შევცვალოთ როგორც Enabled, ისე Visible თვისებები
                toolsPanel.Enabled = isAuthorized
                toolsPanel.Visible = isAuthorized

                ' ცალკეული ღილაკების ხილვადობის მართვა
                ' მაგალითად, BtnAddAray ხილვადია მხოლოდ ადმინისტრაციული (1) და მენეჯერი (2) როლებისთვის
                Dim addButton = TryCast(homeControl.Controls.Find("BtnAddAray", True).FirstOrDefault(), Button)
                If addButton IsNot Nothing Then
                    addButton.Visible = isAuthorized AndAlso (userRole = "1" OrElse userRole = "2")
                End If

                ' BtnRefresh ღილაკი ხილვადია ყველა ავტორიზებულისთვის
                Dim refreshButton = TryCast(homeControl.Controls.Find("BtnRefresh", True).FirstOrDefault(), Button)
                If refreshButton IsNot Nothing Then
                    refreshButton.Visible = isAuthorized
                End If
            End If

            ' უფრო დაწვრილებითი დებაგირებისთვის დავამატოთ ჩანაწერი
            Debug.WriteLine($"SetToolsVisibility გამოძახებულია: isAuthorized={isAuthorized}, role={userRole}")
            If toolsPanel IsNot Nothing Then
                Debug.WriteLine($"ინსტრუმენტების პანელი ნაპოვნია: Enabled={toolsPanel.Enabled}, Visible={toolsPanel.Visible}")
            Else
                Debug.WriteLine("ინსტრუმენტების პანელი ვერ მოიძებნა!")
            End If
        Catch ex As Exception
            ' ვიჭერთ და ვაღრიცხავთ ნებისმიერ შეცდომას
            Debug.WriteLine($"შეცდომა ინსტრუმენტების პანელის მართვისას: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' PropertyChanged Handler: UI და მენიუს განახლება ViewModel-დან
    ''' </summary>
    Private Sub OnViewModelPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
        ' თავიდან ავიცილოთ UI-თრედის დაბლოკვა
        If Me.InvokeRequired Then
            Me.Invoke(Sub() OnViewModelPropertyChanged(sender, e))
            Return
        End If

        ' პრევენცია ციკლური განახლებებისთვის
        Static isUpdating As Boolean = False
        If isUpdating Then Return

        isUpdating = True
        Try
            Debug.WriteLine($"OnViewModelPropertyChanged: პროპტი {e.PropertyName} შეიცვალა")

            Select Case e.PropertyName
                Case NameOf(viewModel.Email)
                    LUser.Text = If(String.IsNullOrEmpty(viewModel.Email),
                               "გთხოვთ გავიაროთ ავტორიზაცია",
                               viewModel.Email)
                    Debug.WriteLine($"OnViewModelPropertyChanged: LUser.Text განახლდება: {LUser.Text}")

                Case NameOf(viewModel.IsAuthorized)
                    BtnLogin.Text = If(viewModel.IsAuthorized, "გასვლა", "ავტორიზაცია")
                    Debug.WriteLine($"OnViewModelPropertyChanged: BtnLogin.Text განახლდება: {BtnLogin.Text}")

                    SetToolsVisibility(viewModel.IsAuthorized)
                    Debug.WriteLine($"OnViewModelPropertyChanged: ინსტრუმენტების ხილვადობა განახლდება")

                    If Not viewModel.IsAuthorized Then
                        menuMgr.ShowOnlyHomeMenu()
                        Debug.WriteLine("OnViewModelPropertyChanged: მენიუ განახლდა (მხოლოდ საწყისი)")
                    End If

                Case NameOf(viewModel.Role)
                    If String.IsNullOrEmpty(viewModel.Role) Then
                        menuMgr.ShowOnlyHomeMenu()
                        Debug.WriteLine("OnViewModelPropertyChanged: მენიუ განახლდა (მხოლოდ საწყისი, ცარიელი როლი)")
                    Else
                        menuMgr.ShowMenuByRole(viewModel.Role)
                        Debug.WriteLine($"OnViewModelPropertyChanged: მენიუ განახლდა როლით: {viewModel.Role}")
                    End If
            End Select
        Catch ex As Exception
            Debug.WriteLine($"OnViewModelPropertyChanged: შეცდომა - {ex.Message}")
        Finally
            isUpdating = False
        End Try
    End Sub
    ''' <summary>
    ''' მთავარი გვერდის ჩვენება UC_Home-ის გამოყენებით
    ''' </summary>
    Private Sub ShowHome()
        EnsureHomeControl()
        ShowUserControl(Function()
                            Return homeControl
                        End Function)
        SetToolsVisibility(viewModel.IsAuthorized, viewModel.Role)
    End Sub

    ''' <summary>
    ''' კალენდრის გვერდის ჩვენება
    ''' </summary>
    Private Sub ShowCalendar()
        ShowUserControl(Function()
                            Dim vm As New CalendarViewModel()
                            Dim c As New UC_Calendar(vm)
                            If dataService IsNot Nothing Then c.SetDataService(dataService)
                            c.SetUserEmail(If(viewModel?.Email, "უცნობი"))
                            Return c
                        End Function)
    End Sub

    ''' <summary>
    ''' განრიგის გვერდის ჩვენება - გაუმჯობესებული ვერსია
    ''' მომხმარებლის informacionიის სწორი გადაცემით
    ''' </summary>
    Private Sub ShowSchedule()
        ShowUserControl(Function()
                            Dim c As New UC_Schedule()
                            If dataService IsNot Nothing Then c.SetDataService(dataService)
                            Dim userEmail As String = If(String.IsNullOrEmpty(viewModel?.Email), "user@example.com", viewModel.Email)
                            Dim userRole As String = If(String.IsNullOrEmpty(viewModel?.Role), "6", viewModel.Role)
                            c.SetUserInfo(userEmail, userRole)
                            Return c
                        End Function)
    End Sub

    ''' <summary>
    ''' ✨ ბენეფიციარის ანგარიშის მენიუს არჩევისას გამოძახებული მეთოდი
    ''' ხელმისაწვდომია როლებისთვის: 1, 2, 3
    ''' </summary>
    Private Sub OnBeneficiaryReportSelected(sender As Object, e As EventArgs)
        Debug.WriteLine("OnBeneficiaryReportSelected: 'ბენეფიციარის ანგარიში' არჩეულია მენიუში")

        ' როლის შემოწმება
        If Not IsAccessAllowed({"1", "2", "3"}) Then
            MessageBox.Show("თქვენ არ გაქვთ ბენეფიციარის ანგარიშზე წვდომის უფლება", "წვდომა შეზღუდული არის",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ShowBeneficiaryReport()
    End Sub

    ''' <summary>
    ''' ✨ თერაპევტის ანგარიშის მენიუს არჩევისას გამოძახებული მეთოდი
    ''' ხელმისაწვდომია მხოლოდ როლებისთვის: 1, 2
    ''' </summary>
    Private Sub OnTherapistReportSelected(sender As Object, e As EventArgs)
        Debug.WriteLine("OnTherapistReportSelected: 'თერაპევტის ანგარიში' არჩეულია მენიუში")

        ' როლის შემოწმება
        If Not IsAccessAllowed({"1", "2"}) Then
            MessageBox.Show("თქვენ არ გაქვთ თერაპევტის ანგარიშზე წვდომის უფლება", "წვდომა შეზღუდული არის",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ShowTherapistReport()
    End Sub

    ''' <summary>
    ''' ✨ წვდომის შემოწმება - νέα დამხმარე მეთოდი
    ''' </summary>
    ''' <param name="allowedRoles">ნებადართული როლების მასივი</param>
    ''' <returns>True თუ წვდომა ნებადართულია</returns>
    Private Function IsAccessAllowed(allowedRoles As String()) As Boolean
        Try
            ' ავტორიზაციის შემოწმება
            If Not viewModel.IsAuthorized Then
                Debug.WriteLine("IsAccessAllowed: მომხმარებელი არ არის ავტორიზებული")
                Return False
            End If

            ' როლის შემოწმება
            Dim userRole As String = viewModel.Role
            If String.IsNullOrEmpty(userRole) Then
                Debug.WriteLine("IsAccessAllowed: მომხმარებლის როლი არ არის მითითებული")
                Return False
            End If

            ' შემოწმება ნებადართულ როლებში
            Dim hasAccess As Boolean = allowedRoles.Contains(userRole.Trim())
            Debug.WriteLine($"IsAccessAllowed: მომხმარებლის როლი '{userRole}', წვდომა: {hasAccess}")

            Return hasAccess

        Catch ex As Exception
            Debug.WriteLine($"IsAccessAllowed: შეცდომა - {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ✨ ბენეფიციარის ანგარიშის გვერდის_show_authorization_failure
    ''' </summary>
    Private Sub ShowBeneficiaryReport()
        ShowUserControl(Function()
                            Dim c As New UC_BeneficiaryReport()
                            If dataService IsNot Nothing Then c.SetDataService(dataService)
                            Dim userEmail As String = If(String.IsNullOrEmpty(viewModel?.Email), "user@example.com", viewModel.Email)
                            Dim userRole As String = If(String.IsNullOrEmpty(viewModel?.Role), "6", viewModel.Role)
                            c.SetUserInfo(userEmail, userRole)
                            Return c
                        End Function)
    End Sub

    ''' <summary>
    ''' ✨ თერაპევტის ანგარიშის გვერდის_show_authorization_failure
    ''' </summary>
    Private Sub ShowTherapistReport()
        ShowUserControl(Function()
                            Dim c As New UC_TherapistReport()
                            If dataService IsNot Nothing Then c.SetDataService(dataService)
                            Dim userEmail As String = If(String.IsNullOrEmpty(viewModel?.Email), "user@example.com", viewModel.Email)
                            Dim userRole As String = If(String.IsNullOrEmpty(viewModel?.Role), "6", viewModel.Role)
                            c.SetUserInfo(userEmail, userRole)
                            Return c
                        End Function)
    End Sub
    ' Form1.vb ფაილში დაამატეთ ეს ღილაკზე დაჭერის ფუნქცია
    Private Sub TestBirthdaysDirect()
        Try
            Debug.WriteLine("TestBirthdaysDirect: დაწყება")

            ' სწორად DB-Personal ცხრილიდან მონაცემები
            Dim personalData As IList(Of IList(Of Object)) = dataService.GetData("DB-Personal!B2:E")
            Debug.WriteLine($"TestBirthdaysDirect: მიღებულია {If(personalData Is Nothing, 0, personalData.Count)} მწკრივი")

            ' გამოვიტანოთ ყველა დაბადების თარიხი დებაგირებისთვის
            If personalData IsNot Nothing Then
                For i As Integer = 0 To personalData.Count - 1
                    Dim row = personalData(i)
                    If row.Count >= 3 AndAlso row(2) IsNot Nothing Then
                        Debug.WriteLine($"TestBirthdaysDirect: მწკრივი {i + 2}, სახელი={row(0)}, გვარი={row(1)}, დაბადების თარიხი={row(2)}")
                    End If
                Next
            End If

            Debug.WriteLine("TestBirthdaysDirect: დასრულება")
        Catch ex As Exception
            Debug.WriteLine($"TestBirthdaysDirect: შეცდომა - {ex.Message}")
        End Try
    End Sub

    Private Sub pnlMain_Paint(sender As Object, e As PaintEventArgs) Handles pnlMain.Paint

    End Sub

    ' Helper methods added (if missing after refactor)
    Private Async Function ReloadHomeDataAsync() As Task
        Await LoadHomeDataAsyncInternal()
    End Function

    Private Async Function LoadHomeDataAsyncInternal() As Task
        If dataService Is Nothing Then Return
        Try
            Dim pendingTask = dataService.GetPendingSessionsAsync()
            Dim overdueTask = dataService.GetOverdueSessionsAsync()
            Dim allTask = dataService.GetAllSessionsAsync()
            Dim tasksTask = dataService.GetActiveTasksAsync()
            Dim birthdaysTask = dataService.GetUpcomingBirthdaysAsync(7)
            Await Task.WhenAll(pendingTask, overdueTask, allTask, tasksTask, birthdaysTask)
            Dim today = DateTime.Today
            Dim todaySessions = (Await allTask).Where(Function(s) s.DateTime.Date = today).ToList()
            Me.Invoke(Sub()
                          homeViewModel.PendingSessions.Clear()
                          For Each s In pendingTask.Result : homeViewModel.PendingSessions.Add(s) : Next
                          homeViewModel.PendingSessionsCount = pendingTask.Result.Count
                          homeViewModel.OverdueSessions.Clear()
                          For Each s In overdueTask.Result : homeViewModel.OverdueSessions.Add(s) : Next
                          homeViewModel.UpcomingBirthdays.Clear()
                          For Each b In birthdaysTask.Result : homeViewModel.UpcomingBirthdays.Add(b) : Next
                          homeViewModel.ActiveTasks.Clear()
                          For Each t In tasksTask.Result : homeViewModel.ActiveTasks.Add(t) : Next
                          If homeControl IsNot Nothing AndAlso Not homeControl.IsDisposed Then
                              ' Removed UpdateTodaySessionsStatistics – home control now recalculates internally on data bind
                              homeControl.PopulateOverdueSessions(overdueTask.Result, viewModel.IsAuthorized, viewModel.Role)
                              homeControl.PopulateUpcomingBirthdays(birthdaysTask.Result)
                          End If
                      End Sub)
        Catch ex As Exception
            Debug.WriteLine($"LoadHomeDataAsyncInternal Async შეცდომა: {ex.Message}")
        End Try
    End Function

    Private Sub ShowUserControl(factory As Func(Of UserControl))
        pnlMain.SuspendLayout()
        Try
            pnlMain.Controls.Clear()
            Dim ctrl = factory()
            If ctrl IsNot Nothing Then
                ctrl.Dock = DockStyle.Fill
                pnlMain.Controls.Add(ctrl)
                ctrl.BringToFront()
            End If
        Finally
            pnlMain.ResumeLayout()
        End Try
    End Sub

    ' MenuStrip ItemClicked handler (reintroduced after refactor so that calendar opens)
    Private Sub mainMenu_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles mainMenu.ItemClicked
        Try
            Dim txt = e.ClickedItem.Text.Trim()
            Select Case txt
                Case "საწყისი"
                    ShowHome()
                Case "კალენდარი"
                    ShowCalendar()
                Case "განრიგი"
                    ShowSchedule()
                Case "ბენეფიციარის ანგარიში"
                    OnBeneficiaryReportSelected(Me, EventArgs.Empty)
                Case "თერაპევტის ანგარიში"
                    OnTherapistReportSelected(Me, EventArgs.Empty)
                    ' სხვა დროებითი მენიუს პუნქტების საჭიროებისამებრ დამატება
            End Select
        Catch ex As Exception
            Debug.WriteLine($"mainMenu_ItemClicked შეცდომა: {ex.Message}")
        End Try
    End Sub
End Class