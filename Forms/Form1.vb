﻿' ===========================================
' 📄 Forms/Form1.vb (განახლებული MVVM-თვის და სერვისების დამოკიდებულების ინვერსიისთვის)
' -------------------------------------------
' მიზანი: UI გამიჯნული ViewModel-ზე, დამოკიდებულების ინვერსიით
' ===========================================
Imports System.IO
Imports System.ComponentModel
Imports Google.Apis.Auth.OAuth2
'Imports Scheduler_v8_8a.Services
'Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services
Imports System.Text

Public Class Form1

    ' ViewModel-ები - აღარ გამოვაცხადოთ როგორც ReadOnly
    Private viewModel As MainViewModel
    Private homeViewModel As HomeViewModel

    ' სერვისები - აღარ გამოვაცხადოთ როგორც ReadOnly
    Private authService As GoogleOAuthService
    Private dataService As IDataService

    ' UI კომპონენტები - აღარ გამოვაცხადოთ როგორც ReadOnly
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
    ''' კონსტრუქტორი - განახლებული ივენთების მიმსმენები ჩართებულია
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

        ' UC_Home-ის შექმნისას SheetDataService-ის მიბმა
        homeControl = New UC_Home(homeViewModel)
        homeControl.Dock = DockStyle.Fill
        homeControl.SetDataService(dataService)
        pnlMain.Controls.Add(homeControl)
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

        ' ViewModel-ების ინიციალიზაცია
        viewModel = New MainViewModel()
        homeViewModel = New HomeViewModel()

        ' ივენთის ჰენდლერის დამატება
        AddHandler viewModel.PropertyChanged, AddressOf OnViewModelPropertyChanged

        ' GoogleServiceAccountClient-ის ინიციალიზაცია და მონაცემების სერვისის შექმნა
        Try
            ' შევამოწმოთ არსებობს თუ არა სერვის აკაუნტის ფაილი
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
            MessageBox.Show($"შეცდომა სერვისების ინიციალიზაციისას: {ex.Message}",
                          "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' საწყისი Home View
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

    ''' <summary>
    ''' BtnLogin Click: მხოლოდ მომხმარებლის ავტორიზაცია მისი ვინაობის დასადგენად
    ''' </summary>
    Private Async Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        ' დავბლოკოთ ღილაკი, რომ თავიდან ავირიდოთ მრავალჯერადი დაჭერა
        BtnLogin.Enabled = False

        Try
            If Not viewModel.IsAuthorized Then
                Try
                    ' 1) Google OAuth ავტორიზაცია ᲛᲮᲝᲚᲝᲓ მომხმარებლის ინფორმაციისთვის
                    Dim credential = Await authService.AuthorizeAsync(New String() {
                    Google.Apis.Oauth2.v2.Oauth2Service.Scope.UserinfoEmail,
                    Google.Apis.Oauth2.v2.Oauth2Service.Scope.UserinfoProfile
                })

                    ' 2) OAuth სერვისის შექმნა მომხმარებლის ინფორმაციის მისაღებად
                    Dim oauthService = New Google.Apis.Oauth2.v2.Oauth2Service(
                    New Google.Apis.Services.BaseClientService.Initializer() With {
                        .HttpClientInitializer = credential,
                        .ApplicationName = "Scheduler_v8.8a"
                    })

                    ' 3) მომხმარებლის ინფორმაციის მიღება
                    Dim userInfo = Await oauthService.Userinfo.Get().ExecuteAsync()
                    Dim email = userInfo.Email
                    Dim name = userInfo.Name

                    ' 4) მომხმარებლის როლის მიღება dataService-დან (რომელიც იყენებს სერვის აკაუნტს)
                    Dim role = dataService.GetOrCreateUserRole(email)

                    ' 5) ViewModel განახლება
                    viewModel.Email = email
                    viewModel.Role = role
                    viewModel.IsAuthorized = True

                    ' 6) მხოლოდ სახელის გამოყოფა
                    Dim firstName = GetFirstName(name)
                    homeViewModel.UserName = If(String.IsNullOrEmpty(firstName), email, firstName)

                    ' 7) UI და მენიუს განახლება
                    BtnLogin.Text = "გასვლა"
                    LUser.Text = email
                    menuMgr.ShowMenuByRole(role)

                    ' 8) Home გვერდის ჩვენება და მონაცემების ჩატვირთვა
                    ShowHome()

                    ' 9) პირდაპირ განვაახლოთ მომხმარებლის სახელი (დავამატოთ ეს კოდი)
                    If homeControl IsNot Nothing AndAlso Not homeControl.IsDisposed Then
                        ' დავაყოვნოთ ცოტა, რომ დარწმუნებული ვიყოთ UI-ს აქვს დრო განახლებისთვის
                        Application.DoEvents()
                        System.Threading.Thread.Sleep(200)
                        ' გამოვიძახოთ UpdateUserName მეთოდი
                        homeControl.UpdateUserName(homeViewModel.UserName)
                        Debug.WriteLine($"BtnLogin_Click: მომხმარებლის სახელი განახლდა = '{homeViewModel.UserName}'")
                    End If
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
    ''' სრული სახელიდან პირველი სახელის გამოყოფა
    ''' </summary>
    ''' <param name="fullName">სრული სახელი (სახელი და გვარი)</param>
    ''' <returns>მხოლოდ პირველი სახელი</returns>
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
                ' შევცვალოთ როგორც Enabled, ასევე Visible თვისებები
                toolsPanel.Enabled = isAuthorized
                toolsPanel.Visible = isAuthorized

                ' ცალკეული ღილაკების ხილვადობის მართვა
                ' მაგალითად, BtnAddAray ხილვადია მხოლოდ ადმინისტრატორებისა (1) და მენეჯერებისთვის (2)
                Dim addButton = TryCast(homeControl.Controls.Find("BtnAddAray", True).FirstOrDefault(), Button)
                If addButton IsNot Nothing Then
                    addButton.Visible = isAuthorized AndAlso (userRole = "1" OrElse userRole = "2")
                End If

                ' BtnRefresh ღილაკი ხილვადია ყველასთვის, ვინც ავტორიზებულია
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
    ''' PropertyChanged Handler: UI და მენიუს განახლება ViewModel-იდან
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
            Debug.WriteLine($"OnViewModelPropertyChanged: პროპერთი {e.PropertyName} შეიცვალა")

            Select Case e.PropertyName
                Case NameOf(viewModel.Email)
                    LUser.Text = If(String.IsNullOrEmpty(viewModel.Email),
                               "გთხოვთ გაიაროთ ავტორიზაცია",
                               viewModel.Email)
                    Debug.WriteLine($"OnViewModelPropertyChanged: LUser.Text განახლდა: {LUser.Text}")

                Case NameOf(viewModel.IsAuthorized)
                    BtnLogin.Text = If(viewModel.IsAuthorized, "გასვლა", "ავტორიზაცია")
                    Debug.WriteLine($"OnViewModelPropertyChanged: BtnLogin.Text განახლდა: {BtnLogin.Text}")

                    SetToolsVisibility(viewModel.IsAuthorized)
                    Debug.WriteLine($"OnViewModelPropertyChanged: ინსტრუმენტების ხილვადობა განახლდა")

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
        Try
            Debug.WriteLine("ShowHome: დაიწყო - ამჟამინდელი UC_Home მდგომარეობა:")

            ' ამჟამინდელი home კონტროლის დიაგნოსტიკა
            If homeControl IsNot Nothing Then
                Debug.WriteLine($"ShowHome: არსებული homeControl - Disposed={homeControl.IsDisposed}, Visible={homeControl.Visible}")
            Else
                Debug.WriteLine("ShowHome: homeControl არის Nothing")
            End If

            ' შევამოწმოთ pnlMain
            Debug.WriteLine($"ShowHome: pnlMain - Controls={pnlMain.Controls.Count}, Visible={pnlMain.Visible}")

            ' პირველ რიგში გავასუფთავოთ მთავარი პანელი
            pnlMain.Controls.Clear()

            ' ამოცანა: UC_Home კონტროლის ჩვენება და განახლება
            ' ცვლილება: ყოველთვის შევქმნათ ახალი homeControl
            ' მიუხედავად მისი არსებული მდგომარეობისა
            Debug.WriteLine("ShowHome: ახალი homeControl-ის შექმნა")
            homeControl = New UC_Home(homeViewModel)
            homeControl.Dock = DockStyle.Fill
            pnlMain.Controls.Add(homeControl)
            Debug.WriteLine($"ShowHome: შეიქმნა ახალი homeControl - Disposed={homeControl.IsDisposed}, Visible={homeControl.Visible}")

            ' გავხადოთ homeControl წინა პლანზე
            homeControl.BringToFront()
            Debug.WriteLine($"ShowHome: homeControl.BringToFront() გამოძახებულია")

            ' შეგვიძლია ასევე გამოვიძახოთ Refresh
            homeControl.Refresh()
            Debug.WriteLine($"ShowHome: homeControl.Refresh() გამოძახებულია")

            ' ინსტრუმენტების ხილვადობის განახლება
            SetToolsVisibility(viewModel.IsAuthorized, viewModel.Role)

            ' დავამატოთ მონაცემთა სერვისის მითითება
            If dataService IsNot Nothing Then
                homeControl.SetDataService(dataService)
                Debug.WriteLine("ShowHome: მონაცემთა სერვისი გადაეცა homeControl-ს")
            Else
                Debug.WriteLine("ShowHome: dataService არის Nothing, homeControl-ს არ აქვს მონაცემთა წყარო")
            End If

            ' ყველა შემთხვევაში ვცდილობთ მონაცემების ჩატვირთვას
            If dataService IsNot Nothing Then
                Try
                    Debug.WriteLine("ShowHome: LoadHomeDataAsync() გამოძახება...")
                    LoadHomeDataAsync()
                Catch ex As Exception
                    Debug.WriteLine($"ShowHome: მონაცემების დატვირთვის შეცდომა: {ex.Message}")
                End Try
            Else
                Debug.WriteLine("ShowHome: dataService არის Nothing, მონაცემების დატვირთვა შეუძლებელია")
            End If

            ' ბოლოს, დავრწმუნდეთ რომ მომხმარებლის სახელი განახლებულია
            If homeControl IsNot Nothing AndAlso Not homeControl.IsDisposed AndAlso viewModel IsNot Nothing Then
                ' დავაყოვნოთ ცოტა, რომ დარწმუნებული ვიყოთ UI-ს აქვს დრო განახლებისთვის
                Application.DoEvents()

                ' გამოვიძახოთ UpdateUserName მეთოდი
                homeControl.UpdateUserName(homeViewModel.UserName)
                Debug.WriteLine($"ShowHome: მომხმარებლის სახელი განახლდა = '{homeViewModel.UserName}'")
            End If

            Debug.WriteLine("ShowHome: დასრულებულია")
        Catch ex As Exception
            Debug.WriteLine($"ShowHome: შეცდომა - {ex.Message}")
            Debug.WriteLine($"ShowHome: StackTrace - {ex.StackTrace}")
            MessageBox.Show($"მთავარი გვერდის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' HomeViewModel-ისთვის მონაცემების ასინქრონული დატვირთვა
    ''' </summary>
    Private Async Sub LoadHomeDataAsync()
        Try
            ' დავიცადოთ დატვირთვამდე
            Await Task.Delay(100)

            Debug.WriteLine("LoadHomeDataAsync: დაიწყო მონაცემების დატვირთვა")

            ' დატვირთვის ლოგიკა გავიტანოთ ცალკე თრედზე
            Await Task.Run(Sub()
                               Try
                                   ' მონაცემების წამოღება სხვა თრედზე
                                   Dim pendingSessions = dataService.GetPendingSessions()
                                   Debug.WriteLine($"LoadHomeDataAsync: წამოღებულია {pendingSessions.Count} მოლოდინში სესია")

                                   Dim overdueSessions = dataService.GetOverdueSessions()
                                   Debug.WriteLine($"LoadHomeDataAsync: წამოღებულია {overdueSessions.Count} ვადაგადაცილებული სესია")

                                   ' წამოვიღოთ ყველა სესია დღევანდელი სტატისტიკისთვის
                                   Dim allSessions = dataService.GetAllSessions()
                                   Debug.WriteLine($"LoadHomeDataAsync: წამოღებულია {allSessions.Count} სესია სულ")

                                   ' ვიპოვოთ დღევანდელი სესიები სტატისტიკისთვის - უფრო მკაცრი ფილტრაცია
                                   Dim currentDate As DateTime = DateTime.Today
                                   Dim todaySessions = New List(Of SessionModel)()

                                   ' ვამოწმებთ თითოეულ სესიას ცალ-ცალკე, მკაცრი შედარებით
                                   For Each session In allSessions
                                       Dim sessionDate = session.DateTime.Date
                                       Dim isSameDay = (sessionDate.Year = currentDate.Year) AndAlso
                    (sessionDate.Month = currentDate.Month) AndAlso
                    (sessionDate.Day = currentDate.Day)

                                       If isSameDay Then
                                           todaySessions.Add(session)
                                           Debug.WriteLine($"LoadHomeDataAsync: დამატებულია დღევანდელი სესია - ID={session.Id}, თარიღი={session.DateTime:dd.MM.yyyy}")
                                       End If
                                   Next

                                   Debug.WriteLine($"LoadHomeDataAsync: დღევანდელი სესიების რაოდენობა: {todaySessions.Count}")
                                   ' პირველი 3 ვადაგადაცილებული სესიის დებაგინგი
                                   For i As Integer = 0 To Math.Min(2, overdueSessions.Count - 1)
                                       Debug.WriteLine($"LoadHomeDataAsync: ვადაგადაცილებული სესია #{i + 1} - " &
                                              $"ID={overdueSessions(i).Id}, " &
                                              $"თარიღი={overdueSessions(i).DateTime:dd.MM.yyyy HH:mm}, " &
                                              $"სტატუსი='{overdueSessions(i).Status}'")
                                   Next

                                   Dim tasks = dataService.GetActiveTasks()
                                   Debug.WriteLine($"LoadHomeDataAsync: წამოღებულია {tasks.Count} აქტიური დავალება")

                                   ' დაბადების დღეების მონაცემების წამოღება
                                   Dim birthdays = dataService.GetUpcomingBirthdays(7) ' 7 დღე
                                   Debug.WriteLine($"LoadHomeDataAsync: წამოღებულია {birthdays.Count} მოახლოებული დაბადების დღე")

                                   ' თუ ბაზიდან არ მოვიდა დაბადების დღეები, შევქმნათ საცდელი მონაცემები
                                   If birthdays Is Nothing OrElse birthdays.Count = 0 Then
                                       birthdays = New List(Of BirthdayModel)()
                                       Debug.WriteLine("LoadHomeDataAsync: ბაზიდან არ მოვიდა დაბადების დღეები, ვქმნით საცდელ მონაცემებს")

                                       ' პირდაპირ ვიღებთ ცხრილიდან
                                       Dim personalData As IList(Of IList(Of Object)) = dataService.GetData("DB-Personal!B2:E")
                                       Debug.WriteLine($"LoadHomeDataAsync: DB-Personal-დან მიღებულია {If(personalData Is Nothing, 0, personalData.Count)} მწკრივი")

                                       If personalData IsNot Nothing AndAlso personalData.Count > 0 Then
                                           ' დღევანდელი თარიღი
                                           Dim today As DateTime = DateTime.Today

                                           ' გადავირბინოთ ყველა მწკრივი
                                           For Each row As IList(Of Object) In personalData
                                               If row.Count >= 3 Then
                                                   Try
                                                       ' შევამოწმოთ არის თუ არა დაბადების თარიღი
                                                       If row(2) IsNot Nothing AndAlso Not String.IsNullOrEmpty(row(2).ToString()) Then
                                                           Dim birthDateStr As String = row(2).ToString()
                                                           Dim birthDate As DateTime

                                                           ' სხვადასხვა ფორმატების მცდელობა პარსინგისთვის
                                                           If DateTime.TryParseExact(birthDateStr,
                                                                        New String() {"dd.MM.yyyy", "dd,MM,yyyy", "dd/MM/yyyy", "dd-MM-yyyy"},
                                                                        System.Globalization.CultureInfo.InvariantCulture,
                                                                        System.Globalization.DateTimeStyles.None,
                                                                        birthDate) OrElse
                                                      DateTime.TryParse(birthDateStr, birthDate) Then

                                                               ' შემდეგი დაბადების დღის გამოთვლა
                                                               Dim nextBirthday As DateTime = New DateTime(today.Year, birthDate.Month, birthDate.Day)
                                                               If nextBirthday < today Then
                                                                   nextBirthday = nextBirthday.AddYears(1)
                                                               End If

                                                               ' რამდენი დღე რჩება
                                                               Dim daysLeft As Integer = (nextBirthday - today).Days

                                                               ' თუ 7 დღეზე ნაკლები რჩება, დავამატოთ
                                                               If daysLeft <= 7 Then
                                                                   Dim birthday As New BirthdayModel()
                                                                   birthday.Id = birthdays.Count + 1
                                                                   birthday.PersonName = If(row(0) IsNot Nothing, row(0).ToString(), "")
                                                                   birthday.PersonSurname = If(row(1) IsNot Nothing, row(1).ToString(), "")
                                                                   birthday.BirthDate = birthDate
                                                                   birthdays.Add(birthday)

                                                                   Debug.WriteLine($"LoadHomeDataAsync: დაემატა დაბადების დღე - ID={birthday.Id}, " &
                                                                          $"სახელი={birthday.PersonName}, გვარი={birthday.PersonSurname}, " &
                                                                          $"თარიღი={birthday.BirthDate:dd.MM.yyyy}, დარჩა={daysLeft} დღე")
                                                               End If
                                                           End If
                                                       End If
                                                   Catch ex As Exception
                                                       Debug.WriteLine($"LoadHomeDataAsync: მწკრივის დამუშავების შეცდომა - {ex.Message}")
                                                   End Try
                                               End If
                                           Next
                                       End If

                                       ' თუ მაინც ვერ ვიპოვეთ, შევქმნათ საცდელი
                                       If birthdays.Count = 0 Then
                                           ' ხელოვნური საცდელი მონაცემი, რომ დავრწმუნდეთ UI მუშაობს
                                           Dim testBirthday As New BirthdayModel()
                                           testBirthday.Id = 1
                                           testBirthday.PersonName = "საცდელი"
                                           testBirthday.PersonSurname = "მომხმარებელი"
                                           testBirthday.BirthDate = DateTime.Today.AddDays(3)
                                           birthdays.Add(testBirthday)

                                           Debug.WriteLine("LoadHomeDataAsync: დაემატა საცდელი დაბადების დღე")
                                       End If
                                   End If

                                   ' UI განახლება მთავარ თრედზე
                                   Me.Invoke(Sub()
                                                 Try
                                                     ' სესიების განახლება
                                                     homeViewModel.PendingSessions.Clear()
                                                     For Each session In pendingSessions
                                                         homeViewModel.PendingSessions.Add(session)
                                                     Next
                                                     homeViewModel.PendingSessionsCount = pendingSessions.Count
                                                     Debug.WriteLine($"LoadHomeDataAsync: ViewModel-ში ჩაემატა {pendingSessions.Count} მოლოდინში სესია")

                                                     ' ვადაგადაცილებული სესიების განახლება
                                                     homeViewModel.OverdueSessions.Clear()
                                                     For Each session In overdueSessions
                                                         homeViewModel.OverdueSessions.Add(session)
                                                     Next
                                                     Debug.WriteLine($"LoadHomeDataAsync: ViewModel-ში ჩაემატა {overdueSessions.Count} ვადაგადაცილებული სესია")

                                                     ' ვადაგადაცილებული სესიების ბარათების შექმნა
                                                     If homeControl IsNot Nothing AndAlso Not homeControl.IsDisposed Then
                                                         Debug.WriteLine($"LoadHomeDataAsync: ვიძახებთ PopulateOverdueSessions - " &
                                                            $"სესიების რაოდენობა: {overdueSessions.Count}")

                                                         ' თუ ეს მოქმედი homeControl-ია
                                                         If homeControl.Visible AndAlso homeControl.IsHandleCreated Then
                                                             ' განვაახლოთ დღევანდელი სესიების სტატისტიკა
                                                             homeControl.UpdateTodaySessionsStatistics(todaySessions)
                                                             Debug.WriteLine($"LoadHomeDataAsync: დღევანდელი სესიების სტატისტიკა განახლდა, რაოდენობა: {todaySessions.Count}")

                                                             ' ვადაგადაცილებული სესიების შევსება
                                                             homeControl.PopulateOverdueSessions(overdueSessions.ToList(),
                                                                                viewModel.IsAuthorized,
                                                                                viewModel.Role)

                                                             Debug.WriteLine("LoadHomeDataAsync: PopulateOverdueSessions გამოძახება დასრულდა")

                                                             ' დავამატოთ დაბადების დღეების განახლების გამოძახება
                                                             homeControl.PopulateUpcomingBirthdays(birthdays)
                                                             Debug.WriteLine("LoadHomeDataAsync: PopulateUpcomingBirthdays გამოძახება დასრულდა")

                                                             ' განახლების შემდეგ კიდევ ერთხელ Application.DoEvents
                                                             Application.DoEvents()
                                                         End If
                                                     Else
                                                         Debug.WriteLine($"LoadHomeDataAsync: homeControl არის None ან Disposed!")
                                                     End If

                                                     ' დაბადების დღეების განახლება
                                                     homeViewModel.UpcomingBirthdays.Clear()
                                                     For Each birthday In birthdays
                                                         homeViewModel.UpcomingBirthdays.Add(birthday)
                                                     Next
                                                     Debug.WriteLine($"LoadHomeDataAsync: ViewModel-ში ჩაემატა {birthdays.Count} დაბადების დღე")

                                                     ' დავალებების განახლება
                                                     homeViewModel.ActiveTasks.Clear()
                                                     For Each task In tasks
                                                         homeViewModel.ActiveTasks.Add(task)
                                                     Next
                                                     Debug.WriteLine($"LoadHomeDataAsync: ViewModel-ში ჩაემატა {tasks.Count} დავალება")

                                                     ' Application.DoEvents() გამოძახება, რათა UI-მ მოასწროს განახლება
                                                     Application.DoEvents()
                                                     Debug.WriteLine("LoadHomeDataAsync: Application.DoEvents() გამოძახებულია UI-ის განახლებისთვის")

                                                 Catch uiEx As Exception
                                                     Debug.WriteLine($"LoadHomeDataAsync: შეცდომა UI განახლებისას: {uiEx.Message}")
                                                     Debug.WriteLine($"LoadHomeDataAsync: Stack Trace: {uiEx.StackTrace}")
                                                 End Try
                                             End Sub)
                               Catch threadEx As Exception
                                   Debug.WriteLine($"LoadHomeDataAsync: შეცდომა მონაცემების დამუშავების თრედზე: {threadEx.Message}")
                                   Debug.WriteLine($"LoadHomeDataAsync: Stack Trace: {threadEx.StackTrace}")
                               End Try
                           End Sub)

            Debug.WriteLine("LoadHomeDataAsync: მონაცემების დატვირთვა დასრულდა")
        Catch ex As Exception
            Debug.WriteLine($"LoadHomeDataAsync: საერთო შეცდომა: {ex.Message}")
            Debug.WriteLine($"LoadHomeDataAsync: Stack Trace: {ex.StackTrace}")
        End Try
    End Sub
    ''' <summary>
    ''' როლის ცვლილებისას ხილვადობის მართვა
    ''' </summary>
    ''' <param name="role">მომხმარებლის როლი</param>
    Private Sub ManageToolsVisibility(role As String)
        ' თუ homeControl არ არის შექმნილი, გამოვტოვოთ
        If homeControl Is Nothing Then
            Return
        End If

        ' გადავამოწმოთ ავტორიზაციის სტატუსი და როლი
        If viewModel.IsAuthorized Then
            ' როლის მიხედვით გადაწყვეტილება
            Dim hasAccess As Boolean = role = "1" OrElse role = "2" OrElse role = "3"
            homeControl.SetToolsVisibility(hasAccess)
        Else
            ' ავტორიზაციის გარეშე ყოველთვის დამალულია
            homeControl.SetToolsVisibility(False)
        End If
    End Sub

    ''' <summary>
    ''' Menu ItemClicked: MenuStrip-ის ივენთის დამმუშავებელი
    ''' </summary>
    Private Sub mainMenu_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles mainMenu.ItemClicked
        Try
            ' ჯერ ჩავინიშნოთ, რომელი პუნქტი აირჩია მომხმარებელმა
            Dim menuItemText = e.ClickedItem.Text
            Debug.WriteLine($"mainMenu_ItemClicked: მომხმარებელმა აირჩია '{menuItemText}'")

            ' კონკრეტული მენიუს პუნქტების დაჭერის დამუშავება
            Select Case menuItemText
                Case "საწყისი"
                    ShowHome()
                Case "კალენდარი"
                    ShowCalendar()
                Case "ბაზები"
                    Debug.WriteLine("'ბაზები' არჩეულია, ველოდებით ქვემენიუს არჩევას")
                Case "განრიგი"
                    Debug.WriteLine("'განრიგი' არჩეულია, ვაჩვენებთ განრიგის გვერდს")
                    ShowSchedule()
                Case "ბენეფიციარები", "თერაპევტები", "თერაპიები", "დაფინანსება"
                    ShowTemporaryDatabasesUI(menuItemText)
                Case "გრაფიკები"
                    ShowTemporaryGraphsUI()
                Case "დოკუმენტები"
                    Debug.WriteLine("'დოკუმენტები' არჩეულია, ველოდებით ქვემენიუს არჩევას")
                ' ✨ ახალი ქვემენიუს პუნქტები
                Case "ბენეფიციარის ანგარიში"
                    Debug.WriteLine("'ბენეფიციარის ანგარიში' არჩეულია")
                    OnBeneficiaryReportSelected(Me, EventArgs.Empty)
                Case "თერაპევტის ანგარიში"
                    Debug.WriteLine("'თერაპევტის ანგარიში' არჩეულია")
                    OnTherapistReportSelected(Me, EventArgs.Empty)
                Case "ფინანსები"
                    ShowTemporaryFinancesUI()
                Case "ადმინისტრირება", "მომხმარებელთა რეგისტრაცია"
                    ShowTemporaryAdminUI(menuItemText)
            End Select
        Catch ex As Exception
            Debug.WriteLine($"mainMenu_ItemClicked: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბაზების დროებითი ინტერფეისის ჩვენება (შეტყობინების გარეშე)
    ''' </summary>
    Private Sub ShowTemporaryDatabasesUI(menuItemText As String)
        Try
            ' გავასუფთავოთ მთავარი პანელი
            pnlMain.Controls.Clear()

            ' შევქმნათ დროებითი პანელი
            Dim tempPanel As New Panel()
            tempPanel.Dock = DockStyle.Fill
            tempPanel.BackColor = Color.FromArgb(245, 245, 250)

            ' დავამატოთ სათაური
            Dim titleLabel As New Label()
            titleLabel.Text = $"{menuItemText} - მუშავდება"
            titleLabel.Font = New Font("Sylfaen", 18, FontStyle.Bold)
            titleLabel.AutoSize = True
            titleLabel.Location = New Point(20, 20)
            tempPanel.Controls.Add(titleLabel)

            ' დავამატოთ აღწერა
            Dim descLabel As New Label()
            descLabel.Text = $"'{menuItemText}'-ს ფუნქციონალი ამჟამად მუშავდება და მალე იქნება ხელმისაწვდომი." &
                          Environment.NewLine & Environment.NewLine &
                          "გთხოვთ სცადოთ მოგვიანებით."
            descLabel.Font = New Font("Sylfaen", 12, FontStyle.Regular)
            descLabel.AutoSize = True
            descLabel.Location = New Point(20, 60)
            tempPanel.Controls.Add(descLabel)

            ' დროებითი მონაცემების ჩვენება
            Dim dataPanel As New Panel()
            dataPanel.BorderStyle = BorderStyle.FixedSingle
            dataPanel.Width = tempPanel.Width - 40
            dataPanel.Height = 300
            dataPanel.Location = New Point(20, 120)
            dataPanel.BackColor = Color.White
            dataPanel.AutoScroll = True
            tempPanel.Controls.Add(dataPanel)

            ' დავამატოთ პანელი მთავარ კონტეინერზე
            pnlMain.Controls.Add(tempPanel)

            Debug.WriteLine($"ShowTemporaryDatabasesUI: დროებითი ინტერფეისი გამოჩნდა - {menuItemText}")
        Catch ex As Exception
            Debug.WriteLine($"ShowTemporaryDatabasesUI: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' გრაფიკების დროებითი ინტერფეისის ჩვენება
    ''' </summary>
    Private Sub ShowTemporaryGraphsUI()
        Try
            ' გავასუფთავოთ მთავარი პანელი
            pnlMain.Controls.Clear()

            ' შევქმნათ დროებითი პანელი
            Dim tempPanel As New Panel()
            tempPanel.Dock = DockStyle.Fill
            tempPanel.BackColor = Color.FromArgb(230, 245, 230)

            ' დავამატოთ სათაური
            Dim titleLabel As New Label()
            titleLabel.Text = "გრაფიკები - მუშავდება"
            titleLabel.Font = New Font("Sylfaen", 18, FontStyle.Bold)
            titleLabel.AutoSize = True
            titleLabel.Location = New Point(20, 20)
            tempPanel.Controls.Add(titleLabel)

            ' დავამატოთ აღწერა
            Dim descLabel As New Label()
            descLabel.Text = "გრაფიკების ფუნქციონალი ამჟამად მუშავდება და მალე იქნება ხელმისაწვდომი." &
                          Environment.NewLine & Environment.NewLine &
                          "გთხოვთ სცადოთ მოგვიანებით."
            descLabel.Font = New Font("Sylfaen", 12, FontStyle.Regular)
            descLabel.AutoSize = True
            descLabel.Location = New Point(20, 60)
            tempPanel.Controls.Add(descLabel)

            ' დავამატოთ პანელი მთავარ კონტეინერზე
            pnlMain.Controls.Add(tempPanel)

            Debug.WriteLine("ShowTemporaryGraphsUI: დროებითი ინტერფეისი გამოჩნდა - გრაფიკები")
        Catch ex As Exception
            Debug.WriteLine($"ShowTemporaryGraphsUI: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' დოკუმენტების დროებითი ინტერფეისის ჩვენება
    ''' </summary>
    Private Sub ShowTemporaryDocumentsUI()
        Try
            Debug.WriteLine("ShowTemporaryDocumentsUI: აღარ გამოიყენება - ნამდვილი ქვემენიუები ხელმისაწვდომია")

            ' თუ რაიმე მიზეზით აქ მოვიდა, მაშინ უფრო ინფორმაციული შეტყობინება ვაჩვენოთ
            MessageBox.Show("დოკუმენტების სექცია გამოიყენეთ ქვემენიუდან:" & Environment.NewLine &
                          "• ბენეფიციარის ანგარიში" & Environment.NewLine &
                          "• თერაპევტის ანგარიში",
                          "დოკუმენტები", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Debug.WriteLine($"ShowTemporaryDocumentsUI: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფინანსების დროებითი ინტერფეისის ჩვენება
    ''' </summary>
    Private Sub ShowTemporaryFinancesUI()
        Try
            ' გავასუფთავოთ მთავარი პანელი
            pnlMain.Controls.Clear()

            ' შევქმნათ დროებითი პანელი
            Dim tempPanel As New Panel()
            tempPanel.Dock = DockStyle.Fill
            tempPanel.BackColor = Color.FromArgb(230, 230, 245)

            ' დავამატოთ სათაური
            Dim titleLabel As New Label()
            titleLabel.Text = "ფინანსები - მუშავდება"
            titleLabel.Font = New Font("Sylfaen", 18, FontStyle.Bold)
            titleLabel.AutoSize = True
            titleLabel.Location = New Point(20, 20)
            tempPanel.Controls.Add(titleLabel)

            ' დავამატოთ აღწერა
            Dim descLabel As New Label()
            descLabel.Text = "ფინანსების ფუნქციონალი ამჟამად მუშავდება და მალე იქნება ხელმისაწვდომი." &
                          Environment.NewLine & Environment.NewLine &
                          "გთხოვთ სცადოთ მოგვიანებით."
            descLabel.Font = New Font("Sylfaen", 12, FontStyle.Regular)
            descLabel.AutoSize = True
            descLabel.Location = New Point(20, 60)
            tempPanel.Controls.Add(descLabel)

            ' დავამატოთ პანელი მთავარ კონტეინერზე
            pnlMain.Controls.Add(tempPanel)

            Debug.WriteLine("ShowTemporaryFinancesUI: დროებითი ინტერფეისი გამოჩნდა - ფინანსები")
        Catch ex As Exception
            Debug.WriteLine($"ShowTemporaryFinancesUI: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ადმინისტრირების დროებითი ინტერფეისის ჩვენება
    ''' </summary>
    Private Sub ShowTemporaryAdminUI(menuItemText As String)
        Try
            ' გავასუფთავოთ მთავარი პანელი
            pnlMain.Controls.Clear()

            ' შევქმნათ დროებითი პანელი
            Dim tempPanel As New Panel()
            tempPanel.Dock = DockStyle.Fill
            tempPanel.BackColor = Color.FromArgb(245, 230, 230)

            ' დავამატოთ სათაური
            Dim titleLabel As New Label()
            titleLabel.Text = $"{menuItemText} - მუშავდება"
            titleLabel.Font = New Font("Sylfaen", 18, FontStyle.Bold)
            titleLabel.AutoSize = True
            titleLabel.Location = New Point(20, 20)
            tempPanel.Controls.Add(titleLabel)

            ' დავამატოთ აღწერა
            Dim descLabel As New Label()
            descLabel.Text = $"'{menuItemText}'-ს ფუნქციონალი ამჟამად მუშავდება და მალე იქნება ხელმისაწვდომი." &
                          Environment.NewLine & Environment.NewLine &
                          "გთხოვთ სცადოთ მოგვიანებით."
            descLabel.Font = New Font("Sylfaen", 12, FontStyle.Regular)
            descLabel.AutoSize = True
            descLabel.Location = New Point(20, 60)
            tempPanel.Controls.Add(descLabel)

            ' დავამატოთ პანელი მთავარ კონტეინერზე
            pnlMain.Controls.Add(tempPanel)

            Debug.WriteLine($"ShowTemporaryAdminUI: დროებითი ინტერფეისი გამოჩნდა - {menuItemText}")
        Catch ex As Exception
            Debug.WriteLine($"ShowTemporaryAdminUI: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კალენდრის გვერდის ჩვენება
    ''' </summary>
    Private Sub ShowCalendar()
        Try
            Debug.WriteLine("ShowCalendar: დაიწყო")

            ' გავასუფთავოთ მთავარი პანელი
            pnlMain.Controls.Clear()

            ' ახალი CalendarViewModel-ის შექმნა
            Dim calendarViewModel As New CalendarViewModel()

            ' UC_Calendar კონტროლის შექმნა
            Dim calendarControl As New UC_Calendar(calendarViewModel)
            calendarControl.Dock = DockStyle.Fill

            ' მონაცემთა სერვისის მითითება
            If dataService IsNot Nothing Then
                calendarControl.SetDataService(dataService)
            End If

            ' მომხმარებლის ელფოსტის მითითება
            calendarControl.SetUserEmail(If(viewModel?.Email, "უცნობი"))

            ' დავამატოთ კონტროლი პანელზე
            pnlMain.Controls.Add(calendarControl)

            Debug.WriteLine("ShowCalendar: კალენდრის გვერდი წარმატებით გამოჩნდა")
        Catch ex As Exception
            Debug.WriteLine($"ShowCalendar: შეცდომა - {ex.Message}")
            MessageBox.Show($"კალენდრის გვერდის ჩვენების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' განრიგის გვერდის ჩვენება - გაუმჯობესებული ვერსია
    ''' მომხმარებლის ინფორმაციის სწორი გადაცემით
    ''' </summary>
    Private Sub ShowSchedule()
        Try
            Debug.WriteLine("ShowSchedule: დაიწყო")

            ' გავასუფთავოთ მთავარი პანელი
            pnlMain.Controls.Clear()

            ' UC_Schedule კონტროლის შექმნა
            Dim scheduleControl As New UC_Schedule()
            scheduleControl.Dock = DockStyle.Fill

            ' 🔧 ჯერ მონაცემთა სერვისი
            If dataService IsNot Nothing Then
                scheduleControl.SetDataService(dataService)
                Debug.WriteLine("ShowSchedule: მონაცემთა სერვისი გადაცემულია")
            Else
                Debug.WriteLine("ShowSchedule: ❌ dataService არის Nothing")
            End If

            ' 🔧 შემდეგ მომხმარებლის ინფორმაცია
            If viewModel IsNot Nothing Then
                Dim userEmail As String = If(String.IsNullOrEmpty(viewModel.Email), "user@example.com", viewModel.Email)
                Dim userRole As String = If(String.IsNullOrEmpty(viewModel.Role), "6", viewModel.Role)

                scheduleControl.SetUserInfo(userEmail, userRole)
                Debug.WriteLine($"ShowSchedule: მომხმარებლის ინფორმაცია გადაცემულია - ელფოსტა: '{userEmail}', როლი: '{userRole}'")
            Else
                Debug.WriteLine("ShowSchedule: ❌ viewModel არის Nothing")
                ' ნაგულისხმევი მნიშვნელობები
                scheduleControl.SetUserInfo("user@example.com", "6")
            End If

            ' კონტროლის დამატება პანელზე
            pnlMain.Controls.Add(scheduleControl)

            Debug.WriteLine("ShowSchedule: განრიგის გვერდი წარმატებით გამოჩნდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowSchedule: შეცდომა - {ex.Message}")
            MessageBox.Show($"განრიგის გვერდის ჩვენების შეცდომა: {ex.Message}", "შეცდომა",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ამოწმებს და აღწერს კონტროლის ხილვადობის მდგომარეობას დებაგირებისთვის
    ''' </summary>
    Private Sub DiagnoseControlVisibility(control As Control, controlName As String)
        If control Is Nothing Then
            Debug.WriteLine($"DiagnoseControlVisibility: {controlName} არის Nothing")
            Return
        End If

        Debug.WriteLine($"DiagnoseControlVisibility: {controlName}")
        Debug.WriteLine($"  Visible: {control.Visible}")
        Debug.WriteLine($"  Created: {control.IsHandleCreated}")
        Debug.WriteLine($"  Enabled: {control.Enabled}")
        Debug.WriteLine($"  Bounds: {control.Bounds}")
        Debug.WriteLine($"  Size: {control.Size}")

        If control.Parent IsNot Nothing Then
            Debug.WriteLine($"  Parent: {control.Parent.Name}, Visible: {control.Parent.Visible}")
        Else
            Debug.WriteLine($"  Parent: None")
        End If

        Debug.WriteLine($"  Controls Count: {control.Controls.Count}")
    End Sub

    ''' <summary>
    ''' ✨ ბენეფიციარის ანგარიშის მენიუს არჩევისას გამოძახებული მეთოდი
    ''' ხელმისაწვდომია როლებისთვის: 1, 2, 3
    ''' </summary>
    Private Sub OnBeneficiaryReportSelected(sender As Object, e As EventArgs)
        Debug.WriteLine("OnBeneficiaryReportSelected: 'ბენეფიციარის ანგარიში' არჩეულია მენიუში")

        ' როლის შემოწმება
        If Not IsAccessAllowed({"1", "2", "3"}) Then
            MessageBox.Show("თქვენ არ გაქვთ ბენეფიციარის ანგარიშზე წვდომის უფლება", "წვდომა შეზღუდულია",
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
            MessageBox.Show("თქვენ არ გაქვთ თერაპევტის ანგარიშზე წვდომის უფლება", "წვდომა შეზღუდულია",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ShowTherapistReport()
    End Sub

    ''' <summary>
    ''' ✨ წვდომის შემოწმება - ახალი დამხმარე მეთოდი
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
    ''' ✨ ბენეფიციარის ანგარიშის გვერდის ჩვენება
    ''' </summary>
    Private Sub ShowBeneficiaryReport()
        Try
            Debug.WriteLine("ShowBeneficiaryReport: დაიწყო")

            ' გავასუფთავოთ მთავარი პანელი
            pnlMain.Controls.Clear()

            ' UC_BeneficiaryReport კონტროლის შექმნა
            Dim beneficiaryReportControl As New UC_BeneficiaryReport()
            beneficiaryReportControl.Dock = DockStyle.Fill

            ' მონაცემთა სერვისის მითითება
            If dataService IsNot Nothing Then
                beneficiaryReportControl.SetDataService(dataService)
                Debug.WriteLine("ShowBeneficiaryReport: მონაცემთა სერვისი გადაცემულია")
            Else
                Debug.WriteLine("ShowBeneficiaryReport: ❌ dataService არის Nothing")
            End If

            ' მომხმარებლის ინფორმაციის გადაცემა
            If viewModel IsNot Nothing Then
                Dim userEmail As String = If(String.IsNullOrEmpty(viewModel.Email), "user@example.com", viewModel.Email)
                Dim userRole As String = If(String.IsNullOrEmpty(viewModel.Role), "6", viewModel.Role)

                beneficiaryReportControl.SetUserInfo(userEmail, userRole)
                Debug.WriteLine($"ShowBeneficiaryReport: მომხმარებლის ინფორმაცია გადაცემულია - ელფოსტა: '{userEmail}', როლი: '{userRole}'")
            Else
                Debug.WriteLine("ShowBeneficiaryReport: ❌ viewModel არის Nothing")
                beneficiaryReportControl.SetUserInfo("user@example.com", "6")
            End If

            ' კონტროლის დამატება პანელზე
            pnlMain.Controls.Add(beneficiaryReportControl)

            Debug.WriteLine("ShowBeneficiaryReport: ბენეფიციარის ანგარიშის გვერდი წარმატებით გამოჩნდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowBeneficiaryReport: შეცდომა - {ex.Message}")
            MessageBox.Show($"ბენეფიციარის ანგარიშის გვერდის ჩვენების შეცდომა: {ex.Message}", "შეცდომა",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ✨ თერაპევტის ანგარიშის გვერდის ჩვენება
    ''' </summary>
    Private Sub ShowTherapistReport()
        Try
            Debug.WriteLine("ShowTherapistReport: დაიწყო")

            ' გავასუფთავოთ მთავარი პანელი
            pnlMain.Controls.Clear()

            ' UC_TherapistReport კონტროლის შექმნა
            Dim therapistReportControl As New UC_TherapistReport()
            therapistReportControl.Dock = DockStyle.Fill

            ' მონაცემთა სერვისის მითითება
            If dataService IsNot Nothing Then
                therapistReportControl.SetDataService(dataService)
                Debug.WriteLine("ShowTherapistReport: მონაცემთა სერვისი გადაცემულია")
            Else
                Debug.WriteLine("ShowTherapistReport: ❌ dataService არის Nothing")
            End If

            ' მომხმარებლის ინფორმაციის გადაცემა
            If viewModel IsNot Nothing Then
                Dim userEmail As String = If(String.IsNullOrEmpty(viewModel.Email), "user@example.com", viewModel.Email)
                Dim userRole As String = If(String.IsNullOrEmpty(viewModel.Role), "6", viewModel.Role)

                therapistReportControl.SetUserInfo(userEmail, userRole)
                Debug.WriteLine($"ShowTherapistReport: მომხმარებლის ინფორმაცია გადაცემულია - ელფოსტა: '{userEmail}', როლი: '{userRole}'")
            Else
                Debug.WriteLine("ShowTherapistReport: ❌ viewModel არის Nothing")
                therapistReportControl.SetUserInfo("user@example.com", "6")
            End If

            ' კონტროლის დამატება პანელზე
            pnlMain.Controls.Add(therapistReportControl)

            Debug.WriteLine("ShowTherapistReport: თერაპევტის ანგარიშის გვერდი წარმატებით გამოჩნდა")

        Catch ex As Exception
            Debug.WriteLine($"ShowTherapistReport: შეცდომა - {ex.Message}")
            MessageBox.Show($"თერაპევტის ანგარიშის გვერდის ჩვენების შეცდომა: {ex.Message}", "შეცდომა",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Form1.vb ფაილში დაამატეთ ეს ღილაკზე დაჭერის ფუნქცია
    Private Sub TestBirthdaysDirect()
        Try
            Debug.WriteLine("TestBirthdaysDirect: დაწყება")

            ' პირდაპირ DB-Personal ცხრილიდან მონაცემები
            Dim personalData As IList(Of IList(Of Object)) = dataService.GetData("DB-Personal!B2:E")
            Debug.WriteLine($"TestBirthdaysDirect: მიღებულია {If(personalData Is Nothing, 0, personalData.Count)} მწკრივი")

            ' გამოვიტანოთ ყველა დაბადების თარიღი დებაგირებისთვის
            If personalData IsNot Nothing Then
                For i As Integer = 0 To personalData.Count - 1
                    Dim row = personalData(i)
                    If row.Count >= 3 AndAlso row(2) IsNot Nothing Then
                        Debug.WriteLine($"TestBirthdaysDirect: მწკრივი {i + 2}, სახელი={row(0)}, გვარი={row(1)}, დაბადების თარიღი={row(2)}")
                    End If
                Next
            End If

            Debug.WriteLine("TestBirthdaysDirect: დასრულება")
        Catch ex As Exception
            Debug.WriteLine($"TestBirthdaysDirect: შეცდომა - {ex.Message}")
        End Try
    End Sub
End Class