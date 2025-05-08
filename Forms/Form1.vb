' ===========================================
' 📄 Forms/Form1.vb (განახლებული MVVM-თვის და სერვისების დამოკიდებულების ინვერსიისთვის)
' -------------------------------------------
' მიზანი: UI გამიჯნული ViewModel-ზე, დამოკიდებულების ინვერსიით
' ===========================================
Imports System.IO
Imports System.ComponentModel
Imports Google.Apis.Auth.OAuth2
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services
Imports System.Text

Public Class Form1

    ' ViewModel-ები
    Private ReadOnly viewModel As MainViewModel
    Private ReadOnly homeViewModel As HomeViewModel

    ' სერვისები (Dependency Injection)
    Private ReadOnly authService As GoogleOAuthService
    Private serviceAccountService As GoogleServiceAccountService
    Private dataService As IDataService
    Private userService As UserService

    ' UI კომპონენტები
    Private ReadOnly menuMgr As MenuManager
    Private homeControl As UC_Home

    ' კონფიგურაცია
    Private ReadOnly spreadsheetId As String = "1SrBc4vLKPui6467aNmF5Hw-WZEd7dfGhkeFjfcnUqog"
    Private ReadOnly utilsFolder As String = Path.Combine(Application.StartupPath, "Utils")
    Private ReadOnly serviceAccountKeyPath As String = Path.Combine(utilsFolder, "google-service-account-key8_7a.json")
    Private ReadOnly secretsFile As String = Path.Combine(utilsFolder, "client_secret_v8_7.json")
    Private ReadOnly tokenStorePath As String = Path.Combine(utilsFolder, "TokenStore")

    ''' <summary>
    ''' კონსტრუქტორი: InitializeComponent და ViewModel-ების ინიციალიზაცია
    ''' </summary>
    Public Sub New()
        InitializeComponent()

        ' საქაღალდების შექმნა თუ არ არსებობს
        If Not Directory.Exists(utilsFolder) Then Directory.CreateDirectory(utilsFolder)
        If Not Directory.Exists(tokenStorePath) Then Directory.CreateDirectory(tokenStorePath)

        ' ViewModel-ების ინიციალიზაცია
        viewModel = New MainViewModel()
        homeViewModel = New HomeViewModel()

        ' ავტორიზაციის სერვისის ინიციალიზაცია
        authService = New GoogleOAuthService(secretsFile, tokenStorePath)

        ' სერვის ანგარიშის გამოყენება ცხრილებთან დასაკავშირებლად - პროგრამის გაშვებისთანავე
        Try
            Debug.WriteLine("სერვის ანგარიშით დაკავშირების მცდელობა...")
            Dim serviceAccountPath = Path.Combine(utilsFolder, "google-service-account-key8_7a.json")

            ' შემოწმება არსებობს თუ არა ფაილი
            If Not File.Exists(serviceAccountPath) Then
                Debug.WriteLine($"სერვის ანგარიშის ფაილი არ არსებობს: {serviceAccountPath}")
                MessageBox.Show("სერვის ანგარიშის ფაილი არ მოიძებნა. პროგრამა ვერ შეძლებს მონაცემებთან წვდომას.", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim serviceAccountService = New GoogleServiceAccountService(serviceAccountPath)
            Dim credential = serviceAccountService.AuthorizeAsync()
            'Dim sheetsService = serviceAccountService.CreateSheetsService()

            ' შევქმნათ dataService - სერვის ანგარიშის გამოყენებით
            ' შექმენით SheetsService ობიექტი
            Dim sheetsService = New Google.Apis.Sheets.v4.SheetsService(
    New Google.Apis.Services.BaseClientService.Initializer() With {
        .HttpClientInitializer = authService.Credential,
        .ApplicationName = "Scheduler_v8.8a"
    }
)

            ' და შემდეგ გამოიყენეთ ის SheetDataService-ში
            dataService = New SheetDataService(sheetsService, spreadsheetId)
            Debug.WriteLine("სერვის ანგარიშით დაკავშირება წარმატებულია")
        Catch ex As Exception
            Debug.WriteLine($"სერვის ანგარიშით დაკავშირების შეცდომა: {ex.Message}")
            MessageBox.Show($"სერვის ანგარიშით დაკავშირების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' მენიუს მენეჯერის ინიციალიზაცია
        menuMgr = New MenuManager(mainMenu)

        ' PropertyChanged ივენთის ჰენდლერი
        AddHandler viewModel.PropertyChanged, AddressOf OnViewModelPropertyChanged
        ' საწყისი მდგომარეობაში ინსტრუმენტები დამალულია
        SetToolsVisibility(False)
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

        ' GoogleServiceAccountClient-ის ინიციალიზაცია და მონაცემების სერვისის შექმნა
        Try
            ' შევამოწმოთ, არსებობს თუ არა საქაღალდე და საჭირო ფაილები
            If Not Directory.Exists(utilsFolder) Then
                Directory.CreateDirectory(utilsFolder)
            End If

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

        ' ივენთის ჰენდლერის დამატება
        AddHandler viewModel.PropertyChanged, AddressOf OnViewModelPropertyChanged

        ' საწყისი Home View
        ShowHome()

        ' UI-ის საწყისი ინსტრუქციები ViewModel-იდან
        LUser.Text = If(String.IsNullOrEmpty(viewModel.Email), "გთხოვთ გაიაროთ ავტორიზაცია", viewModel.Email)
        BtnLogin.Text = If(viewModel.IsAuthorized, "გასვლა", "ავტორიზაცია")

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
        ' ამოცანა: UC_Home კონტროლის ჩვენება და განახლება
        If homeControl Is Nothing Then
            homeControl = New UC_Home(homeViewModel)
            homeControl.Dock = DockStyle.Fill
            pnlMain.Controls.Add(homeControl)
            Debug.WriteLine("შეიქმნა ახალი homeControl")
        End If

        homeControl.BringToFront()
        Debug.WriteLine($"homeControl გახდა წინ: IsDisposed={homeControl.IsDisposed}, Visible={homeControl.Visible}")

        ' ინსტრუმენტების ხილვადობის განახლება
        SetToolsVisibility(viewModel.IsAuthorized, viewModel.Role)

        ' ყველა შემთხვევაში ვცდილობთ მონაცემების ჩატვირთვას, რადგან dataService უკვე ინიციალიზებულია სერვის ანგარიშით
        If dataService IsNot Nothing Then
            Try
                ' ახალი მეთოდი: დატვირთვა ასინქრონულად
                LoadHomeDataAsync()
            Catch ex As Exception
                Debug.WriteLine($"მონაცემების დატვირთვის შეცდომა: {ex.Message}")
            End Try
        End If
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

                                   Dim birthdays = dataService.GetUpcomingBirthdays(7)
                                   Dim tasks = dataService.GetActiveTasks()

                                   ' UI განახლება მთავარ თრედზე
                                   Me.Invoke(Sub()
                                                 Try
                                                     ' სესიების განახლება
                                                     homeViewModel.PendingSessions.Clear()
                                                     For Each session In pendingSessions
                                                         homeViewModel.PendingSessions.Add(session)
                                                     Next
                                                     homeViewModel.PendingSessionsCount = pendingSessions.Count

                                                     ' ვადაგადაცილებული სესიების განახლება
                                                     homeViewModel.OverdueSessions.Clear()
                                                     For Each session In overdueSessions
                                                         homeViewModel.OverdueSessions.Add(session)
                                                     Next
                                                     Debug.WriteLine($"LoadHomeDataAsync: ViewModel-ში ჩაემატა {overdueSessions.Count} ვადაგადაცილებული სესია")

                                                     ' ვადაგადაცილებული სესიების ბარათების შექმნა
                                                     If homeControl IsNot Nothing Then
                                                         Debug.WriteLine("LoadHomeDataAsync: ვიძახებთ PopulateOverdueSessions")
                                                         homeControl.PopulateOverdueSessions(overdueSessions.ToList(),
                                                          viewModel.IsAuthorized,
                                                          viewModel.Role)
                                                     Else
                                                         Debug.WriteLine("LoadHomeDataAsync: homeControl არის Nothing!")
                                                     End If

                                                     ' დაბადების დღეების განახლება
                                                     homeViewModel.UpcomingBirthdays.Clear()
                                                     For Each birthday In birthdays
                                                         homeViewModel.UpcomingBirthdays.Add(birthday)
                                                     Next

                                                     ' დავალებების განახლება
                                                     homeViewModel.ActiveTasks.Clear()
                                                     For Each task In tasks
                                                         homeViewModel.ActiveTasks.Add(task)
                                                     Next
                                                 Catch uiEx As Exception
                                                     Debug.WriteLine($"LoadHomeDataAsync: შეცდომა UI განახლებისას: {uiEx.Message}")
                                                 End Try
                                             End Sub)
                               Catch threadEx As Exception
                                   Debug.WriteLine($"LoadHomeDataAsync: შეცდომა მონაცემების დამუშავების თრედზე: {threadEx.Message}")
                               End Try
                           End Sub)

            Debug.WriteLine("LoadHomeDataAsync: მონაცემების დატვირთვა დასრულდა")
        Catch ex As Exception
            Debug.WriteLine($"LoadHomeDataAsync: საერთო შეცდომა: {ex.Message}")
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
        ' კონკრეტული მენიუს პუნქტების დაჭერის დამუშავება
        If e.ClickedItem.Text = "საწყისი" Then
            ShowHome()
        ElseIf e.ClickedItem.Text = "კალენდარი" Then
            ' TODO: კალენდრის გვერდის ჩვენება
            MessageBox.Show("კალენდრის ფუნქციონალი ჯერ არ არის იმპლემენტირებული", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf e.ClickedItem.Text = "ბაზები" Then
            ' TODO: ბაზების გვერდის ჩვენება
            MessageBox.Show("ბაზების ფუნქციონალი ჯერ არ არის იმპლემენტირებული", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
    ' დებაგის ღილაკის დაჭერაზე რეაქცია
    Private Sub btnDebug_Click_1(sender As Object, e As EventArgs) Handles btnDebug.Click
        Try
            ' შევამოწმოთ გვაქვს თუ არა dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული!", "შეტყობინება", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' მივიღოთ ყველა სესია პირდაპირ
            Dim rows = dataService.GetData("DB-Schedule!B2:P") ' ყველა სესიის მონაცემები

            If rows Is Nothing OrElse rows.Count = 0 Then
                MessageBox.Show("ვერ მოიძებნა არცერთი სესია!", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' დავამუშავოთ ყველა მწკრივი
            Dim allSessions As New List(Of SessionModel)()
            For Each row In rows
                Try
                    ' მინიმუმ 12 სვეტი გვჭირდება
                    If row.Count < 12 Then Continue For

                    ' შევქმნათ სესიის ობიექტი
                    Dim session = SessionModel.FromSheetRow(row)
                    allSessions.Add(session)
                Catch ex As Exception
                    ' თუ რომელიმე მწკრივის დამუშავება ვერ მოხერხდა, გავაგრძელოთ შემდეგით
                    Continue For
                End Try
            Next

            ' შევქმნათ ტექსტური ანგარიში
            Dim report As New StringBuilder()
            report.AppendLine($"===== სულ {allSessions.Count} სესია =====")
            report.AppendLine($"მიმდინარე დრო: {DateTime.Now:dd.MM.yyyy HH:mm:ss}")
            report.AppendLine()

            ' ვადაგადაცილებული სესიების მთვლელი
            Dim overdueCount As Integer = 0

            ' თითოეული სესიის ანალიზი
            For Each session In allSessions
                Dim currentDate = DateTime.Today
                Dim sessionDate = session.DateTime.Date
                Dim statusLower = session.Status.Trim().ToLower()

                Dim isPastDue = sessionDate < currentDate
                Dim isPlanned = statusLower = "დაგეგმილი"
                Dim isOverdue = isPlanned AndAlso isPastDue

                report.AppendLine($"ID: {session.Id}")
                report.AppendLine($"  თარიღი: {session.DateTime:dd.MM.yyyy HH:mm}")
                report.AppendLine($"  სტატუსი: '{session.Status}'")
                report.AppendLine($"  დაგეგმილია: {isPlanned}")
                report.AppendLine($"  ვადა გასულია: {isPastDue}")
                report.AppendLine($"  ვადაგადაცილებულია: {isOverdue}")
                report.AppendLine()

                If isOverdue Then
                    overdueCount += 1
                End If
            Next

            report.AppendLine($"===== სულ {overdueCount} ვადაგადაცილებული სესია =====")

            ' გამოვაჩინოთ ანგარიში
            MessageBox.Show(report.ToString(), "სესიების ანალიზი", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"დებაგ რეჟიმის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class