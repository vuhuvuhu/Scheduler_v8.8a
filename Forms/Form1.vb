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

    ''' <summary>
    ''' კონსტრუქტორი: ძირითადი ინიციალიზაცია
    ''' </summary>
    Public Sub New()
        InitializeComponent()

        ' საქაღალდეების შექმნა თუ არ არსებობს
        If Not Directory.Exists(utilsFolder) Then Directory.CreateDirectory(utilsFolder)
        If Not Directory.Exists(tokenStorePath) Then Directory.CreateDirectory(tokenStorePath)

        ' მენიუს მენეჯერის ინიციალიზაცია
        menuMgr = New MenuManager(mainMenu)
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
        Debug.WriteLine("ShowHome: დაიწყო - ამჟამინდელი UC_Home მდგომარეობა:")

        ' ამჟამინდელი home კონტროლის დიაგნოსტიკა
        If homeControl IsNot Nothing Then
            Debug.WriteLine($"ShowHome: არსებული homeControl - Disposed={homeControl.IsDisposed}, Visible={homeControl.Visible}")
        Else
            Debug.WriteLine("ShowHome: homeControl არის Nothing")
        End If

        ' შევამოწმოთ pnlMain
        Debug.WriteLine($"ShowHome: pnlMain - Controls={pnlMain.Controls.Count}, Visible={pnlMain.Visible}")

        ' ამოცანა: UC_Home კონტროლის ჩვენება და განახლება
        If homeControl Is Nothing OrElse homeControl.IsDisposed Then
            Debug.WriteLine("ShowHome: ახალი homeControl-ის შექმნა")
            homeControl = New UC_Home(homeViewModel)
            homeControl.Dock = DockStyle.Fill
            pnlMain.Controls.Add(homeControl)
            Debug.WriteLine($"ShowHome: შეიქმნა ახალი homeControl - Disposed={homeControl.IsDisposed}, Visible={homeControl.Visible}")
        Else
            Debug.WriteLine("ShowHome: homeControl უკვე არსებობს და ვალიდურია")
        End If

        ' გავხადოთ homeControl წინა პლანზე
        homeControl.BringToFront()
        Debug.WriteLine($"ShowHome: homeControl.BringToFront() გამოძახებულია")

        ' შეგვიძლია ასევე გამოვიძახოთ Refresh
        homeControl.Refresh()
        Debug.WriteLine($"ShowHome: homeControl.Refresh() გამოძახებულია")

        ' ინსტრუმენტების ხილვადობის განახლება
        SetToolsVisibility(viewModel.IsAuthorized, viewModel.Role)

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

        Debug.WriteLine("ShowHome: დასრულებულია")
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

                                   ' პირველი 3 ვადაგადაცილებული სესიის დებაგინგი
                                   For i As Integer = 0 To Math.Min(2, overdueSessions.Count - 1)
                                       Debug.WriteLine($"LoadHomeDataAsync: ვადაგადაცილებული სესია #{i + 1} - " &
                                                  $"ID={overdueSessions(i).Id}, " &
                                                  $"თარიღი={overdueSessions(i).DateTime:dd.MM.yyyy HH:mm}, " &
                                                  $"სტატუსი='{overdueSessions(i).Status}'")
                                   Next

                                   Dim birthdays = dataService.GetUpcomingBirthdays(7)
                                   Debug.WriteLine($"LoadHomeDataAsync: წამოღებულია {birthdays.Count} მოახლოებული დაბადების დღე")

                                   Dim tasks = dataService.GetActiveTasks()
                                   Debug.WriteLine($"LoadHomeDataAsync: წამოღებულია {tasks.Count} აქტიური დავალება")

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
                                                         Debug.WriteLine($"LoadHomeDataAsync: ViewModel-ში ჩაემატა სესია - ID={session.Id}, სტატუსი={session.Status}")
                                                     Next
                                                     Debug.WriteLine($"LoadHomeDataAsync: ViewModel-ში ჩაემატა {overdueSessions.Count} ვადაგადაცილებული სესია")

                                                     ' ვადაგადაცილებული სესიების ბარათების შექმნა
                                                     If homeControl IsNot Nothing AndAlso Not homeControl.IsDisposed Then
                                                         Debug.WriteLine($"LoadHomeDataAsync: ვიძახებთ PopulateOverdueSessions - " &
                                                                    $"სესიების რაოდენობა: {overdueSessions.Count}, " &
                                                                    $"homeControl.Visible: {homeControl.Visible}, " &
                                                                    $"homeControl.IsHandleCreated: {homeControl.IsHandleCreated}")

                                                         ' შევამოწმოთ თუ ეს "ლაივ" homeControl ობიექტია და არა გასუფთავებული მიმართვა
                                                         Dim isHomeControlVisible As Boolean = False
                                                         Try
                                                             isHomeControlVisible = homeControl.Visible AndAlso homeControl.IsHandleCreated
                                                             Debug.WriteLine($"LoadHomeDataAsync: homeControl.Visible={isHomeControlVisible}")
                                                         Catch ex As Exception
                                                             Debug.WriteLine($"LoadHomeDataAsync: ვერ შემოწმდა homeControl.Visible: {ex.Message}")
                                                         End Try

                                                         ' PopulateOverdueSessions-ის გამოძახება, თუ ის მოქმედი homeControl-ია
                                                         homeControl.PopulateOverdueSessions(overdueSessions.ToList(),
                                                                                         viewModel.IsAuthorized,
                                                                                         viewModel.Role)

                                                         Debug.WriteLine("LoadHomeDataAsync: PopulateOverdueSessions გამოძახება დასრულდა")
                                                     Else
                                                         Debug.WriteLine($"LoadHomeDataAsync: homeControl არის None ან Disposed! IsNothing={homeControl Is Nothing}, IsDisposed={If(homeControl IsNot Nothing, homeControl.IsDisposed, True)}")
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

            ' მონაცემების წამოღება მთლიანი დიაპაზონიდან
            Dim rows = dataService.GetData("DB-Schedule!A2:O") ' სრული დიაპაზონი

            If rows Is Nothing OrElse rows.Count = 0 Then
                MessageBox.Show("ვერ მოიძებნა არცერთი სესია!", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' სტატუსების ანალიზი
            Dim statusCounts As New Dictionary(Of String, Integer)
            Dim overdueCount As Integer = 0

            For Each row In rows
                If row.Count > 12 Then
                    ' M სვეტი (სტატუსი) არის ინდექსით 12
                    Dim status = row(12).ToString().Trim()

                    ' დავამატოთ სტატუსი ლექსიკონში
                    If statusCounts.ContainsKey(status) Then
                        statusCounts(status) += 1
                    Else
                        statusCounts.Add(status, 1)
                    End If

                    ' შევამოწმოთ არის თუ არა ვადაგადაცილებული
                    If status.ToLower() = "დაგეგმილი" Then
                        ' F სვეტი (თარიღი) არის ინდექსით 5
                        Dim dateStr = If(row.Count > 5, row(5).ToString(), String.Empty)
                        Dim sessionDate As DateTime

                        If DateTime.TryParseExact(dateStr, "dd.MM.yyyy HH:mm", System.Globalization.CultureInfo.InvariantCulture,
                                             System.Globalization.DateTimeStyles.None, sessionDate) Then
                            If sessionDate.Date < DateTime.Today Then
                                overdueCount += 1
                            End If
                        End If
                    End If
                End If
            Next

            ' შედეგების გამოტანა
            Dim statusInfo As New StringBuilder()
            statusInfo.AppendLine("==== სტატუსების ანალიზი (M სვეტი) ====")
            statusInfo.AppendLine($"სულ {rows.Count} მწკრივია ცხრილში.")
            statusInfo.AppendLine()

            For Each pair In statusCounts
                statusInfo.AppendLine($"სტატუსი '{pair.Key}': {pair.Value} სესია")
            Next

            statusInfo.AppendLine()
            statusInfo.AppendLine($"სულ ვადაგადაცილებული სესიები: {overdueCount}")
            statusInfo.AppendLine("(დაგეგმილი სტატუსით და წარსული თარიღით)")

            ' პირველი რამდენიმე მწკრივის სტრუქტურა
            statusInfo.AppendLine()
            statusInfo.AppendLine("==== პირველი 3 სესიის სტრუქტურა ====")

            For i As Integer = 0 To Math.Min(2, rows.Count - 1)
                statusInfo.AppendLine($"მწკრივი {i + 2}:")
                statusInfo.AppendLine($"  ID (A): '{rows(i)(0)}'")
                If rows(i).Count > 5 Then statusInfo.AppendLine($"  თარიღი (F): '{rows(i)(5)}'")
                If rows(i).Count > 12 Then statusInfo.AppendLine($"  სტატუსი (M): '{rows(i)(12)}'")
                statusInfo.AppendLine()
            Next

            MessageBox.Show(statusInfo.ToString(), "სესიების ანალიზი", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"დებაგ რეჟიმის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
End Class