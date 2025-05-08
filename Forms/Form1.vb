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

Public Class Form1

    ' ViewModel-ები
    Private ReadOnly viewModel As MainViewModel
    Private ReadOnly homeViewModel As HomeViewModel

    ' სერვისები (Dependency Injection) - აღარ არის ReadOnly
    Private dataService As IDataService
    Private userService As UserService
    Private ReadOnly authService As GoogleOAuthService

    ' UI კომპონენტები
    Private ReadOnly menuMgr As MenuManager
    Private homeControl As UC_Home

    ' კონფიგურაცია
    Private ReadOnly spreadsheetId As String = "1SrBc4vLKPui6467aNmF5Hw-WZEd7dfGhkeFjfcnUqog"
    Private ReadOnly utilsFolder As String = Path.Combine(Application.StartupPath, "Utils")
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

        ' მენიუს მენეჯერის ინიციალიზაცია
        menuMgr = New MenuManager(mainMenu)

        ' PropertyChanged ივენთის ჰენდლერი
        AddHandler viewModel.PropertyChanged, AddressOf OnViewModelPropertyChanged
        ' საწყისი მდგომარეობაში ინსტრუმენტები დამალულია
        SetToolsVisibility(False)
    End Sub

    ''' <summary>
    ''' Form Load: სერვისების დაწყება, მენიუს კონფიგურაცია და UI-ის ინიციალიზაცია
    ''' </summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' მენიუს საწყისი მდგომარეობა - მხოლოდ საწყისი
        menuMgr.ShowOnlyHomeMenu()

        ' საწყისი Home View
        ShowHome()

        ' UI-ის საწყისი ინსტრუქციები ViewModel-იდან
        LUser.Text = If(String.IsNullOrEmpty(viewModel.Email), "გთხოვთ გაიაროთ ავტორიზაცია", viewModel.Email)
        BtnLogin.Text = If(viewModel.IsAuthorized, "გასვლა", "ავტორიზაცია")

        ' დავამატოთ მისალმება homeViewModel-ში
        homeViewModel.Greeting = homeViewModel.GetGreetingByTime()
    End Sub
    ''' <summary>
    ''' BtnLogin Click: Login ან Logout ივენთი ViewModel-ით დასმული
    ''' განახლებული ვერსია UserService-ის გამოყენებით
    ''' </summary>
    Private Async Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        ' დავბლოკოთ ღილაკი, რომ თავიდან ავირიდოთ მრავალჯერადი დაჭერა
        BtnLogin.Enabled = False

        Try
            If Not viewModel.IsAuthorized Then
                Try
                    ' 1) Google OAuth ავტორიზაცია
                    Await authService.AuthorizeAsync(New String() {
                    Google.Apis.Oauth2.v2.Oauth2Service.Scope.UserinfoEmail,
                    Google.Apis.Oauth2.v2.Oauth2Service.Scope.UserinfoProfile,
                    Google.Apis.Sheets.v4.SheetsService.Scope.SpreadsheetsReadonly,
                    Google.Apis.Sheets.v4.SheetsService.Scope.Spreadsheets
                })

                    ' 2) ინიციალიზება სერვისების ავტორიზაციის შემდეგ
                    dataService = New GoogleSheetsDataService(authService.Credential, spreadsheetId)
                    userService = New UserService(dataService)

                    ' 3) მომხმარებლის პროფილის მიღება
                    Dim userProfile = Await userService.GetUserProfile(authService.Credential)

                    ' 4) ViewModel განახლება -> იწვევს OnViewModelPropertyChanged
                    viewModel.Email = userProfile.Email
                    viewModel.Role = userProfile.Role
                    viewModel.IsAuthorized = True

                    ' 5) HomeViewModel-ის განახლება
                    homeViewModel.UserName = If(String.IsNullOrEmpty(userProfile.Name), userProfile.Email, userProfile.Name)

                    ' 6) UI განახლება
                    ShowHome()
                Catch ex As Exception
                    MessageBox.Show($"ავტორიზაცია ვერ შესრულდა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Else
                Try
                    ' Logout
                    Await authService.RevokeAsync()
                    viewModel.IsAuthorized = False
                    viewModel.Email = String.Empty
                    viewModel.Role = String.Empty
                    homeViewModel.UserName = String.Empty

                    ' მენიუს განახლება ავტორიზაციის შემდეგ
                    menuMgr.ShowOnlyHomeMenu()
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

        ' ვიხეიროთ მიმდინარე მნიშვნელობები იმისთვის, რომ თავიდან ავიცილოთ ციკლური განახლებები
        Static isUpdating As Boolean = False
        If isUpdating Then Return

        isUpdating = True
        Try
            Select Case e.PropertyName
                Case NameOf(viewModel.Email)
                    LUser.Text = viewModel.Email
                Case NameOf(viewModel.IsAuthorized)
                    BtnLogin.Text = If(viewModel.IsAuthorized, "გასვლა", "ავტორიზაცია")
                    SetToolsVisibility(viewModel.IsAuthorized)
                Case NameOf(viewModel.Role)
                    menuMgr.ShowMenuByRole(viewModel.Role)
                    ' როლის ცვლილება შეიძლება არ საჭიროებდეს ინსტრუმენტების ხელახალ კონფიგურაციას
                    ' თუ მხოლოდ მენიუებს ცვლის
            End Select
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

        ' ინსტრუმენტების ხილვადობის განახლება - მნიშვნელოვანია ამ ადგილას!
        SetToolsVisibility(viewModel.IsAuthorized, viewModel.Role)

        ' თუ ავტორიზებულია მომხმარებელი, ვცდილობთ განვაახლოთ მონაცემები
        If viewModel.IsAuthorized AndAlso dataService IsNot Nothing Then
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

            ' დატვირთვის ლოგიკა გავიტანოთ ცალკე თრედზე
            Await Task.Run(Sub()
                               Try
                                   ' მონაცემების წამოღება სხვა თრედზე
                                   Dim pendingSessions = dataService.GetPendingSessions()
                                   Dim overdueSessions = dataService.GetOverdueSessions() ' დავამატოთ ეს ხაზი
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

                                                     ' ვადაგადაცილებული სესიების ბარათების შექმნა
                                                     If homeControl IsNot Nothing Then
                                                         homeControl.PopulateOverdueSessions(overdueSessions.ToList(),
                                                              viewModel.IsAuthorized,
                                                              viewModel.Role)
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
                                                     Debug.WriteLine($"შეცდომა UI განახლებისას: {uiEx.Message}")
                                                 End Try
                                             End Sub)
                               Catch threadEx As Exception
                                   Debug.WriteLine($"შეცდომა მონაცემების დამუშავების თრედზე: {threadEx.Message}")
                               End Try
                           End Sub)
        Catch ex As Exception
            Debug.WriteLine($"საერთო შეცდომა მონაცემების დატვირთვისას: {ex.Message}")
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

End Class