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
        LUser.Text = viewModel.Email
        BtnLogin.Text = If(viewModel.IsAuthorized, "გასვლა", "ავტორიზაცია")

        ' დავამატოთ მისალმება homeViewModel-ში
        homeViewModel.Greeting = homeViewModel.GetGreetingByTime()
    End Sub

    ''' <summary>
    ''' BtnLogin Click: Login ან Logout ივენთი ViewModel-ით დასმული
    ''' განახლებული ვერსია UserService-ის გამოყენებით
    ''' </summary>
    Private Async Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
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
            Catch ex As Exception
                MessageBox.Show($"გასვლა ვერ განხორციელდა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' PropertyChanged Handler: UI და მენიუს განახლება ViewModel-იდან
    ''' </summary>
    Private Sub OnViewModelPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
        ' ამოცანა: UI-ის განახლება ViewModel-ის ცვლილებების შესაბამისად
        Select Case e.PropertyName
            Case NameOf(viewModel.Email)
                LUser.Text = viewModel.Email
            Case NameOf(viewModel.IsAuthorized)
                BtnLogin.Text = If(viewModel.IsAuthorized, "გასვლა", "ავტორიზაცია")
            Case NameOf(viewModel.Role)
                menuMgr.ShowMenuByRole(viewModel.Role)
        End Select
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
        End If

        homeControl.BringToFront()

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

            ' მოლოდინში არსებული სესიების დატვირთვა
            If dataService IsNot Nothing Then
                ' მოლოდინში არსებული სესიები
                Dim pendingSessions = dataService.GetPendingSessions()
                homeViewModel.PendingSessions.Clear()
                For Each session In pendingSessions
                    homeViewModel.PendingSessions.Add(session)
                Next

                ' სესიების რაოდენობის დაყენება
                homeViewModel.PendingSessionsCount = pendingSessions.Count

                ' მომავალი დაბადების დღეები (7 დღით)
                Dim birthdays = dataService.GetUpcomingBirthdays(7)
                homeViewModel.UpcomingBirthdays.Clear()
                For Each birthday In birthdays
                    homeViewModel.UpcomingBirthdays.Add(birthday)
                Next

                ' აქტიური დავალებები
                Dim tasks = dataService.GetActiveTasks()
                homeViewModel.ActiveTasks.Clear()
                For Each task In tasks
                    homeViewModel.ActiveTasks.Add(task)
                Next
            End If
        Catch ex As Exception
            Debug.WriteLine($"მონაცემების ასინქრონული დატვირთვის შეცდომა: {ex.Message}")
        End Try
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