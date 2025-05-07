' ===========================================
' 📄 Forms/Form1.vb (Refactored with MVVM and UC_Home integration)
' -------------------------------------------
' მიზანი: UI გამიჯნული ViewModel-ზე, GoogleOAuthService და SheetDataService გამოყენებით,
'       UC_Home ასაჩვენებლად გამარტივებული ლოჯიკით
' ===========================================
Imports System.IO
Imports System.ComponentModel
Imports Google.Apis.Oauth2.v2
Imports Google.Apis.Services
Imports Google.Apis.Sheets.v4
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8_8a.Models
Imports Utils
Imports System.Drawing
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class Form1

    ' ViewModel-ები და სერვისები
    Private viewModel As MainViewModel
    Private homeViewModel As HomeViewModel
    Private authService As GoogleOAuthService
    Private menuMgr As MenuManager
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
        viewModel = New MainViewModel()
        homeViewModel = New HomeViewModel()
        AddHandler viewModel.PropertyChanged, AddressOf OnViewModelPropertyChanged
    End Sub

    ''' <summary>
    ''' Form Load: სერვისების დაწყება, მენიუს კონფიგურაცია, და მთავარი View (UC_Home)
    ''' </summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' საქაღალდების შექმნა თუ არ არსებობს
        If Not Directory.Exists(utilsFolder) Then Directory.CreateDirectory(utilsFolder)
        If Not Directory.Exists(tokenStorePath) Then Directory.CreateDirectory(tokenStorePath)

        ' ავტორიზაცია სერვისი
        authService = New GoogleOAuthService(secretsFile, tokenStorePath)

        ' მენიუს მენეჯერი - მხოლოდ საწყისი
        menuMgr = New MenuManager(mainMenu)
        menuMgr.ShowOnlyHomeMenu()

        ' საწყისი Home View
        ShowHome()

        ' UI-ის საწყისი ინსტრუქციები ViewModel-იდან
        LUser.Text = viewModel.Email
        BtnLogin.Text = If(viewModel.IsAuthorized, "გასვლა", "ავტორიზაცია")
    End Sub

    ''' <summary>
    ''' BtnLogin Click: Login ან Logout ივენთი ViewModel-ით დასმული
    ''' </summary>
    Private Async Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        If Not viewModel.IsAuthorized Then
            Try
                '1) Google OAuth ავტორიზაცია
                Await authService.AuthorizeAsync(New String() {
                    Oauth2Service.Scope.UserinfoEmail,
                    Oauth2Service.Scope.UserinfoProfile,
                    SheetsService.Scope.SpreadsheetsReadonly,
                    SheetsService.Scope.Spreadsheets
                })
                Dim cred = authService.Credential

                '2) Email წამოღება OAuth2 Service-ით
                Dim oauthSvc = New Oauth2Service(New BaseClientService.Initializer() With {
                    .HttpClientInitializer = cred,
                    .ApplicationName = "Scheduler_v8.8a"
                })
                Dim email = oauthSvc.Userinfo.Get().Execute().Email

                '3) როლის მიღება/დამატება SheetDataService-ით
                Dim sheetService = New SheetDataService(cred, spreadsheetId)
                Dim role = sheetService.GetOrCreateUserRole(email)

                '4) ViewModel განახლება -> იწვევს OnViewModelPropertyChanged
                viewModel.Email = email
                viewModel.Role = role
                viewModel.IsAuthorized = True
                ' ორ ადგილი: DB-Personal-დან სახელი წამოღება და HomeControl-ში ჩვენება
                Dim personalService = New SheetDataService(authService.Credential, spreadsheetId)
                Dim personalRows = personalService.ReadRange("DB-Personal!B2:G")
                Dim foundName As String = String.Empty
                If personalRows IsNot Nothing Then
                    For Each prow As IList(Of Object) In personalRows
                        If prow.Count >= 6 AndAlso String.Equals(prow(5).ToString(), viewModel.Email, StringComparison.OrdinalIgnoreCase) Then
                            foundName = prow(0).ToString()
                            Exit For
                        End If
                    Next
                End If
                ' Ensure Home control is visible and set the name or email with style
                ShowHome()
                If Not String.IsNullOrEmpty(foundName) Then
                    homeControl.LUserName.Text = foundName
                    homeControl.LUserName.Font = New Font(homeControl.LUserName.Font.FontFamily, 12, FontStyle.Bold)
                Else
                    homeControl.LUserName.Text = viewModel.Email
                    homeControl.LUserName.Font = New Font(homeControl.LUserName.Font.FontFamily, 8, FontStyle.Regular)
                End If
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
            Catch ex As Exception
                MessageBox.Show($"გასვლა ვერ განხორციელდა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' PropertyChanged Handler: UI და მენიუს განახლება ViewModel-იდან
    ''' </summary>
    Private Sub OnViewModelPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
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
    ''' Shows UC_Home in pnlMain
    ''' </summary>
    Private Sub ShowHome()
        If homeControl Is Nothing Then
            homeControl = New UC_Home(homeViewModel)
            homeControl.Dock = DockStyle.Fill
            pnlMain.Controls.Add(homeControl)
        End If
        homeControl.BringToFront()
    End Sub

    ''' <summary>
    ''' Menu ItemClicked: საწყისის ღილაკზე ShowHome იძახება
    ''' </summary>
    Private Sub mainMenu_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles mainMenu.ItemClicked
        If e.ClickedItem.Text = "საწყისი" Then ShowHome()
        ' სხვა მენიუთიემები...
    End Sub
End Class
