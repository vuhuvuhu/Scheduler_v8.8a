' ===========================================
' 📄 Form1.vb (Forms\Form1)
' -------------------------------------------
' მიზანი:
'   ● Google OAuth 2.0 ავტორიზაცია BtnLogin–ით
'   ● LUser-ში ელ.ფოსტა და როლის ჩვენება
'   ● მენიუს მართვა როლის მიხედვით MenuManager–ით
' ===========================================

Imports System.IO
Imports System.Threading
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports Google.Apis.Oauth2.v2
Imports Google.Apis.Oauth2.v2.Data
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data
Imports Google.Apis.Util.Store

' Utils კლასების ფოლდერში დამატებული MenuManager კლასის გამოყენება
Imports Utils

Public Class Form1

    ' 🔐 ავტორიზაციის ველები
    Private userEmail As String = String.Empty
    Private isAuthorized As Boolean = False
    Private cred As UserCredential = Nothing
    Private userRole As String = String.Empty

    ' მენიუს მართვის ობიექტი
    Private menuMgr As MenuManager

    ' Resources
    Private ReadOnly spreadsheetId As String = "1SrBc4vLKPui6467aNmF5Hw-WZEd7dfGhkeFjfcnUqog"
    Private ReadOnly utilsFolder As String = Path.Combine(Application.StartupPath, "Utils")
    Private ReadOnly secretsFile As String = Path.Combine(utilsFolder, "client_secret_v8_7.json")

    ' ===========================================
    ' ფორმის ჩატვირთვის UI და მენიუს ინიციალიზაცია
    ' ===========================================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Utils საქაღალდის შექმნა
        If Not Directory.Exists(utilsFolder) Then Directory.CreateDirectory(utilsFolder)

        ' UI საწყისი
        BtnLogin.Text = "ავტორიზაცია"
        LUser.Text = "გთხოვთ გაიაროთ ავტორიზაცია"

        ' MenuManager-ის შექმნა და დასაწყისი მენიუს კონფიგურაცია
        menuMgr = New MenuManager(mainMenu)
        menuMgr.ShowOnlyHomeMenu()
    End Sub

    ' ===========================================
    ' BtnLogin-ის ჰენდლერი: ავტორიზაცია და მენიუს განახლება
    ' ===========================================
    Private Async Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        If Not isAuthorized Then
            Try
                ' წინა ტოკენების წაშლა ახალი scopes–ით
                Dim tokenPath = Path.Combine(utilsFolder, "TokenStore")
                If Directory.Exists(tokenPath) Then Directory.Delete(tokenPath, True)

                ' Client Secrets
                Dim secrets = GoogleClientSecrets.Load(
                    New FileStream(secretsFile, FileMode.Open, FileAccess.Read)
                ).Secrets

                ' OAuth ავტორიზაცია Scopes–ით და FileDataStore-ით
                Dim dataStore = New FileDataStore(tokenPath, True)
                cred = Await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    secrets,
                    New String() {Oauth2Service.Scope.UserinfoEmail, Oauth2Service.Scope.UserinfoProfile, SheetsService.Scope.SpreadsheetsReadonly, SheetsService.Scope.Spreadsheets},
                    "user",
                    CancellationToken.None,
                    dataStore
                )

                ' მომხმარებლის ელ.ფოსტის მიღება
                Dim oauthService = New Oauth2Service(New BaseClientService.Initializer() With {
                    .HttpClientInitializer = cred,
                    .ApplicationName = "Scheduler_v8.8a"
                })
                Dim userInfo = oauthService.Userinfo.Get().Execute()
                userEmail = userInfo.Email

                ' Sheets API და DB-Users დეტექცია
                Dim sheetsClient = New SheetsService(New BaseClientService.Initializer() With {
                    .HttpClientInitializer = cred,
                    .ApplicationName = "Scheduler_v8.8a"
                })
                Dim getReq = sheetsClient.Spreadsheets.Values.Get(spreadsheetId, "DB-Users!B2:C")
                Dim rows = getReq.Execute().Values

                ' როლის გამოკითხვა ან დამატება
                Dim found As Boolean = False
                If rows IsNot Nothing Then
                    For Each row As IList(Of Object) In rows
                        If row.Count >= 2 AndAlso String.Equals(row(0).ToString(), userEmail, StringComparison.OrdinalIgnoreCase) Then
                            userRole = row(1).ToString()
                            found = True
                            Exit For
                        End If
                    Next
                End If
                If Not found Then
                    userRole = "6"
                    Dim vr = New ValueRange With {.Values = New List(Of IList(Of Object)) From {New List(Of Object) From {userEmail, userRole}}}
                    Dim appendReq = sheetsClient.Spreadsheets.Values.Append(vr, spreadsheetId, "DB-Users!B:C")
                    appendReq.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
                    appendReq.Execute()
                End If

                ' UI განახლება
                LUser.Text = $"{userEmail} (როლი: {userRole})"
                BtnLogin.Text = "გასვლა"
                isAuthorized = True

                ' მენიუს ჩვენება როლის მიხედვით
                menuMgr.ShowMenuByRole(userRole)

            Catch ex As Exception
                MessageBox.Show($"ავტორიზაცია ან მონაცემთა განახლება ვერ დასრულდა: {ex.Message}",
                                "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            ' გასვლა
            Try
                If cred IsNot Nothing Then Await cred.RevokeTokenAsync(CancellationToken.None)
                cred = Nothing
                isAuthorized = False
                userEmail = String.Empty
                userRole = String.Empty

                ' UI და მენიუს დაწყობა ისევ მხოლოდ მთავარი
                LUser.Text = "გთხოვთ გაიაროთ ავტორიზაცია"
                BtnLogin.Text = "ავტორიზაცია"
                menuMgr.ShowOnlyHomeMenu()

            Catch ex As Exception
                MessageBox.Show($"გასვლისას მოხდა შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
End Class
