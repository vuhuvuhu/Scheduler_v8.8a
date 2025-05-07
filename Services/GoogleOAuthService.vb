' ===========================================
' 📄 Services/GoogleOAuthService.vb
' -------------------------------------------
' პასუხისმგებელი Google OAuth 2.0 ავტორიზაციისთვის და ტოკენის მართვისთვის
' ===========================================
Imports System.IO
Imports System.Threading
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Util.Store

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' GoogleOAuthService ახორციელებს Google OAuth 2.0 ავტორიზაციას
    ''' და უზრუნველყოფს კრედენშალის შენახვას FileDataStore–ში.
    ''' </summary>
    Public Class GoogleOAuthService

        ' JSON Client Secret–ის ფაილის მისამართი
        Private ReadOnly secretsFilePath As String
        ' FileDataStore საქაღალდე ტოკენისთვის
        Private ReadOnly tokenStorePath As String
        ' დამოწმებული UserCredential ობიექტი
        Public Property Credential As UserCredential
            Get
                Return _credential
            End Get
            Private Set(value As UserCredential)
                _credential = value
            End Set
        End Property
        Private _credential As UserCredential

        ''' <summary>
        ''' კონსტრუქტორი: იღებს Client Secret JSON ფაილის და TokenStore საქაღალდის ადგილს.
        ''' </summary>
        ''' <param name="secretsFilePath">client_secret.json ფაილის სრული გზა</param>
        ''' <param name="tokenStorePath">TokenStore საქაღალდის სრული გზა</param>
        Public Sub New(secretsFilePath As String, tokenStorePath As String)
            Me.secretsFilePath = secretsFilePath
            Me.tokenStorePath = tokenStorePath
        End Sub

        ''' <summary>
        ''' აკეთებს ავტორიზაციას საჭირო Scopes–ით და ინახავს ტოკენს FileDataStore–ში
        ''' </summary>
        ''' <param name="scopes">OAuth სქოფების მარყუჟი</param>
        ''' <returns>Task, რომელიც ერიცხავს UserCredential ობიექტს</returns>
        Public Async Function AuthorizeAsync(scopes As String()) As Task(Of UserCredential)
            ' თუ არსებობს ძველი ტოკენი, წავშალოთ ახალი კონსენტის მისაღებად
            If Directory.Exists(tokenStorePath) Then Directory.Delete(tokenStorePath, True)

            Dim secrets = GoogleClientSecrets.Load(
                New FileStream(secretsFilePath, FileMode.Open, FileAccess.Read)
            ).Secrets

            Dim dataStore = New FileDataStore(tokenStorePath, True)
            Credential = Await GoogleWebAuthorizationBroker.AuthorizeAsync(
                secrets,
                scopes,
                "user",
                CancellationToken.None,
                dataStore
            )
            Return Credential
        End Function

        ''' <summary>
        ''' Revokes existing token და წაშლის Local FileDataStore მონაცემებს
        ''' </summary>
        Public Async Function RevokeAsync() As Task
            If Credential IsNot Nothing Then
                Await Credential.RevokeTokenAsync(CancellationToken.None)
            End If
            ' წაშალეთ საქაღალდე, რომ მომავალში ახალი ავტორიზაცია განხორციელდეს
            If Directory.Exists(tokenStorePath) Then Directory.Delete(tokenStorePath, True)
        End Function

    End Class

End Namespace
