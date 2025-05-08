Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Services
Imports System.IO

Namespace Scheduler_v8_8a.Services
    Public Class GoogleServiceAccountService
        Private ReadOnly serviceAccountKeyPath As String
        Public Property Credential As GoogleCredential
            Get
                Return _credential
            End Get
            Private Set(value As GoogleCredential)
                _credential = value
            End Set
        End Property
        Private _credential As GoogleCredential

        ''' <summary>
        ''' კონსტრუქტორი: სერვის ანგარიშის ფაილის მისამართს ითხოვს
        ''' </summary>
        ''' <param name="serviceAccountKeyPath">სერვის ანგარიშის JSON ფაილის მისამართი</param>
        Public Sub New(serviceAccountKeyPath As String)
            Me.serviceAccountKeyPath = serviceAccountKeyPath
        End Sub

        ''' <summary>
        ''' ავტორიზაციის განხორციელება სერვის ანგარიშის გამოყენებით
        ''' </summary>
        ''' <returns>GoogleCredential ობიექტი</returns>
        Public Function AuthorizeAsync() As GoogleCredential
            Try
                ' წავიკითხოთ JSON ფაილი და შევქმნათ credential ობიექტი
                Using stream = New FileStream(serviceAccountKeyPath, FileMode.Open, FileAccess.Read)
                    Credential = GoogleCredential.FromStream(stream).CreateScoped({SheetsService.Scope.Spreadsheets})
                End Using

                Return Credential
            Catch ex As Exception
                Throw New Exception("სერვის ანგარიშის ავტორიზაციის შეცდომა: " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' შექმენით SheetsService ობიექტი
        ''' </summary>
        ''' <returns>SheetsService ობიექტი</returns>
        Public Function CreateSheetsService() As SheetsService
            If Credential Is Nothing Then
                AuthorizeAsync()
            End If

            Return New SheetsService(New BaseClientService.Initializer() With {
                .HttpClientInitializer = Credential,
                .ApplicationName = "Scheduler_v8.8a"
            })
        End Function
    End Class
End Namespace