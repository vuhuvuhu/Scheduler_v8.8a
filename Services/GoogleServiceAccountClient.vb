' ===========================================
' 📄 Services/GoogleServiceAccountClient.vb
' -------------------------------------------
' Google Sheets-თან წვდომა სერვის აკაუნტით
' ===========================================
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data
Imports System.IO

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' Google Service Account-ის კლიენტი Sheets API-სთან სამუშაოდ
    ''' </summary>
    Public Class GoogleServiceAccountClient
        Private ReadOnly sheetsService As SheetsService
        Private ReadOnly spreadsheetId As String

        ' თუკი რაიმე შეცდომა მოხდა ინიციალიზაციისას
        Private _initializationError As String = Nothing
        Public ReadOnly Property InitializationError As String
            Get
                Return _initializationError
            End Get
        End Property

        ' ფლაგი ინიციალიზაციის სტატუსის შესამოწმებლად
        Private _isInitialized As Boolean = False
        Public ReadOnly Property IsInitialized As Boolean
            Get
                Return _isInitialized
            End Get
        End Property

        ''' <summary>
        ''' კონსტრუქტორი: ქმნის SheetsService სერვის აკაუნტის კრედენშალებით
        ''' </summary>
        ''' <param name="serviceAccountKeyPath">სერვის აკაუნტის JSON ფაილის მისამართი</param>
        ''' <param name="spreadsheetId">Google Spreadsheet ID</param>
        Public Sub New(serviceAccountKeyPath As String, spreadsheetId As String)
            Me.spreadsheetId = spreadsheetId

            Try
                ' სერვის აკაუნტის კრედენშალების შექმნა
                Dim credential As GoogleCredential = Nothing

                Using stream As New FileStream(serviceAccountKeyPath, FileMode.Open, FileAccess.Read)
                    credential = GoogleCredential.FromStream(stream).CreateScoped({SheetsService.Scope.Spreadsheets})
                End Using

                ' SheetsService-ის ინიციალიზაცია
                sheetsService = New SheetsService(New BaseClientService.Initializer() With {
                    .HttpClientInitializer = credential,
                    .ApplicationName = "Scheduler_v8.8a"
                })

                _isInitialized = True
                _initializationError = Nothing

                Debug.WriteLine("GoogleServiceAccountClient: წარმატებით ინიციალიზებულია")
            Catch ex As Exception
                _initializationError = ex.Message
                _isInitialized = False
                Debug.WriteLine($"GoogleServiceAccountClient: შეცდომა ინიციალიზაციისას - {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მონაცემების წაკითხვა მითითებული დიაპაზონიდან
        ''' </summary>
        ''' <param name="range">დიაპაზონი (მაგ: "Sheet1!A1:C10")</param>
        ''' <returns>მწკრივების სია, თითოეული შეიცავს მნიშვნელობების სიას</returns>
        Public Function ReadRange(range As String) As IList(Of IList(Of Object))
            Try
                If Not _isInitialized Then
                    'Debug.WriteLine("GoogleServiceAccountClient.ReadRange: სერვისი არ არის ინიციალიზებული")
                    Return New List(Of IList(Of Object))()
                End If

                Dim request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range)
                Dim response = request.Execute()
                Dim values = response.Values

                'Debug.WriteLine($"GoogleServiceAccountClient.ReadRange: წაკითხულია {If(values Is Nothing, 0, values.Count)} მწკრივი დიაპაზონიდან {range}")
                Return If(values, New List(Of IList(Of Object))())
            Catch ex As Exception
                'Debug.WriteLine($"GoogleServiceAccountClient.ReadRange: შეცდომა - {ex.Message}")
                Return New List(Of IList(Of Object))()
            End Try
        End Function

        ''' <summary>
        ''' მონაცემების დამატება მითითებულ დიაპაზონში
        ''' </summary>
        ''' <param name="range">დიაპაზონი (მაგ: "Sheet1!A:C")</param>
        ''' <param name="values">მნიშვნელობების სია დასამატებლად</param>
        ''' <returns>დამატებული მწკრივების რაოდენობა</returns>
        Public Function AppendValues(range As String, values As IList(Of Object)) As Integer
            Try
                If Not _isInitialized Then
                    Debug.WriteLine("GoogleServiceAccountClient.AppendValues: სერვისი არ არის ინიციალიზებული")
                    Return 0
                End If

                Dim valueRange As New ValueRange With {
                    .Values = New List(Of IList(Of Object)) From {values}
                }

                Dim request = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, range)
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
                Dim response = request.Execute()

                Debug.WriteLine($"GoogleServiceAccountClient.AppendValues: დაემატა მონაცემები დიაპაზონში {range}")
                Return 1
            Catch ex As Exception
                Debug.WriteLine($"GoogleServiceAccountClient.AppendValues: შეცდომა - {ex.Message}")
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' მონაცემების განახლება მითითებულ დიაპაზონში
        ''' </summary>
        ''' <param name="range">დიაპაზონი (მაგ: "Sheet1!A1:C1")</param>
        ''' <param name="values">მნიშვნელობების სია განახლებისთვის</param>
        ''' <returns>განახლებული მწკრივების რაოდენობა</returns>
        Public Function UpdateValues(range As String, values As IList(Of Object)) As Integer
            Try
                If Not _isInitialized Then
                    Debug.WriteLine("GoogleServiceAccountClient.UpdateValues: სერვისი არ არის ინიციალიზებული")
                    Return 0
                End If

                Dim valueRange As New ValueRange With {
                    .Values = New List(Of IList(Of Object)) From {values}
                }

                Dim request = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, range)
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED
                Dim response = request.Execute()

                Debug.WriteLine($"GoogleServiceAccountClient.UpdateValues: განახლდა მონაცემები დიაპაზონში {range}")
                Return 1
            Catch ex As Exception
                Debug.WriteLine($"GoogleServiceAccountClient.UpdateValues: შეცდომა - {ex.Message}")
                Return 0
            End Try
        End Function
    End Class
End Namespace