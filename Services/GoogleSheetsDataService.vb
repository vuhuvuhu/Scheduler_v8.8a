' ===========================================
' 📄 Services/GoogleSheetsDataService.vb
' -------------------------------------------
' Google Sheets-დან მონაცემთა წამოღებასა და ჩანაწერების დამატების სერვისი 
' IDataService ინტერფეისის იმპლემენტაციით
' ===========================================
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data
Imports Google.Apis.Services
Imports Google.Apis.Auth.OAuth2
Imports System.IO
Imports Newtonsoft.Json
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' GoogleSheetsDataService ახორციელებს მონაცემთა წაკითხვასა და განახლებას
    ''' Google Sheets–დან და ახდენს IDataService ინტერფეისის იმპლემენტაციას
    ''' ამავდროულად უზრუნველყოფს მონაცემთა ქეშირებას და ოფლაინ მუშაობას
    ''' </summary>
    Public Class GoogleSheetsDataService
        Implements IDataService

        Private ReadOnly sheetsService As SheetsService
        Private ReadOnly spreadsheetId As String

        ' კონსტანტები ხშირად გამოყენებადი დიაპაზონებისთვის
        Private Const usersRange As String = "DB-Users!B2:C"
        Private Const appendUsersRange As String = "DB-Users!B:C"
        Private Const birthdaysRange As String = "DB-Personal!B2:G"
        Private Const sessionsRange As String = "DB-Schedule!B2:P"
        Private Const tasksRange As String = "DB-Tasks!B2:M"
        Private Const defaultRole As String = "6"

        ' ქეშირების სისტემა
        Private ReadOnly cache As New Dictionary(Of String, CacheEntry)
        Private ReadOnly cacheFolder As String
        Private ReadOnly cacheFile As String
        Private ReadOnly cacheDuration As TimeSpan = TimeSpan.FromMinutes(15) ' 15 წუთი

        ''' <summary>
        ''' ქეშის ჩანაწერი
        ''' </summary>
        Private Class CacheEntry
            Public Property Data As Object
            Public Property Timestamp As DateTime

            Public Function IsExpired() As Boolean
                Return DateTime.Now.Subtract(Timestamp) > TimeSpan.FromMinutes(15)
            End Function
        End Class

        ''' <summary>
        ''' კონსტრუქტორი: აგებს SheetsService–ს OAUTHCredential–ით და spreadsheetId–თ
        ''' </summary>
        ''' <param name="credential">Google OAuth2 UserCredential</param>
        ''' <param name="spreadsheetId">Spreadsheet-ის ID</param>
        Public Sub New(credential As UserCredential, spreadsheetId As String)
            Me.sheetsService = New SheetsService(New BaseClientService.Initializer() With {
                .HttpClientInitializer = credential,
                .ApplicationName = "Scheduler_v8.8a"
            })
            Me.spreadsheetId = spreadsheetId

            ' ქეშის საქაღალდის ინიციალიზაცია
            cacheFolder = Path.Combine(Application.StartupPath, "Cache")
            cacheFile = Path.Combine(cacheFolder, "data_cache.json")

            If Not Directory.Exists(cacheFolder) Then
                Directory.CreateDirectory(cacheFolder)
            End If

            ' ქეშის ჩატვირთვა თუ არსებობს
            LoadCache()
        End Sub

        ''' <summary>
        ''' ქეშის ჩატვირთვა ფაილიდან
        ''' </summary>
        Private Sub LoadCache()
            Try
                If File.Exists(cacheFile) Then
                    Dim json = File.ReadAllText(cacheFile)
                    Dim savedCache = JsonConvert.DeserializeObject(Of Dictionary(Of String, CacheEntry))(json)
                    If savedCache IsNot Nothing Then
                        For Each entry In savedCache
                            ' მხოლოდ არავადაგასული ჩანაწერების ჩატვირთვა
                            If Not entry.Value.IsExpired() Then
                                cache(entry.Key) = entry.Value
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception
                ' ქეშის ჩატვირთვის შეცდომის დროს ცარიელი ქეშით გაგრძელება
                Debug.WriteLine($"ქეშის ჩატვირთვის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ქეშის შენახვა ფაილში
        ''' </summary>
        Private Sub SaveCache()
            Try
                ' გავასუფთაოთ ქეში ვადაგასული ჩანაწერებისგან
                Dim validEntries = cache.Where(Function(e) Not e.Value.IsExpired()).ToDictionary(Function(e) e.Key, Function(e) e.Value)

                ' შევინახოთ ქეში
                Dim json = JsonConvert.SerializeObject(validEntries)
                File.WriteAllText(cacheFile, json)
            Catch ex As Exception
                Debug.WriteLine($"ქეშის შენახვის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მონაცემთა შენახვა ქეშში
        ''' </summary>
        Private Sub CacheData(key As String, data As Object)
            cache(key) = New CacheEntry With {
                .Data = data,
                .Timestamp = DateTime.Now
            }
            SaveCache() ' ყოველი განახლებისას ვინახავთ ქეშს
        End Sub

        ''' <summary>
        ''' შემოწმება თუ არსებობს ქეშში მონაცემები
        ''' </summary>
        Private Function TryGetFromCache(Of T)(key As String, ByRef result As T) As Boolean
            If cache.ContainsKey(key) AndAlso Not cache(key).IsExpired() Then
                Try
                    result = CType(cache(key).Data, T)
                    Return True
                Catch
                    ' თუ ტიპების გარდაქმნა ვერ ხერხდება, ვაბრუნებთ false
                    Return False
                End Try
            End If
            Return False
        End Function

        ''' <summary>
        ''' IDataService.GetData იმპლემენტაცია
        ''' </summary>
        Public Function GetData(range As String) As IList(Of IList(Of Object)) Implements IDataService.GetData
            ' შევამოწმოთ ქეში
            Dim cacheKey = $"data_{range}"
            Dim cachedData As IList(Of IList(Of Object)) = Nothing

            If TryGetFromCache(cacheKey, cachedData) Then
                Return cachedData
            End If

            Try
                ' მონაცემთა წამოღება Google Sheets-დან
                Dim request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range)
                Dim response = request.Execute()
                Dim values = response.Values

                ' ქეშში შენახვა
                If values IsNot Nothing Then
                    CacheData(cacheKey, values)
                End If

                Return values
            Catch ex As Exception
                Debug.WriteLine($"მონაცემების წამოღების შეცდომა: {ex.Message}")
                Return New List(Of IList(Of Object))()
            End Try
        End Function

        ''' <summary>
        ''' IDataService.AppendData იმპლემენტაცია
        ''' </summary>
        Public Sub AppendData(range As String, values As IList(Of Object)) Implements IDataService.AppendData
            Try
                ' შევქმნათ ValueRange ობიექტი
                Dim valueRange As New ValueRange With {
                    .Values = New List(Of IList(Of Object)) From {values}
                }

                ' გამოვიძახოთ Append მეთოდი
                Dim request = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, range)
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
                request.Execute()

                ' წავშალოთ ქეში ამ range-სთვის
                Dim cacheKey = $"data_{range}"
                If cache.ContainsKey(cacheKey) Then
                    cache.Remove(cacheKey)
                    SaveCache()
                End If
            Catch ex As Exception
                Debug.WriteLine($"მონაცემების დამატების შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' IDataService.UpdateData იმპლემენტაცია
        ''' </summary>
        Public Sub UpdateData(range As String, values As IList(Of Object)) Implements IDataService.UpdateData
            Try
                ' შევქმნათ ValueRange ობიექტი
                Dim valueRange As New ValueRange With {
                    .Values = New List(Of IList(Of Object)) From {values}
                }

                ' გამოვიძახოთ Update მეთოდი
                Dim request = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, range)
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED
                request.Execute()

                ' წავშალოთ ქეში ამ range-სთვის
                Dim cacheKey = $"data_{range}"
                If cache.ContainsKey(cacheKey) Then
                    cache.Remove(cacheKey)
                    SaveCache()
                End If
            Catch ex As Exception
                Debug.WriteLine($"მონაცემების განახლების შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' IDataService.GetUserRole იმპლემენტაცია
        ''' </summary>
        Public Function GetUserRole(email As String) As String Implements IDataService.GetUserRole
            ' შევამოწმოთ ქეში
            Dim cacheKey = $"user_role_{email}"
            Dim cachedRole As String = Nothing

            If TryGetFromCache(cacheKey, cachedRole) Then
                Return cachedRole
            End If

            Try
                ' მომხმარებლის როლის წამოღება
                Dim rows = GetData(usersRange)
                If rows IsNot Nothing Then
                    For Each row As IList(Of Object) In rows
                        If row.Count >= 2 AndAlso String.Equals(row(0).ToString(), email, StringComparison.OrdinalIgnoreCase) Then
                            Dim role = row(1).ToString()

                            ' ქეშში შენახვა
                            CacheData(cacheKey, role)

                            Return role
                        End If
                    Next
                End If

                ' მომხმარებელი ვერ მოიძებნა
                Return String.Empty
            Catch ex As Exception
                Debug.WriteLine($"მომხმარებლის როლის წამოღების შეცდომა: {ex.Message}")
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' IDataService.GetOrCreateUserRole იმპლემენტაცია
        ''' </summary>
        Public Function GetOrCreateUserRole(email As String) As String Implements IDataService.GetOrCreateUserRole
            ' ჯერ ვცდილობთ მოვძებნოთ
            Dim role = GetUserRole(email)

            ' თუ ვერ მოვძებნეთ, ვქმნით ახალს
            If String.IsNullOrEmpty(role) Then
                Try
                    ' ვქმნით ახალ ჩანაწერს
                    Dim newRow As IList(Of Object) = New List(Of Object) From {email, defaultRole}
                    AppendData(appendUsersRange, newRow)

                    ' ქეშში ვინახავთ
                    CacheData($"user_role_{email}", defaultRole)

                    Return defaultRole
                Catch ex As Exception
                    Debug.WriteLine($"მომხმარებლის როლის შექმნის შეცდომა: {ex.Message}")
                    Return defaultRole ' მაინც ვაბრუნებთ ნაგულისხმევ როლს
                End Try
            End If

            Return role
        End Function

        ''' <summary>
        ''' IDataService.GetUpcomingBirthdays იმპლემენტაცია
        ''' </summary>
        Public Function GetUpcomingBirthdays(Optional days As Integer = 7) As List(Of BirthdayModel) Implements IDataService.GetUpcomingBirthdays
            ' შევამოწმოთ ქეში
            Dim cacheKey = $"upcoming_birthdays_{days}"
            Dim cachedBirthdays As List(Of BirthdayModel) = Nothing

            If TryGetFromCache(cacheKey, cachedBirthdays) Then
                Return cachedBirthdays
            End If

            Try
                Dim birthdays As New List(Of BirthdayModel)()
                Dim rows = GetData(birthdaysRange)

                If rows IsNot Nothing Then
                    For Each row As IList(Of Object) In rows
                        Try
                            ' შევქმნათ BirthdayModel
                            Dim birthday = BirthdayModel.CreateFromSheetRow(row)

                            ' შევამოწმოთ არის თუ არა მოახლოებული
                            If birthday.DaysUntilBirthday <= days AndAlso birthday.DaysUntilBirthday >= 0 Then
                                birthdays.Add(birthday)
                            End If
                        Catch ex As Exception
                            ' ვაგრძელებთ შემდეგი ჩანაწერით
                            Continue For
                        End Try
                    Next
                End If

                ' ქეშში შენახვა
                CacheData(cacheKey, birthdays)

                Return birthdays
            Catch ex As Exception
                Debug.WriteLine($"დაბადების დღეების წამოღების შეცდომა: {ex.Message}")
                Return New List(Of BirthdayModel)()
            End Try
        End Function

        ''' <summary>
        ''' IDataService.GetPendingSessions იმპლემენტაცია
        ''' </summary>
        Public Function GetPendingSessions() As List(Of SessionModel) Implements IDataService.GetPendingSessions
            ' შევამოწმოთ ქეში
            Dim cacheKey = "pending_sessions"
            Dim cachedSessions As List(Of SessionModel) = Nothing

            If TryGetFromCache(cacheKey, cachedSessions) Then
                Return cachedSessions
            End If

            Try
                Dim sessions As New List(Of SessionModel)()
                Dim rows = GetData(sessionsRange)

                If rows IsNot Nothing Then
                    For Each row As IList(Of Object) In rows
                        Try
                            ' შევქმნათ SessionModel
                            Dim session = SessionModel.FromSheetRow(row)

                            ' შევამოწმოთ არის თუ არა სესია მოლოდინში
                            ' მოლოდინში არის სესია, რომელიც ჯერ არ არის შესრულებული და დაგეგმილია მომავალში
                            If session.Status = "დაგეგმილი" AndAlso session.DateTime > DateTime.Now Then
                                sessions.Add(session)
                            End If
                        Catch ex As Exception
                            ' ვაგრძელებთ შემდეგი ჩანაწერით
                            Continue For
                        End Try
                    Next
                End If

                ' დავალაგოთ სესიები თარიღის მიხედვით
                sessions = sessions.OrderBy(Function(s) s.DateTime).ToList()

                ' ქეშში შენახვა
                CacheData(cacheKey, sessions)

                Return sessions
            Catch ex As Exception
                Debug.WriteLine($"მოლოდინში არსებული სესიების წამოღების შეცდომა: {ex.Message}")
                Return New List(Of SessionModel)()
            End Try
        End Function

        ''' <summary>
        ''' IDataService.GetActiveTasks იმპლემენტაცია
        ''' </summary>
        Public Function GetActiveTasks() As List(Of TaskModel) Implements IDataService.GetActiveTasks
            ' შევამოწმოთ ქეში
            Dim cacheKey = "active_tasks"
            Dim cachedTasks As List(Of TaskModel) = Nothing

            If TryGetFromCache(cacheKey, cachedTasks) Then
                Return cachedTasks
            End If

            Try
                Dim tasks As New List(Of TaskModel)()
                Dim rows = GetData(tasksRange)

                If rows IsNot Nothing Then
                    For Each row As IList(Of Object) In rows
                        Try
                            ' შევქმნათ TaskModel
                            Dim task = TaskModel.FromSheetRow(row)

                            ' შევამოწმოთ არის თუ არა დავალება აქტიური
                            ' აქტიურია დავალება, რომელიც არ არის დასრულებული ან გაუქმებული
                            If task.Status <> TaskModel.TaskStatus.Completed AndAlso
                               task.Status <> TaskModel.TaskStatus.Cancelled Then
                                tasks.Add(task)
                            End If
                        Catch ex As Exception
                            ' ვაგრძელებთ შემდეგი ჩანაწერით
                            Continue For
                        End Try
                    Next
                End If

                ' დავალაგოთ დავალებები პრიორიტეტისა და ვადის მიხედვით
                tasks = tasks.OrderByDescending(Function(t) t.Priority) _
                             .ThenBy(Function(t) t.DueDate) _
                             .ToList()

                ' ქეშში შენახვა
                CacheData(cacheKey, tasks)

                Return tasks
            Catch ex As Exception
                Debug.WriteLine($"აქტიური დავალებების წამოღების შეცდომა: {ex.Message}")
                Return New List(Of TaskModel)()
            End Try
        End Function
    End Class
End Namespace
