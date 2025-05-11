' ===========================================
' 📄 Services/SheetDataService.vb
' -------------------------------------------
' გაუმჯობესებული ვერსია, რომელიც უზრუნველყოფს მონაცემების ქეშირებას
' და API-ს გამოძახებების რაოდენობის კონტროლს
' ===========================================
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Sheets.v4
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' SheetDataService - მონაცემთა წაკითხვა და ჩაწერა სერვის აკაუნტის გამოყენებით
    ''' API-ს მოთხოვნების რაოდენობის შეზღუდვით და მონაცემების ქეშირებით
    ''' </summary>
    Public Class SheetDataService
        Implements IDataService

        Private ReadOnly serviceClient As GoogleServiceAccountClient

        ' მუდმივები ხშირად გამოყენებული დიაპაზონებისთვის
        Private Const usersRange As String = "DB-Users!B2:C"
        Private Const appendRange As String = "DB-Users!B:C"
        Private Const defaultRole As String = "6"
        Private Const sessionsRange As String = "DB-Schedule!A2:O"
        Private Const birthdaysRange As String = "DB-Personal!B2:G"
        Private Const tasksRange As String = "DB-Tasks!B2:M"

        ' ქეშირების ცვლადები
        Private cachedTodaySessions As List(Of SessionModel) = Nothing
        Private cachedSessionsTime As DateTime = DateTime.MinValue
        Private cachedOverdueSessions As List(Of SessionModel) = Nothing
        Private cachedOverdueTime As DateTime = DateTime.MinValue
        Private cachedPendingSessions As List(Of SessionModel) = Nothing
        Private cachedPendingTime As DateTime = DateTime.MinValue
        Private cachedBirthdays As List(Of BirthdayModel) = Nothing
        Private cachedBirthdaysTime As DateTime = DateTime.MinValue

        ' მოთხოვნების შეზღუდვის პარამეტრები - ქეშის ვადა წუთებში
        Private Const CacheExpirationMinutes As Integer = 5

        ' უკანასკნელი API მოთხოვნის დრო
        Private lastApiCallTime As DateTime = DateTime.MinValue
        ' მინიმალური დრო API მოთხოვნებს შორის (მილიწამებში)
        Private Const MinApiCallIntervalMs As Integer = 1000

        ''' <summary>
        ''' კონსტრუქტორი: შექმნის სერვის კლიენტს
        ''' </summary>
        ''' <param name="serviceAccountKeyPath">სერვის აკაუნტის key ფაილის მისამართი</param>
        ''' <param name="spreadsheetId">სპრედშიტის ID</param>
        Public Sub New(serviceAccountKeyPath As String, spreadsheetId As String)
            ' შევქმნათ სერვის კლიენტი
            serviceClient = New GoogleServiceAccountClient(serviceAccountKeyPath, spreadsheetId)

            ' შევამოწმოთ წარმატებით შეიქმნა თუ არა
            If Not serviceClient.IsInitialized Then
                Throw New Exception($"ვერ მოხერხდა Google Sheets სერვისის ინიციალიზაცია: {serviceClient.InitializationError}")
            End If
        End Sub

        ''' <summary>
        ''' შეამოწმებს საჭიროა თუ არა ქეშის განახლება
        ''' </summary>
        ''' <param name="cachedTime">ქეშის შენახვის დრო</param>
        ''' <returns>True თუ ქეშის განახლება საჭიროა</returns>
        Private Function IsCacheExpired(cachedTime As DateTime) As Boolean
            Return DateTime.Now.Subtract(cachedTime).TotalMinutes >= CacheExpirationMinutes
        End Function

        ''' <summary>
        ''' API-ს მოთხოვნების შეზღუდვა გარკვეული ინტერვალით
        ''' </summary>
        Private Sub ThrottleApiCall()
            Dim timeSinceLastCall = DateTime.Now.Subtract(lastApiCallTime).TotalMilliseconds

            ' თუ ბოლო მოთხოვნიდან საკმარისი დრო არ გასულა, დაველოდოთ
            If timeSinceLastCall < MinApiCallIntervalMs Then
                Dim waitTime = CInt(MinApiCallIntervalMs - timeSinceLastCall)
                System.Threading.Thread.Sleep(waitTime)
            End If

            ' განვაახლოთ API მოთხოვნის დრო
            lastApiCallTime = DateTime.Now
        End Sub

        ''' <summary>
        ''' მონაცემების მიღება მითითებული დიაპაზონიდან შეზღუდვის გათვალისწინებით
        ''' </summary>
        Public Function GetData(range As String) As IList(Of IList(Of Object)) Implements IDataService.GetData
            ThrottleApiCall()
            Return serviceClient.ReadRange(range)
        End Function

        ''' <summary>
        ''' მონაცემების დამატება მითითებულ დიაპაზონში
        ''' </summary>
        Public Sub AppendData(range As String, values As IList(Of Object)) Implements IDataService.AppendData
            ThrottleApiCall()
            serviceClient.AppendValues(range, values)

            ' ქეშის გაუქმება შესაბამისი დიაპაზონისთვის
            InvalidateCacheForRange(range)
        End Sub

        ''' <summary>
        ''' მონაცემების განახლება მითითებულ დიაპაზონში
        ''' </summary>
        Public Sub UpdateData(range As String, values As IList(Of Object)) Implements IDataService.UpdateData
            ThrottleApiCall()
            serviceClient.UpdateValues(range, values)

            ' ქეშის გაუქმება შესაბამისი დიაპაზონისთვის
            InvalidateCacheForRange(range)
        End Sub

        ''' <summary>
        ''' ქეშის გაუქმება კონკრეტული დიაპაზონისთვის
        ''' </summary>
        Private Sub InvalidateCacheForRange(range As String)
            If range.Contains("DB-Schedule") Then
                cachedTodaySessions = Nothing
                cachedSessionsTime = DateTime.MinValue
                cachedOverdueSessions = Nothing
                cachedOverdueTime = DateTime.MinValue
                cachedPendingSessions = Nothing
                cachedPendingTime = DateTime.MinValue
            ElseIf range.Contains("DB-Personal") Then
                cachedBirthdays = Nothing
                cachedBirthdaysTime = DateTime.MinValue
            End If
        End Sub

        ''' <summary>
        ''' მთლიანი ქეშის გაუქმება - განახლების ღილაკისთვის 
        ''' </summary>
        Public Sub InvalidateAllCache()
            cachedTodaySessions = Nothing
            cachedSessionsTime = DateTime.MinValue
            cachedOverdueSessions = Nothing
            cachedOverdueTime = DateTime.MinValue
            cachedPendingSessions = Nothing
            cachedPendingTime = DateTime.MinValue
            cachedBirthdays = Nothing
            cachedBirthdaysTime = DateTime.MinValue

            Debug.WriteLine("SheetDataService: მთლიანი ქეში გაუქმებულია (ხელით განახლება)")
        End Sub

        ''' <summary>
        ''' მომხმარებლის როლის მიღება ელფოსტის მიხედვით
        ''' </summary>
        Public Function GetUserRole(email As String) As String Implements IDataService.GetUserRole
            Dim rows = GetData(usersRange)

            If rows IsNot Nothing Then
                For Each row In rows
                    If row.Count >= 2 AndAlso String.Equals(row(0).ToString(), email, StringComparison.OrdinalIgnoreCase) Then
                        Return row(1).ToString()
                    End If
                Next
            End If

            Return String.Empty
        End Function

        ''' <summary>
        ''' მომხმარებლის როლის მიღება ან შექმნა
        ''' </summary>
        Public Function GetOrCreateUserRole(email As String) As String Implements IDataService.GetOrCreateUserRole
            ' ჯერ ვცდილობთ არსებული როლის მოძიებას
            Dim role = GetUserRole(email)

            ' თუ ვერ ვიპოვეთ, ვქმნით ახალს
            If String.IsNullOrEmpty(role) Then
                Dim newRow As New List(Of Object) From {email, defaultRole}
                AppendData(appendRange, newRow)
                role = defaultRole
            End If

            Return role
        End Function

        ''' <summary>
        ''' დღევანდელი სესიების მიღება ქეშირებით
        ''' </summary>
        Public Function GetTodaySessions() As List(Of Models.SessionModel) Implements IDataService.GetTodaySessions
            ' თუ ქეში ვალიდურია, გამოვიყენოთ ის
            If cachedTodaySessions IsNot Nothing AndAlso Not IsCacheExpired(cachedSessionsTime) Then
                Debug.WriteLine("GetTodaySessions: გამოყენებულია ქეშირებული მონაცემები")
                Return cachedTodaySessions
            End If

            Try
                Dim sessions As New List(Of Models.SessionModel)()
                Dim rows = GetData(sessionsRange)

                ' მიმდინარე თარიღი - მხოლოდ თარიღის კომპონენტი (დრო 00:00)
                Dim today = DateTime.Today

                If rows IsNot Nothing Then
                    For Each row As IList(Of Object) In rows
                        Try
                            ' შევქმნათ სესიის ობიექტი
                            Dim session = SessionModel.FromSheetRow(row)

                            ' შევამოწმოთ არის თუ არა დღევანდელი
                            If session.DateTime.Date = today Then
                                sessions.Add(session)
                            End If

                        Catch ex As Exception
                            ' ცარიელი ან არასწორი მწკრივის შემთხვევაში გავაგრძელოთ
                            Continue For
                        End Try
                    Next
                End If

                ' ქეშის განახლება
                cachedTodaySessions = sessions
                cachedSessionsTime = DateTime.Now

                Debug.WriteLine($"GetTodaySessions: ნაპოვნია {sessions.Count} დღევანდელი სესია")
                Return sessions

            Catch ex As Exception
                Debug.WriteLine($"GetTodaySessions: შეცდომა - {ex.Message}")
                Return New List(Of Models.SessionModel)()
            End Try
        End Function

        ''' <summary>
        ''' წამოიღებს მოახლოებულ დაბადების დღეებს ქეშირებით
        ''' </summary>
        Public Function GetUpcomingBirthdays(Optional days As Integer = 7) As List(Of Models.BirthdayModel) Implements IDataService.GetUpcomingBirthdays
            ' თუ ქეში ვალიდურია, გამოვიყენოთ ის
            If cachedBirthdays IsNot Nothing AndAlso Not IsCacheExpired(cachedBirthdaysTime) Then
                Debug.WriteLine("GetUpcomingBirthdays: გამოყენებულია ქეშირებული მონაცემები")
                Return cachedBirthdays.Where(Function(b) b.DaysUntilBirthday <= days).ToList()
            End If

            Try
                Dim birthdays As New List(Of Models.BirthdayModel)()
                Dim rows = GetData("DB-Personal!B2:E") ' B-დან E-მდე სვეტები - სახელი, გვარი, ტელეფონი, დაბადების თარიღი

                Debug.WriteLine($"GetUpcomingBirthdays: მიღებულია {If(rows Is Nothing, 0, rows.Count)} მწკრივი DB-Personal ფურცლიდან")

                If rows IsNot Nothing Then
                    ' დღევანდელი თარიღი
                    Dim today As DateTime = DateTime.Today

                    ' მწკრივების გადარჩევა
                    For i As Integer = 0 To rows.Count - 1
                        Try
                            Dim row = rows(i)

                            ' შევამოწმოთ საკმარისი სვეტები
                            If row.Count >= 4 AndAlso row(3) IsNot Nothing AndAlso Not String.IsNullOrEmpty(row(3).ToString()) Then
                                Dim firstName = If(row(0) IsNot Nothing, row(0).ToString(), "")
                                Dim lastName = If(row(1) IsNot Nothing, row(1).ToString(), "")
                                Dim birthDateStr = row(3).ToString() ' ინდექსი 3 = E სვეტი (არა ინდექსი 2 = D სვეტი)

                                Debug.WriteLine($"GetUpcomingBirthdays: მწკრივი {i + 2}, სახელი={firstName}, გვარი={lastName}, დაბადების თარიღი='{birthDateStr}'")

                                ' დავამატოთ მცირე ვალიდაცია - თუ დაბადების თარიღი არის მხოლოდ ციფრები და 8+ სიმბოლო,
                                ' სავარაუდოდ ეს არ არის ვალიდური თარიღი ფორმატი
                                If birthDateStr.Length > 8 AndAlso birthDateStr.All(Function(c) Char.IsDigit(c)) Then
                                    Debug.WriteLine($"GetUpcomingBirthdays: სავარაუდოდ არასწორი ფორმატის თარიღი: '{birthDateStr}' - გადავახტეთ")
                                    Continue For
                                End If

                                ' სხვადასხვა ფორმატების მცდელობა
                                Dim birthDate As DateTime
                                Dim formats() As String = {"dd.MM.yyyy", "dd,MM,yyyy", "dd/MM/yyyy", "dd-MM-yyyy",
                                                "yyyy.MM.dd", "yyyy,MM,dd", "yyyy/MM/dd", "yyyy-MM-dd",
                                                "MM/dd/yyyy", "MM-dd-yyyy", "MM.dd.yyyy", "dd.MM", "dd,MM"}

                                If DateTime.TryParseExact(birthDateStr, formats, Globalization.CultureInfo.InvariantCulture,
                                               Globalization.DateTimeStyles.None, birthDate) OrElse
                           DateTime.TryParse(birthDateStr, birthDate) Then

                                    Debug.WriteLine($"GetUpcomingBirthdays: დაბადების თარიღი დაპარსულია: {birthDate:dd.MM.yyyy}")

                                    ' თუ მხოლოდ დღე და თვე იყო მითითებული (წელი არის 1), ვიყენებთ მიმდინარე წელს
                                    If birthDate.Year = 1 Then
                                        birthDate = New DateTime(today.Year, birthDate.Month, birthDate.Day)
                                    End If

                                    ' შევქმნათ დაბადების დღის მოდელი
                                    Dim birthday As New BirthdayModel()
                                    birthday.Id = i + 1
                                    birthday.PersonName = firstName
                                    birthday.PersonSurname = lastName
                                    birthday.BirthDate = birthDate

                                    ' დავამატოთ ყველა დაბადების თარიღი, ქეშისთვის
                                    birthdays.Add(birthday)

                                    Debug.WriteLine($"GetUpcomingBirthdays: დაემატა დაბადების დღე - ID={birthday.Id}, " &
                                             $"სახელი={birthday.PersonName}, გვარი={birthday.PersonSurname}, " &
                                             $"თარიღი={birthday.BirthDate:dd.MM.yyyy}, დარჩა={birthday.DaysUntilBirthday} დღე")

                                Else
                                    Debug.WriteLine($"GetUpcomingBirthdays: ვერ მოხერხდა დაბადების თარიღის '{birthDateStr}' პარსინგი")
                                End If
                            End If
                        Catch ex As Exception
                            Debug.WriteLine($"GetUpcomingBirthdays: შეცდომა მწკრივის დამუშავებისას: {ex.Message}")
                        End Try
                    Next
                End If

                ' დავალაგოთ დაბადების დღეები დღეების რაოდენობის მიხედვით
                birthdays = birthdays.OrderBy(Function(b) b.DaysUntilBirthday).ToList()

                ' შევინახოთ ქეშში
                cachedBirthdays = birthdays
                cachedBirthdaysTime = DateTime.Now

                ' დავაბრუნოთ მხოლოდ ის დაბადების დღეები, რომლებიც მოთხოვნილ დღეებში ჯდება
                Dim filteredBirthdays = birthdays.Where(Function(b) b.DaysUntilBirthday <= days).ToList()

                Debug.WriteLine($"GetUpcomingBirthdays: საბოლოოდ ნაპოვნია {filteredBirthdays.Count} დაბადების დღე")
                Return filteredBirthdays

            Catch ex As Exception
                Debug.WriteLine($"GetUpcomingBirthdays: საერთო შეცდომა: {ex.Message}")
                Return New List(Of Models.BirthdayModel)()
            End Try
        End Function

        ''' <summary>
        ''' წამოიღებს მოლოდინში არსებულ სესიებს ქეშირებით
        ''' </summary>
        Public Function GetPendingSessions() As List(Of Models.SessionModel) Implements IDataService.GetPendingSessions
            ' თუ ქეში ვალიდურია, გამოვიყენოთ ის
            If cachedPendingSessions IsNot Nothing AndAlso Not IsCacheExpired(cachedPendingTime) Then
                Debug.WriteLine("GetPendingSessions: გამოყენებულია ქეშირებული მონაცემები")
                Return cachedPendingSessions
            End If

            Dim sessions As New List(Of Models.SessionModel)()
            Dim rows = GetData(sessionsRange)

            If rows IsNot Nothing Then
                For Each row In rows
                    Try
                        If row.Count < 12 Then Continue For

                        Dim session = SessionModel.FromSheetRow(row)

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

            ' შევინახოთ ქეშში
            cachedPendingSessions = sessions
            cachedPendingTime = DateTime.Now

            Return sessions
        End Function

        ''' <summary>
        ''' წამოიღებს აქტიურ დავალებებს
        ''' </summary>
        Public Function GetActiveTasks() As List(Of Models.TaskModel) Implements IDataService.GetActiveTasks
            ' აქ შეგიძლიათ დაამატოთ აქტიური დავალებების მიღების ლოგიკა
            Return New List(Of Models.TaskModel)()
        End Function

        ''' <summary>
        ''' წამოიღებს ვადაგადაცილებულ სესიებს ქეშირებით
        ''' </summary>
        Public Function GetOverdueSessions() As List(Of Models.SessionModel) Implements IDataService.GetOverdueSessions
            ' თუ ქეში ვალიდურია, გამოვიყენოთ ის
            If cachedOverdueSessions IsNot Nothing AndAlso Not IsCacheExpired(cachedOverdueTime) Then
                Debug.WriteLine("GetOverdueSessions: გამოყენებულია ქეშირებული მონაცემები")
                Return cachedOverdueSessions
            End If

            Dim overdueSessions As New List(Of Models.SessionModel)()
            Dim rows = GetData(sessionsRange)

            If rows IsNot Nothing Then
                For Each row In rows
                    Try
                        ' მინიმუმ 12 სვეტი გვჭირდება
                        If row.Count < 12 Then Continue For

                        ' სესიის ობიექტის შექმნა
                        Dim session = SessionModel.FromSheetRow(row)

                        ' ვადაგადაცილებულობის შემოწმება
                        If session.IsOverdue Then
                            overdueSessions.Add(session)
                        End If
                    Catch ex As Exception
                        ' ვაგრძელებთ შემდეგი მწკრივით
                        Continue For
                    End Try
                Next
            End If

            ' დავალაგოთ თარიღის მიხედვით
            overdueSessions = overdueSessions.OrderBy(Function(s) s.DateTime).ToList()

            ' შევინახოთ ქეშში
            cachedOverdueSessions = overdueSessions
            cachedOverdueTime = DateTime.Now

            Return overdueSessions
        End Function
        ''' <summary>
        ''' წამოიღებს ყველა სესიას
        ''' </summary>
        Public Function GetAllSessions() As List(Of Models.SessionModel) Implements IDataService.GetAllSessions
            Dim allSessions As New List(Of Models.SessionModel)()
            Dim rows = GetData(sessionsRange)

            If rows IsNot Nothing Then
                For Each row In rows
                    Try
                        ' მინიმუმ 12 სვეტი გვჭირდება
                        If row.Count < 12 Then Continue For

                        ' სესიის ობიექტის შექმნა
                        Dim session = SessionModel.FromSheetRow(row)
                        allSessions.Add(session)
                    Catch ex As Exception
                        ' ვაგრძელებთ შემდეგი მწკრივით
                        Continue For
                    End Try
                Next
            End If

            Return allSessions
        End Function
    End Class
End Namespace