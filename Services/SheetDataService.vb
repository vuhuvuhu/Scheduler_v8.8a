' ===========================================
' 📄 Services/SheetDataService.vb
' -------------------------------------------
' შეცვლილი კლასი, რომელიც იყენებს GoogleServiceAccountClient-ს
' ===========================================
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Sheets.v4
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' SheetDataService - მონაცემთა წაკითხვა და ჩაწერა სერვის აკაუნტის გამოყენებით
    ''' </summary>
    Public Class SheetDataService
        Implements IDataService

        Private ReadOnly serviceClient As GoogleServiceAccountClient
        Private Const usersRange As String = "DB-Users!B2:C"
        Private Const appendRange As String = "DB-Users!B:C"
        Private Const defaultRole As String = "6"
        Private Const sessionsRange As String = "DB-Schedule!A2:O"
        Private Const birthdaysRange As String = "DB-Personal!B2:G"
        Private Const tasksRange As String = "DB-Tasks!B2:M"

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
        ''' მონაცემების მიღება მითითებული დიაპაზონიდან
        ''' </summary>
        Public Function GetData(range As String) As IList(Of IList(Of Object)) Implements IDataService.GetData
            Return serviceClient.ReadRange(range)
        End Function

        ''' <summary>
        ''' მონაცემების დამატება მითითებულ დიაპაზონში
        ''' </summary>
        Public Sub AppendData(range As String, values As IList(Of Object)) Implements IDataService.AppendData
            serviceClient.AppendValues(range, values)
        End Sub

        ''' <summary>
        ''' მონაცემების განახლება მითითებულ დიაპაზონში
        ''' </summary>
        Public Sub UpdateData(range As String, values As IList(Of Object)) Implements IDataService.UpdateData
            serviceClient.UpdateValues(range, values)
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
        ''' წამოიღებს მოახლოებულ დაბადების დღეებს
        ''' </summary>
        Public Function GetUpcomingBirthdays(Optional days As Integer = 7) As List(Of Models.BirthdayModel) Implements IDataService.GetUpcomingBirthdays
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

                                    ' შემდეგი დაბადების დღის გამოთვლა
                                    Dim nextBirthday As DateTime = New DateTime(today.Year, birthDate.Month, birthDate.Day)

                                    ' თუ დაბადების დღე უკვე გავიდა წელს, გადავიდეთ მომავალ წელზე
                                    If nextBirthday < today Then
                                        nextBirthday = nextBirthday.AddYears(1)
                                    End If

                                    ' დარჩენილი დღეების რაოდენობა
                                    Dim daysLeft = (nextBirthday - today).Days

                                    ' დავამატოთ სიაში, თუ დარჩა 7 დღე ან ნაკლები
                                    If daysLeft <= days Then
                                        Dim birthday As New BirthdayModel()
                                        birthday.Id = i + 1
                                        birthday.PersonName = firstName
                                        birthday.PersonSurname = lastName
                                        birthday.BirthDate = birthDate

                                        birthdays.Add(birthday)
                                        Debug.WriteLine($"GetUpcomingBirthdays: დაემატა დაბადების დღე - {firstName} {lastName}, " &
                                             $"თარიღი={birthDate:dd.MM.yyyy}, დარჩენილია {daysLeft} დღე")
                                    End If
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

                Debug.WriteLine($"GetUpcomingBirthdays: საბოლოოდ ნაპოვნია {birthdays.Count} დაბადების დღე")
                Return birthdays

            Catch ex As Exception
                Debug.WriteLine($"GetUpcomingBirthdays: საერთო შეცდომა: {ex.Message}")
                Return New List(Of Models.BirthdayModel)()
            End Try
        End Function

        ''' <summary>
        ''' წამოიღებს მოლოდინში არსებულ სესიებს
        ''' </summary>
        Public Function GetPendingSessions() As List(Of Models.SessionModel) Implements IDataService.GetPendingSessions
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
        ''' წამოიღებს ვადაგადაცილებულ სესიებს
        ''' </summary>
        Public Function GetOverdueSessions() As List(Of Models.SessionModel) Implements IDataService.GetOverdueSessions
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

            Return overdueSessions
        End Function
        ''' <summary>
        ''' წამოიღებს დღევანდელი სესიების სიას
        ''' </summary>
        Public Function GetTodaySessions() As List(Of Models.SessionModel) Implements IDataService.GetTodaySessions
            Dim todaySessions As New List(Of Models.SessionModel)()
            Dim rows = GetData(sessionsRange)
            Dim today As DateTime = DateTime.Today

            If rows IsNot Nothing Then
                For Each row In rows
                    Try
                        ' მინიმუმ 6 სვეტი გვჭირდება
                        If row.Count < 6 Then Continue For

                        ' სესიის ობიექტის შექმნა
                        Dim session = Models.SessionModel.FromSheetRow(row)

                        ' ვფილტრავთ მხოლოდ დღევანდელი სესიებით
                        If session.DateTime.Date = today Then
                            todaySessions.Add(session)
                        End If
                    Catch ex As Exception
                        ' გავაგრძელოთ შემდეგი ჩანაწერით
                        Continue For
                    End Try
                Next
            End If

            Debug.WriteLine($"SheetDataService.GetTodaySessions: ნაპოვნია {todaySessions.Count} დღევანდელი სესია")
            Return todaySessions
        End Function
    End Class

End Namespace