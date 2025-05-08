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
        Private Const sessionsRange As String = "DB-Schedule!B2:P"
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
            ' აქ შეგიძლიათ დაამატოთ დაბადების დღეების მიღების ლოგიკა
            Return New List(Of Models.BirthdayModel)()
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
    End Class
End Namespace