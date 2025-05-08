' ===========================================
' 📄 Services/SheetDataService.vb
' -------------------------------------------
' პასუხისმგებელი Google Sheets–დან მონაცემთა წამოღებასა და ჩანაწერების დამატებაზე
' ===========================================
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data
Imports Google.Apis.Services
Imports Google.Apis.Auth.OAuth2
Imports System.IO
Imports Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' SheetDataService ახორციელებს მონაცემთა წაკითხვასა და განახლებას
    ''' Google Sheets–ის "DB-Users" ფურცელზე (B სვეტი = ელ.ფოსტა, C სვეტი = როლი).
    ''' </summary>
    Public Class SheetDataService
        Implements IDataService

        Private ReadOnly sheetsService As SheetsService
        Private ReadOnly spreadsheetId As String
        Private Const usersRange As String = "DB-Users!B2:C"
        Private Const appendRange As String = "DB-Users!B:C"
        Private Const defaultRole As String = "6"
        Private Const sessionsRange As String = "DB-Schedule!B2:P"

        ' ქეშირებისთვის საჭირო ცვლადები
        Private ReadOnly cacheFolder As String
        Private ReadOnly cacheFile As String

        ''' <summary>
        ''' კონსტრუქტორი: აგებს SheetsService–ს OAUTHCredential–ით და spreadsheetId–თ.
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
            cacheFile = Path.Combine(cacheFolder, "sheetdata_cache.json")

            If Not Directory.Exists(cacheFolder) Then
                Directory.CreateDirectory(cacheFolder)
            End If

            ' ქეშის ჩატვირთვა თუ არსებობს
            LoadCache()
        End Sub

        ''' <summary>
        ''' კონსტრუქტორი: SheetsService-ით
        ''' </summary>
        ''' <param name="service">Google SheetsService ობიექტი</param>
        ''' <param name="spreadsheetId">Spreadsheet-ის ID</param>
        Public Sub New(service As SheetsService, spreadsheetId As String)
            Me.sheetsService = service
            Me.spreadsheetId = spreadsheetId

            ' ქეშის საქაღალდის ინიციალიზაცია
            cacheFolder = Path.Combine(Application.StartupPath, "Cache")
            cacheFile = Path.Combine(cacheFolder, "sheetdata_cache.json")

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
            ' ქეშის ჩატვირთვის ლოგიკა
            Debug.WriteLine("ქეშის ჩატვირთვის მცდელობა...")
        End Sub

        ''' <summary>
        ''' IDataService.GetData იმპლემენტაცია
        ''' </summary>
        Public Function GetData(range As String) As IList(Of IList(Of Object)) Implements IDataService.GetData
            Dim request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range)
            Return request.Execute().Values
        End Function

        ''' <summary>
        ''' IDataService.AppendData იმპლემენტაცია
        ''' </summary>
        Public Sub AppendData(range As String, values As IList(Of Object)) Implements IDataService.AppendData
            Dim valueRange As New ValueRange With {
                .Values = New List(Of IList(Of Object)) From {values}
            }
            Dim request = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, range)
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
            request.Execute()
        End Sub

        ''' <summary>
        ''' IDataService.UpdateData იმპლემენტაცია
        ''' </summary>
        Public Sub UpdateData(range As String, values As IList(Of Object)) Implements IDataService.UpdateData
            Dim valueRange As New ValueRange With {
                .Values = New List(Of IList(Of Object)) From {values}
            }
            Dim request = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, range)
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED
            request.Execute()
        End Sub

        ''' <summary>
        ''' აბრუნებს მომხმარებლის როლს, ან ხსნის ახალ ჩანაწერს თუ არ არსებობს (role=6).
        ''' </summary>
        ''' <param name="userEmail">მომხმარებლის ელ.ფოსტა</param>
        ''' <returns>C სვეტში მდებარე როლი როგორც string</returns>
        Public Function GetOrCreateUserRole(userEmail As String) As String Implements IDataService.GetOrCreateUserRole
            ' 1) წაკითხვა
            Dim getReq = sheetsService.Spreadsheets.Values.Get(spreadsheetId, usersRange)
            Dim rows As IList(Of IList(Of Object)) = getReq.Execute().Values
            If rows IsNot Nothing Then
                For Each row As IList(Of Object) In rows
                    If row.Count >= 2 AndAlso String.Equals(row(0).ToString(), userEmail, StringComparison.OrdinalIgnoreCase) Then
                        Return row(1).ToString()
                    End If
                Next
            End If
            ' 2) თუ არ მოიძებნა, დავამატოთ ახალი role=6 და დავუბრუნოთ defaultRole
            Dim newRow As IList(Of Object) = New List(Of Object) From {userEmail, defaultRole}
            Dim valueRange As New ValueRange With {.Values = New List(Of IList(Of Object)) From {newRow}}
            Dim appendReq = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, appendRange)
            appendReq.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
            appendReq.Execute()
            Return defaultRole
        End Function

        ''' <summary>
        ''' IDataService.GetUserRole იმპლემენტაცია
        ''' </summary>
        Public Function GetUserRole(email As String) As String Implements IDataService.GetUserRole
            ' მოვძებნოთ მომხმარებლის როლი
            Dim rows = GetData(usersRange)
            If rows IsNot Nothing Then
                For Each row As IList(Of Object) In rows
                    If row.Count >= 2 AndAlso String.Equals(row(0).ToString(), email, StringComparison.OrdinalIgnoreCase) Then
                        Return row(1).ToString()
                    End If
                Next
            End If
            Return String.Empty
        End Function

        ''' <summary>
        ''' ვადაგადაცილებული სესიების მიღება
        ''' </summary>
        Public Function GetOverdueSessions() As List(Of Models.SessionModel) Implements IDataService.GetOverdueSessions
            Dim sessions As New List(Of Models.SessionModel)()

            Try
                ' მონაცემების წამოღება
                Dim rows = GetData(sessionsRange)

                If rows IsNot Nothing Then
                    Debug.WriteLine($"DB-Schedule-დან წამოღებულია {rows.Count} ჩანაწერი")

                    For Each row As IList(Of Object) In rows
                        Try
                            ' ვქმნით SessionModel-ს row-დან
                            Dim session = Models.SessionModel.FromSheetRow(row)

                            ' ვამოწმებთ არის თუ არა სესია ვადაგადაცილებული
                            If session.Status = "დაგეგმილი" AndAlso session.DateTime < DateTime.Now Then
                                sessions.Add(session)
                                Debug.WriteLine($"ნაპოვნია ვადაგადაცილებული სესია: {session.Id}, {session.DateTime}, {session.BeneficiaryName}")
                            End If
                        Catch ex As Exception
                            Debug.WriteLine($"შეცდომა სესიის შექმნისას: {ex.Message}")
                            Continue For
                        End Try
                    Next
                End If

                Debug.WriteLine($"სულ ვადაგადაცილებული სესია: {sessions.Count}")

                Return sessions
            Catch ex As Exception
                Debug.WriteLine($"შეცდომა GetOverdueSessions-ში: {ex.Message}")
                Return New List(Of Models.SessionModel)()
            End Try
        End Function

        ''' <summary>
        ''' მოლოდინში სესიების მიღება
        ''' </summary>
        Public Function GetPendingSessions() As List(Of Models.SessionModel) Implements IDataService.GetPendingSessions
            Dim sessions As New List(Of Models.SessionModel)()

            Try
                ' მონაცემების წამოღება
                Dim rows = GetData(sessionsRange)

                If rows IsNot Nothing Then
                    Debug.WriteLine($"DB-Schedule-დან წამოღებულია {rows.Count} ჩანაწერი")

                    For Each row As IList(Of Object) In rows
                        Try
                            ' ვქმნით SessionModel-ს row-დან
                            Dim session = Models.SessionModel.FromSheetRow(row)

                            ' ვამოწმებთ არის თუ არა სესია მოლოდინში
                            If session.Status = "დაგეგმილი" AndAlso session.DateTime > DateTime.Now Then
                                sessions.Add(session)
                            End If
                        Catch ex As Exception
                            Debug.WriteLine($"შეცდომა სესიის შექმნისას: {ex.Message}")
                            Continue For
                        End Try
                    Next
                End If

                Return sessions
            Catch ex As Exception
                Debug.WriteLine($"შეცდომა GetPendingSessions-ში: {ex.Message}")
                Return New List(Of Models.SessionModel)()
            End Try
        End Function

        ''' <summary>
        ''' დაბადების დღეების მიღება
        ''' </summary>
        Public Function GetUpcomingBirthdays(Optional days As Integer = 7) As List(Of Models.BirthdayModel) Implements IDataService.GetUpcomingBirthdays
            ' მოკლედ შევავსოთ ცარიელი სიით
            Return New List(Of Models.BirthdayModel)()
        End Function

        ''' <summary>
        ''' აქტიური დავალებების მიღება
        ''' </summary>
        Public Function GetActiveTasks() As List(Of Models.TaskModel) Implements IDataService.GetActiveTasks
            ' მოკლედ შევავსოთ ცარიელი სიით
            Return New List(Of Models.TaskModel)()
        End Function
    End Class
End Namespace