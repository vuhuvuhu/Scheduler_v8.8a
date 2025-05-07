' ===========================================
' 📄 Services/UserService.vb
' -------------------------------------------
' მომხმარებლის სერვისი: პასუხიმგებელია მომხმარებლის მონაცემების მართვაზე
' ===========================================
Imports Google.Apis.Auth.OAuth2
Imports System.ComponentModel

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' UserService - აწარმოებს მომხმარებლის მონაცემებთან დაკავშირებულ ოპერაციებს
    ''' </summary>
    Public Class UserService

        ' მონაცემთა სერვისი (მომხმარებლის მონაცემების წყარო)
        Private ReadOnly dataService As IDataService

        ' მომხმარებლის მოდელი
        Public Class UserProfile
            ''' <summary>მომხმარებლის ელფოსტა</summary>
            Public Property Email As String

            ''' <summary>მომხმარებლის სახელი</summary>
            Public Property Name As String

            ''' <summary>მომხმარებლის როლი</summary>
            Public Property Role As String

            ''' <summary>მომხმარებლის ტელეფონი</summary>
            Public Property Phone As String

            ''' <summary>მომხმარებლის ფოტო URL</summary>
            Public Property PhotoUrl As String

            ''' <summary>მომხმარებლის დამატებითი ინფორმაცია</summary>
            Public Property AdditionalInfo As Dictionary(Of String, String)

            ''' <summary>
            ''' არის თუ არა ადმინისტრატორი
            ''' </summary>
            Public ReadOnly Property IsAdmin As Boolean
                Get
                    Return Role = "1"
                End Get
            End Property

            ''' <summary>
            ''' არის თუ არა მენეჯერი
            ''' </summary>
            Public ReadOnly Property IsManager As Boolean
                Get
                    Return Role = "2" OrElse IsAdmin
                End Get
            End Property
        End Class

        ''' <summary>
        ''' კონსტრუქტორი: იღებს IDataService-ს
        ''' </summary>
        ''' <param name="dataService">მონაცემთა სერვისი</param>
        Public Sub New(dataService As IDataService)
            Me.dataService = dataService
        End Sub

        ''' <summary>
        ''' იღებს მომხმარებლის პროფილს ელფოსტის მიხედვით
        ''' </summary>
        ''' <param name="email">მომხმარებლის ელფოსტა</param>
        ''' <returns>მომხმარებლის პროფილი ან Nothing</returns>
        Public Function GetUserByEmail(email As String) As UserProfile
            Try
                ' მომხმარებლის როლის წამოღება
                Dim role = dataService.GetUserRole(email)

                ' თუ როლი ვერ მოიძებნა, ვაბრუნებთ ნალს
                If String.IsNullOrEmpty(role) Then
                    Return Nothing
                End If

                ' პერსონალური მონაცემების წამოღება
                Dim personalRows = dataService.GetData("DB-Personal!B2:G")
                Dim name As String = String.Empty
                Dim phone As String = String.Empty

                If personalRows IsNot Nothing Then
                    For Each row As IList(Of Object) In personalRows
                        ' შევამოწმოთ საკმარისი სვეტებია თუ არა და მოვძებნოთ ელფოსტით
                        If row.Count >= 6 AndAlso
                           String.Equals(row(5).ToString(), email, StringComparison.OrdinalIgnoreCase) Then
                            name = row(0).ToString() ' სახელი პირველ სვეტშია

                            ' ტელეფონი თუ არის
                            If row.Count >= 4 Then
                                phone = row(3).ToString()
                            End If

                            Exit For
                        End If
                    Next
                End If

                ' პროფილის შექმნა და დაბრუნება
                Return New UserProfile With {
                    .Email = email,
                    .Name = name,
                    .Role = role,
                    .Phone = phone,
                    .PhotoUrl = "", ' ამჟამად არ გვაქვს ფოტო
                    .AdditionalInfo = New Dictionary(Of String, String)()
                }
            Catch ex As Exception
                Debug.WriteLine($"მომხმარებლის მოძიების შეცდომა: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' ახალი მომხმარებლის შექმნა ან არსებულის განახლება
        ''' </summary>
        ''' <param name="profile">მომხმარებლის პროფილი</param>
        ''' <returns>შენახვის ოპერაცია წარმატებულია თუ არა</returns>
        Public Function SaveUserProfile(profile As UserProfile) As Boolean
            Try
                ' მომხმარებლის როლის შენახვა/განახლება
                ' თუ არსებობს, მოვძებნოთ და განვაახლოთ
                Dim existingRole = dataService.GetUserRole(profile.Email)

                If String.IsNullOrEmpty(existingRole) Then
                    ' ახალი მომხმარებლის შექმნა
                    Dim newUserRow As New List(Of Object) From {profile.Email, profile.Role}
                    dataService.AppendData("DB-Users!B:C", newUserRow)
                Else
                    ' არსებული მომხმარებლის როლის განახლება
                    ' ჯერ უნდა მოვძებნოთ მწკრივის ნომერი
                    Dim usersRows = dataService.GetData("DB-Users!A2:C")
                    Dim rowIndex As Integer = -1

                    For i As Integer = 0 To usersRows.Count - 1
                        Dim row = usersRows(i)
                        If row.Count >= 2 AndAlso
                           String.Equals(row(1).ToString(), profile.Email, StringComparison.OrdinalIgnoreCase) Then
                            rowIndex = i + 2 ' +2 რადგან მწკრივები 1-დან იწყება და სათაურის მწკრივიც არის
                            Exit For
                        End If
                    Next

                    If rowIndex > 0 Then
                        ' მწკრივი მოიძებნა, განვაახლოთ
                        Dim updatedUserRow As New List(Of Object) From {profile.Email, profile.Role}
                        dataService.UpdateData($"DB-Users!B{rowIndex}:C{rowIndex}", updatedUserRow)
                    End If
                End If

                ' პერსონალური მონაცემების შენახვა/განახლება
                ' შევამოწმოთ არსებობს თუ არა უკვე
                Dim personalRows = dataService.GetData("DB-Personal!B2:G")
                Dim existingRowIndex As Integer = -1

                If personalRows IsNot Nothing Then
                    For i As Integer = 0 To personalRows.Count - 1
                        Dim row = personalRows(i)
                        If row.Count >= 6 AndAlso
                           String.Equals(row(5).ToString(), profile.Email, StringComparison.OrdinalIgnoreCase) Then
                            existingRowIndex = i + 2 ' +2 იგივე მიზეზით
                            Exit For
                        End If
                    Next
                End If

                If existingRowIndex > 0 Then
                    ' არსებული ჩანაწერის განახლება
                    Dim updatedPersonalRow As New List(Of Object) From {
                        profile.Name,
                        "", ' გვარი (ამჟამად არ გამოიყენება)
                        "", ' დაბადების თარიღი (ამჟამად არ გამოიყენება)
                        profile.Phone,
                        "", ' თანამდებობა (ამჟამად არ გამოიყენება)
                        profile.Email
                    }
                    dataService.UpdateData($"DB-Personal!B{existingRowIndex}:G{existingRowIndex}", updatedPersonalRow)
                Else
                    ' ახალი ჩანაწერის შექმნა
                    Dim newPersonalRow As New List(Of Object) From {
                        profile.Name,
                        "", ' გვარი (ამჟამად არ გამოიყენება)
                        "", ' დაბადების თარიღი (ამჟამად არ გამოიყენება)
                        profile.Phone,
                        "", ' თანამდებობა (ამჟამად არ გამოიყენება)
                        profile.Email
                    }
                    dataService.AppendData("DB-Personal!B:G", newPersonalRow)
                End If

                Return True
            Catch ex As Exception
                Debug.WriteLine($"მომხმარებლის პროფილის შენახვის შეცდომა: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' მომხმარებლის პროფილის მიღება OAuth და SheetDataService-ის გამოყენებით
        ''' </summary>
        ''' <param name="credential">OAuth Credential</param>
        ''' <returns>მომხმარებლის პროფილი</returns>
        Public Async Function GetUserProfile(credential As UserCredential) As Task(Of UserProfile)
            Try
                ' OAuth სერვისის შექმნა პროფილის წამოსაღებად
                Dim oauthService = New Google.Apis.Oauth2.v2.Oauth2Service(
                    New Google.Apis.Services.BaseClientService.Initializer() With {
                        .HttpClientInitializer = credential,
                        .ApplicationName = "Scheduler_v8.8a"
                    })

                ' მომხმარებლის ინფორმაციის წამოღება
                Dim userInfo = Await oauthService.Userinfo.Get().ExecuteAsync()

                ' როლის წამოღება ან შექმნა
                Dim role = dataService.GetOrCreateUserRole(userInfo.Email)

                ' პერსონალური ინფორმაციის წამოღება
                Dim existingUser = GetUserByEmail(userInfo.Email)

                ' ახალი პროფილის შექმნა
                Dim profile As New UserProfile With {
                    .Email = userInfo.Email,
                    .Name = If(existingUser?.Name, userInfo.Name),
                    .Role = role,
                    .Phone = If(existingUser?.Phone, ""),
                    .PhotoUrl = userInfo.Picture,
                    .AdditionalInfo = New Dictionary(Of String, String)()
                }

                ' დამატებითი ინფორმაცია
                If Not String.IsNullOrEmpty(userInfo.Locale) Then
                    profile.AdditionalInfo("Locale") = userInfo.Locale
                End If

                If Not String.IsNullOrEmpty(userInfo.Gender) Then
                    profile.AdditionalInfo("Gender") = userInfo.Gender
                End If

                Return profile
            Catch ex As Exception
                Debug.WriteLine($"მომხმარებლის პროფილის წამოღების შეცდომა: {ex.Message}")
                Throw New Exception("მომხმარებლის პროფილის წამოღება ვერ მოხერხდა", ex)
            End Try
        End Function
    End Class
End Namespace